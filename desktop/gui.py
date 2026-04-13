"""
gui.py — CustomTkinter 3-tab UI for the PowerUp Portal desktop app.

Design goals (Level-1 polish pass — Apr 2026):
  * Cleaner palette (slate + indigo, no default-blue chrome)
  * Card-based layout — each section sits in a white card with a subtle border
  * Typography hierarchy (title / heading / body / caption) via ui_theme.Font
  * Progress bar (indeterminate) visible during generation
  * Subtler status bar at the bottom
  * Fewer, more intentional emojis
  * No animations Tk can't do — no fake drop-shadows or gradients

Workers (M1/M2/M3) are unchanged. Only the GUI layer is touched.
"""
from __future__ import annotations

import threading
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Optional

import customtkinter as ctk

# Bootstrap portal + credentials BEFORE any worker imports
import app_config                                       # noqa: F401
from workers import common
from workers import m1_worker, m2_worker, m3_worker, agreement_worker

from ui_theme import Color, Font, Space, Radius, BUTTON_H, BUTTON_H_LG, INPUT_H, BANNER_H, STATUS_H


# ── Appearance ────────────────────────────────────────────────
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")   # overridden per-widget via fg_color

# Default window size: 960x680 fits comfortably on a 1366x768 laptop
# (taskbar leaves ~720px usable). Content scrolls so even shorter screens work.
WINDOW_W, WINDOW_H = 960, 680


# ═══════════════════════════════════════════════════════════════
# Reusable styled widgets
# ═══════════════════════════════════════════════════════════════
def Card(parent, **kwargs) -> ctk.CTkFrame:
    """White rounded card with a 1px slate-200 border."""
    return ctk.CTkFrame(
        parent,
        fg_color=Color.BG_CARD,
        border_color=Color.BORDER,
        border_width=1,
        corner_radius=Radius.CARD,
        **kwargs,
    )


def PrimaryButton(parent, **kwargs) -> ctk.CTkButton:
    """Indigo-filled primary action button. All style defaults go through
    setdefault so callers (e.g. HeroButton) can override any of them without
    causing 'multiple values for keyword argument' TypeErrors."""
    kwargs.setdefault('height', BUTTON_H)
    kwargs.setdefault('font', Font.BUTTON)
    kwargs.setdefault('fg_color', Color.PRIMARY)
    kwargs.setdefault('hover_color', Color.PRIMARY_HOVER)
    kwargs.setdefault('text_color', Color.TEXT_ON_DARK)
    kwargs.setdefault('text_color_disabled', "#A5B4FC")
    kwargs.setdefault('corner_radius', Radius.BUTTON)
    return ctk.CTkButton(parent, **kwargs)


def HeroButton(parent, **kwargs) -> ctk.CTkButton:
    """Extra-tall version of the primary button for the main CTA."""
    kwargs.setdefault('height', BUTTON_H_LG)
    kwargs.setdefault('font', Font.BUTTON_LG)
    return PrimaryButton(parent, **kwargs)


def SecondaryButton(parent, **kwargs) -> ctk.CTkButton:
    """White outlined ghost-style button."""
    kwargs.setdefault('height', BUTTON_H)
    kwargs.setdefault('font', Font.BUTTON)
    kwargs.setdefault('fg_color', Color.SECONDARY)
    kwargs.setdefault('hover_color', Color.SECONDARY_HOVER)
    kwargs.setdefault('text_color', Color.TEXT_PRIMARY)
    kwargs.setdefault('border_color', Color.SECONDARY_BORDER)
    kwargs.setdefault('border_width', 1)
    kwargs.setdefault('corner_radius', Radius.BUTTON)
    return ctk.CTkButton(parent, **kwargs)


def Label(parent, text: str, **kwargs) -> ctk.CTkLabel:
    kwargs.setdefault('font', Font.BODY)
    kwargs.setdefault('text_color', Color.TEXT_PRIMARY)
    kwargs.setdefault('anchor', 'w')
    return ctk.CTkLabel(parent, text=text, **kwargs)


def MutedLabel(parent, text: str, **kwargs) -> ctk.CTkLabel:
    kwargs.setdefault('font', Font.LABEL)
    kwargs.setdefault('text_color', Color.TEXT_SECONDARY)
    kwargs.setdefault('anchor', 'w')
    return ctk.CTkLabel(parent, text=text, **kwargs)


def TextInput(parent, **kwargs) -> ctk.CTkEntry:
    kwargs.setdefault('height', INPUT_H)
    kwargs.setdefault('font', Font.BODY)
    return ctk.CTkEntry(
        parent,
        fg_color=Color.BG_INPUT,
        border_color=Color.BORDER,
        border_width=1,
        text_color=Color.TEXT_PRIMARY,
        placeholder_text_color=Color.TEXT_MUTED,
        corner_radius=Radius.SM,
        **kwargs,
    )


def Dropdown(parent, **kwargs) -> ctk.CTkOptionMenu:
    kwargs.setdefault('height', INPUT_H)
    kwargs.setdefault('font', Font.BODY)
    return ctk.CTkOptionMenu(
        parent,
        fg_color=Color.BG_INPUT,
        button_color=Color.BG_SUBTLE,
        button_hover_color=Color.BORDER,
        text_color=Color.TEXT_PRIMARY,
        dropdown_fg_color=Color.BG_CARD,
        dropdown_text_color=Color.TEXT_PRIMARY,
        dropdown_hover_color=Color.BG_SUBTLE,
        dynamic_resizing=False,
        corner_radius=Radius.SM,
        **kwargs,
    )


# ═══════════════════════════════════════════════════════════════
# Base tab — shared layout for M1 / M2 / M3
# ═══════════════════════════════════════════════════════════════
class _BaseTab(ctk.CTkFrame):
    TITLE      = "Base"
    SUBTITLE   = ""
    GENERATE_LABEL = "Generate"

    def __init__(self, parent, status_cb):
        # Outer frame fills the tab; an inner scrollable frame holds all
        # the content so the user can still see the "Open in Drive" button
        # on a 1366x768 laptop even when every card is expanded.
        super().__init__(parent, fg_color=Color.BG_APP)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._scroll = ctk.CTkScrollableFrame(
            self, fg_color=Color.BG_APP, corner_radius=0,
            scrollbar_button_color=Color.BORDER,
            scrollbar_button_hover_color=Color.SECONDARY_BORDER,
        )
        self._scroll.grid(row=0, column=0, sticky="nsew")
        self._scroll.grid_columnconfigure(0, weight=1)

        self.status_cb = status_cb
        self.xlsx_path: str | None = None
        self.clients: list[tuple[str, str]] = []
        self.selected_pf_id: str | None = None
        self._busy = False
        self._build_ui()

    # ── Scaffolding ──────────────────────────────────────────
    def _build_ui(self):
        # All grid() calls below put their widgets inside self._scroll
        # rather than self, so the whole tab content scrolls.
        self._scroll.grid_columnconfigure(0, weight=1)

        s = self._scroll  # shorthand — every widget below lives inside it

        # Header (no card — sits directly on the app bg)
        header = ctk.CTkFrame(s, fg_color="transparent")
        header.grid(row=0, column=0, padx=Space.XXL, pady=(Space.XL, Space.MD),
                    sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            header, text=self.TITLE, font=Font.HEADING,
            text_color=Color.TEXT_PRIMARY, anchor="w",
        ).grid(row=0, column=0, sticky="ew")
        if self.SUBTITLE:
            ctk.CTkLabel(
                header, text=self.SUBTITLE, font=Font.SUBHEAD,
                text_color=Color.TEXT_SECONDARY, anchor="w",
                wraplength=820, justify="left",
            ).grid(row=1, column=0, sticky="ew", pady=(Space.XS, 0))

        # Upload card
        upload_card = Card(s)
        upload_card.grid(row=1, column=0, padx=Space.XXL, pady=(Space.MD, Space.MD),
                         sticky="ew")
        upload_card.grid_columnconfigure(1, weight=1)

        PrimaryButton(
            upload_card, text="Upload Excel", command=self._pick_file,
            width=180,
        ).grid(row=0, column=0, padx=Space.LG, pady=Space.LG, sticky="w")

        self.file_label = ctk.CTkLabel(
            upload_card, text="No file selected", font=Font.BODY,
            text_color=Color.TEXT_MUTED, anchor="w",
        )
        self.file_label.grid(row=0, column=1, padx=(Space.SM, Space.LG),
                             pady=Space.LG, sticky="ew")

        # Client picker (shown only when >1 clients)
        self.picker_frame = ctk.CTkFrame(s, fg_color="transparent")
        self.picker_frame.grid(row=2, column=0, padx=Space.XXL, pady=(0, Space.SM),
                               sticky="ew")
        self.picker_frame.grid_columnconfigure(1, weight=1)

        self.client_label = MutedLabel(self.picker_frame, text="Client")
        self.client_dropdown = Dropdown(
            self.picker_frame, values=["(upload a file first)"],
            command=self._on_client_changed,
        )
        self.client_dropdown.set("(upload a file first)")
        self._hide_picker()

        # Subclass extras (questionnaire picker for M2, name entry for M3)
        self.extra_frame = ctk.CTkFrame(s, fg_color="transparent")
        self.extra_frame.grid(row=3, column=0, padx=Space.XXL, pady=(0, Space.LG),
                              sticky="ew")
        self.extra_frame.grid_columnconfigure(1, weight=1)
        self._build_extra(self.extra_frame)

        # Generate button (hero size)
        self.generate_btn = HeroButton(
            s, text=self.GENERATE_LABEL, command=self._on_generate_click,
            state="disabled",
        )
        self.generate_btn.grid(row=4, column=0, padx=Space.XXL,
                               pady=(Space.MD, Space.MD), sticky="ew")

        # Progress bar (hidden until working)
        self.progress = ctk.CTkProgressBar(
            s, mode='indeterminate', height=4,
            progress_color=Color.PRIMARY, fg_color=Color.BORDER,
            corner_radius=Radius.PILL,
        )
        # Not grid'd until we start working

        # Result card (populated on success, hidden otherwise)
        self.result_frame = Card(s)
        self.result_frame.grid_columnconfigure(0, weight=1)
        # Not grid'd yet

    def _build_extra(self, parent): pass

    # ── File pick ────────────────────────────────────────────
    def _pick_file(self):
        path = filedialog.askopenfilename(
            title="Select client Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not path:
            return
        self.xlsx_path = path
        self.file_label.configure(
            text=Path(path).name, text_color=Color.TEXT_PRIMARY,
        )
        self.status_cb(f"Parsing {Path(path).name}...")

        def _work():
            try:
                clients = self._parse_clients(path)
            except Exception as e:
                self.after(0, lambda: self._parse_failed(str(e)))
                return
            self.after(0, lambda: self._parse_done(clients))
        threading.Thread(target=_work, daemon=True).start()

    def _parse_clients(self, path):
        return common.list_clients_in_excel(path)

    def _parse_failed(self, err: str):
        self.status_cb(f"Error: {err}")
        messagebox.showerror("Could not read file", err)
        self.file_label.configure(text="No file selected",
                                  text_color=Color.TEXT_MUTED)
        self.xlsx_path = None
        self._update_generate_state()

    def _parse_done(self, clients):
        self.clients = clients
        if not clients:
            self.status_cb("No PF_IDs found in the uploaded file.")
            messagebox.showwarning(
                "No clients found",
                "The uploaded file does not appear to contain a PF_ID column "
                "in any known tab (PF_level, Scheme_level).",
            )
            self._hide_picker()
            self.selected_pf_id = None
        elif len(clients) == 1:
            self.selected_pf_id = clients[0][0]
            self.status_cb(f"Found 1 client: {clients[0][1]}")
            self._hide_picker()
        else:
            names = [self._format_client(pid, n) for pid, n in clients]
            self.client_dropdown.configure(values=names)
            self.client_dropdown.set(names[0])
            self.selected_pf_id = clients[0][0]
            self.status_cb(f"Found {len(clients)} clients — pick one to continue.")
            self._show_picker()
        self._update_generate_state()

    @staticmethod
    def _format_client(pid: str, name: str) -> str:
        return f"{name}  ·  {pid[:12]}…" if len(pid) > 14 else f"{name}  ·  {pid}"

    def _on_client_changed(self, display_name: str):
        for pid, name in self.clients:
            if self._format_client(pid, name) == display_name:
                self.selected_pf_id = pid
                self.status_cb(f"Selected: {name}")
                break
        self._update_generate_state()

    def _show_picker(self):
        self.client_label.grid(row=0, column=0, padx=(0, Space.MD), sticky="w")
        self.client_dropdown.grid(row=0, column=1, sticky="ew")

    def _hide_picker(self):
        self.client_label.grid_remove()
        self.client_dropdown.grid_remove()

    # ── Generation ───────────────────────────────────────────
    def _can_generate(self) -> bool:
        return self.xlsx_path is not None and self.selected_pf_id is not None

    def _update_generate_state(self):
        if self._busy:
            return
        self.generate_btn.configure(
            state="normal" if self._can_generate() else "disabled"
        )

    def _on_generate_click(self):
        if self._busy:
            return
        self._busy = True
        self.generate_btn.configure(state="disabled", text="Working…")
        self._show_progress()
        self._hide_result()

        def _work():
            import traceback as _tb
            try:
                result = self._run_generation()
            except Exception as e:
                tb_str = _tb.format_exc()
                self.after(0, lambda: self._on_failure(str(e), tb_str))
                return
            self.after(0, lambda: self._on_success(result))
        threading.Thread(target=_work, daemon=True).start()

    def _on_success(self, result: dict):
        self._busy = False
        self.generate_btn.configure(text=self.GENERATE_LABEL)
        self._hide_progress()
        self._update_generate_state()
        self.status_cb("Done.")
        self._show_result(result.get('name', 'Done'), result.get('url', ''))

    def _on_failure(self, err: str, tb: str = ""):
        self._busy = False
        self.generate_btn.configure(text=self.GENERATE_LABEL)
        self._hide_progress()
        self._update_generate_state()
        self.status_cb(f"Failed: {err}")
        body = err
        if tb:
            body += "\n\n─── traceback ───\n" + tb
        messagebox.showerror("Generation failed", body)

    def _show_progress(self):
        # Lives inside self._scroll, so positions are relative to that.
        self.progress.grid(row=5, column=0, padx=Space.XXL, pady=(0, Space.MD),
                           sticky="ew")
        self.progress.start()

    def _hide_progress(self):
        self.progress.stop()
        self.progress.grid_remove()

    def _show_result(self, title: str, url: str):
        for w in self.result_frame.winfo_children():
            w.destroy()
        self.result_frame.grid(row=6, column=0, padx=Space.XXL,
                               pady=(Space.SM, Space.XL), sticky="ew")
        # Auto-scroll to the result card so users on small laptop screens
        # don't have to manually scroll down to see "Open in Drive".
        try:
            self._scroll._parent_canvas.yview_moveto(1.0)
        except Exception:
            pass

        # Green "success" pill
        pill = ctk.CTkLabel(
            self.result_frame, text="✓  GENERATED",
            font=Font.TAG, text_color=Color.SUCCESS,
            fg_color=Color.SUCCESS_BG, corner_radius=Radius.PILL,
            padx=Space.MD, pady=Space.XS,
        )
        pill.grid(row=0, column=0, padx=Space.LG, pady=(Space.LG, Space.SM),
                  sticky="w")

        ctk.CTkLabel(
            self.result_frame, text=title, font=Font.BODY_BOLD,
            text_color=Color.TEXT_PRIMARY, anchor="w",
            wraplength=820, justify="left",
        ).grid(row=1, column=0, padx=Space.LG, pady=(0, Space.MD), sticky="w")

        if url:
            PrimaryButton(
                self.result_frame, text="Open in Google Drive",
                command=lambda: webbrowser.open(url), width=200,
            ).grid(row=2, column=0, padx=Space.LG,
                   pady=(0, Space.LG), sticky="w")

    def _hide_result(self):
        self.result_frame.grid_remove()

    def _run_generation(self) -> dict:
        raise NotImplementedError


# ═══════════════════════════════════════════════════════════════
# M1 tab
# ═══════════════════════════════════════════════════════════════
class M1Tab(_BaseTab):
    TITLE    = "M1 — Client Report Sheet"
    SUBTITLE = ("Upload a client Excel and generate their M1 report as a Google "
                "Sheet. Syncs the uploaded data to the Main Data spreadsheet first, "
                "then triggers the Apps Script web app.")
    GENERATE_LABEL = "Generate M1 Report"

    def _run_generation(self) -> dict:
        assert self.xlsx_path and self.selected_pf_id
        result = m1_worker.generate(self.xlsx_path, self.selected_pf_id)
        return {'url': result['url'], 'name': result['title']}


# ═══════════════════════════════════════════════════════════════
# M2 tab
# ═══════════════════════════════════════════════════════════════
class M2Tab(_BaseTab):
    TITLE    = "M2 — Strategy Deck"
    SUBTITLE = ("Upload a client Excel. Client data is read locally from the "
                "file; only the questionnaire response is fetched from Google "
                "Sheets.")
    GENERATE_LABEL = "Generate M2 Deck"

    _Q_AUTO    = "(auto-match by name)"
    _Q_LOADING = "(loading…)"
    _Q_FAILED  = "(fetch failed — see status bar)"

    def _build_extra(self, parent):
        parent.grid_columnconfigure(1, weight=1)

        MutedLabel(parent, text="Questionnaire").grid(
            row=0, column=0, padx=(0, Space.MD), sticky="w",
        )
        self.q_dropdown = Dropdown(
            parent, values=[self._Q_LOADING], command=self._on_q_selected,
            width=420,
        )
        self.q_dropdown.set(self._Q_LOADING)
        self.q_dropdown.grid(row=0, column=1, sticky="ew")

        self.q_match_label = ctk.CTkLabel(
            parent, text="", font=Font.SMALL,
            text_color=Color.TEXT_SECONDARY, anchor="w",
            wraplength=600, justify="left",
        )
        self.q_match_label.grid(row=1, column=1, sticky="ew",
                                pady=(Space.XS, 0))

        self._q_names: list[str] = []
        self._fetch_questionnaire_names()

    def _fetch_questionnaire_names(self):
        def _work():
            try:
                qdf = common.fetch_questionnaire()
                if qdf.empty:
                    self.after(0, lambda: self._on_q_loaded(
                        [], "Questionnaire sheet is empty."))
                    return
                name_col = next((c for c in qdf.columns
                                 if c.lower() == "name"), None)
                if not name_col:
                    self.after(0, lambda: self._on_q_loaded(
                        [], "Questionnaire sheet has no 'Name' column."))
                    return
                names = sorted(set(
                    str(n).strip() for n in qdf[name_col]
                    if str(n).strip() and str(n).strip().lower() != 'nan'
                ))
                self.after(0, lambda: self._on_q_loaded(names, None))
            except Exception as e:
                err = f"Could not fetch questionnaire: {type(e).__name__}: {e}"
                self.after(0, lambda: self._on_q_loaded([], err))
        threading.Thread(target=_work, daemon=True).start()

    def _on_q_loaded(self, names: list[str], err: str | None):
        self._q_names = names
        if err or not names:
            self.q_dropdown.configure(values=[self._Q_FAILED])
            self.q_dropdown.set(self._Q_FAILED)
            self.q_match_label.configure(
                text=err or "No questionnaire responses found.",
                text_color=Color.ERROR,
            )
            self.status_cb(err or "No questionnaire responses found.")
        else:
            self.q_dropdown.configure(values=[self._Q_AUTO] + names)
            self.q_dropdown.set(self._Q_AUTO)
            self.q_match_label.configure(
                text=f"{len(names)} responses loaded from Google Sheets.",
                text_color=Color.TEXT_SECONDARY,
            )
            self.status_cb(f"Loaded {len(names)} questionnaire responses.")
        self._update_match_preview()

    def _on_q_selected(self, _v):
        self._update_match_preview()
        self._update_generate_state()

    def _on_client_changed(self, display_name: str):
        super()._on_client_changed(display_name)
        self._update_match_preview()

    def _parse_done(self, clients):
        super()._parse_done(clients)
        self._update_match_preview()

    def _update_match_preview(self):
        if not hasattr(self, 'q_match_label') or not self._q_names:
            return

        selected = self.q_dropdown.get()
        client_name = next(
            (n for pid, n in self.clients if pid == self.selected_pf_id),
            None,
        )

        if selected and selected != self._Q_AUTO and selected not in (
                self._Q_LOADING, self._Q_FAILED):
            self.q_match_label.configure(
                text=f"✓  Using: {selected}", text_color=Color.SUCCESS,
            )
            return

        if not client_name:
            self.q_match_label.configure(
                text="Upload an Excel and pick a client to see the auto-match.",
                text_color=Color.TEXT_SECONDARY,
            )
            return

        from difflib import SequenceMatcher
        cl = client_name.lower().strip()
        best, score = None, 0.0
        for n in self._q_names:
            s = SequenceMatcher(None, cl, n.lower().strip()).ratio()
            if s > score:
                score, best = s, n

        if best and score >= 0.5:
            pct = int(round(score * 100))
            self.q_match_label.configure(
                text=f"✓  Auto-matched: {best}   ({pct}% similarity)",
                text_color=Color.SUCCESS,
            )
        elif best and score >= 0.3:
            pct = int(round(score * 100))
            self.q_match_label.configure(
                text=(f"!  Weak auto-match: {best}  ({pct}% similarity). "
                      f"If this is wrong, pick the correct response above."),
                text_color=Color.WARNING,
            )
        else:
            self.q_match_label.configure(
                text=(f"✗  No match for '{client_name}'. "
                      f"Pick the correct response from the dropdown above."),
                text_color=Color.ERROR,
            )

    def _run_generation(self) -> dict:
        assert self.xlsx_path and self.selected_pf_id
        customer_name = next(
            (n for pid, n in self.clients if pid == self.selected_pf_id),
            self.selected_pf_id,
        )
        selected = self.q_dropdown.get()
        if selected in (self._Q_AUTO, self._Q_LOADING, self._Q_FAILED):
            q_name = None
        else:
            q_name = selected
        return m2_worker.generate(
            self.xlsx_path, self.selected_pf_id, customer_name,
            questionnaire_name=q_name,
        )


# ═══════════════════════════════════════════════════════════════
# M3 tab
# ═══════════════════════════════════════════════════════════════
class M3Tab(_BaseTab):
    TITLE    = "M3 — Portfolio Transition Deck"
    SUBTITLE = ("Upload a client's curation/master-plan Excel. Monthly reference "
                "data (AUM, Powerranking, etc.) is fetched from Google Sheets.")
    GENERATE_LABEL = "Generate M3 Deck"

    def _build_extra(self, parent):
        parent.grid_columnconfigure(1, weight=1)
        MutedLabel(parent, text="Client name").grid(
            row=0, column=0, padx=(0, Space.MD), sticky="w",
        )
        self.name_entry = TextInput(
            parent, placeholder_text="As it should appear on the deck cover",
        )
        self.name_entry.grid(row=0, column=1, sticky="ew")
        self.name_entry.bind("<KeyRelease>", lambda _: self._update_generate_state())

    def _parse_clients(self, path: str):
        self.selected_pf_id = "__m3_no_pf_id__"
        return []

    def _parse_done(self, clients):
        self.status_cb("Ready. Enter the client name and click Generate.")
        self._hide_picker()
        self.selected_pf_id = "__m3_no_pf_id__"
        self._update_generate_state()

    def _can_generate(self) -> bool:
        return (self.xlsx_path is not None
                and bool(self.name_entry.get().strip()))

    def _run_generation(self) -> dict:
        assert self.xlsx_path
        return m3_worker.generate(self.xlsx_path, self.name_entry.get().strip())


# ═══════════════════════════════════════════════════════════════
# Agreement tab — pure form, no Excel upload
# ═══════════════════════════════════════════════════════════════
class AgreementTab(ctk.CTkFrame):
    """Generates an Elite or Non-Elite agreement DOCX, uploads it as a
    Google Doc to the agreement output folder, and shows the user a link.
    Different shape from M1/M2/M3 tabs — no file upload, no client picker.
    Pure form input + dynamic fee-slab rows."""

    TITLE = "Agreement"
    SUBTITLE = ("Generate a client agreement (Elite or Non-Elite) with default "
                "or custom fee slabs.")
    GENERATE_LABEL = "Generate Agreement"

    def __init__(self, parent, status_cb):
        super().__init__(parent, fg_color=Color.BG_APP)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._scroll = ctk.CTkScrollableFrame(
            self, fg_color=Color.BG_APP, corner_radius=0,
            scrollbar_button_color=Color.BORDER,
            scrollbar_button_hover_color=Color.SECONDARY_BORDER,
        )
        self._scroll.grid(row=0, column=0, sticky="nsew")
        self._scroll.grid_columnconfigure(0, weight=1)

        self.status_cb = status_cb
        self._busy = False
        self._slab_rows: list[tuple[ctk.CTkEntry, ctk.CTkEntry]] = []
        self._build_ui()

    def _build_ui(self):
        s = self._scroll

        # Header
        header = ctk.CTkFrame(s, fg_color="transparent")
        header.grid(row=0, column=0, padx=Space.XXL, pady=(Space.XL, Space.MD),
                    sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            header, text=self.TITLE, font=Font.HEADING,
            text_color=Color.TEXT_PRIMARY, anchor="w",
        ).grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(
            header, text=self.SUBTITLE, font=Font.SUBHEAD,
            text_color=Color.TEXT_SECONDARY, anchor="w",
            wraplength=820, justify="left",
        ).grid(row=1, column=0, sticky="ew", pady=(Space.XS, 0))

        # Type card — radio between Elite / Non-Elite
        type_card = Card(s)
        type_card.grid(row=1, column=0, padx=Space.XXL,
                       pady=(Space.MD, Space.MD), sticky="ew")
        type_card.grid_columnconfigure(0, weight=1)

        Label(type_card, text="Agreement type",
              font=Font.BODY_BOLD).grid(
            row=0, column=0, padx=Space.LG,
            pady=(Space.LG, Space.SM), sticky="w",
        )
        self.type_var = ctk.StringVar(value="Elite")
        type_row = ctk.CTkFrame(type_card, fg_color="transparent")
        type_row.grid(row=1, column=0, padx=Space.LG,
                      pady=(0, Space.LG), sticky="w")
        for v, label in (("Elite",     "Elite (PowerUp Infinite)"),
                         ("Non-Elite", "Non-Elite (UW IA Client Agreement)")):
            ctk.CTkRadioButton(
                type_row, text=label, variable=self.type_var, value=v,
                font=Font.BODY, text_color=Color.TEXT_PRIMARY,
                fg_color=Color.PRIMARY, hover_color=Color.PRIMARY_HOVER,
                border_color=Color.SECONDARY_BORDER,
                command=self._on_type_changed,
            ).grid(row=0, column=("Elite", "Non-Elite").index(v),
                   padx=(0, Space.XL), pady=Space.XS, sticky="w")

        # Client details card
        client_card = Card(s)
        client_card.grid(row=2, column=0, padx=Space.XXL,
                         pady=(0, Space.MD), sticky="ew")
        client_card.grid_columnconfigure(1, weight=1)

        Label(client_card, text="Client details",
              font=Font.BODY_BOLD).grid(
            row=0, column=0, columnspan=2, padx=Space.LG,
            pady=(Space.LG, Space.SM), sticky="w",
        )

        MutedLabel(client_card, text="Client name").grid(
            row=1, column=0, padx=(Space.LG, Space.MD),
            pady=Space.SM, sticky="w",
        )
        self.name_entry = TextInput(
            client_card,
            placeholder_text="Full name as it should appear on the agreement",
        )
        self.name_entry.grid(row=1, column=1, padx=(0, Space.LG),
                             pady=Space.SM, sticky="ew")
        self.name_entry.bind("<KeyRelease>", lambda _: self._refresh_state())

        # Email / phone — only when Non-Elite. Slots reserved up-front so
        # later layout doesn't shuffle.
        MutedLabel(client_card, text="Email").grid(
            row=2, column=0, padx=(Space.LG, Space.MD), pady=Space.SM, sticky="w",
        )
        self.email_entry = TextInput(
            client_card, placeholder_text="Client email (or NA to skip)",
        )
        self.email_entry.grid(row=2, column=1, padx=(0, Space.LG),
                              pady=Space.SM, sticky="ew")

        MutedLabel(client_card, text="Phone").grid(
            row=3, column=0, padx=(Space.LG, Space.MD), pady=Space.SM, sticky="w",
        )
        self.phone_entry = TextInput(
            client_card, placeholder_text="Phone number (or NA to skip)",
        )
        self.phone_entry.grid(row=3, column=1, padx=(0, Space.LG),
                              pady=(Space.SM, Space.LG), sticky="ew")

        # Date — only when Elite
        self.date_label = MutedLabel(client_card, text="Agreement date")
        self.date_entry = TextInput(
            client_card, placeholder_text="YYYY-MM-DD (default: today)",
        )

        # Fee slabs card
        slab_card = Card(s)
        slab_card.grid(row=3, column=0, padx=Space.XXL,
                       pady=(0, Space.MD), sticky="ew")
        slab_card.grid_columnconfigure(0, weight=1)

        Label(slab_card, text="Fee slabs",
              font=Font.BODY_BOLD).grid(
            row=0, column=0, padx=Space.LG,
            pady=(Space.LG, Space.SM), sticky="w",
        )

        self.slab_mode = ctk.StringVar(value="Default")
        slab_mode_row = ctk.CTkFrame(slab_card, fg_color="transparent")
        slab_mode_row.grid(row=1, column=0, padx=Space.LG,
                           pady=(0, Space.SM), sticky="w")
        for v, label in (("Default", "Default (3 standard slabs)"),
                         ("Custom",  "Custom")):
            ctk.CTkRadioButton(
                slab_mode_row, text=label, variable=self.slab_mode, value=v,
                font=Font.BODY, text_color=Color.TEXT_PRIMARY,
                fg_color=Color.PRIMARY, hover_color=Color.PRIMARY_HOVER,
                border_color=Color.SECONDARY_BORDER,
                command=self._on_slab_mode_changed,
            ).grid(row=0, column=("Default", "Custom").index(v),
                   padx=(0, Space.XL), pady=Space.XS, sticky="w")

        # Custom-slabs sub-area (hidden until Custom is picked)
        self.custom_frame = ctk.CTkFrame(slab_card, fg_color="transparent")
        self.custom_frame.grid_columnconfigure(0, weight=1)
        # not grid'd until Custom mode

        # Inside custom_frame: count selector, dynamic rows, bottom padding
        self.slab_count_row = ctk.CTkFrame(self.custom_frame, fg_color="transparent")
        self.slab_count_row.grid(row=0, column=0, padx=Space.LG,
                                  pady=(Space.SM, Space.SM), sticky="w")
        MutedLabel(self.slab_count_row, text="Number of slabs").grid(
            row=0, column=0, padx=(0, Space.MD), sticky="w",
        )
        self.slab_count_var = ctk.StringVar(value="1")
        # Max 3 slabs — agreements never have more than 3 fee tiers.
        self.slab_count_dd = Dropdown(
            self.slab_count_row,
            values=["1", "2", "3"],
            command=self._on_slab_count_changed,
            width=80,
        )
        self.slab_count_dd.set("1")
        self.slab_count_dd.grid(row=0, column=1, sticky="w")

        self.slab_rows_frame = ctk.CTkFrame(self.custom_frame, fg_color="transparent")
        self.slab_rows_frame.grid(row=1, column=0, padx=Space.LG,
                                   pady=(0, Space.LG), sticky="ew")
        self.slab_rows_frame.grid_columnconfigure(0, weight=1)
        self.slab_rows_frame.grid_columnconfigure(1, weight=1)

        # Generate button
        self.generate_btn = HeroButton(
            s, text=self.GENERATE_LABEL, command=self._on_generate_click,
            state="disabled",
        )
        self.generate_btn.grid(row=4, column=0, padx=Space.XXL,
                               pady=(Space.MD, Space.MD), sticky="ew")

        # Progress + result (same pattern as _BaseTab)
        self.progress = ctk.CTkProgressBar(
            s, mode='indeterminate', height=4,
            progress_color=Color.PRIMARY, fg_color=Color.BORDER,
            corner_radius=Radius.PILL,
        )
        self.result_frame = Card(s)
        self.result_frame.grid_columnconfigure(0, weight=1)

        # Initial state: Elite is the default → show date row, hide email/phone
        self._on_type_changed()

    # ── Type / mode changes ───────────────────────────────────
    def _on_type_changed(self):
        is_elite = self.type_var.get() == "Elite"
        # Email + phone rows live at rows 2 and 3 — show or hide them
        if is_elite:
            self.email_entry.master.grid_slaves(row=2, column=0)
            for r in (2, 3):
                for w in self.email_entry.master.grid_slaves(row=r):
                    w.grid_remove()
            self.date_label.grid(row=2, column=0, padx=(Space.LG, Space.MD),
                                 pady=(Space.SM, Space.LG), sticky="w")
            self.date_entry.grid(row=2, column=1, padx=(0, Space.LG),
                                 pady=(Space.SM, Space.LG), sticky="ew")
        else:
            self.date_label.grid_remove()
            self.date_entry.grid_remove()
            # Re-show email/phone (row 2, 3)
            for r in (2, 3):
                for w in self.email_entry.master.grid_slaves(row=r):
                    if w in (self.date_label, self.date_entry):
                        continue
                    w.grid()
        self._refresh_state()

    def _on_slab_mode_changed(self):
        if self.slab_mode.get() == "Custom":
            self.custom_frame.grid(row=2, column=0, sticky="ew")
            self._on_slab_count_changed(self.slab_count_dd.get())
        else:
            self.custom_frame.grid_remove()
        self._refresh_state()

    def _on_slab_count_changed(self, value: str):
        try:
            n = int(value)
        except ValueError:
            n = 1
        # Rebuild the rows
        for w in self.slab_rows_frame.winfo_children():
            w.destroy()
        self._slab_rows = []

        # Header row
        MutedLabel(self.slab_rows_frame, text="Fee % (e.g. 0.75% p.a.)").grid(
            row=0, column=0, padx=(0, Space.MD), pady=(Space.SM, Space.XS),
            sticky="w",
        )
        MutedLabel(self.slab_rows_frame, text="AUA range").grid(
            row=0, column=1, padx=(Space.MD, 0), pady=(Space.SM, Space.XS),
            sticky="w",
        )

        for i in range(n):
            fee = TextInput(self.slab_rows_frame,
                            placeholder_text="e.g. 0.75% p.a.")
            fee.grid(row=i + 1, column=0, padx=(0, Space.MD),
                     pady=Space.XS, sticky="ew")
            aua = TextInput(self.slab_rows_frame,
                            placeholder_text="e.g. less than 50 lakhs")
            aua.grid(row=i + 1, column=1, padx=(Space.MD, 0),
                     pady=Space.XS, sticky="ew")
            self._slab_rows.append((fee, aua))
        self._refresh_state()

    # ── Generate flow ─────────────────────────────────────────
    def _refresh_state(self):
        if self._busy:
            return
        ok = bool(self.name_entry.get().strip())
        self.generate_btn.configure(state="normal" if ok else "disabled")

    def _collect_custom_slabs(self) -> Optional[list[dict]]:
        if self.slab_mode.get() != "Custom":
            return None
        slabs = []
        for fee, aua in self._slab_rows:
            f, a = fee.get().strip(), aua.get().strip()
            if not f:
                continue
            if "p.a" not in f.lower():
                f += " p.a."
            slabs.append({"fee": f, "aua": a})
        return slabs or None

    def _on_generate_click(self):
        if self._busy:
            return
        self._busy = True
        self.generate_btn.configure(state="disabled", text="Working…")
        self._show_progress()
        self._hide_result()

        is_elite      = self.type_var.get() == "Elite"
        client_name   = self.name_entry.get().strip()
        email         = self.email_entry.get().strip() if not is_elite else ""
        phone         = self.phone_entry.get().strip() if not is_elite else ""
        date_text     = self.date_entry.get().strip() if is_elite else ""
        custom_slabs  = self._collect_custom_slabs()

        # Parse the date if provided
        from datetime import date as _date, datetime as _dt
        agreement_date = None
        if is_elite:
            if date_text:
                try:
                    agreement_date = _dt.strptime(date_text, "%Y-%m-%d").date()
                except ValueError:
                    pass  # fall through — engine defaults to today
            if agreement_date is None:
                agreement_date = _date.today()

        def _work():
            import traceback as _tb
            try:
                result = agreement_worker.generate(
                    is_elite=is_elite,
                    client_name=client_name,
                    email=email,
                    phone=phone,
                    agreement_date=agreement_date,
                    custom_slabs=custom_slabs,
                )
            except Exception as e:
                tb_str = _tb.format_exc()
                self.after(0, lambda: self._on_failure(str(e), tb_str))
                return
            self.after(0, lambda: self._on_success(result))
        threading.Thread(target=_work, daemon=True).start()

    def _on_success(self, result: dict):
        self._busy = False
        self.generate_btn.configure(text=self.GENERATE_LABEL)
        self._hide_progress()
        self._refresh_state()
        self.status_cb("Done.")
        self._show_result(result.get('name', 'Done'), result.get('url', ''))

    def _on_failure(self, err: str, tb: str = ""):
        self._busy = False
        self.generate_btn.configure(text=self.GENERATE_LABEL)
        self._hide_progress()
        self._refresh_state()
        self.status_cb(f"Failed: {err}")
        body = err
        if tb:
            body += "\n\n─── traceback ───\n" + tb
        messagebox.showerror("Generation failed", body)

    def _show_progress(self):
        self.progress.grid(row=5, column=0, padx=Space.XXL,
                           pady=(0, Space.MD), sticky="ew")
        self.progress.start()

    def _hide_progress(self):
        self.progress.stop()
        self.progress.grid_remove()

    def _show_result(self, title: str, url: str):
        for w in self.result_frame.winfo_children():
            w.destroy()
        self.result_frame.grid(row=6, column=0, padx=Space.XXL,
                               pady=(Space.SM, Space.XL), sticky="ew")
        try:
            self._scroll._parent_canvas.yview_moveto(1.0)
        except Exception:
            pass

        ctk.CTkLabel(
            self.result_frame, text="✓  GENERATED",
            font=Font.TAG, text_color=Color.SUCCESS,
            fg_color=Color.SUCCESS_BG, corner_radius=Radius.PILL,
            padx=Space.MD, pady=Space.XS,
        ).grid(row=0, column=0, padx=Space.LG,
               pady=(Space.LG, Space.SM), sticky="w")

        ctk.CTkLabel(
            self.result_frame, text=title, font=Font.BODY_BOLD,
            text_color=Color.TEXT_PRIMARY, anchor="w",
            wraplength=820, justify="left",
        ).grid(row=1, column=0, padx=Space.LG,
               pady=(0, Space.MD), sticky="w")

        if url:
            PrimaryButton(
                self.result_frame, text="Open in Google Drive",
                command=lambda: webbrowser.open(url), width=200,
            ).grid(row=2, column=0, padx=Space.LG,
                   pady=(0, Space.LG), sticky="w")

    def _hide_result(self):
        self.result_frame.grid_remove()


# ═══════════════════════════════════════════════════════════════
# Main window
# ═══════════════════════════════════════════════════════════════
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("PowerUp Portal (Local)")
        self.geometry(f"{WINDOW_W}x{WINDOW_H}")
        # Min size kept generous-enough that the upload + generate button are
        # always visible without scrolling. Below that, scroll handles it.
        self.minsize(820, 540)
        self.configure(fg_color=Color.BG_APP)

        # Row layout:
        #   0  banner (fixed height)
        #   1  accent stripe (2px)
        #   2  tabview (expandable)
        #   3  status bar (fixed height)
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Top banner — near-black with a subtle violet accent stripe under it.
        banner = ctk.CTkFrame(
            self, fg_color=Color.BG_BANNER, corner_radius=0, height=BANNER_H,
        )
        banner.grid(row=0, column=0, sticky="ew")
        banner.grid_columnconfigure(0, weight=1)
        banner.grid_propagate(False)

        # Logo dot + wordmark — more "branded" than plain text.
        logo_row = ctk.CTkFrame(banner, fg_color="transparent")
        logo_row.grid(row=0, column=0, padx=Space.XXL, pady=Space.MD, sticky="w")
        ctk.CTkLabel(
            logo_row, text="●", font=(Font.TITLE[0], 22, "bold"),
            text_color=Color.PRIMARY, anchor="w",
        ).grid(row=0, column=0, padx=(0, Space.SM))
        ctk.CTkLabel(
            logo_row, text="PowerUp Portal", font=Font.TITLE,
            text_color=Color.TEXT_ON_DARK, anchor="w",
        ).grid(row=0, column=1, sticky="w")

        # Right-side meta tag
        meta = ctk.CTkLabel(
            banner, text="LOCAL  ·  v1", font=Font.TAG,
            text_color="#7A7A8A", anchor="e",
            fg_color="#1A1A22", corner_radius=Radius.PILL,
            padx=Space.MD, pady=Space.XS,
        )
        meta.grid(row=0, column=1, padx=Space.XXL, pady=Space.MD, sticky="e")

        # 1px violet accent stripe below the banner
        accent = ctk.CTkFrame(self, fg_color=Color.PRIMARY,
                              corner_radius=0, height=2)
        accent.grid(row=1, column=0, sticky="ew")
        accent.grid_propagate(False)

        # Tabview — customtkinter's default is fine once we theme the segment
        self.tabview = ctk.CTkTabview(
            self, corner_radius=Radius.LG, fg_color=Color.BG_APP,
            segmented_button_fg_color=Color.BG_SUBTLE,
            segmented_button_selected_color=Color.PRIMARY,
            segmented_button_selected_hover_color=Color.PRIMARY_HOVER,
            segmented_button_unselected_color=Color.BG_SUBTLE,
            segmented_button_unselected_hover_color=Color.BORDER,
            text_color=Color.TEXT_PRIMARY,
        )
        self.tabview.grid(row=2, column=0, padx=0, pady=0, sticky="nsew")
        self.tabview.add("M1 Report")
        self.tabview.add("M2 Deck")
        self.tabview.add("M3 Deck")
        self.tabview.add("Agreement")

        # Status bar — separated from content by a thin top border.
        status_border = ctk.CTkFrame(self, fg_color=Color.BORDER,
                                     corner_radius=0, height=1)
        status_border.grid(row=3, column=0, sticky="ew")

        status_frame = ctk.CTkFrame(
            self, fg_color=Color.BG_SUBTLE, corner_radius=0, height=STATUS_H,
        )
        status_frame.grid(row=4, column=0, sticky="ew")
        status_frame.grid_columnconfigure(1, weight=1)
        status_frame.grid_propagate(False)

        # Subtle bullet that gets coloured per-state could go here later;
        # for now, a static muted dot reads as a "status indicator".
        ctk.CTkLabel(
            status_frame, text="●", font=(Font.SMALL[0], 9),
            text_color=Color.TEXT_MUTED,
        ).grid(row=0, column=0, padx=(Space.LG, Space.SM), pady=Space.XS, sticky="w")

        self.status_var = ctk.StringVar(value="Ready.")
        ctk.CTkLabel(
            status_frame, textvariable=self.status_var, font=Font.SMALL,
            text_color=Color.TEXT_SECONDARY, anchor="w",
        ).grid(row=0, column=1, sticky="ew", padx=(0, Space.LG), pady=Space.XS)

        # Mount the four tabs
        for name, cls in (("M1 Report",  M1Tab),
                          ("M2 Deck",    M2Tab),
                          ("M3 Deck",    M3Tab),
                          ("Agreement",  AgreementTab)):
            tab = self.tabview.tab(name)
            tab.grid_columnconfigure(0, weight=1)
            tab.grid_rowconfigure(0, weight=1)
            frame = cls(tab, status_cb=self._set_status)
            frame.grid(row=0, column=0, sticky="nsew")

        # Route worker PROGRESS messages into our status bar
        common.PROGRESS.set(self._set_status)

    def _set_status(self, msg: str):
        try:
            self.after(0, lambda: self.status_var.set(msg))
        except Exception:
            pass


def run():
    app = App()
    app.mainloop()
