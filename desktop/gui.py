"""
gui.py — CustomTkinter 3-tab UI for the PowerUp Portal desktop app.

Layout:
  Window (900 x 680)
  ├── Top banner: "PowerUp Portal" title
  ├── Tabview: [M1] [M2] [M3]
  └── Bottom status bar: wraps the last progress message

Each tab:
  • Big "Upload Excel" button → opens file picker
  • Filename + client dropdown (auto-populated after parse, shown only if multi-client)
  • Questionnaire-name dropdown (M2 only; optional, for matching form responses)
  • Big "Generate" button, greyed until an upload + client is chosen
  • Result area: green success banner + clickable Drive link

Every heavy operation runs on a background thread to keep the window responsive.
"""
from __future__ import annotations

import os
import sys
import threading
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

# Bootstrap portal/ + credentials before any worker imports
import app_config                                # noqa: F401
from workers import common
from workers import m1_worker, m2_worker, m3_worker


# ── Appearance ────────────────────────────────────────────────
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

WINDOW_W, WINDOW_H = 900, 680

FONT_TITLE = ("Segoe UI", 24, "bold")
FONT_HEAD  = ("Segoe UI", 18, "bold")
FONT_BODY  = ("Segoe UI", 13)
FONT_BTN   = ("Segoe UI", 16, "bold")
FONT_SMALL = ("Segoe UI", 11)


# ═══════════════════════════════════════════════════════════════
# Base tab — shared UI for M1/M2/M3
# ═══════════════════════════════════════════════════════════════
class _BaseTab(ctk.CTkFrame):
    """Common upload + client-picker + generate skeleton."""
    TITLE      = "Base"
    SUBTITLE   = ""
    GENERATE_LABEL = "Generate"

    def __init__(self, parent, status_cb):
        super().__init__(parent, fg_color="transparent")
        self.status_cb = status_cb
        self.xlsx_path: str | None = None
        self.clients: list[tuple[str, str]] = []
        self.selected_pf_id: str | None = None
        self._busy = False
        self._build_ui()

    # ── UI scaffolding (subclasses may extend) ────────────────
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)

        # Header
        ctk.CTkLabel(self, text=self.TITLE, font=FONT_HEAD, anchor="w"
                     ).grid(row=0, column=0, padx=30, pady=(24, 2), sticky="ew")
        if self.SUBTITLE:
            ctk.CTkLabel(self, text=self.SUBTITLE, font=FONT_SMALL,
                         text_color="#555", anchor="w"
                         ).grid(row=1, column=0, padx=30, pady=(0, 18), sticky="ew")

        # Upload card
        upload_card = ctk.CTkFrame(self, corner_radius=12)
        upload_card.grid(row=2, column=0, padx=30, pady=(10, 10), sticky="ew")
        upload_card.grid_columnconfigure(1, weight=1)

        self.upload_btn = ctk.CTkButton(
            upload_card, text="📂  Upload Excel", font=FONT_BTN,
            width=220, height=60, command=self._pick_file,
        )
        self.upload_btn.grid(row=0, column=0, padx=20, pady=20, sticky="w")

        self.file_label = ctk.CTkLabel(
            upload_card, text="No file selected", font=FONT_BODY,
            text_color="#666", anchor="w",
        )
        self.file_label.grid(row=0, column=1, padx=(10, 20), pady=20, sticky="ew")

        # Client picker (shown only when parse yields clients)
        self.picker_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.picker_frame.grid(row=3, column=0, padx=30, pady=(0, 10), sticky="ew")
        self.picker_frame.grid_columnconfigure(1, weight=1)

        self.client_label = ctk.CTkLabel(
            self.picker_frame, text="Client:", font=FONT_BODY,
        )
        self.client_dropdown = ctk.CTkOptionMenu(
            self.picker_frame, values=["(upload a file first)"],
            font=FONT_BODY, height=40, dynamic_resizing=False,
            command=self._on_client_changed,
        )
        self.client_dropdown.set("(upload a file first)")
        # Hidden until there are clients
        self._hide_picker()

        # Subclass hook for extra rows (e.g. questionnaire dropdown for M2,
        # client-name textbox for M3)
        self.extra_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.extra_frame.grid(row=4, column=0, padx=30, pady=(0, 10), sticky="ew")
        self.extra_frame.grid_columnconfigure(1, weight=1)
        self._build_extra(self.extra_frame)

        # Generate button
        self.generate_btn = ctk.CTkButton(
            self, text=self.GENERATE_LABEL, font=FONT_BTN,
            height=70, command=self._on_generate_click,
            state="disabled",
        )
        self.generate_btn.grid(row=5, column=0, padx=30, pady=(20, 10), sticky="ew")

        # Result area (shown after success)
        self.result_frame = ctk.CTkFrame(self, fg_color="#e8f5e9", corner_radius=12)
        self.result_frame.grid_columnconfigure(0, weight=1)
        # (not .grid()'d yet — appears on success)

    # ── Hooks for subclasses ──────────────────────────────────
    def _build_extra(self, parent):
        """Subclasses override to add their own rows inside `parent`."""
        pass

    def _run_generation(self) -> dict:
        """Subclasses implement to call their worker. Return {'url','name'}."""
        raise NotImplementedError

    def _can_generate(self) -> bool:
        """Subclasses may extend to add extra preconditions."""
        return self.xlsx_path is not None and self.selected_pf_id is not None

    # ── File pick + parse ─────────────────────────────────────
    def _pick_file(self):
        path = filedialog.askopenfilename(
            title="Select client Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not path:
            return
        self.xlsx_path = path
        self.file_label.configure(text=Path(path).name, text_color="#1a2e3b")
        self.status_cb(f"Parsing {Path(path).name}...")

        # Run parse on a thread — multi-MB files can take a second or two
        def _work():
            try:
                clients = self._parse_clients(path)
            except Exception as e:
                self.after(0, lambda: self._parse_failed(str(e)))
                return
            self.after(0, lambda: self._parse_done(clients))

        threading.Thread(target=_work, daemon=True).start()

    def _parse_clients(self, path: str) -> list[tuple[str, str]]:
        """Subclasses may override (e.g. M3 doesn't use PF_ID)."""
        return common.list_clients_in_excel(path)

    def _parse_failed(self, err: str):
        self.status_cb(f"Error: {err}")
        messagebox.showerror("Could not read file", err)
        self.file_label.configure(text="No file selected", text_color="#666")
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
            # Don't show the dropdown when there's only one — cleaner UI
            self._hide_picker()
        else:
            names = [f"{n}  —  {pid[:12]}..." if len(pid) > 14 else f"{n}  —  {pid}"
                     for pid, n in clients]
            self.client_dropdown.configure(values=names)
            self.client_dropdown.set(names[0])
            self.selected_pf_id = clients[0][0]
            self.status_cb(f"Found {len(clients)} clients — pick one to continue.")
            self._show_picker()
        self._update_generate_state()

    def _on_client_changed(self, display_name: str):
        # Map the selected display back to a pf_id
        for i, (pid, name) in enumerate(self.clients):
            shown = f"{name}  —  {pid[:12]}..." if len(pid) > 14 else f"{name}  —  {pid}"
            if shown == display_name:
                self.selected_pf_id = pid
                self.status_cb(f"Selected: {name}")
                break
        self._update_generate_state()

    def _show_picker(self):
        self.client_label.grid(row=0, column=0, padx=(0, 10), sticky="w")
        self.client_dropdown.grid(row=0, column=1, sticky="ew")

    def _hide_picker(self):
        self.client_label.grid_remove()
        self.client_dropdown.grid_remove()

    # ── Generate ──────────────────────────────────────────────
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
        self.generate_btn.configure(state="disabled", text="Working...")
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
        self._update_generate_state()
        self.status_cb("Done. ✅")
        self._show_result(result.get('name', 'Done'), result.get('url', ''))

    def _on_failure(self, err: str, tb: str = ""):
        self._busy = False
        self.generate_btn.configure(text=self.GENERATE_LABEL)
        self._update_generate_state()
        self.status_cb(f"Failed: {err}")
        # Show full traceback in the error dialog so the user can screenshot
        # it and share with us — str(e) alone is rarely enough to debug.
        body = err
        if tb:
            body += "\n\n─── traceback ───\n" + tb
        messagebox.showerror("Generation failed", body)

    def _show_result(self, title: str, url: str):
        for w in self.result_frame.winfo_children():
            w.destroy()
        self.result_frame.grid(row=6, column=0, padx=30, pady=(10, 20), sticky="ew")
        ctk.CTkLabel(
            self.result_frame, text="✅  Deck generated and uploaded",
            font=FONT_BTN, text_color="#2a9c4a", anchor="w",
        ).grid(row=0, column=0, padx=20, pady=(16, 2), sticky="ew")
        ctk.CTkLabel(
            self.result_frame, text=title, font=FONT_BODY, anchor="w",
            text_color="#1a2e3b",
        ).grid(row=1, column=0, padx=20, pady=(0, 10), sticky="ew")
        if url:
            ctk.CTkButton(
                self.result_frame, text="🌐  Open in Google Drive",
                font=FONT_BTN, height=44, fg_color="#1a2e3b", hover_color="#2a4155",
                command=lambda: webbrowser.open(url),
            ).grid(row=2, column=0, padx=20, pady=(0, 18), sticky="w")

    def _hide_result(self):
        self.result_frame.grid_remove()


# ═══════════════════════════════════════════════════════════════
# M1 tab
# ═══════════════════════════════════════════════════════════════
class M1Tab(_BaseTab):
    TITLE    = "M1 — Client Report Sheet"
    SUBTITLE = ("Upload a client Excel and generate their M1 report as a Google "
                "Sheet. Syncs the data to the Main Data spreadsheet first.")
    GENERATE_LABEL = "⚙️  Generate M1 Report"

    def _run_generation(self) -> dict:
        assert self.xlsx_path and self.selected_pf_id
        result = m1_worker.generate(self.xlsx_path, self.selected_pf_id)
        return {'url': result['url'], 'name': result['title']}


# ═══════════════════════════════════════════════════════════════
# M2 tab
# ═══════════════════════════════════════════════════════════════
class M2Tab(_BaseTab):
    TITLE    = "M2 — Strategy Deck"
    SUBTITLE = ("Upload a client Excel and generate their personalised strategy "
                "deck. Reads client data locally; fetches questionnaire response "
                "from Google Sheets.")
    GENERATE_LABEL = "⚙️  Generate M2 Deck"

    _Q_AUTO = "(auto-match by name)"
    _Q_LOADING = "(loading questionnaire...)"
    _Q_FAILED = "(fetch failed — see status)"

    def _build_extra(self, parent):
        parent.grid_columnconfigure(1, weight=1)

        # Row 0: label + dropdown
        ctk.CTkLabel(parent, text="Questionnaire response:",
                     font=FONT_BODY).grid(row=0, column=0, padx=(0, 10),
                                          sticky="w")
        self.q_dropdown = ctk.CTkOptionMenu(
            parent, values=[self._Q_LOADING],
            font=FONT_BODY, height=40, dynamic_resizing=False, width=400,
            command=self._on_q_selected,
        )
        self.q_dropdown.set(self._Q_LOADING)
        self.q_dropdown.grid(row=0, column=1, sticky="ew")

        # Row 1: live match-status indicator (shows auto-matched name OR error)
        self.q_match_label = ctk.CTkLabel(
            parent, text="", font=FONT_SMALL, text_color="#666", anchor="w",
            wraplength=500, justify="left",
        )
        self.q_match_label.grid(row=1, column=1, sticky="ew", pady=(6, 0))

        self._q_names: list[str] = []
        self._fetch_questionnaire_names()

    def _fetch_questionnaire_names(self) -> None:
        """Background fetch of all Name values in the questionnaire sheet."""
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
                text_color="#cc0000",
            )
            self.status_cb(err or "No questionnaire responses found.")
        else:
            self.q_dropdown.configure(values=[self._Q_AUTO] + names)
            self.q_dropdown.set(self._Q_AUTO)
            self.q_match_label.configure(
                text=f"{len(names)} responses loaded from Google Sheets.",
                text_color="#666",
            )
            self.status_cb(f"Loaded {len(names)} questionnaire responses.")
        self._update_match_preview()

    def _on_q_selected(self, _value: str) -> None:
        self._update_match_preview()
        self._update_generate_state()

    # Override so the match preview refreshes whenever the client changes.
    def _on_client_changed(self, display_name: str):
        super()._on_client_changed(display_name)
        self._update_match_preview()

    def _parse_done(self, clients):
        super()._parse_done(clients)
        self._update_match_preview()

    def _update_match_preview(self) -> None:
        """Show next to the dropdown which questionnaire row will actually
        be used for the currently-selected client."""
        if not hasattr(self, 'q_match_label'):
            return
        if not self._q_names:
            return  # loaded state; error already shown

        selected = self.q_dropdown.get()
        client_name = next(
            (n for pid, n in self.clients if pid == self.selected_pf_id),
            None,
        )

        if selected and selected != self._Q_AUTO and selected not in (
                self._Q_LOADING, self._Q_FAILED):
            # Manual choice
            self.q_match_label.configure(
                text=f"✓ Using: {selected}", text_color="#2a9c4a",
            )
            return

        if not client_name:
            self.q_match_label.configure(
                text="Upload an Excel and pick a client to see the auto-match.",
                text_color="#666",
            )
            return

        # Auto-match by fuzzy similarity
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
                text=f"✓ Auto-matched: {best}   ({pct}% similarity)",
                text_color="#2a9c4a",
            )
        elif best and score >= 0.3:
            pct = int(round(score * 100))
            self.q_match_label.configure(
                text=(f"⚠ Weak auto-match: {best}  ({pct}% similarity). "
                      f"Pick the correct response from the dropdown if this is wrong."),
                text_color="#cc8800",
            )
        else:
            self.q_match_label.configure(
                text=(f"✗ No match for '{client_name}'. "
                      f"Pick the correct response from the dropdown above."),
                text_color="#cc0000",
            )

    def _run_generation(self) -> dict:
        assert self.xlsx_path and self.selected_pf_id
        # Customer name = the NAME from the Excel's PF_level row
        customer_name = next(
            (n for pid, n in self.clients if pid == self.selected_pf_id),
            self.selected_pf_id,
        )
        # If the user explicitly picked a questionnaire response, pass that
        # name through so generate_deck matches the exact row. Otherwise
        # leave it None and generate_deck falls back to fuzzy matching on
        # the customer_name.
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
    GENERATE_LABEL = "⚙️  Generate M3 Deck"

    def _build_extra(self, parent):
        ctk.CTkLabel(parent, text="Client name:",
                     font=FONT_BODY).grid(row=0, column=0, padx=(0, 10), sticky="w")
        self.name_entry = ctk.CTkEntry(
            parent, placeholder_text="As it should appear on the deck cover",
            font=FONT_BODY, height=40,
        )
        self.name_entry.grid(row=0, column=1, sticky="ew")
        self.name_entry.bind("<KeyRelease>", lambda _: self._update_generate_state())

    def _parse_clients(self, path: str):
        # M3 workbooks don't use PF_ID — no client picker needed.
        self.selected_pf_id = "__m3_no_pf_id__"
        return []

    def _parse_done(self, clients):
        # Override — we always proceed (no picker), just enable generate.
        self.status_cb("Ready. Enter the client name and click Generate.")
        self._hide_picker()
        self.selected_pf_id = "__m3_no_pf_id__"
        self._update_generate_state()

    def _can_generate(self) -> bool:
        return (self.xlsx_path is not None
                and bool(self.name_entry.get().strip()))

    def _run_generation(self) -> dict:
        assert self.xlsx_path
        name = self.name_entry.get().strip()
        return m3_worker.generate(self.xlsx_path, name)


# ═══════════════════════════════════════════════════════════════
# Main window
# ═══════════════════════════════════════════════════════════════
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("PowerUp Portal (Local)")
        self.geometry(f"{WINDOW_W}x{WINDOW_H}")
        self.minsize(800, 600)

        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Top banner
        banner = ctk.CTkFrame(self, fg_color="#1a2e3b", corner_radius=0, height=72)
        banner.grid(row=0, column=0, sticky="ew")
        banner.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            banner, text="PowerUp Portal",
            font=FONT_TITLE, text_color="white", anchor="w",
        ).grid(row=0, column=0, padx=30, pady=18, sticky="ew")

        # Tabview
        self.tabview = ctk.CTkTabview(self, corner_radius=8)
        self.tabview.grid(row=1, column=0, padx=20, pady=(20, 0), sticky="nsew")
        self.tabview.add("M1 Report")
        self.tabview.add("M2 Deck")
        self.tabview.add("M3 Deck")

        # Status bar
        self.status_var = ctk.StringVar(value="Ready.")
        status = ctk.CTkLabel(
            self, textvariable=self.status_var, font=FONT_SMALL,
            text_color="#555", anchor="w", fg_color="#f3f4f6", height=28,
        )
        status.grid(row=2, column=0, sticky="ew", padx=0, pady=0)

        # Instantiate tabs
        for name, cls in (("M1 Report", M1Tab),
                          ("M2 Deck",   M2Tab),
                          ("M3 Deck",   M3Tab)):
            tab = self.tabview.tab(name)
            tab.grid_columnconfigure(0, weight=1)
            tab.grid_rowconfigure(0, weight=1)
            frame = cls(tab, status_cb=self._set_status)
            frame.grid(row=0, column=0, sticky="nsew")

        # Wire the PROGRESS sink so worker messages appear in the status bar
        common.PROGRESS.set(self._set_status)

    def _set_status(self, msg: str):
        # Called from any thread via .after() inside tabs, but also directly
        # from workers — guard with a main-thread hop.
        try:
            self.after(0, lambda: self.status_var.set(msg))
        except Exception:
            pass


def run():
    app = App()
    app.mainloop()
