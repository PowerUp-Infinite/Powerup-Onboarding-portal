"""
ui_theme.py — single source of truth for colors, fonts, spacing.

Palette inspired by modern clean interfaces (Linear / Vercel / Raycast):
slate-gray neutrals, deep indigo primary, semantic colors (success/warn/err)
that pop against white cards. No gradients, no drop-shadows (Tkinter can't
render them anyway) — relying on good colors + spacing + typography.

Change a value here and it propagates across every tab and widget.
"""
from __future__ import annotations

import platform


# ── Colors ────────────────────────────────────────────────────
# Tailwind CSS slate/indigo/emerald families — well-tested in real apps.
class Color:
    # Backgrounds
    BG_APP       = "#F1F5F9"   # slate-100 — window background
    BG_BANNER    = "#0F172A"   # slate-900 — top banner (deep navy)
    BG_CARD      = "#FFFFFF"   # pure white — cards float above the bg
    BG_INPUT     = "#FFFFFF"
    BG_SUBTLE    = "#F8FAFC"   # slate-50  — for inline highlights

    # Borders
    BORDER       = "#E2E8F0"   # slate-200 — subtle card borders
    BORDER_FOCUS = "#C7D2FE"   # indigo-200 — when something's focused

    # Text
    TEXT_PRIMARY   = "#0F172A"   # slate-900 — headings, values
    TEXT_SECONDARY = "#475569"   # slate-600 — subtitles, field labels
    TEXT_MUTED     = "#94A3B8"   # slate-400 — placeholders, meta info
    TEXT_ON_DARK   = "#FFFFFF"
    TEXT_LINK      = "#4F46E5"   # indigo-600

    # Primary action — deep indigo
    PRIMARY         = "#4F46E5"   # indigo-600
    PRIMARY_HOVER   = "#4338CA"   # indigo-700
    PRIMARY_PRESSED = "#3730A3"   # indigo-800
    PRIMARY_DISABLED = "#C7D2FE"  # indigo-200

    # Secondary (ghost / outline button)
    SECONDARY        = "#FFFFFF"
    SECONDARY_HOVER  = "#F1F5F9"
    SECONDARY_BORDER = "#CBD5E1"   # slate-300

    # Semantic
    SUCCESS      = "#059669"   # emerald-600
    SUCCESS_BG   = "#ECFDF5"   # emerald-50
    WARNING      = "#D97706"   # amber-600
    WARNING_BG   = "#FFFBEB"   # amber-50
    ERROR        = "#DC2626"   # red-600
    ERROR_BG     = "#FEF2F2"   # red-50


# ── Fonts ─────────────────────────────────────────────────────
# Use system UI fonts — Tkinter can't bundle custom TTFs reliably.
# Windows → Segoe UI, Mac → SF Pro (Helvetica Neue fallback), Linux → DejaVu.
def _system_font() -> str:
    s = platform.system()
    if s == "Windows":
        return "Segoe UI"
    if s == "Darwin":
        return "SF Pro Display"
    return "DejaVu Sans"


_F = _system_font()


class Font:
    TITLE      = (_F, 22, "bold")    # top banner
    HEADING    = (_F, 18, "bold")    # tab section title
    SUBHEAD    = (_F, 13)            # tab subtitle (muted)
    BODY       = (_F, 13)
    BODY_BOLD  = (_F, 13, "bold")
    BUTTON     = (_F, 14, "bold")
    BUTTON_LG  = (_F, 15, "bold")
    LABEL      = (_F, 12)
    SMALL      = (_F, 11)
    TAG        = (_F, 11, "bold")    # for pill badges


# ── Spacing scale ─────────────────────────────────────────────
# 4px base unit, powers-of-2 steps. Keeps margins/paddings consistent.
class Space:
    XS  = 4
    SM  = 8
    MD  = 12
    LG  = 16
    XL  = 24
    XXL = 32
    XXX = 48


# ── Radii ─────────────────────────────────────────────────────
class Radius:
    SM     = 6
    MD     = 10
    LG     = 14
    BUTTON = 10
    CARD   = 12
    PILL   = 999


# ── Geometry ──────────────────────────────────────────────────
BUTTON_H       = 44
BUTTON_H_LG    = 54   # hero "Generate" button
INPUT_H        = 40
BANNER_H       = 72
STATUS_H       = 32
