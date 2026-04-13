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
# Light theme inspired by Linear / Vercel — warm-tinted whites, soft
# violet primary, very subtle borders. Goal: feels premium without
# leaving Tk's rendering envelope.
class Color:
    # Backgrounds — a touch warmer than pure slate, less "Bootstrap-ey"
    BG_APP       = "#FAFAFB"   # warm off-white window background
    BG_BANNER    = "#0B0B12"   # near-black with a hint of violet
    BG_CARD      = "#FFFFFF"   # pure white cards float above bg
    BG_INPUT     = "#FFFFFF"
    BG_SUBTLE    = "#F4F4F7"   # for inline highlights / hover surfaces

    # Borders — softer, more subtle than slate-200
    BORDER       = "#EAEAEF"   # default card/input border
    BORDER_HOVER = "#D8D8E0"   # slightly stronger on hover
    BORDER_FOCUS = "#C7C2F4"   # violet-tinted when focused

    # Text
    TEXT_PRIMARY   = "#0B0B12"   # near-black headings/values
    TEXT_SECONDARY = "#4B4B58"   # body text
    TEXT_MUTED     = "#8E8E9A"   # placeholders, meta, captions
    TEXT_ON_DARK   = "#FFFFFF"
    TEXT_LINK      = "#5B5BD6"

    # Primary action — soft violet (Linear-style), warmer than indigo-600
    PRIMARY         = "#5B5BD6"
    PRIMARY_HOVER   = "#4848C2"
    PRIMARY_PRESSED = "#3C3CA8"
    PRIMARY_DISABLED = "#C9C9F0"

    # Secondary (ghost / outline button)
    SECONDARY        = "#FFFFFF"
    SECONDARY_HOVER  = "#F4F4F7"
    SECONDARY_BORDER = "#D8D8E0"

    # Semantic — slightly desaturated for a calmer feel
    SUCCESS      = "#10A37F"   # softer emerald
    SUCCESS_BG   = "#E8F8F2"
    WARNING      = "#D97706"
    WARNING_BG   = "#FFF8EE"
    ERROR        = "#DC2626"
    ERROR_BG     = "#FEF1F1"


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
    TITLE      = (_F, 20, "bold")    # top banner
    HEADING    = (_F, 22, "bold")    # tab section title — more visual weight
    SUBHEAD    = (_F, 13)            # tab subtitle (muted)
    BODY       = (_F, 13)
    BODY_BOLD  = (_F, 13, "bold")
    BUTTON     = (_F, 13, "bold")
    BUTTON_LG  = (_F, 14, "bold")
    LABEL      = (_F, 12)
    SMALL      = (_F, 11)
    TAG        = (_F, 10, "bold")    # for pill badges (smaller, tighter)


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
# Slightly larger radii than v1 — Linear/Vercel use 8–12 for inputs and
# 14–16 for cards. Reads as more "premium" without going full pill.
class Radius:
    SM     = 8
    MD     = 12
    LG     = 16
    BUTTON = 12
    CARD   = 14
    PILL   = 999


# ── Geometry ──────────────────────────────────────────────────
BUTTON_H       = 42
BUTTON_H_LG    = 56   # hero "Generate" button — slightly taller, more confident
INPUT_H        = 42
BANNER_H       = 64   # tighter banner — less navy, more content
STATUS_H       = 30
