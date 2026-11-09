"""
eFBR CPR Downloader — Themes
============================
Colour palettes for every Design Mode x Appearance combination.

Design modes:
  classic     the original look (coloured header bar, flat grey panels)
  glass       glassmorphism  — gradient backdrop, frosted translucent cards
  neumorphic  neumorphism    — one soft surface, cards raised by twin shadows
  deep        dark colours   — flat cards, dark teal / emerald / jungle fills

Appearance:  light | dark

Tkinter has no real transparency or blur, so "glass" cards are colours
pre-blended against the gradient, and neumorphic shadows are painted on
a canvas behind the cards (see gui.py).
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Dict, Tuple

from constants import ACCENT, ERROR_CLR, WARN_CLR, get_settings_path

# ── Brand colours ─────────────────────────────────────────────────────────────
TEAL         = "#008080"    # primary actions, progress
TEAL_BRIGHT  = "#2DD4BF"    # teal on dark surfaces
EMERALD      = "#50C878"    # Start button, success
JUNGLE       = ACCENT       # tropical jungle green — header, hovers, deep tones
INK          = "#04201D"    # text on bright teal / emerald fills
START_BTN    = (EMERALD, "#3DB868", INK, 0)

# Darker shades of the same three, used by the "Deep" design
DEEP_TEAL    = "#006666"
DEEP_EMERALD = "#0B6B43"
DEEP_JUNGLE  = "#00483A"

DESIGNS = {"Classic": "classic", "Glass": "glass",
           "Neumorphic": "neumorphic", "Deep": "deep"}
MODES = {"Light": "light", "Dark": "dark"}

DEFAULT_DESIGN = "glass"
DEFAULT_MODE = "dark"


# ── Colour helpers ────────────────────────────────────────────────────────────

def _rgb(hex_color: str) -> Tuple[int, int, int]:
    h = hex_color.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def blend(base: str, over: str, t: float) -> str:
    """Mix `over` into `base`; t=0 gives base, t=1 gives over."""
    t = max(0.0, min(1.0, t))
    b, o = _rgb(base), _rgb(over)
    return "#{:02X}{:02X}{:02X}".format(
        *(round(b[i] + (o[i] - b[i]) * t) for i in range(3))
    )


# ── Palette ───────────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class Palette:
    design: str
    mode: str

    # Window backdrop (top == bottom → flat; otherwise a vertical gradient)
    bg_top: str
    bg_bottom: str

    # Cards
    card: str
    border: str
    border_width: int
    radius: int

    # Text
    text: str
    muted: str

    # Accent / interactive
    accent: str
    accent_hover: str
    on_accent: str
    hover: str                  # neutral hover (dropdown rows, secondary buttons)

    # Inputs, progress track, log
    input_bg: str
    input_border: str
    track: str
    log_bg: str

    # Semantic
    ok: str
    err: str
    warn: str
    info: str

    # Header (classic uses a coloured bar; the others reuse the card colours)
    header_bg: str
    header_text: str
    header_muted: str
    seg_bg: str
    seg_selected: str
    seg_hover: str

    # Neumorphic shadows ("" → no shadows)
    shadow_dark: str = ""
    shadow_light: str = ""

    # kind -> (fill, hover, text, border_width)
    buttons: Dict[str, Tuple[str, str, str, int]] = field(default_factory=dict)

    @property
    def bg_mid(self) -> str:
        return blend(self.bg_top, self.bg_bottom, 0.5)

    def color(self, role: str) -> str:
        """Resolve a semantic text role to a colour."""
        return {
            "text": self.text, "muted": self.muted, "accent": self.accent,
            "ok": self.ok, "err": self.err, "warn": self.warn,
            "h_text": self.header_text, "h_muted": self.header_muted,
        }[role]


def _classic(mode: str) -> Palette:
    dark = mode == "dark"
    bg = "#242424" if dark else "#EBEBEB"
    card = "#2B2B2B" if dark else "#DBDBDB"
    text = "#DCE4EE" if dark else "#1A1A1A"
    primary, primary_hover = TEAL, JUNGLE
    return Palette(
        design="classic", mode=mode,
        bg_top=bg, bg_bottom=bg,
        card=card, border=card, border_width=0, radius=6,
        text=text, muted="#9A9A9A" if dark else "#5F5F5F",
        accent="#14B8A6" if dark else primary, accent_hover=primary_hover, on_accent="#FFFFFF",
        hover="#3A3D40" if dark else "#C9CDD1",
        input_bg="#343638" if dark else "#F9F9FA",
        input_border="#565B5E" if dark else "#979DA2",
        track="#4A4D50" if dark else "#B4BABF",
        log_bg="#1D1E1E" if dark else "#F9F9FA",
        ok=EMERALD if dark else "#0B7A4B",
        err=ERROR_CLR if dark else "#C0392B",
        warn=WARN_CLR if dark else "#8A6100",
        info="#AAAAAA" if dark else "#555555",
        header_bg=JUNGLE, header_text="#FFFFFF", header_muted="#CDEBE6",
        seg_bg="#005F4C", seg_selected="#004638", seg_hover="#006B56",
        buttons={
            "primary":   (primary, primary_hover, "#FFFFFF", 0),
            "secondary": (card, "#3A3D40" if dark else "#C9CDD1", text, 1),
            "start":     START_BTN,
            "pause":     (WARN_CLR, "#B8860B", "#1A1A1A", 0),
            "cancel":    (ERROR_CLR, "#B03A3A", "#FFFFFF", 0),
        },
    )


def _glass(mode: str) -> Palette:
    dark = mode == "dark"
    if dark:
        top, bottom = "#052B30", "#03382C"
        tint, text, muted = "#FFFFFF", "#EFFBF8", "#9FC2BD"
        accent, accent_hover, on_accent = TEAL_BRIGHT, "#5EEAD4", INK
        ok, err, warn, info = EMERALD, "#FF7A7A", "#F5C451", "#9FC2BD"
        card_t, border_t, input_t = 0.09, 0.24, 0.05
    else:
        top, bottom = "#CDEFEF", "#D3F3DF"
        tint, text, muted = "#FFFFFF", "#0F2A28", "#48635F"
        accent, accent_hover, on_accent = TEAL, JUNGLE, "#FFFFFF"
        ok, err, warn, info = "#0B7A4B", "#C0392B", "#8A6100", "#48635F"
        card_t, border_t, input_t = 0.55, 1.0, 0.80
    mid = blend(top, bottom, 0.5)
    card = blend(mid, tint, card_t)
    border = blend(mid, tint, border_t)
    input_bg = blend(mid, "#000000", 0.22) if dark else blend(mid, tint, input_t)
    hover = blend(card, tint if dark else accent, 0.12)
    return Palette(
        design="glass", mode=mode,
        bg_top=top, bg_bottom=bottom,
        card=card, border=border, border_width=1, radius=16,
        text=text, muted=muted,
        accent=accent, accent_hover=accent_hover, on_accent=on_accent,
        hover=hover,
        input_bg=input_bg, input_border=border,
        track=blend(mid, "#000000", 0.30) if dark else blend(mid, accent, 0.18),
        log_bg=input_bg,
        ok=ok, err=err, warn=warn, info=info,
        header_bg=card, header_text=text, header_muted=muted,
        seg_bg=input_bg,
        seg_selected=blend(card, tint, 0.16) if dark else "#FFFFFF",
        seg_hover=hover,
        buttons={
            "primary":   (accent, accent_hover, on_accent, 0),
            "secondary": (blend(card, tint, 0.05 if dark else 0.5), hover, text, 1),
            "start":     START_BTN,
            "pause":     (blend(card, tint, 0.05 if dark else 0.5), hover, text, 1),
            "cancel":    (blend(card, tint, 0.05 if dark else 0.5), hover, err, 1),
        },
    )


def _neumorphic(mode: str) -> Palette:
    dark = mode == "dark"
    if dark:
        bg, text, muted = "#232B2B", "#E2EEEC", "#93A8A4"
        shadow_dark, shadow_light = "#131818", "#343F3E"
        accent, accent_hover, on_accent = TEAL_BRIGHT, "#5EEAD4", INK
        ok, err, warn, info = EMERALD, "#FF7373", "#F2C14E", "#93A8A4"
        inset = "#1C2323"
    else:
        bg, text, muted = "#E4EDEB", "#213533", "#566C68"
        shadow_dark, shadow_light = "#B0C2BE", "#FFFFFF"
        accent, accent_hover, on_accent = TEAL, JUNGLE, "#FFFFFF"
        ok, err, warn, info = "#0B7A4B", "#C0392B", "#8A6100", "#566C68"
        inset = "#D7E2DF"
    edge = blend(bg, shadow_dark, 0.45)
    hover = blend(bg, shadow_dark, 0.25)
    return Palette(
        design="neumorphic", mode=mode,
        bg_top=bg, bg_bottom=bg,
        card=bg, border=bg, border_width=0, radius=18,
        text=text, muted=muted,
        accent=accent, accent_hover=accent_hover, on_accent=on_accent,
        hover=hover,
        input_bg=inset, input_border=edge,
        track=inset, log_bg=inset,
        ok=ok, err=err, warn=warn, info=info,
        header_bg=bg, header_text=text, header_muted=muted,
        seg_bg=inset, seg_selected=blend(bg, shadow_light, 0.6), seg_hover=hover,
        shadow_dark=shadow_dark, shadow_light=shadow_light,
        buttons={
            "primary":   (accent, accent_hover, on_accent, 0),
            "secondary": (bg, hover, text, 1),
            "start":     START_BTN,
            "pause":     (bg, hover, text, 1),
            "cancel":    (bg, hover, err, 1),
        },
    )


def _deep(mode: str) -> Palette:
    dark = mode == "dark"
    if dark:
        bg, card, border = "#03181A", "#0A2B2B", "#174642"
        text, muted, hover = "#E6F4F1", "#8FB3AD", "#123A38"
        accent = "#26A69A"                  # readable teal on dark surfaces
        input_bg, input_border, track, log_bg = "#061F20", "#1C4A47", "#061F20", "#041517"
        ok, err, warn = EMERALD, "#FF7A7A", "#F5C451"
    else:
        bg, card, border = "#E8EFED", "#FFFFFF", "#C9D6D3"
        text, muted, hover = "#0F2A28", "#4C635F", "#E3ECEA"
        accent = DEEP_TEAL
        input_bg, input_border, track, log_bg = "#F2F6F5", "#B5C5C1", "#D3DEDB", "#F7FAF9"
        ok, err, warn = DEEP_EMERALD, "#A93226", "#7A5600"
    return Palette(
        design="deep", mode=mode,
        bg_top=bg, bg_bottom=bg,
        card=card, border=border, border_width=1, radius=10,
        text=text, muted=muted,
        accent=accent, accent_hover="#004C4C", on_accent="#FFFFFF",
        hover=hover,
        input_bg=input_bg, input_border=input_border,
        track=track, log_bg=log_bg,
        ok=ok, err=err, warn=warn, info=muted,
        header_bg=DEEP_JUNGLE, header_text="#FFFFFF", header_muted="#B7D8CF",
        seg_bg="#00352B", seg_selected="#006652", seg_hover="#005443",
        buttons={
            "primary":   (DEEP_TEAL, "#004C4C", "#FFFFFF", 0),
            "secondary": (card, hover, text, 1),
            "start":     (DEEP_EMERALD, "#09573A", "#FFFFFF", 0),
            "pause":     (card, hover, text, 1),
            "cancel":    (card, hover, err, 1),
        },
    )


_BUILDERS = {"classic": _classic, "glass": _glass,
             "neumorphic": _neumorphic, "deep": _deep}


def get_palette(design: str, mode: str) -> Palette:
    return _BUILDERS.get(design, _glass)(mode if mode in ("light", "dark") else "dark")


# ── Saved preferences ─────────────────────────────────────────────────────────

def load_prefs() -> Tuple[str, str]:
    """Return the saved (design, mode), falling back to the defaults."""
    try:
        data = json.loads(get_settings_path().read_text(encoding="utf-8"))
        design, mode = data.get("design"), data.get("mode")
    except Exception:
        design = mode = None
    if design not in _BUILDERS:
        design = DEFAULT_DESIGN
    if mode not in ("light", "dark"):
        mode = DEFAULT_MODE
    return design, mode


def save_prefs(design: str, mode: str) -> None:
    try:
        get_settings_path().write_text(
            json.dumps({"design": design, "mode": mode}, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass                    # preferences are a nicety, never fatal
