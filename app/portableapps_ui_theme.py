import tkinter as tk
from tkinter import ttk

from PIL import Image, ImageDraw, ImageTk


UI_COLORS = {
    "page": "#eef2f7",
    "surface": "#ffffff",
    "surface_alt": "#f8fafc",
    "toolbar": "#f7f9fc",
    "border": "#d8e0ea",
    "field": "#eef5fc",
    "field_soft": "#f3f8fe",
    "field_focus_fill": "#f7fbff",
    "field_border": "#b6c4d4",
    "field_border_strong": "#a8b6c7",
    "field_focus": "#7fa991",
    "button_fill": "#ffffff",
    "button_fill_hover": "#eef4fb",
    "button_border": "#c1ccd9",
    "button_border_hover": "#aebaca",
    "text": "#17212f",
    "muted": "#627386",
    "accent": "#1f7a57",
    "accent_hover": "#186346",
    "accent_soft": "#e6f4ee",
    "accent_line": "#b7dcc8",
    "card_header": "#f8fafc",
    "card_header_active": "#eef5fb",
    "danger": "#b42318",
    "danger_hover": "#912018",
    "danger_soft": "#fbe9e7",
    "danger_line": "#ebbbb6",
    "soft": "#edf4fa",
    "warn": "#9a5b12",
    "warn_soft": "#fff6e8",
    "warn_line": "#f2d3a0",
}


def build_checkbox_style_images(colors):
    def make_checkbox(fill, border, check=None):
        scale = 4
        image = Image.new("RGBA", (24 * scale, 20 * scale), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        draw.ellipse((1 * scale, 1 * scale, 19 * scale, 19 * scale), fill=fill, outline=border, width=2 * scale)
        if check:
            draw.line((5 * scale, 10 * scale, 8 * scale, 13 * scale), fill=check, width=2 * scale)
            draw.line((8 * scale, 13 * scale, 14 * scale, 7 * scale), fill=check, width=2 * scale)
        image = image.resize((24, 20), Image.Resampling.LANCZOS)
        return ImageTk.PhotoImage(image)

    return {
        "unchecked": make_checkbox("#ffffff", colors["border"]),
        "unchecked_hover": make_checkbox(colors["surface_alt"], colors["accent_line"]),
        "unchecked_disabled": make_checkbox(colors["surface_alt"], colors["border"]),
        "checked": make_checkbox(colors["accent"], colors["accent"], "#ffffff"),
        "checked_hover": make_checkbox(colors["accent_hover"], colors["accent_hover"], "#ffffff"),
        "checked_disabled": make_checkbox(colors["accent_line"], colors["accent_line"], "#ffffff"),
    }


def create_root_window():
    return tk.Tk()


def setup_ttk_styles(colors):
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    checkbox_style_images = build_checkbox_style_images(colors)

    style.configure("TFrame", background=colors["page"])
    style.configure("Surface.TFrame", background=colors["surface"])
    style.configure("PanelBody.TFrame", background=colors["surface"])
    style.configure("TLabel", background=colors["page"], foreground=colors["text"], font=("Segoe UI", 9))
    style.configure("Muted.TLabel", background=colors["page"], foreground=colors["muted"], font=("Segoe UI", 9))
    style.configure("Surface.TLabel", background=colors["surface"], foreground=colors["text"], font=("Segoe UI", 9))
    style.configure("PanelTitle.TLabel", background=colors["surface"], foreground=colors["text"], font=("Segoe UI Semibold", 10))
    style.configure("PanelNote.TLabel", background=colors["surface"], foreground=colors["muted"], font=("Segoe UI", 8))
    style.configure("Title.TLabel", background=colors["page"], foreground=colors["text"], font=("Segoe UI Semibold", 16))
    style.configure(
        "TButton",
        font=("Segoe UI Semibold", 9),
        padding=(14, 7),
        background=colors["button_fill"],
        foreground=colors["text"],
        borderwidth=1,
        focusthickness=0,
        relief="solid",
        focuscolor=colors["button_fill"],
        bordercolor=colors["button_border"],
        lightcolor=colors["button_border"],
        darkcolor=colors["button_border"],
    )
    style.map(
        "TButton",
        background=[("active", colors["button_fill_hover"]), ("pressed", colors["button_fill_hover"])],
        foreground=[("disabled", colors["muted"])],
        bordercolor=[("active", colors["button_border_hover"]), ("pressed", colors["button_border_hover"])],
        lightcolor=[("active", colors["button_border_hover"]), ("pressed", colors["button_border_hover"])],
        darkcolor=[("active", colors["button_border_hover"]), ("pressed", colors["button_border_hover"])],
    )
    style.configure(
        "Accent.TButton",
        background=colors["accent"],
        foreground="#ffffff",
        borderwidth=1,
        focusthickness=0,
        focuscolor=colors["accent"],
        relief="solid",
        bordercolor=colors["accent"],
        lightcolor=colors["accent"],
        darkcolor=colors["accent"],
    )
    style.map(
        "Accent.TButton",
        background=[("active", colors["accent_hover"]), ("pressed", colors["accent_hover"])],
        foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
    )
    style.configure(
        "Danger.TButton",
        background=colors["danger_soft"],
        foreground=colors["danger"],
        padding=(12, 8),
        borderwidth=1,
        focusthickness=0,
        focuscolor=colors["danger_line"],
        relief="solid",
        bordercolor=colors["danger_line"],
        lightcolor=colors["danger_line"],
        darkcolor=colors["danger_line"],
    )
    style.map(
        "Danger.TButton",
        background=[("active", colors["danger"]), ("pressed", colors["danger"])],
        foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
    )

    if "Web.Checkbutton.indicator" not in style.element_names():
        style.element_create(
            "Web.Checkbutton.indicator",
            "image",
            checkbox_style_images["unchecked"],
            ("disabled", "selected", checkbox_style_images["checked_disabled"]),
            ("disabled", checkbox_style_images["unchecked_disabled"]),
            ("pressed", "selected", checkbox_style_images["checked_hover"]),
            ("active", "selected", checkbox_style_images["checked_hover"]),
            ("selected", checkbox_style_images["checked"]),
            ("pressed", checkbox_style_images["unchecked_hover"]),
            ("active", checkbox_style_images["unchecked_hover"]),
            border=0,
            sticky="w",
        )
    style.layout(
        "TCheckbutton",
        [
            (
                "Checkbutton.padding",
                {
                    "sticky": "nswe",
                    "children": [
                        ("Web.Checkbutton.indicator", {"side": "left", "sticky": "w"}),
                        ("Checkbutton.label", {"side": "left", "sticky": "w"}),
                    ],
                },
            )
        ],
    )
    style.configure("TCheckbutton", background=colors["surface"], foreground=colors["text"], font=("Segoe UI", 9), padding=(0, 2))
    style.map("TCheckbutton", foreground=[("disabled", colors["muted"])], background=[("active", colors["surface"])])

    style.configure(
        "Web.TEntry",
        fieldbackground=colors["field_soft"],
        foreground=colors["text"],
        padding=(8, 7),
        borderwidth=1,
        relief="solid",
        bordercolor=colors["field_border"],
        lightcolor=colors["field_border"],
        darkcolor=colors["field_border"],
        insertcolor=colors["text"],
    )
    style.map(
        "Web.TEntry",
        fieldbackground=[
            ("disabled", colors["surface_alt"]),
            ("readonly", colors["field_soft"]),
            ("focus", colors["field_focus_fill"]),
        ],
        foreground=[("disabled", colors["muted"])],
        bordercolor=[("focus", colors["field_focus"])],
        lightcolor=[("focus", colors["field_focus"])],
        darkcolor=[("focus", colors["field_focus"])],
    )

    style.configure(
        "Web.TCombobox",
        fieldbackground=colors["field_soft"],
        background=colors["field_soft"],
        foreground=colors["text"],
        padding=(8, 7),
        arrowsize=13,
        borderwidth=1,
        relief="solid",
        arrowcolor=colors["muted"],
        bordercolor=colors["field_border"],
        lightcolor=colors["field_border"],
        darkcolor=colors["field_border"],
        selectbackground=colors["field_soft"],
        selectforeground=colors["text"],
    )
    style.map(
        "Web.TCombobox",
        fieldbackground=[
            ("disabled", colors["surface_alt"]),
            ("readonly", colors["field_soft"]),
            ("focus", colors["field_focus_fill"]),
        ],
        background=[
            ("disabled", colors["surface_alt"]),
            ("readonly", colors["field_soft"]),
            ("focus", colors["field_focus_fill"]),
        ],
        foreground=[("disabled", colors["muted"])],
        arrowcolor=[("disabled", colors["muted"]), ("active", colors["text"]), ("focus", colors["text"])],
        bordercolor=[("focus", colors["field_focus"])],
        lightcolor=[("focus", colors["field_focus"])],
        darkcolor=[("focus", colors["field_focus"])],
    )

    style.configure(
        "Vertical.TScrollbar",
        background=colors["surface_alt"],
        troughcolor=colors["toolbar"],
        bordercolor=colors["border"],
        arrowcolor=colors["muted"],
        darkcolor=colors["surface_alt"],
        lightcolor=colors["surface_alt"],
        gripcount=0,
    )
    style.configure("TNotebook", background=colors["surface"], borderwidth=0, tabmargins=(0, 0, 0, 0))
    style.configure("TNotebook.Tab", background=colors["page"], foreground=colors["muted"], padding=(16, 10), borderwidth=0)
    style.map(
        "TNotebook.Tab",
        background=[("selected", colors["surface"]), ("active", colors["surface_alt"])],
        foreground=[("selected", colors["text"]), ("active", colors["text"])],
    )

    return checkbox_style_images


def create_combobox(colors, parent, *, textvariable, values, width=None):
    combo = ttk.Combobox(
        parent,
        textvariable=textvariable,
        values=list(values),
        state="readonly",
        width=width if width is not None else 1,
        style="Web.TCombobox",
    )
    combo.combo = combo
    original_cget = combo.cget
    combo.cget = lambda name: textvariable.get() if name == "value" else original_cget(name)
    return combo


def create_entry(colors, parent, *, textvariable, state="normal", width=None):
    entry = ttk.Entry(
        parent,
        textvariable=textvariable,
        width=width if width is not None else 1,
        state="disabled" if state == "disabled" else "normal",
        style="Web.TEntry",
    )

    if state == "readonly":
        entry.configure(state="readonly")

    return entry


def create_scrollbar(colors, parent, *, orientation, command):
    orient = tk.VERTICAL if orientation == "vertical" else tk.HORIZONTAL
    return ttk.Scrollbar(parent, orient=orient, command=command)


def make_button(colors, parent, *, text, command, style=None, state="normal", width=None):
    options = {
        "text": text,
        "command": command,
        "state": state,
        "style": style or "TButton",
    }
    if width is not None:
        options["width"] = width
    return ttk.Button(parent, **options)
