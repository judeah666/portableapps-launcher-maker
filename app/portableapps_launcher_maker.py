import shutil
import subprocess
import tempfile
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageDraw, ImageOps, ImageTk

from app.portableapps_core import (
    DEFAULT_SPLASH_ASSET,
    HELP_IMAGE_FILENAMES,
    ICON_PREVIEW_DISPLAY_SIZES,
    LAUNCHER_TEMPLATE_FILENAMES,
    PORTABLEAPPS_DEVELOPMENT_DOWNLOADS_URL,
    SOFTWARE_ICON,
    SOFTWARE_ICON_PNG,
    TEMPLATE_ASSET_SPECS,
    ValidationItem,
    LauncherProject,
    app_base_path,
    asset_path,
    bool_to_ini,
    build_appinfo_ini,
    build_help_html,
    build_installer_ini,
    build_launcher_ini,
    build_readme,
    build_registry_key_entries_from_reg_text,
    build_validation_items,
    clean_display_name,
    clean_identifier,
    clean_ini_lines,
    create_help_images,
    create_launcher_project,
    create_launcher_template_assets,
    default_portableapps_output_dir,
    detect_app_name_from_exe,
    extract_embedded_icon,
    find_portableapps_launcher,
    has_ini_lines,
    help_image_asset_path,
    load_icon_image,
    make_fallback_icon,
    merge_ini_line_sets,
    normalize_registry_path,
    parse_registry_paths_from_reg_text,
    render_validation_report,
    resolve_project_tokens,
    splash_asset_path,
    validate_ini_mapping_lines,
    validate_project,
)
from app.portableapps_ui_theme import (
    UI_COLORS,
    create_root_window,
    create_combobox as themed_create_combobox,
    create_entry as themed_create_entry,
    create_scrollbar as themed_create_scrollbar,
    make_button as themed_make_button,
    setup_ttk_styles,
)
from app.version import APP_VERSION


class ScrolledText(tk.Frame):
    """Small Tk text editor wrapper with built-in ttk scrollbars."""

    def __init__(self, parent, colors, *, height=4, background=None, wrap="none", font=("Consolas", 11)):
        super().__init__(
            parent,
            bg=colors["field_border"],
            highlightthickness=1,
            highlightbackground=colors["field_border"],
            bd=0,
        )
        self.colors = colors
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        field_background = background or colors["field_soft"]
        self._text = tk.Text(
            self,
            height=height,
            wrap=wrap,
            bg=field_background,
            fg=colors["text"],
            insertbackground=colors["text"],
            selectbackground=colors["accent_soft"],
            selectforeground=colors["text"],
            relief="flat",
            bd=0,
            highlightthickness=0,
            font=font,
            undo=True,
        )
        self._yscrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self._text.yview)
        self._text.configure(yscrollcommand=self._yscrollbar.set)
        self._text.grid(row=0, column=0, sticky="nsew")
        self._yscrollbar.grid(row=0, column=1, sticky="ns")

    def insert(self, *args, **kwargs):
        return self._text.insert(*args, **kwargs)

    def delete(self, *args, **kwargs):
        return self._text.delete(*args, **kwargs)

    def get(self, *args, **kwargs):
        return self._text.get(*args, **kwargs)

    def bind(self, sequence=None, func=None, add=None):
        return self._text.bind(sequence, func, add)

    def tag_configure(self, *args, **kwargs):
        return self._text.tag_configure(*args, **kwargs)

    def configure(self, cnf=None, **kwargs):
        if cnf:
            kwargs.update(cnf)
        frame_options = {}
        text_options = {}
        option_aliases = {
            "border_width": "bd",
            "text_color": "fg",
            "fg_color": "bg",
        }
        for key, value in kwargs.items():
            key = option_aliases.get(key, key)
            if key in {"bg", "background", "highlightthickness", "highlightbackground"}:
                frame_options[key] = value
            else:
                text_options[key] = value
        if frame_options:
            super().configure(**frame_options)
        if text_options:
            return self._text.configure(**text_options)
        return None

    config = configure

    def cget(self, key):
        if key in {"bg", "background", "highlightthickness", "highlightbackground"}:
            return super().cget(key)
        key = {"border_width": "bd", "text_color": "fg", "fg_color": "bg"}.get(key, key)
        return self._text.cget(key)

    def __getattr__(self, name):
        return getattr(self._text, name)


class PortableAppsLauncherMaker:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PortableApps.com Launcher Maker")
        self.root.geometry("1120x760")
        self.root.minsize(980, 660)
        self.app_icon_image = None

        self.colors = UI_COLORS.copy()

        self.vars = {
            "app_name": tk.StringVar(),
            "package_name": tk.StringVar(),
            "publisher": tk.StringVar(),
            "trademarks": tk.StringVar(),
            "homepage": tk.StringVar(value="https://portableapps.com/"),
            "category": tk.StringVar(value="Utilities"),
            "language": tk.StringVar(value="Multilingual"),
            "description": tk.StringVar(),
            "donate": tk.StringVar(),
            "install_type": tk.StringVar(),
            "version": tk.StringVar(value=APP_VERSION),
            "display_version": tk.StringVar(value=APP_VERSION),
            "app_exe": tk.StringVar(),
            "output_dir": tk.StringVar(value=default_portableapps_output_dir()),
            "command_line": tk.StringVar(),
            "working_directory": tk.StringVar(value="%PAL:AppDir%\\{app_name}"),
            "close_exe": tk.StringVar(),
            "wait_for_other_instances": tk.BooleanVar(value=True),
            "min_os": tk.StringVar(),
            "max_os": tk.StringVar(),
            "run_as_admin": tk.StringVar(),
            "refresh_shell_icons": tk.StringVar(),
            "hide_command_line_window": tk.BooleanVar(value=False),
            "no_spaces_in_path": tk.BooleanVar(value=False),
            "supports_unc": tk.StringVar(),
            "activate_java": tk.StringVar(),
            "activate_xml": tk.BooleanVar(value=False),
            "live_mode_copy_app": tk.BooleanVar(value=False),
            "live_mode_copy_data": tk.BooleanVar(value=False),
            "files_move": tk.StringVar(),
            "directories_move": tk.StringVar(),
            "installer_close_exe": tk.StringVar(),
            "installer_close_name": tk.StringVar(),
            "include_installer_source": tk.BooleanVar(value=False),
            "remove_app_directory": tk.BooleanVar(value=False),
            "remove_data_directory": tk.BooleanVar(value=False),
            "remove_other_directory": tk.BooleanVar(value=False),
            "optional_components_enabled": tk.BooleanVar(value=False),
            "main_section_title": tk.StringVar(),
            "main_section_description": tk.StringVar(),
            "optional_section_title": tk.StringVar(),
            "optional_section_description": tk.StringVar(),
            "optional_section_selected_install_type": tk.StringVar(),
            "optional_section_not_selected_install_type": tk.StringVar(),
            "optional_section_preselected": tk.StringVar(),
            "installer_languages": tk.StringVar(),
            "preserve_directories": tk.StringVar(),
            "remove_directories": tk.StringVar(),
            "preserve_files": tk.StringVar(),
            "remove_files": tk.StringVar(),
            "icon_source": tk.StringVar(),
            "icon_index": tk.StringVar(value="0"),
            "registry_enabled": tk.BooleanVar(value=False),
            "registry_keys": tk.StringVar(),
            "registry_cleanup_if_empty": tk.StringVar(),
            "registry_cleanup_force": tk.StringVar(),
            "copy_app_files": tk.BooleanVar(value=True),
            "wait_for_program": tk.BooleanVar(value=True),
            "license_shareable": tk.BooleanVar(value=True),
            "license_open_source": tk.BooleanVar(value=False),
            "license_freeware": tk.BooleanVar(value=True),
            "license_commercial_use": tk.BooleanVar(value=True),
            "license_eula_version": tk.StringVar(),
            "special_plugins": tk.StringVar(value="NONE"),
            "dependency_uses_ghostscript": tk.StringVar(value="no"),
            "dependency_uses_java": tk.StringVar(value="no"),
            "dependency_uses_dotnet_version": tk.StringVar(),
            "dependency_requires_64bit_os": tk.StringVar(value="no"),
            "dependency_requires_portable_app": tk.StringVar(),
            "dependency_requires_admin": tk.StringVar(value="no"),
            "control_icons": tk.StringVar(value="1"),
            "control_start": tk.StringVar(),
            "control_extract_icon": tk.StringVar(),
            "control_extract_name": tk.StringVar(),
            "control_base_app_id": tk.StringVar(),
            "control_base_app_id_64": tk.StringVar(),
            "control_base_app_id_arm64": tk.StringVar(),
            "control_exit_exe": tk.StringVar(),
            "control_exit_parameters": tk.StringVar(),
            "association_file_types": tk.StringVar(),
            "association_file_type_command_line": tk.StringVar(),
            "association_file_type_command_line_extension": tk.StringVar(),
            "association_protocols": tk.StringVar(),
            "association_protocol_command_line": tk.StringVar(),
            "association_protocol_command_line_protocol": tk.StringVar(),
            "association_send_to": tk.BooleanVar(value=False),
            "association_send_to_command_line": tk.StringVar(),
            "association_shell": tk.BooleanVar(value=False),
            "association_shell_command": tk.StringVar(),
            "file_type_icons": tk.StringVar(),
        }
        self.status_var = tk.StringVar(value="Choose an EXE and output folder, then create the launcher project.")
        self.preview_var = tk.StringVar()
        self.generator_status_var = tk.StringVar()
        self.active_scroll_canvas = None
        self.main_notebook = None
        self.main_tab_buttons = {}
        self.main_tab_frames = {}
        self.current_main_tab = None
        self.hover_main_tab = None
        self.preview_tab_buttons = {}
        self.preview_tab_frames = {}
        self.current_preview_tab = None
        self.hover_preview_tab = None
        self.launcher_tab = None
        self.help_window = None
        self.icon_preview_labels = []
        self.icon_preview_caption = None
        self.icon_preview_images = []
        self.icon_preview_cache_key = None
        self.sidebar_icon_preview_labels = []
        self.sidebar_icon_preview_caption = None
        self.sidebar_icon_preview_images = []
        self.sidebar_splash_preview_label = None
        self.sidebar_splash_preview_caption = None
        self.sidebar_splash_preview_image = None
        self.panel_cards = []
        self.create_button = None
        self.validate_button = None
        self.help_button = None
        self.import_registry_button = None
        self.detected_defaults = {
            "app_name": "",
            "package_name": "",
            "description": "",
            "control_start": "",
        }
        self.bound_text_widgets = {}
        self.template_asset_path_vars = {}
        self.template_splash_label = None
        self.template_splash_caption = None
        self.template_splash_image = None
        self.validation_window = None

        self.setup_styles()
        self.apply_window_icon(self.root)
        self.create_ui()
        self.refresh_generator_status()
        self.update_registry_controls()
        self.bind_preview_updates()
        self.refresh_preview()

    def setup_styles(self):
        self.checkbox_style_images = setup_ttk_styles(self.colors)

    def create_combobox(self, parent, *, textvariable, values, width=None):
        return themed_create_combobox(self.colors, parent, textvariable=textvariable, values=values, width=width)

    def create_entry(self, parent, *, textvariable, state="normal", width=None):
        return themed_create_entry(self.colors, parent, textvariable=textvariable, state=state, width=width)

    def make_button(self, parent, *, text, command, style=None, state="normal", width=None):
        return themed_make_button(self.colors, parent, text=text, command=command, style=style, state=state, width=width)

    def create_scrollbar(self, parent, *, orientation, command):
        return themed_create_scrollbar(self.colors, parent, orientation=orientation, command=command)

    def apply_window_icon(self, window):
        icon_png = asset_path(SOFTWARE_ICON_PNG)
        if icon_png.exists():
            try:
                with Image.open(icon_png) as source:
                    icon_image = ImageTk.PhotoImage(source.convert("RGBA"))
                window.iconphoto(True, icon_image)
                if window is self.root:
                    self.app_icon_image = icon_image
                else:
                    window._app_icon_image = icon_image
            except (tk.TclError, OSError):
                pass

        icon_ico = asset_path(SOFTWARE_ICON)
        if icon_ico.exists():
            try:
                window.iconbitmap(default=str(icon_ico))
            except tk.TclError:
                pass

    def create_ui(self):
        self.root.configure(bg=self.colors["page"])
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        body = ttk.Frame(self.root, padding=(16, 12, 16, 10))
        body.grid(row=0, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)
        body.rowconfigure(1, weight=0)

        self.root.bind_all("<MouseWheel>", self.scroll_form, add="+")
        self.root.bind_all("<Button-4>", self.scroll_form, add="+")
        self.root.bind_all("<Button-5>", self.scroll_form, add="+")

        notebook_shell = self.create_panel(body, "Project Settings", "Build and preview the PortableApps project from one place.", collapsible=False)
        notebook_shell.grid(row=0, column=0, sticky="nsew")
        notebook_shell.header.columnconfigure(1, weight=0)
        notebook_shell.content.columnconfigure(0, weight=1)
        notebook_shell.content.rowconfigure(1, weight=1)

        paths_bar = tk.Frame(
            notebook_shell.header,
            bg=self.colors["card_header"],
            highlightthickness=0,
            bd=0,
        )
        paths_bar.grid(row=0, column=1, sticky="e", padx=(24, 0))
        paths_bar.columnconfigure(0, weight=0)
        paths_bar.columnconfigure(1, weight=1, minsize=360)
        paths_bar.columnconfigure(2, weight=0)

        tk.Label(
            paths_bar,
            text="Application EXE",
            bg=self.colors["card_header"],
            fg=self.colors["text"],
            font=("Segoe UI", 8),
            anchor="e",
        ).grid(
            row=0,
            column=0,
            sticky="e",
            padx=(0, 8),
            pady=(0, 6),
        )
        self.create_entry(paths_bar, textvariable=self.vars["app_exe"]).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(0, 8),
            pady=(0, 6),
        )
        self.make_button(paths_bar, text="Browse", command=self.choose_app_exe).grid(
            row=0,
            column=2,
            sticky="ew",
            pady=(0, 6),
        )

        tk.Label(
            paths_bar,
            text="Output Folder",
            bg=self.colors["card_header"],
            fg=self.colors["text"],
            font=("Segoe UI", 8),
            anchor="e",
        ).grid(
            row=1,
            column=0,
            sticky="e",
            padx=(0, 8),
        )
        self.create_entry(paths_bar, textvariable=self.vars["output_dir"]).grid(
            row=1,
            column=1,
            sticky="ew",
            padx=(0, 8),
        )
        self.make_button(paths_bar, text="Browse", command=self.choose_output_dir).grid(
            row=1,
            column=2,
            sticky="ew",
        )

        tab_bar = tk.Frame(notebook_shell.content, bg=self.colors["toolbar"], highlightthickness=1, highlightbackground=self.colors["border"], bd=0)
        tab_bar.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        tabs_host = tk.Frame(tab_bar, bg=self.colors["toolbar"], highlightthickness=0, bd=0)
        tabs_host.pack(side="left", anchor="w", padx=8, pady=8)
        actions = tk.Frame(tab_bar, bg=self.colors["toolbar"], highlightthickness=0, bd=0)
        actions.pack(side="right", anchor="e", padx=8, pady=8)
        self.help_button = self.make_button(actions, text="Help", style="Danger.TButton", command=self.open_help)
        self.help_button.pack(side="left")

        tab_content = ttk.Frame(notebook_shell.content, style="Surface.TFrame")
        tab_content.grid(row=1, column=0, sticky="nsew")
        tab_content.columnconfigure(0, weight=1)
        tab_content.columnconfigure(1, weight=0)
        tab_content.rowconfigure(0, weight=1)

        editor_host = ttk.Frame(tab_content, style="Surface.TFrame")
        editor_host.grid(row=0, column=0, sticky="nsew")
        editor_host.columnconfigure(0, weight=1)
        editor_host.rowconfigure(0, weight=1)

        preview_sidebar = ttk.Frame(tab_content, style="Surface.TFrame", padding=(12, 0, 0, 0), width=450)
        preview_sidebar.grid(row=0, column=1, sticky="ns")
        preview_sidebar.grid_propagate(False)
        preview_sidebar.columnconfigure(0, weight=1)
        preview_sidebar.rowconfigure(0, weight=1)

        appinfo_tab, appinfo_content = self.create_scrollable_tab(editor_host)
        launcher_tab, launcher_content = self.create_scrollable_tab(editor_host)
        installer_tab, installer_content = self.create_scrollable_tab(editor_host)
        registry_tab, registry_content = self.create_scrollable_tab(editor_host)
        icon_tab, icon_content = self.create_scrollable_tab(editor_host)
        templates_tab, templates_content = self.create_scrollable_tab(editor_host)
        self.create_main_tab(tabs_host, "appinfo", "appinfo.ini", appinfo_tab)
        self.launcher_tab = self.create_main_tab(tabs_host, "launcher", "AppNamePortable.ini", launcher_tab)
        self.create_main_tab(tabs_host, "installer", "installer.ini", installer_tab)
        self.create_main_tab(tabs_host, "registry", "Registry", registry_tab)
        self.create_main_tab(tabs_host, "icon", "Icon", icon_tab)
        self.create_main_tab(tabs_host, "templates", "Splash", templates_tab)
        self.select_main_tab("appinfo")

        self.create_appinfo_editor(appinfo_content)

        self.create_launcher_editor(launcher_content)
        self.create_installer_editor(installer_content)

        self.create_registry_editor(registry_content)
        self.create_icon_editor(icon_content)
        self.create_template_editor(templates_content)

        preview_shell = self.create_panel(
            preview_sidebar,
            "Preview",
            "Live project preview while you edit settings.",
            collapsible=False,
        )
        preview_shell.grid(row=0, column=0, sticky="nsew")
        preview_shell.content.columnconfigure(0, weight=1)
        preview_shell.content.rowconfigure(0, weight=1)

        preview_canvas = tk.Canvas(
            preview_shell.content,
            bg=self.colors["surface"],
            highlightthickness=0,
            bd=0,
            yscrollincrement=24,
        )
        preview_scrollbar = self.create_scrollbar(preview_shell.content, orientation="vertical", command=preview_canvas.yview)
        preview_canvas.configure(yscrollcommand=preview_scrollbar.set)
        preview_canvas.grid(row=0, column=0, sticky="nsew")
        preview_scrollbar.grid(row=0, column=1, sticky="ns")

        preview_body = ttk.Frame(preview_canvas, style="Surface.TFrame")
        preview_body.columnconfigure(0, weight=1)
        preview_window_id = preview_canvas.create_window((0, 0), window=preview_body, anchor="nw")

        preview_body.bind("<Configure>", lambda _event: preview_canvas.configure(scrollregion=preview_canvas.bbox("all")))
        preview_canvas.bind("<Configure>", lambda event: preview_canvas.itemconfigure(preview_window_id, width=event.width))

        preview_canvas._scroll_canvas = preview_canvas
        preview_body._scroll_canvas = preview_canvas
        preview_canvas.bind("<Enter>", lambda _event, target=preview_canvas: self.set_active_scroll_canvas(target), add="+")
        preview_body.bind("<Enter>", lambda _event, target=preview_canvas: self.set_active_scroll_canvas(target), add="+")
        preview_canvas.bind("<Leave>", lambda _event, target=preview_canvas: self.clear_active_scroll_canvas(target), add="+")
        preview_body.bind("<Leave>", lambda _event, target=preview_canvas: self.clear_active_scroll_canvas(target), add="+")

        file_preview_shell = self.create_panel(
            preview_body,
            "File Folder Preview",
            "Generated project folder structure.",
            collapsible=False,
        )
        file_preview_shell.grid(row=0, column=0, sticky="nsew")
        file_preview_shell.content.columnconfigure(0, weight=1)
        file_preview_shell.content.rowconfigure(1, weight=1)

        preview_bar = tk.Frame(
            file_preview_shell.content,
            bg=self.colors["toolbar"],
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            bd=0,
        )
        preview_bar.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        preview_bar.columnconfigure(0, weight=1)
        preview_bar.columnconfigure(1, weight=1)

        preview_content = ttk.Frame(file_preview_shell.content, style="Surface.TFrame")
        preview_content.grid(row=1, column=0, sticky="nsew")
        preview_content.columnconfigure(0, weight=1)
        preview_content.rowconfigure(0, weight=1)

        self.preview_texts = {}
        for key, label in (
            ("folder", "Folder Preview"),
            ("appinfo", "appinfo.ini"),
            ("launcher", "launcher.ini"),
            ("installer", "installer.ini"),
        ):
            tab = ttk.Frame(preview_content, style="Surface.TFrame", padding=0)
            tab.columnconfigure(0, weight=1)
            tab.rowconfigure(0, weight=1)
            preview_shell_widget, text = self.create_preview_text(tab)
            preview_shell_widget.grid(row=0, column=0, sticky="nsew")
            self.preview_texts[key] = text
            self.create_preview_tab(preview_bar, key, label, tab)
        self.select_preview_tab("folder")

        sidebar_icon_shell = self.create_panel(
            preview_body,
            "Icon Preview",
            "Current generated icon set.",
            collapsible=False,
        )
        sidebar_icon_shell.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        sidebar_icon_group = sidebar_icon_shell.content
        sidebar_icon_group.columnconfigure(0, weight=1)
        sidebar_icon_frame = tk.Frame(
            sidebar_icon_group,
            bg=self.colors["surface_alt"],
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            bd=0,
            padx=10,
            pady=10,
        )
        sidebar_icon_frame.grid(row=0, column=0, sticky="ew")
        sidebar_icon_sizes = tk.Frame(sidebar_icon_frame, bg=self.colors["surface_alt"])
        sidebar_icon_sizes.pack(anchor="center")
        self.sidebar_icon_preview_labels = []
        for row, (size_label, display_size) in enumerate(ICON_PREVIEW_DISPLAY_SIZES):
            item_frame = tk.Frame(sidebar_icon_sizes, bg=self.colors["surface_alt"])
            item_frame.grid(row=row, column=0, pady=(0 if row == 0 else 8, 0), sticky="s")
            holder = tk.Frame(item_frame, bg=self.colors["surface_alt"], width=max(display_size + 10, 40), height=max(display_size + 10, 40))
            holder.pack()
            holder.pack_propagate(False)
            icon_label = tk.Label(holder, bg=self.colors["surface_alt"])
            icon_label.pack(expand=True)
            self.sidebar_icon_preview_labels.append(icon_label)
            ttk.Label(item_frame, text=f"{size_label}px", style="PanelNote.TLabel").pack(pady=(4, 0))
        self.sidebar_icon_preview_caption = ttk.Label(
            sidebar_icon_group,
            text="Waiting for icon source...",
            style="PanelNote.TLabel",
            wraplength=320,
        )
        self.sidebar_icon_preview_caption.grid(row=1, column=0, sticky="w", pady=(8, 0))

        sidebar_splash_shell = self.create_panel(
            preview_body,
            "Splash Preview",
            "Current default splash asset.",
            collapsible=False,
        )
        sidebar_splash_shell.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        sidebar_splash_group = sidebar_splash_shell.content
        sidebar_splash_group.columnconfigure(0, weight=1)
        sidebar_splash_frame = tk.Frame(
            sidebar_splash_group,
            bg=self.colors["surface_alt"],
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            bd=0,
            padx=10,
            pady=10,
        )
        sidebar_splash_frame.grid(row=0, column=0, sticky="ew")
        sidebar_splash_frame.columnconfigure(0, weight=1)
        self.sidebar_splash_preview_label = ttk.Label(sidebar_splash_frame, style="Surface.TLabel")
        self.sidebar_splash_preview_label.grid(row=0, column=0, sticky="n")
        self.sidebar_splash_preview_caption = ttk.Label(
            sidebar_splash_group,
            text="Waiting for splash preview...",
            style="PanelNote.TLabel",
            wraplength=320,
        )
        self.sidebar_splash_preview_caption.grid(row=1, column=0, sticky="w", pady=(8, 0))

        actions = ttk.Frame(body, style="TFrame")
        actions.grid(row=1, column=0, sticky="e", pady=(10, 0))
        actions.columnconfigure(0, weight=0)
        actions.columnconfigure(1, weight=0)
        self.validate_button = self.make_button(actions, text="Validate", command=self.validate_current_project)
        self.validate_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.create_button = self.make_button(actions, text="Create Project + EXE", style="Accent.TButton", command=self.create_project)
        self.create_button.grid(row=0, column=1, sticky="ew")

    def create_preview_text(self, parent):
        shell = tk.Frame(
            parent,
            bg=self.colors["field_border"],
            highlightthickness=1,
            highlightbackground=self.colors["field_border"],
            bd=0,
        )
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(0, weight=1)
        text = tk.Text(
            shell,
            height=12,
            wrap="none",
            bg=self.colors["field_soft"],
            fg=self.colors["text"],
            insertbackground=self.colors["text"],
            selectbackground=self.colors["accent_soft"],
            selectforeground=self.colors["text"],
            relief="flat",
            bd=0,
            highlightthickness=0,
            font=("Consolas", 10),
        )
        yscrollbar = self.create_scrollbar(shell, orientation="vertical", command=text.yview)
        text.configure(yscrollcommand=yscrollbar.set)
        text.grid(row=0, column=0, sticky="nsew")
        yscrollbar.grid(row=0, column=1, sticky="ns")
        text.tag_configure("folder", foreground="#2f5f98", font=("Consolas", 10, "bold"))
        text.tag_configure("important", foreground=self.colors["accent_hover"], font=("Consolas", 10, "bold"))
        text.tag_configure("optional", foreground=self.colors["muted"], font=("Consolas", 10, "italic"))
        text.tag_configure("comment", foreground=self.colors["muted"], font=("Consolas", 10))
        text.tag_configure("plain", foreground=self.colors["text"], font=("Consolas", 10))
        text.configure(state="disabled")
        return shell, text

    def create_main_tab(self, parent, key, text, frame):
        frame.grid(row=0, column=0, sticky="nsew")
        frame.grid_remove()

        button = tk.Label(
            parent,
            text=text,
            bg=self.colors["toolbar"],
            fg=self.colors["muted"],
            padx=16,
            pady=9,
            cursor="hand2",
            font=("Segoe UI Semibold", 9),
            bd=0,
            highlightthickness=1,
            highlightbackground=self.colors["toolbar"],
            highlightcolor=self.colors["toolbar"],
            takefocus=0,
        )
        button.pack(side="left", padx=(0, 6))
        button.bind("<Button-1>", lambda _event, selected=key: self.select_main_tab(selected), add="+")
        button.bind("<Enter>", lambda _event, hovered=key: self.set_main_tab_hover(hovered), add="+")
        button.bind("<Leave>", lambda _event: self.set_main_tab_hover(None), add="+")

        self.main_tab_buttons[key] = button
        self.main_tab_frames[key] = frame
        return button

    def set_main_tab_hover(self, key):
        self.hover_main_tab = key
        self.refresh_main_tabs()

    def refresh_main_tabs(self):
        for key, button in self.main_tab_buttons.items():
            selected = key == self.current_main_tab
            hovered = key == self.hover_main_tab
            if selected:
                background = self.colors["accent_soft"]
                foreground = self.colors["accent_hover"]
                border = self.colors["accent_line"]
            elif hovered:
                background = self.colors["surface"]
                foreground = self.colors["text"]
                border = self.colors["border"]
            else:
                background = self.colors["toolbar"]
                foreground = self.colors["muted"]
                border = self.colors["toolbar"]
            button.configure(bg=background, fg=foreground, highlightbackground=border, highlightcolor=border)

    def select_main_tab(self, key):
        if key not in self.main_tab_frames:
            return
        for frame in self.main_tab_frames.values():
            frame.grid_remove()
        self.main_tab_frames[key].grid()
        self.current_main_tab = key
        self.refresh_main_tabs()

    def create_preview_tab(self, parent, key, text, frame):
        frame.grid(row=0, column=0, sticky="nsew")
        frame.grid_remove()

        button_index = len(self.preview_tab_buttons)
        row = button_index // 2
        column = button_index % 2
        button = tk.Label(
            parent,
            text=text,
            bg=self.colors["toolbar"],
            fg=self.colors["muted"],
            padx=14,
            pady=8,
            cursor="hand2",
            font=("Segoe UI Semibold", 9),
            bd=0,
            highlightthickness=1,
            highlightbackground=self.colors["toolbar"],
            highlightcolor=self.colors["toolbar"],
            takefocus=0,
        )
        button.grid(row=row, column=column, sticky="ew", padx=(8, 8), pady=(8, 0 if row == 0 else 8))
        button.bind("<Button-1>", lambda _event, selected=key: self.select_preview_tab(selected), add="+")
        button.bind("<Enter>", lambda _event, hovered=key: self.set_preview_tab_hover(hovered), add="+")
        button.bind("<Leave>", lambda _event: self.set_preview_tab_hover(None), add="+")

        self.preview_tab_buttons[key] = button
        self.preview_tab_frames[key] = frame
        return button

    def set_preview_tab_hover(self, key):
        self.hover_preview_tab = key
        self.refresh_preview_tabs()

    def refresh_preview_tabs(self):
        for key, button in self.preview_tab_buttons.items():
            selected = key == self.current_preview_tab
            hovered = key == self.hover_preview_tab
            if selected:
                background = self.colors["accent_soft"]
                foreground = self.colors["accent_hover"]
                border = self.colors["accent_line"]
            elif hovered:
                background = self.colors["surface"]
                foreground = self.colors["text"]
                border = self.colors["border"]
            else:
                background = self.colors["toolbar"]
                foreground = self.colors["muted"]
                border = self.colors["toolbar"]
            button.configure(bg=background, fg=foreground, highlightbackground=border, highlightcolor=border)

    def select_preview_tab(self, key):
        if key not in self.preview_tab_frames:
            return
        for frame in self.preview_tab_frames.values():
            frame.grid_remove()
        self.preview_tab_frames[key].grid()
        self.current_preview_tab = key
        self.refresh_preview_tabs()

    def create_scrollable_tab(self, notebook):
        outer = ttk.Frame(notebook, style="Surface.TFrame")
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(0, weight=1)

        canvas = tk.Canvas(
            outer,
            bg=self.colors["surface"],
            highlightthickness=0,
            bd=0,
            yscrollincrement=24,
        )
        scrollbar = self.create_scrollbar(outer, orientation="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        content = ttk.Frame(canvas, style="Surface.TFrame", padding=16)
        content.columnconfigure(0, weight=1)
        window_id = canvas.create_window((0, 0), window=content, anchor="nw")

        # Keep the embedded frame sized to the canvas width so card layouts
        # reflow naturally while the canvas handles vertical scrolling.
        content.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(window_id, width=event.width))

        canvas._scroll_canvas = canvas
        content._scroll_canvas = canvas
        canvas.bind("<Enter>", lambda _event, target=canvas: self.set_active_scroll_canvas(target), add="+")
        content.bind("<Enter>", lambda _event, target=canvas: self.set_active_scroll_canvas(target), add="+")
        canvas.bind("<Leave>", lambda _event, target=canvas: self.clear_active_scroll_canvas(target), add="+")
        content.bind("<Leave>", lambda _event, target=canvas: self.clear_active_scroll_canvas(target), add="+")

        return outer, content

    def create_panel(self, parent, title, note="", collapsible=True, expanded=True):
        outer = tk.Frame(parent, bg=self.colors["border"], highlightthickness=0, bd=0)
        inner = tk.Frame(outer, bg=self.colors["surface"], highlightthickness=0, bd=0)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        inner.columnconfigure(0, weight=1)
        inner.rowconfigure(1, weight=1)

        header = tk.Frame(inner, bg=self.colors["card_header"], highlightthickness=0, bd=0, padx=16, pady=12)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        title_block = tk.Frame(header, bg=self.colors["card_header"], highlightthickness=0, bd=0)
        title_block.grid(row=0, column=0, sticky="w")

        title_label = tk.Label(
            title_block,
            text=title,
            bg=self.colors["card_header"],
            fg=self.colors["text"],
            font=("Segoe UI Semibold", 10),
            anchor="w",
        )
        title_label.pack(anchor="w")

        note_label = None
        if note:
            note_label = tk.Label(
                title_block,
                text=note,
                bg=self.colors["card_header"],
                fg=self.colors["muted"],
                font=("Segoe UI", 8),
                justify="left",
                anchor="w",
                wraplength=760,
            )
            note_label.pack(anchor="w", pady=(3, 0))

        toggle_label = None
        if collapsible:
            toggle_label = tk.Label(
                header,
                text="▾" if expanded else "▸",
                bg=self.colors["card_header"],
                fg=self.colors["muted"],
                font=("Segoe UI Semibold", 11),
                padx=4,
                cursor="hand2",
            )
            toggle_label.grid(row=0, column=1, sticky="e", padx=(12, 0))

        body_shell = tk.Frame(inner, bg=self.colors["surface"], highlightthickness=0, bd=0, padx=16, pady=16)
        body_shell.grid(row=1, column=0, sticky="nsew")
        body_shell.columnconfigure(0, weight=1)
        body_shell.rowconfigure(0, weight=1)

        content = ttk.Frame(body_shell, style="PanelBody.TFrame")
        content.grid(row=0, column=0, sticky="nsew")
        content.columnconfigure(0, weight=1)

        outer.content = content
        outer.inner = inner
        outer.header = header
        outer.title_label = title_label
        outer.note_label = note_label
        outer.toggle_label = toggle_label
        outer.body_shell = body_shell
        outer.collapsible = collapsible
        outer.expanded = expanded
        self.panel_cards.append(outer)

        if collapsible:
            # Treat the whole header like a clickable card title bar, similar to
            # a collapsible section in a web settings screen.
            self.set_panel_expanded(outer, expanded)
            for widget in (header, title_block, title_label, note_label, toggle_label):
                if widget is not None:
                    widget.bind("<Button-1>", lambda _event, panel=outer: self.toggle_panel(panel), add="+")
                    widget.bind("<Enter>", lambda _event, panel=outer: self.set_panel_hover(panel, True), add="+")
                    widget.bind("<Leave>", lambda _event, panel=outer: self.set_panel_hover(panel, False), add="+")
        else:
            self.set_panel_hover(outer, False)

        return outer

    def set_panel_hover(self, panel, hovered):
        base = self.colors["card_header_active"] if hovered and panel.collapsible else self.colors["card_header"]
        panel.header.configure(bg=base)
        panel.title_label.configure(bg=base)
        if panel.note_label is not None:
            panel.note_label.configure(bg=base)
        if panel.toggle_label is not None:
            panel.toggle_label.configure(bg=base)
        title_parent = panel.title_label.master
        title_parent.configure(bg=base)

    def set_panel_expanded(self, panel, expanded):
        panel.expanded = expanded
        if expanded:
            panel.body_shell.grid()
        else:
            panel.body_shell.grid_remove()
        if panel.toggle_label is not None:
            panel.toggle_label.configure(text="▾" if expanded else "▸")

    def toggle_panel(self, panel):
        if not panel.collapsible:
            return
        self.set_panel_expanded(panel, not panel.expanded)

    def create_text_editor(self, parent, height=4, background="#ffffff"):
        field_background = self.colors["field_soft"] if background == "#ffffff" else background
        return ScrolledText(
            parent,
            self.colors,
            height=max(3, height),
            background=field_background,
            font=("Consolas", 11),
            wrap="none",
        )

    def create_appinfo_editor(self, parent):
        parent.columnconfigure(0, weight=1)
        row = 0
        categories = (
            "Choose Category...",
            "Accessibility",
            "Development",
            "Education",
            "Games",
            "Graphics & Pictures",
            "Internet",
            "Music & Video",
            "Office",
            "Security",
            "Utilities",
        )
        languages = (
            "Multilingual",
            "English",
            "SimpChinese",
            "TradChinese",
            "Japanese",
            "Korean",
            "German",
            "Spanish",
            "French",
            "Italian",
        )

        def add_group(title, pair_count, note=""):
            nonlocal row
            shell = self.create_panel(parent, title, note)
            shell.grid(row=row, column=0, sticky="ew", pady=(0, 12))
            frame = shell.content
            for index in range(pair_count * 2):
                frame.columnconfigure(index, weight=1)
            row += 1
            return frame

        def add_entry(frame, field_row, column, label, key, value_span=1):
            field = ttk.Frame(frame, style="PanelBody.TFrame")
            field.grid(
                row=field_row,
                column=column,
                columnspan=value_span + 1,
                sticky="ew",
                padx=(0, 10),
                pady=(0, 10),
            )
            field.columnconfigure(0, weight=1)
            ttk.Label(field, text=label, style="Surface.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 6))
            self.create_entry(field, textvariable=self.vars[key]).grid(row=1, column=0, sticky="ew")

        def add_combo(frame, field_row, column, label, key, values, value_span=1):
            field = ttk.Frame(frame, style="PanelBody.TFrame")
            field.grid(
                row=field_row,
                column=column,
                columnspan=value_span + 1,
                sticky="ew",
                padx=(0, 10),
                pady=(0, 10),
            )
            field.columnconfigure(0, weight=1)
            ttk.Label(field, text=label, style="Surface.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 6))
            self.create_combobox(field, textvariable=self.vars[key], values=values).grid(row=1, column=0, sticky="ew")

        details = add_group("Details", 4, "Core app metadata used by PortableApps.com Format.")
        add_entry(details, 0, 0, "App Name", "app_name")
        add_entry(details, 0, 2, "Package ID", "package_name")
        add_entry(details, 0, 4, "Publisher", "publisher")
        add_entry(details, 0, 6, "Trademarks", "trademarks")
        add_combo(details, 1, 0, "Category", "category", categories, value_span=3)
        add_combo(details, 1, 4, "Language", "language", languages, value_span=3)
        add_entry(details, 2, 0, "Description", "description", value_span=7)
        add_entry(details, 3, 0, "Homepage", "homepage", value_span=7)
        add_entry(details, 4, 0, "Donate", "donate", value_span=7)
        add_entry(details, 5, 0, "Install Type", "install_type", value_span=7)

        version = add_group("Version", 2, "Version values shown in the app info and PortableApps metadata.")
        add_entry(version, 0, 0, "Package Version", "version")
        add_entry(version, 0, 2, "Display Version", "display_version")

        special_paths = add_group("SpecialPaths", 1, "Optional special folders used by the portable package.")
        add_entry(special_paths, 0, 0, "Plugins", "special_plugins")

        dependencies = add_group("Dependencies", 3, "Declare runtime requirements and other PortableApps dependencies.")
        add_combo(dependencies, 0, 0, "Uses Ghostscript", "dependency_uses_ghostscript", ("no", "yes", "optional"))
        add_combo(dependencies, 0, 2, "Uses Java", "dependency_uses_java", ("no", "yes", "optional"))
        add_entry(dependencies, 0, 4, ".NET Version", "dependency_uses_dotnet_version")
        add_combo(dependencies, 1, 0, "Requires 64-bit OS", "dependency_requires_64bit_os", ("no", "yes"))
        add_combo(dependencies, 1, 2, "Requires Admin", "dependency_requires_admin", ("no", "yes"))
        add_entry(dependencies, 2, 0, "Requires Portable App", "dependency_requires_portable_app", value_span=5)

        license_group = add_group("License", 2, "Flags saved into the [License] section of appinfo.ini.")
        ttk.Checkbutton(license_group, text="Shareable", variable=self.vars["license_shareable"]).grid(row=0, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Checkbutton(license_group, text="Open Source", variable=self.vars["license_open_source"]).grid(row=0, column=2, columnspan=2, sticky="w", pady=4)
        ttk.Checkbutton(license_group, text="Freeware", variable=self.vars["license_freeware"]).grid(row=1, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Checkbutton(license_group, text="Commercial Use", variable=self.vars["license_commercial_use"]).grid(row=1, column=2, columnspan=2, sticky="w", pady=4)
        add_entry(license_group, 2, 0, "EULA Version", "license_eula_version", value_span=3)

        control_group_shell = self.create_panel(parent, "Control", "Direct editor for the [Control] section in appinfo.ini.")
        control_group_shell.grid(row=row, column=0, sticky="ew", pady=(0, 12))
        control_group = control_group_shell.content
        control_group.columnconfigure(0, weight=1)
        self.add_multiline_control_editor(control_group)
        row += 1

        associations_group_shell = self.create_panel(parent, "Associations", "Edit file associations, protocols, SendTo, and shell behavior.")
        associations_group_shell.grid(row=row, column=0, sticky="ew", pady=(0, 12))
        associations_group = associations_group_shell.content
        associations_group.columnconfigure(0, weight=1)
        self.add_multiline_associations_editor(associations_group)
        row += 1

        file_type_icons_shell = self.create_panel(parent, "FileTypeIcons", "One key=value mapping per line for file type icon overrides.")
        file_type_icons_shell.grid(row=row, column=0, sticky="ew", pady=(0, 12))
        file_type_icons_group = file_type_icons_shell.content
        file_type_icons_group.columnconfigure(0, weight=1)
        self.add_multiline_row(file_type_icons_group, 0, "Entries", "file_type_icons", height=5)

    def create_registry_editor(self, parent):
        parent.columnconfigure(0, weight=1)

        registry_shell = self.create_panel(
            parent,
            "Registry Settings",
            "Controls the [Activate] registry flag and optional launcher registry sections.",
        )
        registry_shell.grid(row=0, column=0, sticky="ew")
        registry_group = registry_shell.content
        registry_group.columnconfigure(0, weight=1)
        registry_group.columnconfigure(1, weight=0)

        ttk.Checkbutton(
            registry_group,
            text="Enable registry handling in [Activate]",
            variable=self.vars["registry_enabled"],
            command=self.update_registry_controls,
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))
        self.import_registry_button = self.make_button(
            registry_group,
            text="Import Saved Registry (.reg)",
            command=self.import_registry_file,
        )
        self.import_registry_button.grid(row=0, column=1, sticky="e", padx=(12, 0), pady=(0, 10))
        self.vars["registry_enabled"].trace_add("write", lambda *_args: self.update_registry_controls())

        keys_shell = self.create_panel(
            parent,
            "RegistryKeys",
            "One key=value mapping per line for [RegistryKeys].",
        )
        keys_shell.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        keys_group = keys_shell.content
        keys_group.columnconfigure(0, weight=1)
        next_row = self.add_multiline_row(keys_group, 0, "Entries", "registry_keys", height=6)
        ttk.Label(
            keys_group,
            text=r"Sample: appname_portable=HKCU\Software\Publisher\AppName",
            style="PanelNote.TLabel",
            wraplength=720,
        ).grid(row=next_row, column=0, sticky="w", pady=(0, 4))

        cleanup_empty_shell = self.create_panel(
            parent,
            "RegistryCleanupIfEmpty",
            "One key=value mapping per line for [RegistryCleanupIfEmpty].",
        )
        cleanup_empty_shell.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        cleanup_empty_group = cleanup_empty_shell.content
        cleanup_empty_group.columnconfigure(0, weight=1)
        self.add_multiline_row(cleanup_empty_group, 0, "Entries", "registry_cleanup_if_empty", height=5)

        cleanup_force_shell = self.create_panel(
            parent,
            "RegistryCleanupForce",
            "One key=value mapping per line for [RegistryCleanupForce].",
        )
        cleanup_force_shell.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        cleanup_force_group = cleanup_force_shell.content
        cleanup_force_group.columnconfigure(0, weight=1)
        self.add_multiline_row(cleanup_force_group, 0, "Entries", "registry_cleanup_force", height=5)

    def create_launcher_editor(self, parent):
        parent.columnconfigure(0, weight=1)

        os_values = ("", "2000", "XP", "2003", "Vista", "2008", "7", "2008 R2")
        run_as_admin_values = ("", "force", "try", "compile-force")
        refresh_shell_values = ("", "before", "after", "both")
        java_values = ("", "find", "require")
        unc_values = ("", "yes", "warn", "no")

        launch_shell = self.create_panel(
            parent,
            "Launch",
            "Core launch settings for AppNamePortable.ini based on the official PortableApps.com Launcher format.",
        )
        launch_shell.grid(row=0, column=0, sticky="ew")
        launch_group = launch_shell.content
        for column in range(3):
            launch_group.columnconfigure(column, weight=1)

        self.launch_program_executable_var = tk.StringVar()
        program_field = ttk.Frame(launch_group, style="PanelBody.TFrame")
        program_field.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        program_field.columnconfigure(0, weight=1)
        ttk.Label(program_field, text="ProgramExecutable", style="Surface.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 6))
        self.create_entry(program_field, textvariable=self.launch_program_executable_var, state="readonly").grid(row=1, column=0, sticky="ew")

        self.add_stacked_entry(launch_group, 1, 0, "Arguments", "command_line", columnspan=3)
        self.add_stacked_entry(launch_group, 2, 0, "Working Dir", "working_directory")
        self.add_stacked_entry(launch_group, 2, 1, "Close EXE", "close_exe")
        self.add_stacked_combo(launch_group, 3, 0, "Min OS", "min_os", os_values, width=16)
        self.add_stacked_combo(launch_group, 3, 1, "Max OS", "max_os", os_values, width=16)
        self.add_stacked_combo(launch_group, 3, 2, "Run As Admin", "run_as_admin", run_as_admin_values, width=16)
        self.add_stacked_combo(launch_group, 4, 0, "Refresh Shell Icons", "refresh_shell_icons", refresh_shell_values, width=16)
        self.add_stacked_combo(launch_group, 4, 1, "Supports UNC", "supports_unc", unc_values, width=16)

        launch_checks = ttk.Frame(launch_group, style="PanelBody.TFrame")
        launch_checks.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(4, 0))
        launch_checks.columnconfigure(0, weight=1)
        launch_checks.columnconfigure(1, weight=1)
        ttk.Checkbutton(launch_checks, text="Copy selected app folder into App folder", variable=self.vars["copy_app_files"]).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(launch_checks, text="Wait for program before cleanup", variable=self.vars["wait_for_program"]).grid(row=0, column=1, sticky="w", padx=(16, 0))
        ttk.Checkbutton(launch_checks, text="Wait for other instances", variable=self.vars["wait_for_other_instances"]).grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Checkbutton(launch_checks, text="Hide command line window", variable=self.vars["hide_command_line_window"]).grid(row=1, column=1, sticky="w", padx=(16, 0), pady=(6, 0))
        ttk.Checkbutton(launch_checks, text="No spaces in path", variable=self.vars["no_spaces_in_path"]).grid(row=2, column=0, sticky="w", pady=(6, 0))

        activate_shell = self.create_panel(
            parent,
            "Activate",
            "Optional PAL features from the [Activate] section. Registry settings are edited in the Registry tab.",
        )
        activate_shell.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        activate_group = activate_shell.content
        activate_group.columnconfigure(0, weight=1)
        self.add_stacked_combo(activate_group, 0, 0, "Java", "activate_java", java_values, width=18)
        ttk.Checkbutton(activate_group, text="Enable XML support", variable=self.vars["activate_xml"]).grid(row=1, column=0, sticky="w", pady=(2, 0))

        live_mode_shell = self.create_panel(
            parent,
            "LiveMode",
            "Optional live-mode copy settings written only when enabled.",
        )
        live_mode_shell.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        live_mode_group = live_mode_shell.content
        ttk.Checkbutton(live_mode_group, text="Copy app to temporary writable location", variable=self.vars["live_mode_copy_app"]).pack(anchor="w")
        ttk.Checkbutton(live_mode_group, text="Copy data while running in live mode", variable=self.vars["live_mode_copy_data"]).pack(anchor="w", pady=(6, 0))

        files_move_shell = self.create_panel(
            parent,
            "FilesMove",
            "One entry per line for [FilesMove] in the form relative-file=target-directory.",
        )
        files_move_shell.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        files_move_group = files_move_shell.content
        files_move_group.columnconfigure(0, weight=1)
        next_row = self.add_multiline_row(files_move_group, 0, "Entries", "files_move", height=6)
        ttk.Label(
            files_move_group,
            text=r"Sample: settings\config.ini=%PAL:AppDir%\YourApp",
            style="PanelNote.TLabel",
            wraplength=720,
        ).grid(row=next_row, column=0, sticky="w", pady=(0, 4))

        directories_move_shell = self.create_panel(
            parent,
            "DirectoriesMove",
            "One entry per line for [DirectoriesMove] in the form relative-directory=target-location.",
        )
        directories_move_shell.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        directories_move_group = directories_move_shell.content
        directories_move_group.columnconfigure(0, weight=1)
        next_row = self.add_multiline_row(directories_move_group, 0, "Entries", "directories_move", height=6)
        ttk.Label(
            directories_move_group,
            text=r"Sample: settings=%APPDATA%\YourApp",
            style="PanelNote.TLabel",
            wraplength=720,
        ).grid(row=next_row, column=0, sticky="w", pady=(0, 4))

    def create_installer_editor(self, parent):
        parent.columnconfigure(0, weight=1)

        check_running_shell = self.create_panel(
            parent,
            "CheckRunning",
            "Optional process checks used by the PortableApps.com Installer during upgrades.",
        )
        check_running_shell.grid(row=0, column=0, sticky="ew")
        check_running_group = check_running_shell.content
        check_running_group.columnconfigure(0, weight=1)
        check_running_group.columnconfigure(1, weight=1)
        self.add_stacked_entry(check_running_group, 0, 0, "CloseEXE", "installer_close_exe", width=28)
        self.add_stacked_entry(check_running_group, 0, 1, "CloseName", "installer_close_name", width=28)

        source_shell = self.create_panel(
            parent,
            "Source",
            "Include the PortableApps.com Installer source when packaging the app.",
        )
        source_shell.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        source_group = source_shell.content
        ttk.Checkbutton(source_group, text="Include installer source", variable=self.vars["include_installer_source"]).pack(anchor="w")

        main_dirs_shell = self.create_panel(
            parent,
            "MainDirectories",
            "Override the default upgrade behavior for App, Data, and Other directories.",
        )
        main_dirs_shell.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        main_dirs_group = main_dirs_shell.content
        for column in range(3):
            main_dirs_group.columnconfigure(column, weight=1)
        ttk.Checkbutton(
            main_dirs_group,
            text="Remove App Directory",
            variable=self.vars["remove_app_directory"],
        ).grid(row=0, column=0, sticky="w", pady=4, padx=(0, 12))
        ttk.Checkbutton(
            main_dirs_group,
            text="Remove Data Directory",
            variable=self.vars["remove_data_directory"],
        ).grid(row=0, column=1, sticky="w", pady=4, padx=(0, 12))
        ttk.Checkbutton(
            main_dirs_group,
            text="Remove Other Directory",
            variable=self.vars["remove_other_directory"],
        ).grid(row=0, column=2, sticky="w", pady=4)

        optional_shell = self.create_panel(
            parent,
            "OptionalComponents",
            "Configure the optional installer section, typically used for extra languages.",
        )
        optional_shell.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        optional_group = optional_shell.content
        for column in range(3):
            optional_group.columnconfigure(column, weight=1)
        ttk.Checkbutton(
            optional_group,
            text="Enable optional components section",
            variable=self.vars["optional_components_enabled"],
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))
        self.add_stacked_entry(optional_group, 1, 0, "Main Title", "main_section_title")
        self.add_stacked_entry(optional_group, 1, 1, "Main Description", "main_section_description", columnspan=2)
        self.add_stacked_entry(optional_group, 2, 0, "Optional Title", "optional_section_title")
        self.add_stacked_entry(optional_group, 2, 1, "Optional Description", "optional_section_description", columnspan=2)
        self.add_stacked_entry(optional_group, 3, 0, "Selected InstallType", "optional_section_selected_install_type")
        self.add_stacked_entry(optional_group, 3, 1, "Not Selected InstallType", "optional_section_not_selected_install_type")
        self.add_stacked_combo(optional_group, 3, 2, "Preselect If Non-English", "optional_section_preselected", ("", "true", "false"))

        languages_shell = self.create_panel(
            parent,
            "Languages",
            "One key=value line per installer language, such as ENGLISH=true.",
        )
        languages_shell.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        languages_group = languages_shell.content
        languages_group.columnconfigure(0, weight=1)
        next_row = self.add_multiline_row(languages_group, 0, "Entries", "installer_languages", height=5)
        ttk.Label(
            languages_group,
            text="Sample: ENGLISH=true",
            style="PanelNote.TLabel",
            wraplength=720,
        ).grid(row=next_row, column=0, sticky="w", pady=(0, 4))

        preserve_dirs_shell = self.create_panel(
            parent,
            "DirectoriesToPreserve",
            "One key=value line per preserved directory, such as PreserveDirectory1=App\\YourApp\\plugins.",
        )
        preserve_dirs_shell.grid(row=5, column=0, sticky="ew", pady=(12, 0))
        preserve_dirs_group = preserve_dirs_shell.content
        preserve_dirs_group.columnconfigure(0, weight=1)
        next_row = self.add_multiline_row(preserve_dirs_group, 0, "Entries", "preserve_directories", height=4)
        ttk.Label(
            preserve_dirs_group,
            text=r"Sample: PreserveDirectory1=App\YourApp\plugins",
            style="PanelNote.TLabel",
            wraplength=720,
        ).grid(row=next_row, column=0, sticky="w", pady=(0, 4))

        remove_dirs_shell = self.create_panel(
            parent,
            "DirectoriesToRemove",
            "One key=value line per removed directory, such as RemoveDirectory1=App\\YourApp\\cache.",
        )
        remove_dirs_shell.grid(row=6, column=0, sticky="ew", pady=(12, 0))
        remove_dirs_group = remove_dirs_shell.content
        remove_dirs_group.columnconfigure(0, weight=1)
        next_row = self.add_multiline_row(remove_dirs_group, 0, "Entries", "remove_directories", height=4)
        ttk.Label(
            remove_dirs_group,
            text=r"Sample: RemoveDirectory1=App\YourApp\cache",
            style="PanelNote.TLabel",
            wraplength=720,
        ).grid(row=next_row, column=0, sticky="w", pady=(0, 4))

        preserve_files_shell = self.create_panel(
            parent,
            "FilesToPreserve",
            "One key=value line per preserved file, such as PreserveFile1=Data\\settings\\custom.ini.",
        )
        preserve_files_shell.grid(row=7, column=0, sticky="ew", pady=(12, 0))
        preserve_files_group = preserve_files_shell.content
        preserve_files_group.columnconfigure(0, weight=1)
        next_row = self.add_multiline_row(preserve_files_group, 0, "Entries", "preserve_files", height=4)
        ttk.Label(
            preserve_files_group,
            text=r"Sample: PreserveFile1=Data\settings\custom.ini",
            style="PanelNote.TLabel",
            wraplength=720,
        ).grid(row=next_row, column=0, sticky="w", pady=(0, 4))

        remove_files_shell = self.create_panel(
            parent,
            "FilesToRemove",
            "One key=value line per removed file, such as RemoveFile1=App\\YourApp\\*.lang.",
        )
        remove_files_shell.grid(row=8, column=0, sticky="ew", pady=(12, 0))
        remove_files_group = remove_files_shell.content
        remove_files_group.columnconfigure(0, weight=1)
        next_row = self.add_multiline_row(remove_files_group, 0, "Entries", "remove_files", height=4)
        ttk.Label(
            remove_files_group,
            text=r"Sample: RemoveFile1=App\YourApp\*.lang",
            style="PanelNote.TLabel",
            wraplength=720,
        ).grid(row=next_row, column=0, sticky="w", pady=(0, 4))

    def create_icon_editor(self, parent):
        parent.columnconfigure(0, weight=1)

        icon_shell = self.create_panel(
            parent,
            "Icon Settings",
            "Choose which embedded icon to extract, or override it with your own icon file.",
        )
        icon_shell.grid(row=0, column=0, sticky="ew")
        icon_group = icon_shell.content
        icon_group.columnconfigure(0, weight=0)
        icon_group.columnconfigure(1, weight=1)
        icon_group.columnconfigure(2, weight=0)
        self.add_stacked_entry(icon_group, 0, 0, "Icon Index", "icon_index", width=5)
        self.add_stacked_entry(icon_group, 0, 1, "Icon Override", "icon_source", width=42)
        browse_field = ttk.Frame(icon_group, style="PanelBody.TFrame")
        browse_field.grid(row=0, column=2, sticky="ew", pady=(0, 10))
        browse_field.columnconfigure(0, weight=1)
        ttk.Label(browse_field, text="Browse", style="Surface.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 6))
        self.make_button(browse_field, text="Choose Icon", command=self.choose_icon).grid(row=1, column=0, sticky="ew")

        preview_shell = self.create_panel(
            parent,
            "Icon Preview",
            "Shows the icon that will be used when the project is generated.",
        )
        preview_shell.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        preview_group = preview_shell.content
        preview_group.columnconfigure(0, weight=1)

        preview_frame = tk.Frame(
            preview_group,
            bg=self.colors["surface_alt"],
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            bd=0,
            padx=16,
            pady=16,
        )
        preview_frame.grid(row=0, column=0, sticky="w")
        preview_sizes_frame = tk.Frame(preview_frame, bg=self.colors["surface_alt"])
        preview_sizes_frame.pack()
        self.icon_preview_labels = []

        for column, (size_label, display_size) in enumerate(ICON_PREVIEW_DISPLAY_SIZES):
            item_frame = tk.Frame(preview_sizes_frame, bg=self.colors["surface_alt"])
            item_frame.grid(row=0, column=column, padx=(0 if column == 0 else 12, 0), sticky="s")

            holder = tk.Frame(
                item_frame,
                bg=self.colors["surface_alt"],
                width=104,
                height=104,
            )
            holder.pack()
            holder.pack_propagate(False)

            icon_label = tk.Label(
                holder,
                bg=self.colors["surface_alt"],
            )
            icon_label.pack(side="bottom")
            self.icon_preview_labels.append(icon_label)

            ttk.Label(
                item_frame,
                text=f"{size_label}px",
                style="PanelNote.TLabel",
            ).pack(pady=(8, 0))

        self.icon_preview_caption = ttk.Label(
            preview_group,
            text="Waiting for icon source...",
            style="PanelNote.TLabel",
            wraplength=520,
        )
        self.icon_preview_caption.grid(row=1, column=0, sticky="w", pady=(10, 0))

    def create_template_editor(self, parent):
        parent.columnconfigure(0, weight=1)

        assets_shell = self.create_panel(
            parent,
            "Splash Asset",
            "Open the assets folder or replace the bundled splash image used for every new project.",
        )
        assets_shell.grid(row=0, column=0, sticky="ew")
        assets_group = assets_shell.content
        assets_group.columnconfigure(1, weight=1)
        assets_toolbar = ttk.Frame(assets_group, style="PanelBody.TFrame")
        assets_toolbar.grid(row=0, column=0, columnspan=4, sticky="ew", pady=(0, 10))
        assets_toolbar.columnconfigure(0, weight=1)
        self.make_button(assets_toolbar, text="Open Assets Folder", command=self.open_template_folder).grid(row=0, column=0, sticky="w")

        for row_index, (relative_path, label, filetypes) in enumerate(TEMPLATE_ASSET_SPECS, start=1):
            display_row = 1 + ((row_index - 1) * 2)
            self.add_template_asset_row(assets_group, display_row, relative_path, label, filetypes)

        splash_shell = self.create_panel(
            parent,
            "Splash Preview",
            "Default preview source used to generate App\\AppInfo\\Launcher\\Splash.jpg.",
        )
        splash_shell.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        splash_group = splash_shell.content
        splash_group.columnconfigure(0, weight=1)

        splash_frame = tk.Frame(
            splash_group,
            bg=self.colors["surface_alt"],
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            bd=0,
            padx=16,
            pady=16,
        )
        splash_frame.grid(row=0, column=0, sticky="ew")
        splash_frame.columnconfigure(0, weight=1)

        self.template_splash_label = ttk.Label(splash_frame, style="Surface.TLabel")
        self.template_splash_label.grid(row=0, column=0, sticky="w")
        self.template_splash_caption = ttk.Label(splash_frame, style="PanelNote.TLabel", wraplength=640)
        self.template_splash_caption.grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.refresh_template_asset_views()

    def add_template_asset_row(self, parent, row, relative_path, label, filetypes):
        ttk.Label(parent, text=label, style="Surface.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 6), pady=(0, 6))
        path_var = tk.StringVar(value=str(asset_path(relative_path)))
        self.template_asset_path_vars[relative_path] = path_var
        field = ttk.Frame(parent, style="PanelBody.TFrame")
        field.grid(row=row + 1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        field.columnconfigure(0, weight=1)
        self.create_entry(field, textvariable=path_var, state="readonly").grid(row=0, column=0, sticky="ew")

        actions = ttk.Frame(parent, style="PanelBody.TFrame")
        actions.grid(row=row + 1, column=2, columnspan=2, sticky="ew", pady=(0, 10))
        actions.columnconfigure(0, weight=1)
        actions.columnconfigure(1, weight=1)
        self.make_button(
            actions,
            text="Open",
            command=lambda target=relative_path: self.open_template_asset(target),
        ).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.make_button(
            actions,
            text="Replace",
            command=lambda target=relative_path, chooser=filetypes: self.replace_template_asset(target, chooser),
        ).grid(row=0, column=1, sticky="ew")

    def open_template_folder(self):
        open_folder_in_explorer(asset_path(""))
        self.status_var.set(f"Opened {asset_path('')}")

    def open_template_asset(self, relative_path):
        open_folder_in_explorer(asset_path(relative_path))
        self.status_var.set(f"Opened {asset_path(relative_path)}")

    def replace_template_asset(self, relative_path, filetypes):
        current_path = asset_path(relative_path)
        selected = filedialog.askopenfilename(
            title=f"Replace {Path(relative_path).name}",
            filetypes=filetypes,
            initialdir=str(current_path.parent),
        )
        if not selected:
            return
        if current_path.suffix.lower() == ".png":
            with Image.open(selected) as source_image:
                source_image.convert("RGBA").save(current_path, format="PNG")
        else:
            shutil.copy2(selected, current_path)
        self.refresh_template_asset_views()
        self.status_var.set(f"Updated {current_path}")

    def refresh_template_asset_views(self):
        for relative_path, path_var in self.template_asset_path_vars.items():
            path_var.set(str(asset_path(relative_path)))

        if self.template_splash_label is None or self.template_splash_caption is None:
            self.update_sidebar_splash_preview()
            return

        splash_path = splash_asset_path()
        if splash_path.exists():
            try:
                with Image.open(splash_path) as splash_image:
                    preview = ImageOps.contain(splash_image.convert("RGBA"), (320, 180), Image.Resampling.LANCZOS)
                self.template_splash_image = ImageTk.PhotoImage(preview)
                self.template_splash_label.configure(image=self.template_splash_image, text="")
                self.template_splash_caption.configure(text=str(splash_path))
                self.update_sidebar_splash_preview()
                return
            except OSError:
                pass

        self.template_splash_image = None
        self.template_splash_label.configure(image="", text="Splash preview unavailable")
        self.template_splash_caption.configure(text=str(splash_path))
        self.update_sidebar_splash_preview()

    def create_help_content(self, parent, padx=0, pady=0):
        shell = self.create_panel(
            parent,
            "PortableApps Launcher Help",
            "Quick reference for common path variables and registry sections used in launcher.ini.",
        )
        shell.grid(row=0, column=0, sticky="nsew", padx=padx, pady=pady)
        content = shell.content
        content.columnconfigure(0, weight=1)
        content.columnconfigure(1, weight=1)
        content.rowconfigure(0, weight=0)
        content.rowconfigure(1, weight=1)

        drive_help, directory_help, partial_and_language_help = self.variable_help_content()
        registry_keys_help, cleanup_if_empty_help, cleanup_force_help, value_write_help, value_backup_delete_help = self.registry_help_content()
        variable_content = drive_help + "\n\n" + directory_help + "\n\n" + partial_and_language_help
        registry_content = (
            registry_keys_help
            + "\n\n"
            + cleanup_if_empty_help
            + "\n\n"
            + cleanup_force_help
            + "\n\n"
            + value_write_help
            + "\n\n"
            + value_backup_delete_help
        )
        toolbar = ttk.Frame(content, style="PanelBody.TFrame")
        toolbar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        toolbar.columnconfigure(0, weight=1)
        toolbar.columnconfigure(1, weight=1)
        self.make_button(toolbar, text="Additional Variable Help", command=self.open_variable_help_site).grid(
            row=1,
            column=0,
            sticky="w",
        )
        self.make_button(toolbar, text="Additional Registry Help", command=self.open_registry_help_site).grid(
            row=1,
            column=1,
            sticky="e",
        )
        blocks = ttk.Frame(content, style="PanelBody.TFrame")
        blocks.grid(row=1, column=0, columnspan=2, sticky="nsew")
        blocks.columnconfigure(0, weight=1)
        blocks.columnconfigure(1, weight=1)
        blocks.rowconfigure(0, weight=1)
        self.add_help_block(blocks, 0, 0, "Variables", variable_content, height=28)
        self.add_help_block(blocks, 0, 1, "Registry", registry_content, height=28)

    def add_inline_entry(self, parent, row, column, label, key, columnspan=1):
        ttk.Label(parent, text=label).grid(row=row, column=column, sticky="w", padx=(0, 6), pady=2)
        self.create_entry(parent, textvariable=self.vars[key]).grid(row=row, column=column + 1, columnspan=columnspan, sticky="ew", pady=2)

    def add_inline_combo(self, parent, row, column, label, key, values):
        ttk.Label(parent, text=label).grid(row=row, column=column, sticky="w", padx=(0, 6), pady=2)
        self.create_combobox(parent, textvariable=self.vars[key], values=values).grid(row=row, column=column + 1, sticky="ew", pady=2)

    def add_stacked_entry(self, parent, row, column, label, key, columnspan=1, width=None):
        field = ttk.Frame(parent, style="PanelBody.TFrame")
        field.grid(row=row, column=column, columnspan=columnspan, sticky="ew", padx=(0, 10), pady=(0, 10))
        field.columnconfigure(0, weight=1)
        ttk.Label(field, text=label, style="Surface.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 6))
        self.create_entry(field, textvariable=self.vars[key], width=width).grid(row=1, column=0, sticky="ew")
        return field

    def add_stacked_combo(self, parent, row, column, label, key, values, columnspan=1, width=None):
        field = ttk.Frame(parent, style="PanelBody.TFrame")
        field.grid(row=row, column=column, columnspan=columnspan, sticky="ew", padx=(0, 10), pady=(0, 10))
        field.columnconfigure(0, weight=1)
        ttk.Label(field, text=label, style="Surface.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 6))
        self.create_combobox(field, textvariable=self.vars[key], values=values, width=width).grid(row=1, column=0, sticky="ew")
        return field

    def add_multiline_control_editor(self, parent):
        text = self.add_bound_text(parent, 0, "control_text", height=4)
        self.control_text = text
        self.refresh_control_text()
        text.bind("<KeyRelease>", lambda _event: self.parse_control_text(), add="+")
        text.bind("<FocusOut>", lambda _event: self.parse_control_text(), add="+")

    def add_multiline_associations_editor(self, parent):
        text = self.add_bound_text(parent, 0, "associations_text", height=8)
        self.associations_text = text
        self.refresh_associations_text()
        text.bind("<KeyRelease>", lambda _event: self.parse_associations_text(), add="+")
        text.bind("<FocusOut>", lambda _event: self.parse_associations_text(), add="+")

    def add_bound_text(self, parent, row, key, height=4):
        text = self.create_text_editor(parent, height=height)
        text.grid(row=row, column=0, sticky="ew")
        return text

    def set_text_value(self, text, value):
        try:
            previous_state = str(text.cget("state"))
        except Exception:
            previous_state = str(getattr(text, "_textbox", text).cget("state"))
        if previous_state == "disabled":
            text.configure(state="normal")
        text.delete("1.0", "end")
        text.insert("1.0", value)
        if previous_state == "disabled":
            text.configure(state=previous_state)

    def refresh_control_text(self):
        if not hasattr(self, "control_text"):
            return
        project = self.current_project()
        lines = [
            f"Icons={project.control_icons.strip() or '1'}",
            f"Start={project.control_start.strip() or project.portable_name + '.exe'}",
        ]
        for key, value in (
            ("ExtractIcon", project.control_extract_icon),
            ("ExtractName", project.control_extract_name),
            ("BaseAppID", project.control_base_app_id),
            ("BaseAppID64", project.control_base_app_id_64),
            ("BaseAppIDARM64", project.control_base_app_id_arm64),
            ("ExitEXE", project.control_exit_exe),
            ("ExitParameters", project.control_exit_parameters),
        ):
            if value.strip():
                lines.append(f"{key}={value.strip()}")
        self.set_text_value(self.control_text, "\n".join(lines))

    def refresh_associations_text(self):
        if not hasattr(self, "associations_text"):
            return
        project = self.current_project()
        lines = [
            f"FileTypes={project.association_file_types}",
            f"FileTypeCommandLine={project.association_file_type_command_line}",
            f"FileTypeCommandLine-extension={project.association_file_type_command_line_extension}",
            f"Protocols={project.association_protocols}",
            f"ProtocolCommandLine={project.association_protocol_command_line}",
            f"ProtocolCommandLine-protocol={project.association_protocol_command_line_protocol}",
            f"SendTo={bool_to_ini(project.association_send_to)}",
            f"SendToCommandLine={project.association_send_to_command_line}",
            f"Shell={bool_to_ini(project.association_shell)}",
            f"ShellCommand={project.association_shell_command}",
        ]
        self.set_text_value(self.associations_text, "\n".join(lines))

    def parse_key_values(self, text):
        values = {}
        for line in text.get("1.0", "end-1c").splitlines():
            if "=" not in line:
                continue
            key, value = line.split("=", 1)
            values[key.strip().casefold()] = value.strip()
        return values

    def parse_control_text(self):
        values = self.parse_key_values(self.control_text)
        key_map = {
            "icons": "control_icons",
            "start": "control_start",
            "extracticon": "control_extract_icon",
            "extractname": "control_extract_name",
            "baseappid": "control_base_app_id",
            "baseappid64": "control_base_app_id_64",
            "baseappidarm64": "control_base_app_id_arm64",
            "exitexe": "control_exit_exe",
            "exitparameters": "control_exit_parameters",
        }
        for key, var_name in key_map.items():
            if key in values:
                self.vars[var_name].set(values[key])

    def parse_associations_text(self):
        values = self.parse_key_values(self.associations_text)
        key_map = {
            "filetypes": "association_file_types",
            "filetypecommandline": "association_file_type_command_line",
            "filetypecommandline-extension": "association_file_type_command_line_extension",
            "protocols": "association_protocols",
            "protocolcommandline": "association_protocol_command_line",
            "protocolcommandline-protocol": "association_protocol_command_line_protocol",
            "sendtocommandline": "association_send_to_command_line",
            "shellcommand": "association_shell_command",
        }
        for key, var_name in key_map.items():
            if key in values:
                self.vars[var_name].set(values[key])
        if "sendto" in values:
            self.vars["association_send_to"].set(values["sendto"].casefold() == "true")
        if "shell" in values:
            self.vars["association_shell"].set(values["shell"].casefold() == "true")

    def set_active_scroll_canvas(self, canvas):
        self.active_scroll_canvas = canvas

    def clear_active_scroll_canvas(self, canvas):
        if self.active_scroll_canvas is canvas:
            self.active_scroll_canvas = None

    def find_scroll_canvas(self, widget):
        current = widget
        while current is not None:
            canvas = getattr(current, "_scroll_canvas", None)
            if canvas is not None:
                return canvas
            parent_name = current.winfo_parent()
            if not parent_name:
                break
            try:
                current = current.nametowidget(parent_name)
            except KeyError:
                break
        return None

    def scroll_form(self, event):
        canvas = self.find_scroll_canvas(getattr(event, "widget", None)) or self.active_scroll_canvas
        if canvas is None:
            return
        if getattr(event, "num", None) == 4:
            canvas.yview_scroll(-1, "units")
        elif getattr(event, "num", None) == 5:
            canvas.yview_scroll(1, "units")
        else:
            delta = int(-1 * (event.delta / 120))
            if delta == 0:
                return
            canvas.yview_scroll(delta, "units")
        return "break"

    def handle_notebook_wheel(self, event):
        self.scroll_form(event)
        return "break"

    def bind_notebook_wheel(self, notebook):
        notebook.bind("<MouseWheel>", self.handle_notebook_wheel, add="+")
        notebook.bind("<Button-4>", self.handle_notebook_wheel, add="+")
        notebook.bind("<Button-5>", self.handle_notebook_wheel, add="+")

    def add_entry_row(self, parent, row, label, key):
        ttk.Label(parent, text=label, style="Surface.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 4), pady=5)
        self.create_entry(parent, textvariable=self.vars[key]).grid(row=row, column=1, columnspan=2, sticky="ew", padx=(0, 6), pady=5)
        return row + 1

    def add_card(self, parent, row, title):
        frame = ttk.LabelFrame(parent, text=title, padding=8)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        return frame

    def configure_card_columns(self, frame, count):
        for column in range(count):
            frame.columnconfigure(column, weight=1 if column % 2 else 0)

    def add_card_entry(self, parent, row, column, label, key, value_span=1):
        ttk.Label(parent, text=label, style="Surface.TLabel").grid(row=row, column=column, sticky="w", padx=(0, 4), pady=4)
        self.create_entry(parent, textvariable=self.vars[key]).grid(row=row, column=column + 1, columnspan=value_span, sticky="ew", padx=(0, 6), pady=4)

    def add_card_combo(self, parent, row, column, label, key, values, value_span=1):
        ttk.Label(parent, text=label, style="Surface.TLabel").grid(row=row, column=column, sticky="w", padx=(0, 4), pady=4)
        self.create_combobox(parent, textvariable=self.vars[key], values=values).grid(row=row, column=column + 1, columnspan=value_span, sticky="ew", padx=(0, 6), pady=4)

    def add_section_label(self, parent, row, label):
        ttk.Label(parent, text=label, style="Surface.TLabel", font=("Segoe UI Semibold", 10)).grid(
            row=row,
            column=0,
            columnspan=3,
            sticky="w",
            pady=(14, 4),
        )
        return row + 1

    def add_multiline_row(self, parent, row, label, key, height=4):
        ttk.Label(parent, text=label, style="Surface.TLabel").grid(row=row, column=0, sticky="w", pady=(0, 6))
        text = self.create_text_editor(parent, height=height)
        text.grid(row=row + 1, column=0, sticky="ew", pady=(0, 4))
        text.insert("1.0", self.vars[key].get())
        self.bound_text_widgets[key] = text

        def sync_var(_event=None):
            self.vars[key].set(text.get("1.0", "end-1c"))

        def sync_text(*_args):
            current = text.get("1.0", "end-1c")
            updated = self.vars[key].get()
            if current != updated:
                self.set_text_value(text, updated)

        text.bind("<KeyRelease>", sync_var)
        text.bind("<FocusOut>", sync_var)
        self.vars[key].trace_add("write", sync_text)
        return row + 2

    def add_help_block(self, parent, row, column, title, content, height=11):
        block = tk.Frame(parent, bg=self.colors["surface_alt"], highlightthickness=1, highlightbackground=self.colors["border"], bd=0)
        block.grid(row=row, column=column, sticky="nsew", padx=(0 if column == 0 else 8, 0), pady=0)
        block.columnconfigure(0, weight=1)
        block.rowconfigure(1, weight=1)

        ttk.Label(block, text=title, style="Surface.TLabel").grid(row=0, column=0, sticky="w", padx=12, pady=(10, 6))
        text = self.create_text_editor(block, height=height, background=self.colors["surface_alt"])
        text.configure(wrap="word", bd=0, highlightthickness=0)
        text.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 10))
        text.insert("1.0", content)
        text.configure(state="disabled")

    def add_combo_row(self, parent, row, label, key, values):
        ttk.Label(parent, text=label, style="Surface.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 4), pady=5)
        self.create_combobox(parent, textvariable=self.vars[key], values=values).grid(row=row, column=1, columnspan=2, sticky="ew", padx=(0, 6), pady=5)
        return row + 1

    def add_path_row(self, parent, row, label, key, command, file_hint):
        ttk.Label(parent, text=label, style="Surface.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 4), pady=5)
        self.create_entry(parent, textvariable=self.vars[key]).grid(row=row, column=1, sticky="ew", padx=(0, 6), pady=5)
        self.make_button(parent, text="Browse", command=command).grid(row=row, column=2, sticky="ew", padx=(8, 0), pady=5)
        return row + 1

    def update_registry_controls(self):
        enabled = self.vars["registry_enabled"].get()
        state = "normal" if enabled else "disabled"
        for key in ("registry_keys", "registry_cleanup_if_empty", "registry_cleanup_force"):
            widget = self.bound_text_widgets.get(key)
            if widget is not None:
                widget.configure(state=state)
        if self.import_registry_button is not None:
            self.import_registry_button.configure(state=state)

    def bind_preview_updates(self):
        for variable in self.vars.values():
            variable.trace_add("write", lambda *_args: self.refresh_preview())

    def load_icon_preview_image(self):
        project = self.current_project()
        app_name = project.app_name
        icon_source = project.icon_source.strip()
        if icon_source:
            return load_icon_image(Path(icon_source), app_name), "Using icon override file."

        app_exe_value = project.app_exe.strip()
        if app_exe_value:
            app_exe = Path(app_exe_value)
            if app_exe.exists() and app_exe.is_file():
                temp_icon = None
                try:
                    # Extract to a temp ICO so the preview uses the same decode path
                    # as the generated PortableApps icon set.
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".ico") as handle:
                        temp_icon = Path(handle.name)
                    if extract_embedded_icon(app_exe, temp_icon, project.icon_index):
                        return (
                            load_icon_image(temp_icon, app_name),
                            f"Using embedded icon #{project.icon_index} from {app_exe.name}.",
                        )
                except OSError:
                    pass
                finally:
                    if temp_icon is not None:
                        try:
                            temp_icon.unlink(missing_ok=True)
                        except OSError:
                            pass
                return make_fallback_icon(app_name), f"Could not extract icon from {app_exe.name}; showing fallback icon."

        return make_fallback_icon(app_name), "No icon source selected yet; showing fallback icon."

    def update_icon_preview(self):
        if (
            not self.icon_preview_labels
            or self.icon_preview_caption is None
            or not self.sidebar_icon_preview_labels
            or self.sidebar_icon_preview_caption is None
        ):
            return

        cache_key = (
            self.vars["app_exe"].get().strip(),
            self.vars["icon_source"].get().strip(),
            self.vars["icon_index"].get().strip(),
            self.vars["app_name"].get().strip(),
        )
        if cache_key == self.icon_preview_cache_key:
            return

        image, caption = self.load_icon_preview_image()
        self.icon_preview_images = []
        self.sidebar_icon_preview_images = []
        for index, (_size_label, display_size) in enumerate(ICON_PREVIEW_DISPLAY_SIZES):
            icon_label = self.icon_preview_labels[index]
            sidebar_icon_label = self.sidebar_icon_preview_labels[index]
            preview = image.resize((display_size, display_size), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(preview)
            self.icon_preview_images.append(photo)
            icon_label.configure(image=photo)
            sidebar_photo = ImageTk.PhotoImage(preview)
            self.sidebar_icon_preview_images.append(sidebar_photo)
            sidebar_icon_label.configure(image=sidebar_photo)
        self.icon_preview_caption.configure(text=caption)
        self.sidebar_icon_preview_caption.configure(text=caption)
        self.icon_preview_cache_key = cache_key

    def update_sidebar_splash_preview(self):
        if self.sidebar_splash_preview_label is None or self.sidebar_splash_preview_caption is None:
            return

        splash_path = splash_asset_path()
        if splash_path.exists():
            try:
                with Image.open(splash_path) as splash_image:
                    preview = ImageOps.contain(splash_image.convert("RGBA"), (300, 170), Image.Resampling.LANCZOS)
                self.sidebar_splash_preview_image = ImageTk.PhotoImage(preview)
                self.sidebar_splash_preview_label.configure(image=self.sidebar_splash_preview_image, text="")
                self.sidebar_splash_preview_caption.configure(text=str(splash_path))
                return
            except OSError:
                pass

        self.sidebar_splash_preview_image = None
        self.sidebar_splash_preview_label.configure(image="", text="Splash preview unavailable")
        self.sidebar_splash_preview_caption.configure(text=str(splash_path))

    def update_launcher_tab_title(self, project):
        if self.launcher_tab is None:
            return
        self.launcher_tab.configure(text=f"{project.portable_name}.ini")
        if hasattr(self, "launch_program_executable_var"):
            self.launch_program_executable_var.set(f"{project.package_name}\\{project.app_exe_name or 'YourApp.exe'}")

    def choose_app_exe(self):
        path = filedialog.askopenfilename(title="Choose application executable", filetypes=[("Executables", "*.exe"), ("All files", "*.*")])
        if not path:
            return
        self.apply_selected_app_exe(path)

    def apply_selected_app_exe(self, path):
        self.vars["app_exe"].set(path)
        app_name = detect_app_name_from_exe(path)
        next_defaults = {
            "app_name": app_name,
            "package_name": clean_identifier(app_name),
            "description": f"{app_name} portable launcher",
            "control_start": f"{clean_identifier(app_name)}Portable.exe",
        }
        for key, detected_value in next_defaults.items():
            current_value = self.vars[key].get().strip()
            if not current_value or current_value == self.detected_defaults.get(key, ""):
                self.vars[key].set(detected_value)
        self.detected_defaults = next_defaults
        self.refresh_control_text()
        self.status_var.set(f"Selected {Path(path).name}")

    def choose_output_dir(self):
        path = filedialog.askdirectory(title="Choose output folder")
        if path:
            self.vars["output_dir"].set(path)

    def refresh_generator_status(self, launcher_path: Path | None = None) -> Path | None:
        launcher_path = launcher_path or find_portableapps_launcher()
        if launcher_path is None:
            self.generator_status_var.set("Generator not found")
        else:
            self.generator_status_var.set(f"Generator ready: {launcher_path.name}")
        return launcher_path

    def set_busy_state(self, busy: bool, status: str | None = None) -> None:
        state = "disabled" if busy else "normal"
        if self.create_button is not None:
            self.create_button.configure(state=state)
        if self.validate_button is not None:
            self.validate_button.configure(state=state)
        if self.help_button is not None:
            self.help_button.configure(state=state)
        if status is not None:
            self.status_var.set(status)
        self.root.update_idletasks()

    def choose_icon(self):
        path = filedialog.askopenfilename(title="Choose app icon", filetypes=[("Icons", "*.ico"), ("All files", "*.*")])
        if path:
            self.vars["icon_source"].set(path)

    def read_text_file_with_fallbacks(self, path: str) -> str:
        encodings = ("utf-16", "utf-8-sig", "utf-8", "cp1252")
        last_error = None
        for encoding in encodings:
            try:
                return Path(path).read_text(encoding=encoding)
            except UnicodeError as exc:
                last_error = exc
            except OSError:
                raise
        if last_error is not None:
            raise last_error
        raise ValueError(f"Could not read {path}")

    def import_registry_file(self):
        path = filedialog.askopenfilename(
            title="Import saved registry file",
            filetypes=[("Registry exports", "*.reg"), ("All files", "*.*")],
        )
        if not path:
            return

        try:
            content = self.read_text_file_with_fallbacks(path)
        except Exception as exc:
            messagebox.showerror("Could Not Read Registry Export", str(exc))
            return

        new_lines = build_registry_key_entries_from_reg_text(content)
        if not new_lines:
            messagebox.showwarning(
                "No Registry Keys Found",
                "No registry key headers were found in that .reg file.",
            )
            return

        existing_text = self.vars["registry_keys"].get()
        replace_existing = True
        if has_ini_lines(existing_text):
            decision = messagebox.askyesnocancel(
                "Update RegistryKeys",
                "Replace the current RegistryKeys entries?\n\nChoose No to append the imported keys.",
            )
            if decision is None:
                return
            replace_existing = decision

        merged_lines = merge_ini_line_sets(existing_text, new_lines, replace_existing)
        self.vars["registry_keys"].set(merged_lines)
        self.vars["registry_enabled"].set(True)
        imported_count = len(new_lines)
        self.status_var.set(f"Imported {imported_count} registry key{'s' if imported_count != 1 else ''} from {Path(path).name}")

    def validate_current_project(self):
        project = self.current_project()
        launcher_path = self.refresh_generator_status()
        items = build_validation_items(project, launcher_path)
        title, _report, status = render_validation_report(items)
        if status == "error":
            self.status_var.set("Validation found issues that need to be fixed.")
        elif status == "warning":
            self.status_var.set("Validation completed with warnings.")
        else:
            self.status_var.set("Validation passed.")
        self.show_validation_popup(title, status, items)

    def close_validation_popup(self):
        if self.validation_window is None:
            return
        try:
            self.validation_window.grab_release()
        except tk.TclError:
            pass
        try:
            self.validation_window.destroy()
        except tk.TclError:
            pass
        self.validation_window = None

    def validation_status_meta(self, status):
        if status == "error":
            return {
                "badge": "Needs fixes",
                "summary": "Fix the blocking issues before building your project.",
                "accent": self.colors["danger"],
                "soft": self.colors["danger_soft"],
                "line": self.colors["danger_line"],
            }
        if status == "warning":
            return {
                "badge": "Needs review",
                "summary": "The project is close, but a few settings still need a quick look.",
                "accent": self.colors["warn"],
                "soft": self.colors["warn_soft"],
                "line": self.colors["warn_line"],
            }
        return {
            "badge": "Ready",
            "summary": "Everything needed for a solid PortableApps project build looks ready.",
            "accent": self.colors["accent"],
            "soft": self.colors["accent_soft"],
            "line": self.colors["accent_line"],
        }

    def draw_validation_status_icon(self, canvas, status, size=46, background=None):
        background = self.colors["surface"] if background is None else background
        meta = self.validation_status_meta(status)
        canvas.configure(width=size, height=size, bg=background, highlightthickness=0, bd=0)
        canvas.delete("all")
        inset = 2
        canvas.create_oval(
            inset,
            inset,
            size - inset,
            size - inset,
            fill=meta["soft"],
            outline=meta["line"],
            width=1,
        )
        if status == "ok":
            points = (
                size * 0.28,
                size * 0.53,
                size * 0.44,
                size * 0.68,
                size * 0.73,
                size * 0.34,
            )
            canvas.create_line(
                *points,
                fill=meta["accent"],
                width=4,
                capstyle=tk.ROUND,
                joinstyle=tk.ROUND,
            )
        elif status == "warning":
            canvas.create_line(
                size * 0.5,
                size * 0.23,
                size * 0.5,
                size * 0.58,
                fill=meta["accent"],
                width=4,
                capstyle=tk.ROUND,
            )
            canvas.create_oval(
                size * 0.46,
                size * 0.7,
                size * 0.54,
                size * 0.78,
                fill=meta["accent"],
                outline=meta["accent"],
            )
        else:
            canvas.create_line(
                size * 0.32,
                size * 0.32,
                size * 0.68,
                size * 0.68,
                fill=meta["accent"],
                width=4,
                capstyle=tk.ROUND,
            )
            canvas.create_line(
                size * 0.68,
                size * 0.32,
                size * 0.32,
                size * 0.68,
                fill=meta["accent"],
                width=4,
                capstyle=tk.ROUND,
            )

    def add_validation_item_row(self, parent, row, item, level):
        meta = self.validation_status_meta(level)
        row_frame = tk.Frame(
            parent,
            bg=self.colors["surface_alt"],
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            bd=0,
            padx=12,
            pady=10,
        )
        row_frame.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        row_frame.columnconfigure(1, weight=1)

        icon_canvas = tk.Canvas(
            row_frame,
            width=18,
            height=18,
            bg=self.colors["surface_alt"],
            highlightthickness=0,
            bd=0,
        )
        icon_canvas.grid(row=0, column=0, rowspan=2, sticky="n", padx=(0, 10), pady=(2, 0))
        self.draw_validation_status_icon(icon_canvas, level, size=18, background=self.colors["surface_alt"])

        tk.Label(
            row_frame,
            text=item.label,
            bg=self.colors["surface_alt"],
            fg=self.colors["text"],
            font=("Segoe UI Semibold", 10),
            anchor="w",
        ).grid(row=0, column=1, sticky="ew")
        tk.Label(
            row_frame,
            text=item.detail,
            bg=self.colors["surface_alt"],
            fg=self.colors["muted"],
            font=("Segoe UI", 9),
            justify="left",
            anchor="w",
            wraplength=680,
        ).grid(row=1, column=1, sticky="ew", pady=(4, 0))

    def add_validation_section(self, parent, row, title, note, items, level):
        section = self.create_panel(parent, title, note)
        section.grid(row=row, column=0, sticky="ew", pady=(0, 12))
        body = section.content
        body.columnconfigure(0, weight=1)
        for index, item in enumerate(items):
            self.add_validation_item_row(body, index, item, level)
        return row + 1

    def show_validation_popup(self, title, status, items):
        self.close_validation_popup()

        errors = [item for item in items if item.level == "error"]
        warnings = [item for item in items if item.level == "warning"]
        oks = [item for item in items if item.level == "ok"]
        meta = self.validation_status_meta(status)

        window = tk.Toplevel(self.root)
        window.title(title)
        window.geometry("880x640")
        window.minsize(760, 520)
        window.configure(bg=self.colors["page"])
        window.transient(self.root)
        window.columnconfigure(0, weight=1)
        window.rowconfigure(0, weight=1)
        window.protocol("WM_DELETE_WINDOW", self.close_validation_popup)
        window.bind("<Escape>", lambda _event: self.close_validation_popup())
        self.apply_window_icon(window)
        self.validation_window = window

        shell = ttk.Frame(window, padding=16)
        shell.grid(row=0, column=0, sticky="nsew")
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(1, weight=1)

        header = tk.Frame(
            shell,
            bg=self.colors["surface"],
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            bd=0,
            padx=18,
            pady=18,
        )
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header.columnconfigure(1, weight=1)

        header_icon = tk.Canvas(header, width=46, height=46, bg=self.colors["surface"], highlightthickness=0, bd=0)
        header_icon.grid(row=0, column=0, rowspan=2, sticky="nw", padx=(0, 14))
        self.draw_validation_status_icon(header_icon, status, size=46, background=self.colors["surface"])

        tk.Label(
            header,
            text=title,
            bg=self.colors["surface"],
            fg=self.colors["text"],
            font=("Segoe UI Semibold", 15),
            anchor="w",
        ).grid(row=0, column=1, sticky="w")

        counts_text = f"{len(errors)} errors   {len(warnings)} warnings   {len(oks)} checks"
        summary_text = meta["summary"] + "  " + counts_text
        tk.Label(
            header,
            text=summary_text,
            bg=self.colors["surface"],
            fg=self.colors["muted"],
            font=("Segoe UI", 9),
            justify="left",
            anchor="w",
            wraplength=620,
        ).grid(row=1, column=1, sticky="ew", pady=(6, 0))

        badge = tk.Label(
            header,
            text=meta["badge"],
            bg=meta["soft"],
            fg=meta["accent"],
            font=("Segoe UI Semibold", 9),
            padx=12,
            pady=6,
            bd=0,
            highlightthickness=1,
            highlightbackground=meta["line"],
        )
        badge.grid(row=0, column=2, rowspan=2, sticky="ne")

        scroll_shell, scroll_content = self.create_scrollable_tab(shell)
        scroll_shell.grid(row=1, column=0, sticky="nsew")
        scroll_content.configure(padding=0)
        scroll_content.columnconfigure(0, weight=1)

        row = 0
        if errors:
            row = self.add_validation_section(
                scroll_content,
                row,
                "Errors",
                "These need to be fixed before you build the project.",
                errors,
                "error",
            )
        if warnings:
            row = self.add_validation_section(
                scroll_content,
                row,
                "Warnings",
                "These are worth reviewing so the generated project behaves the way you expect.",
                warnings,
                "warning",
            )
        if oks:
            row = self.add_validation_section(
                scroll_content,
                row,
                "Checks",
                "These parts already look good.",
                oks,
                "ok",
            )

        footer = ttk.Frame(shell)
        footer.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        footer.columnconfigure(0, weight=1)
        self.make_button(footer, text="Close", command=self.close_validation_popup).grid(row=0, column=1, sticky="e")

        window.grab_set()
        window.focus_force()

    def variable_help_content(self):
        drive_help = "\n".join(
            [
                "Drive variables",
                "%PAL:Drive% - current drive with colon",
                "%PAL:LastDrive% - previous drive with colon",
                "%PAL:DriveLetter% - current drive without colon",
                "%PAL:LastDriveLetter% - previous drive without colon",
                "Examples: %PAL:Drive% -> X:   %PAL:DriveLetter% -> X",
            ]
        )
        directory_help = "\n".join(
            [
                "Directory variables",
                "%PAL:AppDir% - current App directory",
                "%PAL:DataDir% - current Data directory",
                "%PAL:PortableAppsDir% - parent PortableApps directory",
                "%PAL:PortableAppsBaseDir% - root of PortableApps hierarchy",
                "%PAL:LastPortableAppsBaseDir% - previous base directory",
                "%PortableApps.comDocuments% - portable Documents directory",
                "%PortableApps.comPictures% - portable Pictures directory",
                "%PortableApps.comMusic% - portable Music directory",
                "%PortableApps.comVideos% - portable Videos directory",
                "%JAVA_HOME% - Java path when [Activate]:Java=find/require",
                "%USERPROFILE%  %ALLUSERSPROFILE%  %ALLUSERSAPPDATA%",
                "%LOCALAPPDATA%  %APPDATA%  %DOCUMENTS%  %TEMP%",
                "Alternate forms apply to directory vars:",
                ":ForwardSlash  :DoubleBackslash  :java.util.prefs",
                "Example: %PAL:AppDir:ForwardSlash%",
            ]
        )
        partial_and_language_help = "\n".join(
            [
                "Partial directory variables",
                "%PAL:PackagePartialDir% - current package path without drive",
                "%PAL:LastPackagePartialDir% - previous package path without drive",
                "",
                "Language variables",
                "%PortableApps.comLanguageCode%",
                "%PortableApps.comLocaleCode2%",
                "%PortableApps.comLocaleCode3%",
                "%PortableApps.comLocaleglibc%",
                "%PortableApps.comLocaleID%",
                "%PortableApps.comLocaleWinName%",
                "%PortableApps.comLocaleName%",
                "%PAL:LanguageCustom%",
            ]
        )
        return drive_help, directory_help, partial_and_language_help

    def registry_help_content(self):
        registry_keys_help = "\n".join(
            [
                "[Activate]:Registry must be true or registry sections are ignored.",
                "",
                "[RegistryKeys]",
                "Use file-name=registry-key-location.",
                r"Example: appname_portable=HKCU\Software\Publisher\AppName",
                "",
                "The file name becomes Data\\settings\\file-name.reg.",
                "Use -=HKCU\\... if you only want to protect local data and discard changes.",
            ]
        )
        cleanup_if_empty_help = "\n".join(
            [
                "[RegistryCleanupIfEmpty]",
                "Use consecutive integers as keys: 1, 2, 3...",
                r"Example: 1=HKCU\Software\Publisher",
                "",
                "This removes parent keys only if they are empty after cleanup.",
                "Order matters when cleaning nested keys.",
            ]
        )
        cleanup_force_help = "\n".join(
            [
                "[RegistryCleanupForce]",
                "Use consecutive integers as keys: 1, 2, 3...",
                r"Example: 1=HKCU\Software\Publisher\AppName\Temp",
                "",
                "This forcibly removes leftover registry keys after the app exits.",
            ]
        )
        value_write_help = "\n".join(
            [
                "[RegistryValueWrite]",
                r"Format: HKCU\Software\App\Key\Value=REG_SZ:%PAL:DataDir%",
                r"Example: HKCU\Software\App\DisableAssociations=REG_DWORD:1",
                "",
                "REG_TYPE: is optional and defaults to REG_SZ.",
                "Useful for setting values before launch without moving whole keys.",
            ]
        )
        value_backup_delete_help = "\n".join(
            [
                "[RegistryValueBackupDelete]",
                "Use consecutive integers as keys: 1, 2, 3...",
                r"Example: 1=HKCU\Software\Publisher\AppName\DeadValue",
                "",
                "Backs up the value first, restores it later,",
                "and deletes any value written by the portable app while running.",
            ]
        )
        return registry_keys_help, cleanup_if_empty_help, cleanup_force_help, value_write_help, value_backup_delete_help

    def close_help(self):
        if self.help_window is not None:
            try:
                self.help_window.destroy()
            except tk.TclError:
                pass
            self.help_window = None

    def open_variable_help_site(self):
        try:
            webbrowser.open("https://portableapps.com/manuals/PortableApps.comLauncher/ref/envsub.html#ref-envsub")
        except Exception as exc:
            messagebox.showerror("Could Not Open Help", str(exc))

    def open_registry_help_site(self):
        try:
            webbrowser.open("https://portableapps.com/manuals/PortableApps.comLauncher/ref/envsub.html#ref-envsub")
        except Exception as exc:
            messagebox.showerror("Could Not Open Help", str(exc))

    def open_help(self):
        if self.help_window is not None:
            try:
                self.help_window.lift()
                self.help_window.focus_force()
                return
            except tk.TclError:
                self.help_window = None

        window = tk.Toplevel(self.root)
        window.title("PAL Help")
        window.geometry("1100x720")
        window.minsize(980, 620)
        window.configure(bg=self.colors["page"])
        window.transient(self.root)
        window.columnconfigure(0, weight=1)
        window.rowconfigure(0, weight=1)
        window.protocol("WM_DELETE_WINDOW", self.close_help)
        self.apply_window_icon(window)
        self.help_window = window
        self.create_help_content(window, padx=16, pady=16)

    def current_project(self) -> LauncherProject:
        app_name = clean_display_name(self.vars["app_name"].get(), "My App")
        package_name = clean_identifier(self.vars["package_name"].get() or app_name)
        try:
            icon_index = max(0, int(self.vars["icon_index"].get().strip() or "0"))
        except ValueError:
            icon_index = 0
        return LauncherProject(
            app_name=app_name,
            package_name=package_name,
            publisher=self.vars["publisher"].get().strip() or app_name,
            trademarks=self.vars["trademarks"].get().strip(),
            homepage=self.vars["homepage"].get().strip(),
            category=self.vars["category"].get().strip(),
            language=self.vars["language"].get().strip(),
            description=self.vars["description"].get().strip(),
            donate=self.vars["donate"].get().strip(),
            install_type=self.vars["install_type"].get().strip(),
            version=self.vars["version"].get().strip(),
            display_version=self.vars["display_version"].get().strip(),
            app_exe=self.vars["app_exe"].get().strip(),
            output_dir=self.vars["output_dir"].get().strip(),
            command_line=self.vars["command_line"].get().strip(),
            working_directory=self.vars["working_directory"].get().strip() or "%PAL:AppDir%\\{app_name}",
            wait_for_program=self.vars["wait_for_program"].get(),
            close_exe=self.vars["close_exe"].get().strip(),
            wait_for_other_instances=self.vars["wait_for_other_instances"].get(),
            min_os=self.vars["min_os"].get().strip(),
            max_os=self.vars["max_os"].get().strip(),
            run_as_admin=self.vars["run_as_admin"].get().strip(),
            refresh_shell_icons=self.vars["refresh_shell_icons"].get().strip(),
            hide_command_line_window=self.vars["hide_command_line_window"].get(),
            no_spaces_in_path=self.vars["no_spaces_in_path"].get(),
            supports_unc=self.vars["supports_unc"].get().strip(),
            activate_java=self.vars["activate_java"].get().strip(),
            activate_xml=self.vars["activate_xml"].get(),
            live_mode_copy_app=self.vars["live_mode_copy_app"].get(),
            live_mode_copy_data=self.vars["live_mode_copy_data"].get(),
            files_move=self.vars["files_move"].get(),
            directories_move=self.vars["directories_move"].get(),
            installer_close_exe=self.vars["installer_close_exe"].get().strip(),
            installer_close_name=self.vars["installer_close_name"].get().strip(),
            include_installer_source=self.vars["include_installer_source"].get(),
            remove_app_directory=self.vars["remove_app_directory"].get(),
            remove_data_directory=self.vars["remove_data_directory"].get(),
            remove_other_directory=self.vars["remove_other_directory"].get(),
            optional_components_enabled=self.vars["optional_components_enabled"].get(),
            main_section_title=self.vars["main_section_title"].get().strip(),
            main_section_description=self.vars["main_section_description"].get().strip(),
            optional_section_title=self.vars["optional_section_title"].get().strip(),
            optional_section_description=self.vars["optional_section_description"].get().strip(),
            optional_section_selected_install_type=self.vars["optional_section_selected_install_type"].get().strip(),
            optional_section_not_selected_install_type=self.vars["optional_section_not_selected_install_type"].get().strip(),
            optional_section_preselected=self.vars["optional_section_preselected"].get().strip(),
            installer_languages=self.vars["installer_languages"].get(),
            preserve_directories=self.vars["preserve_directories"].get(),
            remove_directories=self.vars["remove_directories"].get(),
            preserve_files=self.vars["preserve_files"].get(),
            remove_files=self.vars["remove_files"].get(),
            copy_app_files=self.vars["copy_app_files"].get(),
            icon_source=self.vars["icon_source"].get().strip(),
            icon_index=icon_index,
            registry_enabled=self.vars["registry_enabled"].get(),
            registry_keys=self.vars["registry_keys"].get(),
            registry_cleanup_if_empty=self.vars["registry_cleanup_if_empty"].get(),
            registry_cleanup_force=self.vars["registry_cleanup_force"].get(),
            license_shareable=self.vars["license_shareable"].get(),
            license_open_source=self.vars["license_open_source"].get(),
            license_freeware=self.vars["license_freeware"].get(),
            license_commercial_use=self.vars["license_commercial_use"].get(),
            license_eula_version=self.vars["license_eula_version"].get().strip(),
            special_plugins=self.vars["special_plugins"].get().strip(),
            dependency_uses_ghostscript=self.vars["dependency_uses_ghostscript"].get().strip(),
            dependency_uses_java=self.vars["dependency_uses_java"].get().strip(),
            dependency_uses_dotnet_version=self.vars["dependency_uses_dotnet_version"].get().strip(),
            dependency_requires_64bit_os=self.vars["dependency_requires_64bit_os"].get().strip(),
            dependency_requires_portable_app=self.vars["dependency_requires_portable_app"].get().strip(),
            dependency_requires_admin=self.vars["dependency_requires_admin"].get().strip(),
            control_icons=self.vars["control_icons"].get().strip(),
            control_start=self.vars["control_start"].get().strip(),
            control_extract_icon=self.vars["control_extract_icon"].get().strip(),
            control_extract_name=self.vars["control_extract_name"].get().strip(),
            control_base_app_id=self.vars["control_base_app_id"].get().strip(),
            control_base_app_id_64=self.vars["control_base_app_id_64"].get().strip(),
            control_base_app_id_arm64=self.vars["control_base_app_id_arm64"].get().strip(),
            control_exit_exe=self.vars["control_exit_exe"].get().strip(),
            control_exit_parameters=self.vars["control_exit_parameters"].get().strip(),
            association_file_types=self.vars["association_file_types"].get().strip(),
            association_file_type_command_line=self.vars["association_file_type_command_line"].get().strip(),
            association_file_type_command_line_extension=self.vars["association_file_type_command_line_extension"].get().strip(),
            association_protocols=self.vars["association_protocols"].get().strip(),
            association_protocol_command_line=self.vars["association_protocol_command_line"].get().strip(),
            association_protocol_command_line_protocol=self.vars["association_protocol_command_line_protocol"].get().strip(),
            association_send_to=self.vars["association_send_to"].get(),
            association_send_to_command_line=self.vars["association_send_to_command_line"].get().strip(),
            association_shell=self.vars["association_shell"].get(),
            association_shell_command=self.vars["association_shell_command"].get().strip(),
            file_type_icons=self.vars["file_type_icons"].get(),
        )

    def refresh_preview(self):
        project = self.current_project()
        self.update_launcher_tab_title(project)
        previews = {
            "folder": self.build_folder_preview_text(project),
            "appinfo": build_appinfo_ini(project),
            "launcher": build_launcher_ini(project),
            "installer": build_installer_ini(project) or "; installer.ini is optional and will only be created when installer options are set.",
        }
        for key, content in previews.items():
            text = self.preview_texts.get(key)
            if text is None:
                continue
            text.configure(state="normal")
            text.delete("1.0", "end")
            if key == "folder":
                self.insert_styled_folder_preview(text, content)
            else:
                text.insert("1.0", content)
            text.configure(state="disabled")
        self.update_icon_preview()
        self.update_sidebar_splash_preview()

    def build_folder_preview_text(self, project):
        installer_enabled = bool(build_installer_ini(project))
        portable_exe = f"{project.portable_name}.exe"
        launcher_ini = f"{project.portable_name}.ini"
        app_exe_name = project.app_exe_name or "YourApp.exe"

        # Each tuple is (tag, line). The text tags let the preview emphasize
        # generated-important files without giving up the lightweight text view.
        return [
            ("folder", f"{project.portable_name}\\\n"),
            ("important", f"|- {portable_exe}\n"),
            ("comment", "   build with PortableApps.com Launcher\n"),
            ("important", "|- help.html\n"),
            ("folder", "|- App\\\n"),
            ("folder", "|  |- AppInfo\\\n"),
            ("important", "|  |  |- appinfo.ini\n"),
            ("important" if installer_enabled else "optional", "|  |  |- installer.ini\n"),
            ("important", "|  |  |- appicon.ico\n"),
            ("plain", "|  |  |- appicon_16.png\n"),
            ("plain", "|  |  |- appicon_32.png\n"),
            ("plain", "|  |  |- appicon_75.png\n"),
            ("plain", "|  |  |- appicon_128.png\n"),
            ("plain", "|  |  |- appicon_256.png\n"),
            ("comment", f"|  |  `- source icon index: {project.icon_index}\n"),
            ("folder", "|  |  `- Launcher\\\n"),
            ("important", f"|  |     |- {launcher_ini}\n"),
            ("important", "|  |     `- Splash.jpg\n"),
            ("folder", f"|  `- {project.package_name}\\\n"),
            ("plain", f"|     `- {app_exe_name}\n"),
            ("folder", "|- Data\\\n"),
            ("folder", "|  `- settings\\\n"),
            ("folder", "`- Other\\\n"),
            ("folder", "   |- Help\\\n"),
            ("folder", "   |  `- Images\\\n"),
            ("plain", "   |     `- appicon_128.png\n"),
            ("folder", "   `- Source\\\n"),
            ("plain", "      `- Readme.txt\n"),
        ]

    def insert_styled_folder_preview(self, text, lines):
        for tag, line in lines:
            text.insert("end", line, tag or "plain")

    def create_project(self):
        project = self.current_project()
        validation_errors = validate_project(project)
        if validation_errors:
            self.status_var.set("Fix the project settings and try again.")
            messagebox.showerror(
                "Project Settings Need Attention",
                "\n".join(f"- {error}" for error in validation_errors),
            )
            return

        launcher_path = self.refresh_generator_status()
        if launcher_path is None:
            self.status_var.set("PortableApps.com Launcher Generator not found.")
            should_open_downloads = messagebox.askyesno(
                "Launcher Generator Not Found",
                "PortableApps.com Launcher Generator is required to create the portable launcher EXE, but it was not found.\n\n"
                "Do you want to open the PortableApps.com development page?",
            )
            if should_open_downloads:
                webbrowser.open(PORTABLEAPPS_DEVELOPMENT_DOWNLOADS_URL)
            return

        self.set_busy_state(True, "Creating project files...")
        try:
            try:
                project_root = create_launcher_project(project)
            except Exception as exc:
                messagebox.showerror("Could Not Create Project", str(exc))
                self.status_var.set("Project creation failed.")
                return

            self.set_busy_state(True, f"Project created at {project_root}. Building portable launcher...")
            try:
                result = subprocess.run(
                    [str(launcher_path), str(project_root)],
                    cwd=str(launcher_path.parent),
                    check=False,
                    capture_output=True,
                    text=True,
                    timeout=300,
                )
            except (OSError, subprocess.SubprocessError) as exc:
                self.status_var.set(f"Created {project_root} (launcher build failed)")
                messagebox.showwarning(
                    "Project Created, Build Failed",
                    "PortableApps.com project created, but the launcher EXE could not be built automatically.\n\n"
                    f"Launcher: {launcher_path}\n"
                    f"Project: {project_root}\n\n"
                    f"{exc}",
                )
                return

            portable_exe = project_root / f"{project.portable_name}.exe"
            if portable_exe.exists():
                self.status_var.set(f"Created {portable_exe}")
                messagebox.showinfo(
                    "Portable EXE Created",
                    "PortableApps.com project created and launcher EXE built.\n\n"
                    f"{portable_exe}",
                )
                open_folder_in_explorer(project_root)
                return

            details = (result.stderr or result.stdout or "").strip()
            self.status_var.set(f"Created {project_root} (launcher build incomplete)")
            messagebox.showwarning(
                "Project Created, EXE Not Found",
                "PortableApps.com project created, but the launcher EXE was not found after running the generator.\n\n"
                f"Launcher: {launcher_path}\n"
                f"Project: {project_root}\n\n"
                f"{details[:1200]}",
            )
        finally:
            self.set_busy_state(False)

def run():
    root = create_root_window()
    PortableAppsLauncherMaker(root)
    root.mainloop()
