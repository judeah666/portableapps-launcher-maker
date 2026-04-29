import base64
import io
import subprocess
import tempfile
import webbrowser
from pathlib import Path

import webview
from PIL import Image, ImageOps

from app.portableapps_core import (
    ICON_PREVIEW_DISPLAY_SIZES,
    PORTABLEAPPS_DEVELOPMENT_DOWNLOADS_URL,
    LauncherProject,
    asset_path,
    bool_to_ini,
    build_appinfo_ini,
    build_installer_ini,
    build_launcher_ini,
    build_registry_key_entries_from_reg_text,
    build_validation_items,
    clean_display_name,
    clean_identifier,
    create_launcher_project,
    default_portableapps_output_dir,
    detect_app_name_from_exe,
    extract_embedded_icon,
    find_portableapps_launcher,
    has_ini_lines,
    load_icon_image,
    make_fallback_icon,
    merge_ini_line_sets,
    open_folder_in_explorer,
    render_validation_report,
    splash_asset_path,
    validate_project,
)
from app.portableapps_ui_theme import UI_COLORS
from app.version import APP_VERSION


APP_SCHEMA = [
    {
        "key": "appinfo",
        "label": "appinfo.ini",
        "sections": [
            {
                "title": "Details",
                "fields": [
                    {"key": "app_name", "label": "App Name", "type": "text"},
                    {"key": "package_name", "label": "Package ID", "type": "text"},
                    {"key": "publisher", "label": "Publisher", "type": "text"},
                    {"key": "trademarks", "label": "Trademarks", "type": "text"},
                    {
                        "key": "category",
                        "label": "Category",
                        "type": "select",
                        "options": [
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
                        ],
                    },
                    {
                        "key": "language",
                        "label": "Language",
                        "type": "select",
                        "options": [
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
                        ],
                    },
                    {"key": "description", "label": "Description", "type": "text", "wide": True},
                    {"key": "homepage", "label": "Homepage", "type": "text", "wide": True},
                    {"key": "donate", "label": "Donate", "type": "text", "wide": True},
                    {"key": "install_type", "label": "Install Type", "type": "text", "wide": True},
                ],
            },
            {
                "title": "Version",
                "fields": [
                    {"key": "version", "label": "Package Version", "type": "text"},
                    {"key": "display_version", "label": "Display Version", "type": "text"},
                ],
            },
            {
                "title": "SpecialPaths",
                "fields": [
                    {"key": "special_plugins", "label": "Plugins", "type": "text", "wide": True},
                ],
            },
            {
                "title": "Dependencies",
                "fields": [
                    {
                        "key": "dependency_uses_ghostscript",
                        "label": "Uses Ghostscript",
                        "type": "select",
                        "options": ["no", "yes", "optional"],
                    },
                    {
                        "key": "dependency_uses_java",
                        "label": "Uses Java",
                        "type": "select",
                        "options": ["no", "yes", "optional"],
                    },
                    {"key": "dependency_uses_dotnet_version", "label": ".NET Version", "type": "text"},
                    {
                        "key": "dependency_requires_64bit_os",
                        "label": "Requires 64-bit OS",
                        "type": "select",
                        "options": ["no", "yes"],
                    },
                    {
                        "key": "dependency_requires_admin",
                        "label": "Requires Admin",
                        "type": "select",
                        "options": ["no", "yes"],
                    },
                    {"key": "dependency_requires_portable_app", "label": "Requires Portable App", "type": "text", "wide": True},
                ],
            },
            {
                "title": "License",
                "fields": [
                    {"key": "license_shareable", "label": "Shareable", "type": "checkbox"},
                    {"key": "license_open_source", "label": "Open Source", "type": "checkbox"},
                    {"key": "license_freeware", "label": "Freeware", "type": "checkbox"},
                    {"key": "license_commercial_use", "label": "Commercial Use", "type": "checkbox"},
                    {"key": "license_eula_version", "label": "EULA Version", "type": "text"},
                ],
            },
            {
                "title": "Control",
                "fields": [
                    {"key": "control_text", "label": "Entries", "type": "textarea", "rows": 5},
                ],
            },
            {
                "title": "Associations",
                "fields": [
                    {"key": "associations_text", "label": "Entries", "type": "textarea", "rows": 8},
                ],
            },
            {
                "title": "FileTypeIcons",
                "fields": [
                    {"key": "file_type_icons", "label": "Entries", "type": "textarea", "rows": 6},
                ],
            },
        ],
    },
    {
        "key": "launcher",
        "label": "AppNamePortable.ini",
        "sections": [
            {
                "title": "Launch",
                "fields": [
                    {"key": "launch_program_executable", "label": "ProgramExecutable", "type": "readonly", "wide": True},
                    {"key": "command_line", "label": "Arguments", "type": "text", "wide": True},
                    {"key": "working_directory", "label": "Working Dir", "type": "text"},
                    {"key": "close_exe", "label": "Close EXE", "type": "text"},
                    {"key": "min_os", "label": "Min OS", "type": "text"},
                    {"key": "max_os", "label": "Max OS", "type": "text"},
                    {"key": "run_as_admin", "label": "Run As Admin", "type": "text"},
                    {"key": "refresh_shell_icons", "label": "Refresh Shell Icons", "type": "text"},
                    {"key": "supports_unc", "label": "Supports UNC", "type": "text"},
                    {"key": "copy_app_files", "label": "Copy selected app folder into App folder", "type": "checkbox", "wide": True},
                    {"key": "wait_for_program", "label": "Wait for program before cleanup", "type": "checkbox", "wide": True},
                    {"key": "wait_for_other_instances", "label": "Wait for other instances", "type": "checkbox"},
                    {"key": "hide_command_line_window", "label": "Hide command line window", "type": "checkbox"},
                    {"key": "no_spaces_in_path", "label": "No spaces in path", "type": "checkbox"},
                ],
            },
            {
                "title": "Activate",
                "fields": [
                    {"key": "activate_java", "label": "Java", "type": "text"},
                    {"key": "activate_xml", "label": "Enable XML support", "type": "checkbox"},
                ],
            },
            {
                "title": "LiveMode",
                "fields": [
                    {"key": "live_mode_copy_app", "label": "Copy App", "type": "checkbox"},
                    {"key": "live_mode_copy_data", "label": "Copy Data", "type": "checkbox"},
                ],
            },
            {"title": "FilesMove", "fields": [{"key": "files_move", "label": "Entries", "type": "textarea", "rows": 5}]},
            {"title": "DirectoriesMove", "fields": [{"key": "directories_move", "label": "Entries", "type": "textarea", "rows": 5}]},
        ],
    },
    {
        "key": "installer",
        "label": "installer.ini",
        "sections": [
            {
                "title": "CheckRunning",
                "fields": [
                    {"key": "installer_close_exe", "label": "CloseEXE", "type": "text"},
                    {"key": "installer_close_name", "label": "CloseName", "type": "text"},
                ],
            },
            {
                "title": "Source",
                "fields": [{"key": "include_installer_source", "label": "Include Installer Source", "type": "checkbox"}],
            },
            {
                "title": "MainDirectories",
                "fields": [
                    {"key": "remove_app_directory", "label": "Remove App Directory", "type": "checkbox"},
                    {"key": "remove_data_directory", "label": "Remove Data Directory", "type": "checkbox"},
                    {"key": "remove_other_directory", "label": "Remove Other Directory", "type": "checkbox"},
                ],
            },
            {
                "title": "OptionalComponents",
                "fields": [
                    {"key": "optional_components_enabled", "label": "Enable Optional Components", "type": "checkbox", "wide": True},
                    {"key": "main_section_title", "label": "Main Section Title", "type": "text", "wide": True},
                    {"key": "main_section_description", "label": "Main Section Description", "type": "text", "wide": True},
                    {"key": "optional_section_title", "label": "Optional Section Title", "type": "text", "wide": True},
                    {"key": "optional_section_description", "label": "Optional Section Description", "type": "text", "wide": True},
                    {"key": "optional_section_selected_install_type", "label": "Selected Install Type", "type": "text"},
                    {"key": "optional_section_not_selected_install_type", "label": "Not Selected Install Type", "type": "text"},
                    {"key": "optional_section_preselected", "label": "Preselected", "type": "text"},
                ],
            },
            {"title": "Languages", "fields": [{"key": "installer_languages", "label": "Entries", "type": "textarea", "rows": 5}]},
            {"title": "DirectoriesToPreserve", "fields": [{"key": "preserve_directories", "label": "Entries", "type": "textarea", "rows": 4}]},
            {"title": "DirectoriesToRemove", "fields": [{"key": "remove_directories", "label": "Entries", "type": "textarea", "rows": 4}]},
            {"title": "FilesToPreserve", "fields": [{"key": "preserve_files", "label": "Entries", "type": "textarea", "rows": 4}]},
            {"title": "FilesToRemove", "fields": [{"key": "remove_files", "label": "Entries", "type": "textarea", "rows": 4}]},
        ],
    },
    {
        "key": "registry",
        "label": "Registry",
        "sections": [
            {
                "title": "Registry",
                "fields": [
                    {"key": "registry_enabled", "label": "Enable registry handling", "type": "checkbox", "wide": True},
                    {"key": "registry_keys", "label": "RegistryKeys", "type": "textarea", "rows": 6},
                    {"key": "registry_cleanup_if_empty", "label": "RegistryCleanupIfEmpty", "type": "textarea", "rows": 4},
                    {"key": "registry_cleanup_force", "label": "RegistryCleanupForce", "type": "textarea", "rows": 4},
                ],
            }
        ],
    },
    {
        "key": "icon",
        "label": "Icon",
        "sections": [
            {
                "title": "Icon Settings",
                "fields": [
                    {"key": "icon_index", "label": "Icon Index", "type": "text"},
                    {"key": "icon_source", "label": "Icon Override", "type": "text", "wide": True},
                ],
            }
        ],
    },
    {
        "key": "splash",
        "label": "Splash",
        "sections": [
            {
                "title": "Splash Asset",
                "fields": [
                    {"key": "splash_asset_path", "label": "Default Splash", "type": "readonly", "wide": True},
                ],
            },
            {"title": "Splash Preview", "fields": []},
        ],
    },
]

PREVIEW_TABS = [
    {"key": "folder", "label": "Folder Preview"},
    {"key": "appinfo", "label": "appinfo.ini"},
    {"key": "launcher", "label": "launcher.ini"},
    {"key": "installer", "label": "installer.ini"},
]


TEXT_KEYS = {
    "app_name": "",
    "package_name": "",
    "publisher": "",
    "trademarks": "",
    "homepage": "https://portableapps.com/",
    "category": "Utilities",
    "language": "Multilingual",
    "description": "",
    "donate": "",
    "install_type": "",
    "version": APP_VERSION,
    "display_version": APP_VERSION,
    "app_exe": "",
    "output_dir": default_portableapps_output_dir(),
    "command_line": "",
    "working_directory": r"%PAL:AppDir%\{app_name}",
    "close_exe": "",
    "min_os": "",
    "max_os": "",
    "run_as_admin": "",
    "refresh_shell_icons": "",
    "supports_unc": "",
    "activate_java": "",
    "files_move": "",
    "directories_move": "",
    "installer_close_exe": "",
    "installer_close_name": "",
    "main_section_title": "",
    "main_section_description": "",
    "optional_section_title": "",
    "optional_section_description": "",
    "optional_section_selected_install_type": "",
    "optional_section_not_selected_install_type": "",
    "optional_section_preselected": "",
    "installer_languages": "",
    "preserve_directories": "",
    "remove_directories": "",
    "preserve_files": "",
    "remove_files": "",
    "icon_source": "",
    "icon_index": "0",
    "registry_keys": "",
    "registry_cleanup_if_empty": "",
    "registry_cleanup_force": "",
    "license_eula_version": "",
    "special_plugins": "NONE",
    "dependency_uses_ghostscript": "no",
    "dependency_uses_java": "no",
    "dependency_uses_dotnet_version": "",
    "dependency_requires_64bit_os": "no",
    "dependency_requires_portable_app": "",
    "dependency_requires_admin": "no",
    "control_icons": "1",
    "control_start": "",
    "control_extract_icon": "",
    "control_extract_name": "",
    "control_base_app_id": "",
    "control_base_app_id_64": "",
    "control_base_app_id_arm64": "",
    "control_exit_exe": "",
    "control_exit_parameters": "",
    "association_file_types": "",
    "association_file_type_command_line": "",
    "association_file_type_command_line_extension": "",
    "association_protocols": "",
    "association_protocol_command_line": "",
    "association_protocol_command_line_protocol": "",
    "association_send_to_command_line": "",
    "association_shell_command": "",
    "file_type_icons": "",
    "control_text": "",
    "associations_text": "",
}

BOOL_KEYS = {
    "wait_for_other_instances": True,
    "hide_command_line_window": False,
    "no_spaces_in_path": False,
    "activate_xml": False,
    "live_mode_copy_app": False,
    "live_mode_copy_data": False,
    "include_installer_source": False,
    "remove_app_directory": False,
    "remove_data_directory": False,
    "remove_other_directory": False,
    "optional_components_enabled": False,
    "copy_app_files": True,
    "wait_for_program": True,
    "registry_enabled": False,
    "license_shareable": True,
    "license_open_source": False,
    "license_freeware": True,
    "license_commercial_use": True,
    "association_send_to": False,
    "association_shell": False,
}


def parse_plain_key_values(text: str) -> dict[str, str]:
    values: dict[str, str] = {}
    for line in (text or "").splitlines():
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key.strip().casefold()] = value.strip()
    return values


def folder_preview_lines(project: LauncherProject) -> list[tuple[str, str]]:
    installer_enabled = bool(build_installer_ini(project))
    portable_exe = f"{project.portable_name}.exe"
    launcher_ini = f"{project.portable_name}.ini"
    app_exe_name = project.app_exe_name or "YourApp.exe"
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


def read_text_file_with_fallbacks(path: str) -> str:
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


def image_to_data_url(image: Image.Image, fmt: str = "PNG") -> str:
    buffer = io.BytesIO()
    image.save(buffer, format=fmt)
    encoded = base64.b64encode(buffer.getvalue()).decode("ascii")
    return f"data:image/{fmt.lower()};base64,{encoded}"


class PortableAppsWebBackend:
    def __init__(self):
        self.window = None
        self.status = "Choose an EXE and output folder, then create the launcher project."
        self.detected_defaults = {
            "app_name": "",
            "package_name": "",
            "description": "",
            "control_start": "",
        }
        self.state = {**TEXT_KEYS, **BOOL_KEYS}
        self.refresh_control_text()
        self.refresh_associations_text()

    def set_window(self, window):
        self.window = window

    def serialize_state(self):
        project = self.current_project()
        serialized = dict(self.state)
        serialized["launch_program_executable"] = f"{project.package_name}\\{project.app_exe_name}" if project.app_exe_name else ""
        serialized["splash_asset_path"] = str(splash_asset_path())
        return serialized

    def refresh_control_text(self):
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
        self.state["control_text"] = "\n".join(lines)

    def refresh_associations_text(self):
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
        self.state["associations_text"] = "\n".join(lines)

    def parse_control_text(self):
        values = parse_plain_key_values(self.state["control_text"])
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
        for key, state_key in key_map.items():
            if key in values:
                self.state[state_key] = values[key]

    def parse_associations_text(self):
        values = parse_plain_key_values(self.state["associations_text"])
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
        for key, state_key in key_map.items():
            if key in values:
                self.state[state_key] = values[key]
        if "sendto" in values:
            self.state["association_send_to"] = values["sendto"].casefold() == "true"
        if "shell" in values:
            self.state["association_shell"] = values["shell"].casefold() == "true"

    def current_project(self) -> LauncherProject:
        app_name = clean_display_name(self.state["app_name"], "My App")
        package_name = clean_identifier(self.state["package_name"] or app_name)
        try:
            icon_index = max(0, int(str(self.state["icon_index"]).strip() or "0"))
        except ValueError:
            icon_index = 0
        return LauncherProject(
            app_name=app_name,
            package_name=package_name,
            publisher=str(self.state["publisher"]).strip() or app_name,
            trademarks=str(self.state["trademarks"]).strip(),
            homepage=str(self.state["homepage"]).strip(),
            category=str(self.state["category"]).strip(),
            language=str(self.state["language"]).strip(),
            description=str(self.state["description"]).strip(),
            donate=str(self.state["donate"]).strip(),
            install_type=str(self.state["install_type"]).strip(),
            version=str(self.state["version"]).strip(),
            display_version=str(self.state["display_version"]).strip(),
            app_exe=str(self.state["app_exe"]).strip(),
            output_dir=str(self.state["output_dir"]).strip(),
            command_line=str(self.state["command_line"]).strip(),
            working_directory=str(self.state["working_directory"]).strip() or r"%PAL:AppDir%\{app_name}",
            wait_for_program=bool(self.state["wait_for_program"]),
            close_exe=str(self.state["close_exe"]).strip(),
            wait_for_other_instances=bool(self.state["wait_for_other_instances"]),
            min_os=str(self.state["min_os"]).strip(),
            max_os=str(self.state["max_os"]).strip(),
            run_as_admin=str(self.state["run_as_admin"]).strip(),
            refresh_shell_icons=str(self.state["refresh_shell_icons"]).strip(),
            hide_command_line_window=bool(self.state["hide_command_line_window"]),
            no_spaces_in_path=bool(self.state["no_spaces_in_path"]),
            supports_unc=str(self.state["supports_unc"]).strip(),
            activate_java=str(self.state["activate_java"]).strip(),
            activate_xml=bool(self.state["activate_xml"]),
            live_mode_copy_app=bool(self.state["live_mode_copy_app"]),
            live_mode_copy_data=bool(self.state["live_mode_copy_data"]),
            files_move=str(self.state["files_move"]),
            directories_move=str(self.state["directories_move"]),
            installer_close_exe=str(self.state["installer_close_exe"]).strip(),
            installer_close_name=str(self.state["installer_close_name"]).strip(),
            include_installer_source=bool(self.state["include_installer_source"]),
            remove_app_directory=bool(self.state["remove_app_directory"]),
            remove_data_directory=bool(self.state["remove_data_directory"]),
            remove_other_directory=bool(self.state["remove_other_directory"]),
            optional_components_enabled=bool(self.state["optional_components_enabled"]),
            main_section_title=str(self.state["main_section_title"]).strip(),
            main_section_description=str(self.state["main_section_description"]).strip(),
            optional_section_title=str(self.state["optional_section_title"]).strip(),
            optional_section_description=str(self.state["optional_section_description"]).strip(),
            optional_section_selected_install_type=str(self.state["optional_section_selected_install_type"]).strip(),
            optional_section_not_selected_install_type=str(self.state["optional_section_not_selected_install_type"]).strip(),
            optional_section_preselected=str(self.state["optional_section_preselected"]).strip(),
            installer_languages=str(self.state["installer_languages"]),
            preserve_directories=str(self.state["preserve_directories"]),
            remove_directories=str(self.state["remove_directories"]),
            preserve_files=str(self.state["preserve_files"]),
            remove_files=str(self.state["remove_files"]),
            copy_app_files=bool(self.state["copy_app_files"]),
            icon_source=str(self.state["icon_source"]).strip(),
            icon_index=icon_index,
            registry_enabled=bool(self.state["registry_enabled"]),
            registry_keys=str(self.state["registry_keys"]),
            registry_cleanup_if_empty=str(self.state["registry_cleanup_if_empty"]),
            registry_cleanup_force=str(self.state["registry_cleanup_force"]),
            license_shareable=bool(self.state["license_shareable"]),
            license_open_source=bool(self.state["license_open_source"]),
            license_freeware=bool(self.state["license_freeware"]),
            license_commercial_use=bool(self.state["license_commercial_use"]),
            license_eula_version=str(self.state["license_eula_version"]).strip(),
            special_plugins=str(self.state["special_plugins"]).strip(),
            dependency_uses_ghostscript=str(self.state["dependency_uses_ghostscript"]).strip(),
            dependency_uses_java=str(self.state["dependency_uses_java"]).strip(),
            dependency_uses_dotnet_version=str(self.state["dependency_uses_dotnet_version"]).strip(),
            dependency_requires_64bit_os=str(self.state["dependency_requires_64bit_os"]).strip(),
            dependency_requires_portable_app=str(self.state["dependency_requires_portable_app"]).strip(),
            dependency_requires_admin=str(self.state["dependency_requires_admin"]).strip(),
            control_icons=str(self.state["control_icons"]).strip(),
            control_start=str(self.state["control_start"]).strip(),
            control_extract_icon=str(self.state["control_extract_icon"]).strip(),
            control_extract_name=str(self.state["control_extract_name"]).strip(),
            control_base_app_id=str(self.state["control_base_app_id"]).strip(),
            control_base_app_id_64=str(self.state["control_base_app_id_64"]).strip(),
            control_base_app_id_arm64=str(self.state["control_base_app_id_arm64"]).strip(),
            control_exit_exe=str(self.state["control_exit_exe"]).strip(),
            control_exit_parameters=str(self.state["control_exit_parameters"]).strip(),
            association_file_types=str(self.state["association_file_types"]).strip(),
            association_file_type_command_line=str(self.state["association_file_type_command_line"]).strip(),
            association_file_type_command_line_extension=str(self.state["association_file_type_command_line_extension"]).strip(),
            association_protocols=str(self.state["association_protocols"]).strip(),
            association_protocol_command_line=str(self.state["association_protocol_command_line"]).strip(),
            association_protocol_command_line_protocol=str(self.state["association_protocol_command_line_protocol"]).strip(),
            association_send_to=bool(self.state["association_send_to"]),
            association_send_to_command_line=str(self.state["association_send_to_command_line"]).strip(),
            association_shell=bool(self.state["association_shell"]),
            association_shell_command=str(self.state["association_shell_command"]).strip(),
            file_type_icons=str(self.state["file_type_icons"]),
        )

    def refresh_generator_status(self, launcher_path: Path | None = None) -> Path | None:
        launcher_path = launcher_path or find_portableapps_launcher()
        if launcher_path is None:
            self.generator_status = "Generator not found"
        else:
            self.generator_status = f"Generator ready: {launcher_path.name}"
        return launcher_path

    def apply_selected_app_exe(self, path: str):
        self.state["app_exe"] = path
        app_name = detect_app_name_from_exe(path)
        next_defaults = {
            "app_name": app_name,
            "package_name": clean_identifier(app_name),
            "description": f"{app_name} portable launcher",
            "control_start": f"{clean_identifier(app_name)}Portable.exe",
        }
        for key, detected_value in next_defaults.items():
            current_value = str(self.state[key]).strip()
            if not current_value or current_value == self.detected_defaults.get(key, ""):
                self.state[key] = detected_value
        self.detected_defaults = next_defaults
        self.refresh_control_text()
        self.status = f"Selected {Path(path).name}"

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

    def get_previews(self):
        project = self.current_project()
        return {
            "folder": "".join(line for _tag, line in folder_preview_lines(project)),
            "appinfo": build_appinfo_ini(project),
            "launcher": build_launcher_ini(project),
            "installer": build_installer_ini(project) or "; installer.ini is optional and will only be created when installer options are set.",
        }

    def get_icon_previews(self):
        image, message = self.load_icon_preview_image()
        items = []
        for size_label, _display_size in ICON_PREVIEW_DISPLAY_SIZES:
            icon_variant = image.copy()
            icon_variant.thumbnail((int(size_label), int(size_label)), Image.Resampling.LANCZOS)
            items.append(
                {
                    "label": f"{size_label}px",
                    "src": image_to_data_url(icon_variant),
                    "width": icon_variant.width,
                    "height": icon_variant.height,
                }
            )
        return {"items": items, "message": message}

    def get_splash_preview(self):
        path = splash_asset_path()
        if not path.exists():
            return None
        with Image.open(path) as image:
            preview = ImageOps.contain(image.convert("RGBA"), (640, 360), Image.Resampling.LANCZOS)
        return image_to_data_url(preview)

    def snapshot(self):
        launcher_path = self.refresh_generator_status()
        project = self.current_project()
        return {
            "state": self.serialize_state(),
            "generatorStatus": self.generator_status,
            "status": self.status,
            "launcherTabLabel": f"{project.portable_name}.ini",
            "previews": self.get_previews(),
            "iconPreviews": self.get_icon_previews(),
            "splashPreview": self.get_splash_preview(),
            "launcherFound": launcher_path is not None,
        }

    def bootstrap(self):
        return {
            "schema": APP_SCHEMA,
            "previewTabs": PREVIEW_TABS,
            "colors": UI_COLORS,
            **self.snapshot(),
        }

    def set_value(self, key, value):
        if key in BOOL_KEYS:
            self.state[key] = bool(value)
        else:
            self.state[key] = value
        if key == "control_text":
            self.parse_control_text()
        elif key == "associations_text":
            self.parse_associations_text()
        elif key.startswith("control_") and key != "control_text":
            self.refresh_control_text()
        elif key.startswith("association_") and key != "associations_text":
            self.refresh_associations_text()
        return self.snapshot()

    def choose_app_exe(self):
        path = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Executable (*.exe)",),
        ) if self.window else None
        if path:
            selected = path[0] if isinstance(path, (list, tuple)) else path
            self.apply_selected_app_exe(selected)
        return self.snapshot()

    def choose_output_dir(self):
        path = self.window.create_file_dialog(webview.FOLDER_DIALOG) if self.window else None
        if path:
            selected = path[0] if isinstance(path, (list, tuple)) else path
            self.state["output_dir"] = selected
        return self.snapshot()

    def choose_icon(self):
        path = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Icons (*.ico)", "All files (*.*)"),
        ) if self.window else None
        if path:
            selected = path[0] if isinstance(path, (list, tuple)) else path
            self.state["icon_source"] = selected
        return self.snapshot()

    def open_assets_folder(self):
        target = asset_path("")
        open_folder_in_explorer(target)
        self.status = f"Opened {target}"
        return self.snapshot()

    def open_splash_asset(self):
        target = splash_asset_path()
        open_folder_in_explorer(target)
        self.status = f"Opened {target}"
        return self.snapshot()

    def replace_splash_asset(self):
        path = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Image files (*.png;*.jpg;*.jpeg)", "All files (*.*)"),
        ) if self.window else None
        if not path:
            return self.snapshot()

        selected = path[0] if isinstance(path, (list, tuple)) else path
        target = splash_asset_path()
        with Image.open(selected) as source_image:
            source_image.convert("RGBA").save(target, format="PNG")
        self.status = f"Updated {target}"
        return self.snapshot()

    def import_registry_file(self):
        path = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Registry exports (*.reg)", "All files (*.*)"),
        ) if self.window else None
        if not path:
            return self.snapshot()

        selected = path[0] if isinstance(path, (list, tuple)) else path
        try:
            content = read_text_file_with_fallbacks(selected)
            new_lines = build_registry_key_entries_from_reg_text(content)
            if new_lines:
                self.state["registry_enabled"] = True
                self.state["registry_keys"] = merge_ini_line_sets("", new_lines, True)
                self.status = f"Imported registry entries from {Path(selected).name}"
        except Exception as exc:
            self.status = f"Registry import failed: {exc}"
        return self.snapshot()

    def validate(self):
        project = self.current_project()
        launcher_path = self.refresh_generator_status()
        items = build_validation_items(project, launcher_path)
        title, report, status = render_validation_report(items)
        if status == "ok":
            self.status = "Validation passed."
        elif status == "warning":
            self.status = "Validation completed with warnings."
        else:
            self.status = "Validation found issues that need to be fixed."
        return {
            "title": title,
            "report": report,
            "status": status,
            "items": [{"label": item.label, "level": item.level, "message": item.detail} for item in items],
            "snapshot": self.snapshot(),
        }

    def create_project(self):
        project = self.current_project()
        errors = validate_project(project)
        if errors:
            return {"ok": False, "message": "\n".join(errors), "snapshot": self.snapshot()}

        launcher_path = self.refresh_generator_status()
        if launcher_path is None:
            return {
                "ok": False,
                "missingGenerator": True,
                "message": "PortableApps.com Launcher Generator is required.",
                "downloadUrl": PORTABLEAPPS_DEVELOPMENT_DOWNLOADS_URL,
                "snapshot": self.snapshot(),
            }

        try:
            project_root = create_launcher_project(project)
        except Exception as exc:
            return {"ok": False, "message": str(exc), "snapshot": self.snapshot()}

        try:
            result = subprocess.run(
                [str(launcher_path), str(project_root)],
                cwd=str(launcher_path.parent),
                check=False,
                capture_output=True,
                text=True,
                timeout=300,
            )
        except Exception as exc:
            return {"ok": False, "message": str(exc), "snapshot": self.snapshot()}

        portable_exe = project_root / f"{project.portable_name}.exe"
        if portable_exe.exists():
            open_folder_in_explorer(project_root)
            self.status = f"Created {portable_exe}"
            return {
                "ok": True,
                "message": f"Created {portable_exe}",
                "path": str(portable_exe),
                "snapshot": self.snapshot(),
            }

        details = (result.stderr or result.stdout or "").strip()
        return {
            "ok": False,
            "message": details[:1200] or "Generator ran but the EXE was not found.",
            "snapshot": self.snapshot(),
        }

    def open_generator_download(self):
        webbrowser.open(PORTABLEAPPS_DEVELOPMENT_DOWNLOADS_URL)
        return True


def run():
    backend = PortableAppsWebBackend()
    index_path = asset_path("web/index.html")
    window = webview.create_window(
        "PortableApps.com Launcher Maker",
        url=index_path.as_uri(),
        js_api=backend,
        width=1280,
        height=860,
        min_size=(980, 700),
    )
    backend.set_window(window)
    webview.start(debug=False, http_server=True)
