import html
import os
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import tkinter as tk
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageDraw, ImageOps, ImageTk


INVALID_FILENAME_CHARS = '<>:"/\\|?*'
PORTABLEAPPS_PNG_ICON_SIZES = (16, 32, 75, 128, 256)
PORTABLEAPPS_ICO_ICON_SIZES = (16, 32, 48, 256)
ICON_PREVIEW_DISPLAY_SIZES = (
    (256, 96),
    (128, 72),
    (75, 56),
    (32, 32),
    (16, 16),
)
DEFAULT_FALLBACK_ICON = "default_portable_icon.png"
SOFTWARE_LOGO = "software_logo.png"
SOFTWARE_ICON = "software_icon.ico"
SOFTWARE_ICON_PNG = "software_icon_256.png"
DEFAULT_SPLASH_ASSET = "default_splash.png"
PORTABLEAPPS_DEVELOPMENT_DOWNLOADS_URL = "https://portableapps.com/development"
REGISTRY_ROOT_ALIASES = {
    "HKEY_CLASSES_ROOT": "HKCR",
    "HKEY_CURRENT_USER": "HKCU",
    "HKEY_LOCAL_MACHINE": "HKLM",
    "HKEY_USERS": "HKU",
    "HKEY_CURRENT_CONFIG": "HKCC",
}
HELP_IMAGE_FILENAMES = (
    "Favicon.ico",
    "Donation_Button.png",
    "Help_Logo_Top.png",
    "Help_Background_Footer.png",
    "Help_Background_Header.png",
)
LAUNCHER_TEMPLATE_FILENAMES = (
    "Splash.jpg",
)
TEMPLATE_ASSET_SPECS = (
    (DEFAULT_SPLASH_ASSET, "Default Splash", (("Image files", "*.png;*.jpg;*.jpeg"), ("All files", "*.*"))),
)


@dataclass
class LauncherProject:
    app_name: str
    package_name: str
    publisher: str
    homepage: str
    category: str
    description: str
    version: str
    display_version: str
    app_exe: str
    output_dir: str
    trademarks: str = ""
    language: str = "Multilingual"
    donate: str = ""
    install_type: str = ""
    command_line: str = ""
    working_directory: str = "%PAL:AppDir%\\{app_name}"
    wait_for_program: bool = True
    close_exe: str = ""
    wait_for_other_instances: bool = True
    min_os: str = ""
    max_os: str = ""
    run_as_admin: str = ""
    refresh_shell_icons: str = ""
    hide_command_line_window: bool = False
    no_spaces_in_path: bool = False
    supports_unc: str = ""
    activate_java: str = ""
    activate_xml: bool = False
    live_mode_copy_app: bool = False
    live_mode_copy_data: bool = False
    files_move: str = ""
    directories_move: str = ""
    installer_close_exe: str = ""
    installer_close_name: str = ""
    include_installer_source: bool = False
    remove_app_directory: bool = False
    remove_data_directory: bool = False
    remove_other_directory: bool = False
    optional_components_enabled: bool = False
    main_section_title: str = ""
    main_section_description: str = ""
    optional_section_title: str = ""
    optional_section_description: str = ""
    optional_section_selected_install_type: str = ""
    optional_section_not_selected_install_type: str = ""
    optional_section_preselected: str = ""
    installer_languages: str = ""
    preserve_directories: str = ""
    remove_directories: str = ""
    preserve_files: str = ""
    remove_files: str = ""
    copy_app_files: bool = True
    icon_source: str = ""
    icon_index: int = 0
    registry_enabled: bool = False
    registry_keys: str = ""
    registry_cleanup_if_empty: str = ""
    registry_cleanup_force: str = ""
    license_shareable: bool = True
    license_open_source: bool = False
    license_freeware: bool = True
    license_commercial_use: bool = True
    license_eula_version: str = ""
    special_plugins: str = "NONE"
    dependency_uses_ghostscript: str = "no"
    dependency_uses_java: str = "no"
    dependency_uses_dotnet_version: str = ""
    dependency_requires_64bit_os: str = "no"
    dependency_requires_portable_app: str = ""
    dependency_requires_admin: str = "no"
    control_icons: str = "1"
    control_start: str = ""
    control_extract_icon: str = ""
    control_extract_name: str = ""
    control_base_app_id: str = ""
    control_base_app_id_64: str = ""
    control_base_app_id_arm64: str = ""
    control_exit_exe: str = ""
    control_exit_parameters: str = ""
    association_file_types: str = ""
    association_file_type_command_line: str = ""
    association_file_type_command_line_extension: str = ""
    association_protocols: str = ""
    association_protocol_command_line: str = ""
    association_protocol_command_line_protocol: str = ""
    association_send_to: bool = False
    association_send_to_command_line: str = ""
    association_shell: bool = False
    association_shell_command: str = ""
    file_type_icons: str = ""

    @property
    def portable_name(self) -> str:
        return f"{self.package_name}Portable"

    @property
    def app_exe_name(self) -> str:
        return Path(self.app_exe).name


@dataclass
class ValidationItem:
    level: str
    label: str
    detail: str


def clean_identifier(value: str, fallback: str = "MyApp") -> str:
    words = re.findall(r"[A-Za-z0-9]+", value or "")
    if not words:
        return fallback
    cleaned = "".join(word[:1].upper() + word[1:] for word in words)
    return cleaned or fallback


def clean_display_name(value: str, fallback: str = "My App") -> str:
    cleaned = "".join(ch for ch in (value or "").strip() if ch not in INVALID_FILENAME_CHARS)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned or fallback


def bool_to_ini(value: bool) -> str:
    return "true" if value else "false"


def clean_ini_lines(text: str) -> list[str]:
    return [line.strip() for line in (text or "").splitlines() if line.strip()]


def has_ini_lines(text: str) -> bool:
    return bool(clean_ini_lines(text))


def normalize_registry_path(path: str) -> str:
    candidate = (path or "").strip().lstrip("\\")
    if not candidate:
        return ""
    for full_name, alias in REGISTRY_ROOT_ALIASES.items():
        if candidate == full_name:
            return alias
        prefix = full_name + "\\"
        if candidate.startswith(prefix):
            return alias + "\\" + candidate[len(prefix):]
    return candidate


def registry_entry_name_for_key(registry_path: str, used_names: set[str]) -> str:
    parts = [part for part in normalize_registry_path(registry_path).split("\\") if part]
    raw_candidates: list[str] = []
    if len(parts) >= 3:
        raw_candidates.append("_".join(parts[-2:]))
    if len(parts) >= 2:
        raw_candidates.append(parts[-1])
    if len(parts) >= 2:
        raw_candidates.append("_".join(parts[1:]))
    raw_candidates.append(parts[-1] if parts else "registry_key")

    base = ""
    for candidate in raw_candidates:
        slug = re.sub(r"[^A-Za-z0-9]+", "_", candidate).strip("_").lower()
        if slug:
            base = slug
            break
    if not base:
        base = "registry_key"
    if base not in used_names:
        used_names.add(base)
        return base

    index = 2
    while f"{base}_{index}" in used_names:
        index += 1
    unique_name = f"{base}_{index}"
    used_names.add(unique_name)
    return unique_name


def parse_registry_paths_from_reg_text(text: str) -> list[str]:
    keys: list[str] = []
    seen: set[str] = set()
    for raw_line in (text or "").splitlines():
        line = raw_line.strip().lstrip("\ufeff")
        if not line or line.startswith(";"):
            continue
        match = re.match(r"^\[(?P<delete>-)?(?P<path>[^\]]+)\]$", line)
        if not match or match.group("delete"):
            continue
        normalized = normalize_registry_path(match.group("path"))
        if normalized and normalized not in seen:
            seen.add(normalized)
            keys.append(normalized)
    return keys


def build_registry_key_entries_from_reg_text(text: str) -> list[str]:
    used_names: set[str] = set()
    return [
        f"{registry_entry_name_for_key(registry_path, used_names)}={registry_path}"
        for registry_path in parse_registry_paths_from_reg_text(text)
    ]


def merge_ini_line_sets(existing_text: str, new_lines: list[str], replace_existing: bool) -> str:
    if replace_existing:
        return "\n".join(new_lines)
    combined: list[str] = []
    seen: set[str] = set()
    for line in clean_ini_lines(existing_text):
        if line not in seen:
            seen.add(line)
            combined.append(line)
    for line in new_lines:
        if line not in seen:
            seen.add(line)
            combined.append(line)
    return "\n".join(combined)


def detect_app_name_from_exe(path: str) -> str:
    stem = Path(path or "").stem
    stem = re.sub(r"[_-]+", " ", stem).strip()
    return clean_display_name(stem, fallback="My App")


def portableapps_launcher_candidates() -> list[Path]:
    candidates = []
    app_dir = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[1]
    drive_root = Path(app_dir.anchor)
    if str(drive_root):
        candidates.append(drive_root / "PortableApps" / "PortableApps.comLauncher" / "PortableApps.comLauncherGenerator.exe")
    for base in (app_dir, *app_dir.parents):
        candidates.extend(
            [
                base / "PortableApps" / "PortableApps.comLauncher" / "PortableApps.comLauncherGenerator.exe",
                base / "PortableApps" / "PortableApps.comLauncher" / "PortableApps.comLauncher.exe",
                base / "PortableApps" / "PortableApps.comLauncher" / "PortableApps.com Launcher.exe",
            ]
        )
    roots = [
        os.environ.get("ProgramFiles"),
        os.environ.get("ProgramFiles(x86)"),
        os.environ.get("LOCALAPPDATA"),
        os.environ.get("USERPROFILE"),
        os.getcwd(),
    ]
    names = [
        r"PortableApps.comLauncher\PortableApps.comLauncher.exe",
        r"PortableApps.comLauncher\PortableApps.comLauncherGenerator.exe",
        r"PortableApps.com Launcher\PortableApps.comLauncher.exe",
        r"PortableApps.com Launcher\PortableApps.com Launcher.exe",
        r"PortableApps\PortableApps.comLauncher\PortableApps.comLauncher.exe",
        r"PortableApps\PortableApps.comLauncher\PortableApps.comLauncherGenerator.exe",
    ]
    for root in roots:
        if not root:
            continue
        for name in names:
            candidates.append(Path(root) / name)
    candidates.append(Path(r"C:\PortableApps\PortableApps.comLauncher\PortableApps.comLauncherGenerator.exe"))
    candidates.append(Path(r"C:\PortableApps\PortableApps.comLauncher\PortableApps.comLauncher.exe"))
    return candidates


def find_portableapps_launcher() -> Path | None:
    for candidate in portableapps_launcher_candidates():
        if candidate.exists():
            return candidate
    return None


def app_base_path() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parents[1]


def default_portableapps_output_dir() -> str:
    base_dir = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[1]
    drive_root = Path(base_dir.anchor)
    if str(drive_root):
        return str(drive_root / "PortableApps")
    return str(Path.cwd() / "PortableApps")


def asset_path(name: str) -> Path:
    return app_base_path() / "app" / "assets" / name


def help_template_path(*parts: str) -> Path:
    return app_base_path() / "app" / "help_template" / Path(*parts)


def help_image_asset_path(name: str) -> Path:
    return help_template_path("Images", name)


def load_help_html_template() -> str:
    template_path = help_template_path("help.html")
    if template_path.exists():
        return template_path.read_text(encoding="utf-8")
    raise FileNotFoundError(f"Help template not found: {template_path}")


def launcher_template_asset_path(name: str) -> Path:
    return help_template_path("Launcher", name)


def splash_asset_path() -> Path:
    return asset_path(DEFAULT_SPLASH_ASSET)


def open_folder_in_explorer(path: Path) -> None:
    target = Path(path)
    try:
        if hasattr(os, "startfile"):
            os.startfile(str(target))
            return
    except OSError:
        pass

    try:
        subprocess.Popen(["explorer", str(target)])
    except OSError:
        pass


def resolve_project_tokens(value: str, project: LauncherProject) -> str:
    resolved = value or ""
    return (
        resolved.replace("{app_name}", project.package_name)
        .replace("{package_name}", project.package_name)
        .replace("{portable_name}", project.portable_name)
    )


def validate_project(project: LauncherProject) -> list[str]:
    errors: list[str] = []
    if not project.app_name.strip():
        errors.append("App Name is required.")
    if not project.package_name.strip():
        errors.append("Package ID is required.")
    if not project.app_exe.strip():
        errors.append("Application EXE is required.")
    else:
        app_exe = Path(project.app_exe)
        if not app_exe.exists():
            errors.append(f"Application EXE was not found: {app_exe}")
        elif not app_exe.is_file():
            errors.append("Application EXE must point to a file.")
        elif app_exe.suffix.lower() != ".exe":
            errors.append("Application EXE must be a .exe file.")
    if not project.output_dir.strip():
        errors.append("Output Folder is required.")
    if project.icon_source.strip() and not Path(project.icon_source).exists():
        errors.append(f"Icon Override was not found: {project.icon_source}")
    if project.version.strip() and not re.fullmatch(r"\d+\.\d+\.\d+\.\d+", project.version.strip()):
        errors.append("Package Version must use the PortableApps format: major.minor.patch.build")
    if (
        not project.registry_enabled
        and (
            has_ini_lines(project.registry_keys)
            or has_ini_lines(project.registry_cleanup_if_empty)
            or has_ini_lines(project.registry_cleanup_force)
        )
    ):
        errors.append("Enable Registry in the Registry tab before using RegistryKeys or cleanup sections.")
    return errors


def validate_ini_mapping_lines(text: str, section_name: str) -> list[str]:
    issues: list[str] = []
    for index, line in enumerate(clean_ini_lines(text), start=1):
        if "=" not in line:
            issues.append(f"{section_name} line {index} must use key=value format.")
    return issues


def build_validation_items(project: LauncherProject, launcher_path: Path | None = None) -> list[ValidationItem]:
    items: list[ValidationItem] = []
    errors = validate_project(project)
    error_map = {
        "App Name is required.": "App Name",
        "Package ID is required.": "Package ID",
        "Application EXE is required.": "Application EXE",
        "Application EXE must point to a file.": "Application EXE",
        "Application EXE must be a .exe file.": "Application EXE",
        "Output Folder is required.": "Output Folder",
        "Package Version must use the PortableApps format: major.minor.patch.build": "Package Version",
        "Enable Registry in the Registry tab before using RegistryKeys or cleanup sections.": "Registry",
    }
    mapped_errors: set[str] = set()
    for error in errors:
        label = error_map.get(error)
        if label is None and error.startswith("Application EXE was not found:"):
            label = "Application EXE"
        elif label is None and error.startswith("Icon Override was not found:"):
            label = "Icon Override"
        if label is not None:
            items.append(ValidationItem("error", label, error))
            mapped_errors.add(error)

    app_name = project.app_name.strip()
    if app_name and "App Name is required." not in mapped_errors:
        items.append(ValidationItem("ok", "App Name", app_name))

    package_name = project.package_name.strip()
    if package_name and "Package ID is required." not in mapped_errors:
        items.append(ValidationItem("ok", "Package ID", package_name))

    app_exe_value = project.app_exe.strip()
    app_exe_path = Path(app_exe_value) if app_exe_value else None
    if (
        app_exe_path is not None
        and app_exe_path.exists()
        and app_exe_path.is_file()
        and app_exe_path.suffix.lower() == ".exe"
    ):
        items.append(ValidationItem("ok", "Application EXE", str(app_exe_path)))

    output_dir = project.output_dir.strip()
    if output_dir:
        output_path = Path(output_dir)
        items.append(ValidationItem("ok", "Output Folder", str(output_path / project.portable_name)))

    icon_override = project.icon_source.strip()
    if icon_override:
        icon_path = Path(icon_override)
        if icon_path.exists():
            items.append(ValidationItem("ok", "Icon Override", str(icon_path)))

    if project.version.strip() and "Package Version must use the PortableApps format: major.minor.patch.build" not in mapped_errors:
        items.append(ValidationItem("ok", "Package Version", project.version.strip()))

    registry_has_sections = any(
        has_ini_lines(value)
        for value in (
            project.registry_keys,
            project.registry_cleanup_if_empty,
            project.registry_cleanup_force,
        )
    )
    if project.registry_enabled:
        detail = "Registry handling is enabled."
        if registry_has_sections:
            detail += " Registry sections will be written."
        items.append(ValidationItem("ok", "Registry", detail))
    elif registry_has_sections:
        items.append(ValidationItem("error", "Registry", "Registry entries exist but registry handling is disabled."))
    else:
        items.append(ValidationItem("ok", "Registry", "Registry handling is disabled and no registry sections are set."))

    for section_name, text in (
        ("FilesMove", project.files_move),
        ("DirectoriesMove", project.directories_move),
        ("RegistryKeys", project.registry_keys),
        ("RegistryCleanupIfEmpty", project.registry_cleanup_if_empty),
        ("RegistryCleanupForce", project.registry_cleanup_force),
        ("Installer Languages", project.installer_languages),
        ("DirectoriesToPreserve", project.preserve_directories),
        ("DirectoriesToRemove", project.remove_directories),
        ("FilesToPreserve", project.preserve_files),
        ("FilesToRemove", project.remove_files),
    ):
        for issue in validate_ini_mapping_lines(text, section_name):
            items.append(ValidationItem("warning", section_name, issue))

    resolved_launcher_path = launcher_path if launcher_path is not None else find_portableapps_launcher()
    if resolved_launcher_path is None or not Path(resolved_launcher_path).exists():
        items.append(
            ValidationItem(
                "warning",
                "PortableApps Generator",
                "PortableApps.com Launcher Generator was not found. Create Project + EXE will not be able to build the launcher.",
            )
        )
    else:
        items.append(ValidationItem("ok", "PortableApps Generator", str(resolved_launcher_path)))

    return items


def render_validation_report(items: list[ValidationItem]) -> tuple[str, str, str]:
    errors = [item for item in items if item.level == "error"]
    warnings = [item for item in items if item.level == "warning"]
    oks = [item for item in items if item.level == "ok"]
    if errors:
        title = "Validation Found Errors"
        status = "error"
    elif warnings:
        title = "Validation Completed With Warnings"
        status = "warning"
    else:
        title = "Validation Passed"
        status = "ok"

    sections: list[str] = []
    if errors:
        sections.append("Errors")
        sections.extend(f"- {item.label}: {item.detail}" for item in errors)
    if warnings:
        if sections:
            sections.append("")
        sections.append("Warnings")
        sections.extend(f"- {item.label}: {item.detail}" for item in warnings)
    if oks:
        if sections:
            sections.append("")
        sections.append("Checks")
        sections.extend(f"- {item.label}: {item.detail}" for item in oks)
    return title, "\n".join(sections), status


def build_appinfo_ini(project: LauncherProject) -> str:
    category = project.category or "Utilities"
    homepage = project.homepage or "https://portableapps.com/"
    publisher = project.publisher or project.app_name
    description = project.description or f"{project.app_name} portable launcher"
    version = project.version or "1.0.0.0"
    display_version = project.display_version or version
    control_start = project.control_start.strip() or f"{project.portable_name}.exe"
    control_icons = project.control_icons.strip() or "1"
    lines = [
        "[Format]",
        "Type=PortableApps.comFormat",
        "Version=3.9",
        "",
        "[Details]",
        f"Name={project.app_name} Portable",
        f"AppID={project.portable_name}",
        f"Publisher={publisher}",
        f"Trademarks={project.trademarks.strip()}",
        f"Homepage={homepage}",
        f"Category={category}",
        f"Description={description}",
        f"Language={project.language.strip() or 'Multilingual'}",
        f"Donate={project.donate.strip()}",
        f"InstallType={project.install_type.strip()}",
        "",
        "[License]",
        f"Shareable={bool_to_ini(project.license_shareable)}",
        f"OpenSource={bool_to_ini(project.license_open_source)}",
        f"Freeware={bool_to_ini(project.license_freeware)}",
        f"CommercialUse={bool_to_ini(project.license_commercial_use)}",
        f"EULAVersion={project.license_eula_version.strip()}",
        "",
        "[Version]",
        f"PackageVersion={version}",
        f"DisplayVersion={display_version}",
        "",
        "[SpecialPaths]",
        f"Plugins={project.special_plugins.strip()}",
        "",
        "[Dependencies]",
        f"UsesGhostscript={project.dependency_uses_ghostscript.strip()}",
        f"UsesJava={project.dependency_uses_java.strip()}",
        f"UsesDotNetVersion={project.dependency_uses_dotnet_version.strip()}",
        f"Requires64bitOS={project.dependency_requires_64bit_os.strip()}",
        f"RequiresPortableApp={project.dependency_requires_portable_app.strip()}",
        f"RequiresAdmin={project.dependency_requires_admin.strip()}",
        "",
        "[Control]",
        f"Icons={control_icons}",
        f"Start={control_start}",
    ]
    optional_control_values = (
        ("ExtractIcon", project.control_extract_icon),
        ("ExtractName", project.control_extract_name),
        ("BaseAppID", project.control_base_app_id),
        ("BaseAppID64", project.control_base_app_id_64),
        ("BaseAppIDARM64", project.control_base_app_id_arm64),
        ("ExitEXE", project.control_exit_exe),
        ("ExitParameters", project.control_exit_parameters),
    )
    for key, value in optional_control_values:
        if value.strip():
            lines.append(f"{key}={value.strip()}")

    association_values = {
        "FileTypes": project.association_file_types.strip(),
        "FileTypeCommandLine": project.association_file_type_command_line.strip(),
        "FileTypeCommandLine-extension": project.association_file_type_command_line_extension.strip(),
        "Protocols": project.association_protocols.strip(),
        "ProtocolCommandLine": project.association_protocol_command_line.strip(),
        "ProtocolCommandLine-protocol": project.association_protocol_command_line_protocol.strip(),
        "SendToCommandLine": project.association_send_to_command_line.strip(),
        "ShellCommand": project.association_shell_command.strip(),
    }
    has_associations = any(association_values.values()) or project.association_send_to or project.association_shell
    if has_associations:
        lines.extend(
            [
                "",
                "[Associations]",
                f"FileTypes={association_values['FileTypes']}",
                f"FileTypeCommandLine={association_values['FileTypeCommandLine']}",
                f"FileTypeCommandLine-extension={association_values['FileTypeCommandLine-extension']}",
                f"Protocols={association_values['Protocols']}",
                f"ProtocolCommandLine={association_values['ProtocolCommandLine']}",
                f"ProtocolCommandLine-protocol={association_values['ProtocolCommandLine-protocol']}",
                f"SendTo={bool_to_ini(project.association_send_to)}",
                f"SendToCommandLine={association_values['SendToCommandLine']}",
                f"Shell={bool_to_ini(project.association_shell)}",
                f"ShellCommand={association_values['ShellCommand']}",
            ]
        )

    file_type_icon_lines = clean_ini_lines(project.file_type_icons)
    if file_type_icon_lines:
        lines.extend(
            [
                "",
                "[FileTypeIcons]",
            ]
        )
        lines.extend(file_type_icon_lines)

    lines.append("")
    return "\n".join(lines)


def build_launcher_ini(project: LauncherProject) -> str:
    working_directory = resolve_project_tokens(project.working_directory, project)
    lines = [
        "[Launch]",
        f"ProgramExecutable={project.package_name}\\{project.app_exe_name}",
        f"CommandLineArguments={project.command_line}",
        f"WorkingDirectory={working_directory}",
        f"WaitForProgram={bool_to_ini(project.wait_for_program)}",
        f"WaitForOtherInstances={bool_to_ini(project.wait_for_other_instances)}",
    ]
    if project.close_exe.strip():
        lines.append(f"CloseEXE={project.close_exe.strip()}")
    if project.min_os.strip():
        lines.append(f"MinOS={project.min_os.strip()}")
    if project.max_os.strip():
        lines.append(f"MaxOS={project.max_os.strip()}")
    if project.run_as_admin.strip():
        lines.append(f"RunAsAdmin={project.run_as_admin.strip()}")
    if project.refresh_shell_icons.strip():
        lines.append(f"RefreshShellIcons={project.refresh_shell_icons.strip()}")
    if project.hide_command_line_window:
        lines.append("HideCommandLineWindow=true")
    if project.no_spaces_in_path:
        lines.append("NoSpacesInPath=true")
    if project.supports_unc.strip():
        lines.append(f"SupportsUNC={project.supports_unc.strip()}")
    lines.extend(
        [
            "",
            "[Activate]",
            f"Registry={bool_to_ini(project.registry_enabled)}",
        ]
    )
    if project.activate_java.strip():
        lines.append(f"Java={project.activate_java.strip()}")
    if project.activate_xml:
        lines.append("XML=true")
    lines.append("")
    if project.live_mode_copy_app or project.live_mode_copy_data:
        lines.extend(
            [
                "[LiveMode]",
                f"CopyApp={bool_to_ini(project.live_mode_copy_app)}",
                f"CopyData={bool_to_ini(project.live_mode_copy_data)}",
                "",
            ]
        )
    if project.files_move.strip():
        lines.append("[FilesMove]")
        lines.extend(clean_ini_lines(project.files_move))
        lines.append("")
    if project.directories_move.strip():
        lines.append("[DirectoriesMove]")
        lines.extend(clean_ini_lines(project.directories_move))
        lines.append("")
    if project.registry_keys.strip():
        lines.append("[RegistryKeys]")
        lines.extend(clean_ini_lines(project.registry_keys))
        lines.append("")
    if project.registry_cleanup_if_empty.strip():
        lines.append("[RegistryCleanupIfEmpty]")
        lines.extend(clean_ini_lines(project.registry_cleanup_if_empty))
        lines.append("")
    if project.registry_cleanup_force.strip():
        lines.append("[RegistryCleanupForce]")
        lines.extend(clean_ini_lines(project.registry_cleanup_force))
        lines.append("")
    return "\n".join(lines)


def build_installer_ini(project: LauncherProject) -> str:
    lines = []

    if project.installer_close_exe.strip() or project.installer_close_name.strip():
        lines.extend(
            [
                "[CheckRunning]",
                f"CloseEXE={project.installer_close_exe.strip()}",
                f"CloseName={project.installer_close_name.strip()}",
                "",
            ]
        )

    if project.include_installer_source:
        lines.extend(
            [
                "[Source]",
                "IncludeInstallerSource=true",
                "",
            ]
        )

    if project.remove_app_directory or project.remove_data_directory or project.remove_other_directory:
        lines.append("[MainDirectories]")
        lines.append(f"RemoveAppDirectory={bool_to_ini(project.remove_app_directory)}")
        lines.append(f"RemoveDataDirectory={bool_to_ini(project.remove_data_directory)}")
        lines.append(f"RemoveOtherDirectory={bool_to_ini(project.remove_other_directory)}")
        lines.append("")

    has_optional_components = (
        project.optional_components_enabled
        or project.main_section_title.strip()
        or project.main_section_description.strip()
        or project.optional_section_title.strip()
        or project.optional_section_description.strip()
        or project.optional_section_selected_install_type.strip()
        or project.optional_section_not_selected_install_type.strip()
        or project.optional_section_preselected.strip()
    )
    if has_optional_components:
        lines.extend(
            [
                "[OptionalComponents]",
                f"OptionalComponents={bool_to_ini(True)}",
            ]
        )
        optional_values = (
            ("MainSectionTitle", project.main_section_title),
            ("MainSectionDescription", project.main_section_description),
            ("OptionalSectionTitle", project.optional_section_title),
            ("OptionalSectionDescription", project.optional_section_description),
            ("OptionalSectionSelectedInstallType", project.optional_section_selected_install_type),
            ("OptionalSectionNotSelectedInstallType", project.optional_section_not_selected_install_type),
            ("OptionalSectionPreSelectedIfNonEnglishInstall", project.optional_section_preselected),
        )
        for key, value in optional_values:
            if value.strip():
                lines.append(f"{key}={value.strip()}")
        lines.append("")

    language_lines = clean_ini_lines(project.installer_languages)
    if language_lines:
        lines.append("[Languages]")
        lines.extend(language_lines)
        lines.append("")

    preserve_directory_lines = clean_ini_lines(project.preserve_directories)
    if preserve_directory_lines:
        lines.append("[DirectoriesToPreserve]")
        lines.extend(preserve_directory_lines)
        lines.append("")

    remove_directory_lines = clean_ini_lines(project.remove_directories)
    if remove_directory_lines:
        lines.append("[DirectoriesToRemove]")
        lines.extend(remove_directory_lines)
        lines.append("")

    preserve_file_lines = clean_ini_lines(project.preserve_files)
    if preserve_file_lines:
        lines.append("[FilesToPreserve]")
        lines.extend(preserve_file_lines)
        lines.append("")

    remove_file_lines = clean_ini_lines(project.remove_files)
    if remove_file_lines:
        lines.append("[FilesToRemove]")
        lines.extend(remove_file_lines)
        lines.append("")

    return "\n".join(lines).strip()


def build_readme(project: LauncherProject) -> str:
    return "\n".join(
        [
            f"{project.app_name} Portable launcher project",
            "",
            "Generated by PortableApps.com Launcher Maker.",
            "",
            "Next steps:",
            "1. Review App\\AppInfo\\appinfo.ini.",
            f"2. Review App\\AppInfo\\Launcher\\{project.portable_name}.ini.",
            "3. Build the launcher with the PortableApps.com Launcher.",
            "4. Package the folder with the PortableApps.com Installer if desired.",
            "",
            "This tool creates the PortableApps.com-format project structure and launcher configuration.",
            "It does not include the official PortableApps.com Launcher compiler.",
            "",
        ]
    )


def build_help_html(project: LauncherProject) -> str:
    return load_help_html_template().replace("**App Name**", html.escape(project.app_name))


def create_help_images(help_images_dir: Path, appinfo_dir: Path) -> None:
    icon_png = appinfo_dir / "appicon_128.png"
    shutil.copy2(icon_png, help_images_dir / "appicon_128.png")

    copied_all_template_assets = True
    for output_name in HELP_IMAGE_FILENAMES:
        source_path = help_image_asset_path(output_name)
        if source_path.exists():
            shutil.copy2(source_path, help_images_dir / output_name)
        else:
            copied_all_template_assets = False

    if copied_all_template_assets:
        return

    with Image.open(icon_png) as source_icon:
        source_icon.convert("RGBA").save(help_images_dir / "Favicon.ico", sizes=[(16, 16), (32, 32), (48, 48)])

    header = Image.new("RGBA", (840, 60), "#b31616")
    ImageDraw.Draw(header).rectangle((0, 44, 840, 60), fill="#8c1010")
    header.save(help_images_dir / "Help_Background_Header.png")

    footer = Image.new("RGBA", (840, 16), "#8c1010")
    footer.save(help_images_dir / "Help_Background_Footer.png")

    logo = Image.new("RGBA", (320, 48), (255, 255, 255, 0))
    logo_draw = ImageDraw.Draw(logo)
    logo_draw.text((8, 12), "PortableApps.com", fill="#ffffff")
    logo.save(help_images_dir / "Help_Logo_Top.png")

    donation = Image.new("RGBA", (150, 32), "#f6c343")
    donation_draw = ImageDraw.Draw(donation)
    donation_draw.rounded_rectangle((0, 0, 149, 31), radius=6, fill="#f6c343", outline="#b88900")
    donation_draw.text((18, 9), "Make a Donation", fill="#5a4200")
    donation.save(help_images_dir / "Donation_Button.png")


def create_launcher_template_assets(launcher_dir: Path) -> None:
    splash_path = splash_asset_path()
    if splash_path.exists():
        with Image.open(splash_path) as splash_image:
            splash_image.convert("RGB").save(launcher_dir / "Splash.jpg", format="JPEG", quality=95)
        return

    for filename in LAUNCHER_TEMPLATE_FILENAMES:
        source_path = launcher_template_asset_path(filename)
        if source_path.exists():
            shutil.copy2(source_path, launcher_dir / filename)


def ensure_empty_or_create(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def is_relative_to_path(path: Path, potential_parent: Path) -> bool:
    try:
        path.resolve().relative_to(potential_parent.resolve())
        return True
    except ValueError:
        return False


def copy_application_folder(source_exe: Path, destination_dir: Path, project_root: Path | None = None) -> None:
    source_dir = source_exe.parent.resolve()
    resolved_destination = destination_dir.resolve()
    resolved_project_root = project_root.resolve() if project_root is not None else resolved_destination
    destination_dir.mkdir(parents=True, exist_ok=True)
    for item in source_dir.iterdir():
        resolved_item = item.resolve()
        if (
            is_relative_to_path(resolved_item, resolved_destination)
            or is_relative_to_path(resolved_item, resolved_project_root)
            or is_relative_to_path(resolved_destination, resolved_item)
            or is_relative_to_path(resolved_project_root, resolved_item)
        ):
            continue
        target = destination_dir / item.name
        if item.is_dir():
            shutil.copytree(item, target, dirs_exist_ok=True)
        elif item.is_file():
            shutil.copy2(item, target)


def make_fallback_icon(app_name: str) -> Image.Image:
    fallback_asset = asset_path(DEFAULT_FALLBACK_ICON)
    if fallback_asset.exists():
        try:
            with Image.open(fallback_asset) as icon:
                normalized = ImageOps.contain(icon.convert("RGBA"), (256, 256), Image.Resampling.LANCZOS)
                canvas = Image.new("RGBA", (256, 256), (0, 0, 0, 0))
                offset = ((256 - normalized.width) // 2, (256 - normalized.height) // 2)
                canvas.alpha_composite(normalized, offset)
                return canvas
        except OSError:
            pass

    image = Image.new("RGBA", (256, 256), "#216b52")
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((18, 18, 238, 238), radius=34, fill="#ffffff")
    draw.rounded_rectangle((34, 34, 222, 222), radius=24, fill="#216b52")
    initials = "".join(part[:1].upper() for part in re.findall(r"[A-Za-z0-9]+", app_name or "")[:2]) or "PA"
    bbox = draw.textbbox((0, 0), initials)
    x = (256 - (bbox[2] - bbox[0])) / 2
    y = (256 - (bbox[3] - bbox[1])) / 2 - 4
    draw.text((x, y), initials, fill="#ffffff")
    return image


def read_uint16(data: bytes, offset: int) -> int:
    return struct.unpack_from("<H", data, offset)[0]


def read_uint32(data: bytes, offset: int) -> int:
    return struct.unpack_from("<I", data, offset)[0]


def get_pe_sections_and_resource_offset(data: bytes):
    if data[:2] != b"MZ":
        raise ValueError("Not a PE executable.")
    pe_offset = read_uint32(data, 0x3C)
    if data[pe_offset:pe_offset + 4] != b"PE\0\0":
        raise ValueError("Invalid PE header.")

    file_header = pe_offset + 4
    section_count = read_uint16(data, file_header + 2)
    optional_header_size = read_uint16(data, file_header + 16)
    optional_header = file_header + 20
    magic = read_uint16(data, optional_header)
    data_directory = optional_header + (112 if magic == 0x20B else 96)
    resource_rva = read_uint32(data, data_directory + 8 * 2)
    section_table = optional_header + optional_header_size

    sections = []
    for index in range(section_count):
        section = section_table + index * 40
        virtual_size = read_uint32(data, section + 8)
        virtual_address = read_uint32(data, section + 12)
        raw_size = read_uint32(data, section + 16)
        raw_pointer = read_uint32(data, section + 20)
        sections.append((virtual_address, max(virtual_size, raw_size), raw_pointer))

    def rva_to_offset(rva: int) -> int:
        for virtual_address, size, raw_pointer in sections:
            if virtual_address <= rva < virtual_address + size:
                return raw_pointer + (rva - virtual_address)
        raise ValueError(f"Could not map RVA {rva}.")

    return sections, rva_to_offset(resource_rva), rva_to_offset


def parse_resource_name(data: bytes, resource_base: int, value: int):
    if value & 0x80000000:
        offset = resource_base + (value & 0x7FFFFFFF)
        length = read_uint16(data, offset)
        raw = data[offset + 2:offset + 2 + length * 2]
        return raw.decode("utf-16le", errors="replace")
    return value & 0xFFFF


def parse_resource_directory(data: bytes, resource_base: int, directory_offset: int):
    named_count = read_uint16(data, directory_offset + 12)
    id_count = read_uint16(data, directory_offset + 14)
    entries = []
    entry_offset = directory_offset + 16
    for index in range(named_count + id_count):
        current = entry_offset + index * 8
        raw_name = read_uint32(data, current)
        raw_target = read_uint32(data, current + 4)
        entries.append(
            (
                parse_resource_name(data, resource_base, raw_name),
                bool(raw_target & 0x80000000),
                resource_base + (raw_target & 0x7FFFFFFF),
            )
        )
    return entries


def collect_resource_data(data: bytes, resource_base: int, rva_to_offset, resource_type: int):
    root_entries = parse_resource_directory(data, resource_base, resource_base)
    type_entry = next((entry for entry in root_entries if entry[0] == resource_type and entry[1]), None)
    if type_entry is None:
        return {}

    resources = {}
    for resource_id, has_child, id_offset in parse_resource_directory(data, resource_base, type_entry[2]):
        if not has_child:
            continue
        language_entries = parse_resource_directory(data, resource_base, id_offset)
        if not language_entries:
            continue
        _language, language_has_child, data_entry_offset = language_entries[0]
        if language_has_child:
            continue
        data_rva = read_uint32(data, data_entry_offset)
        data_size = read_uint32(data, data_entry_offset + 4)
        file_offset = rva_to_offset(data_rva)
        resources[resource_id] = data[file_offset:file_offset + data_size]
    return resources


def extract_icon_group_from_exe(source_exe: Path, destination_ico: Path, icon_index: int = 0) -> bool:
    try:
        data = source_exe.read_bytes()
        _sections, resource_base, rva_to_offset = get_pe_sections_and_resource_offset(data)
        groups = collect_resource_data(data, resource_base, rva_to_offset, 14)
        icons = collect_resource_data(data, resource_base, rva_to_offset, 3)
    except (OSError, struct.error, ValueError):
        return False

    if not groups or not icons:
        return False

    group_keys = sorted(groups, key=lambda value: str(value).casefold())
    selected_key = group_keys[min(max(0, int(icon_index or 0)), len(group_keys) - 1)]
    group_data = groups[selected_key]
    if len(group_data) < 6:
        return False

    reserved, icon_type, icon_count = struct.unpack_from("<HHH", group_data, 0)
    if icon_type != 1 or icon_count == 0:
        return False

    icon_entries = []
    image_chunks = []
    offset = 6
    image_offset = 6 + icon_count * 16
    for _index in range(icon_count):
        if offset + 14 > len(group_data):
            return False
        width, height, color_count, reserved_byte, planes, bit_count, image_size, icon_id = struct.unpack_from(
            "<BBBBHHIH",
            group_data,
            offset,
        )
        image = icons.get(icon_id)
        if not image:
            offset += 14
            continue
        icon_entries.append(
            struct.pack(
                "<BBBBHHII",
                width,
                height,
                color_count,
                reserved_byte,
                planes,
                bit_count,
                len(image),
                image_offset,
            )
        )
        image_chunks.append(image)
        image_offset += len(image)
        offset += 14

    if not image_chunks:
        return False

    destination_ico.write_bytes(
        struct.pack("<HHH", reserved, icon_type, len(image_chunks)) + b"".join(icon_entries) + b"".join(image_chunks)
    )
    return True


def extract_associated_icon(source_exe: Path, destination_ico: Path) -> bool:
    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        (
            "Add-Type -AssemblyName System.Drawing; "
            "$icon = [System.Drawing.Icon]::ExtractAssociatedIcon($args[0]); "
            "if ($null -eq $icon) { exit 2 }; "
            "$stream = [System.IO.File]::Open($args[1], [System.IO.FileMode]::Create); "
            "try { $icon.Save($stream) } finally { $stream.Dispose(); $icon.Dispose() }"
        ),
        str(source_exe),
        str(destination_ico),
    ]
    try:
        result = subprocess.run(command, check=False, capture_output=True, text=True, timeout=20)
    except (OSError, subprocess.SubprocessError):
        return False
    return result.returncode == 0 and destination_ico.exists() and destination_ico.stat().st_size > 0


def extract_embedded_icon(source_exe: Path, destination_ico: Path, icon_index: int = 0) -> bool:
    if extract_icon_group_from_exe(source_exe, destination_ico, icon_index):
        return True

    icon_index = max(0, int(icon_index or 0))
    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        (
            "Add-Type -AssemblyName System.Drawing; "
            "$memberDefinition = '"
            "[DllImport(\"Shell32.dll\", CharSet=CharSet.Unicode)] "
            "public static extern int SHDefExtractIcon(string file, int index, uint flags, IntPtr[] large, IntPtr[] small, uint size); "
            "[DllImport(\"User32.dll\")] "
            "public static extern bool DestroyIcon(IntPtr handle);"
            "'; "
            "Add-Type -Namespace Win32 -Name NativeIcons -MemberDefinition $memberDefinition; "
            "$index = [int]$args[2]; "
            "$large = New-Object IntPtr[] 1; "
            "$small = New-Object IntPtr[] 1; "
            "$size = 0; "
            "$hr = [Win32.NativeIcons]::SHDefExtractIcon($args[0], $index, 0, $large, $small, $size); "
            "if ($hr -ne 0 -or $large[0] -eq [IntPtr]::Zero) { "
            "  $hr = [Win32.NativeIcons]::SHDefExtractIcon($args[0], 0, 0, $large, $small, $size); "
            "} "
            "$handle = $large[0]; "
            "if ($handle -eq [IntPtr]::Zero) { $handle = $small[0] }; "
            "if ($handle -eq [IntPtr]::Zero) { exit 3 }; "
            "$icon = [System.Drawing.Icon]::FromHandle($handle); "
            "$stream = [System.IO.File]::Open($args[1], [System.IO.FileMode]::Create); "
            "try { $icon.Save($stream) } finally { $stream.Dispose(); $icon.Dispose(); [void][Win32.NativeIcons]::DestroyIcon($handle); if ($small[0] -ne [IntPtr]::Zero) { [void][Win32.NativeIcons]::DestroyIcon($small[0]) } }"
        ),
        str(source_exe),
        str(destination_ico),
        str(icon_index),
    ]
    try:
        result = subprocess.run(command, check=False, capture_output=True, text=True, timeout=20)
    except (OSError, subprocess.SubprocessError):
        return False
    if result.returncode == 0 and destination_ico.exists() and destination_ico.stat().st_size > 0:
        return True
    return extract_associated_icon(source_exe, destination_ico)


def load_icon_image(icon_source: Path, app_name: str) -> Image.Image:
    if icon_source.exists():
        try:
            with Image.open(icon_source) as icon:
                best_frame = None
                best_area = -1
                frame_count = getattr(icon, "n_frames", 1)
                for frame_index in range(frame_count):
                    try:
                        icon.seek(frame_index)
                    except EOFError:
                        break
                    frame = icon.convert("RGBA")
                    area = frame.width * frame.height
                    if area > best_area:
                        best_frame = frame.copy()
                        best_area = area
                if best_frame is not None:
                    return best_frame
        except OSError:
            pass
    return make_fallback_icon(app_name)


def save_portableapps_icon_set(source_icon: Path, appinfo_dir: Path, app_name: str) -> None:
    icon = load_icon_image(source_icon, app_name)
    appinfo_dir.mkdir(parents=True, exist_ok=True)
    ico_sizes = [(size, size) for size in PORTABLEAPPS_ICO_ICON_SIZES]
    icon.save(appinfo_dir / "appicon.ico", sizes=ico_sizes)
    for size in PORTABLEAPPS_PNG_ICON_SIZES:
        resized = icon.resize((size, size), Image.Resampling.LANCZOS)
        resized.save(appinfo_dir / f"appicon_{size}.png")


def create_portableapps_icons(project: LauncherProject, appinfo_dir: Path, app_exe: Path) -> None:
    if project.icon_source.strip():
        save_portableapps_icon_set(Path(project.icon_source), appinfo_dir, project.app_name)
        return

    extracted_icon = appinfo_dir / ".extracted-appicon.ico"
    try:
        extract_embedded_icon(app_exe, extracted_icon, project.icon_index)
        save_portableapps_icon_set(extracted_icon, appinfo_dir, project.app_name)
    finally:
        try:
            extracted_icon.unlink()
        except OSError:
            pass


def create_launcher_project(project: LauncherProject) -> Path:
    if not project.app_name.strip():
        raise ValueError("App name is required.")
    if not project.package_name.strip():
        raise ValueError("Package name is required.")
    if not project.app_exe.strip():
        raise ValueError("Application executable is required.")

    app_exe = Path(project.app_exe)
    if not app_exe.exists():
        raise FileNotFoundError(f"Application executable not found: {app_exe}")

    output_root = Path(project.output_dir)
    if not str(output_root).strip():
        raise ValueError("Output folder is required.")

    project_root = output_root / project.portable_name
    app_dir = project_root / "App" / project.package_name
    appinfo_dir = project_root / "App" / "AppInfo"
    launcher_dir = appinfo_dir / "Launcher"
    data_dir = project_root / "Data" / "settings"
    help_images_dir = project_root / "Other" / "Help" / "Images"
    source_dir = project_root / "Other" / "Source"

    for folder in (app_dir, launcher_dir, data_dir, help_images_dir, source_dir):
        ensure_empty_or_create(folder)

    if project.copy_app_files:
        if app_exe.is_file():
            copy_application_folder(app_exe, app_dir, project_root)
        else:
            raise ValueError("Application executable must be a file.")

    create_portableapps_icons(project, appinfo_dir, app_exe)

    (appinfo_dir / "appinfo.ini").write_text(build_appinfo_ini(project), encoding="utf-8")
    installer_ini = build_installer_ini(project)
    if installer_ini:
        (appinfo_dir / "installer.ini").write_text(installer_ini + "\n", encoding="utf-8")
    (launcher_dir / f"{project.portable_name}.ini").write_text(build_launcher_ini(project), encoding="utf-8")
    create_launcher_template_assets(launcher_dir)
    create_help_images(help_images_dir, appinfo_dir)
    (project_root / "help.html").write_text(build_help_html(project), encoding="utf-8")
    (source_dir / "Readme.txt").write_text(build_readme(project), encoding="utf-8")
    return project_root


class PortableAppsLauncherMaker:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PortableApps.com Launcher Maker")
        self.root.geometry("980x700")
        self.root.minsize(840, 620)
        self.app_icon_image = None

        self.colors = {
            "page": "#eef3f7",
            "surface": "#ffffff",
            "surface_alt": "#f6f9fc",
            "toolbar": "#f3f7fb",
            "border": "#d5dee8",
            "text": "#17212f",
            "muted": "#627386",
            "accent": "#1f7a57",
            "accent_hover": "#186346",
            "accent_soft": "#e6f4ee",
            "accent_line": "#b7dcc8",
            "danger": "#b42318",
            "danger_hover": "#912018",
            "danger_soft": "#fbe9e7",
            "danger_line": "#ebbbb6",
            "soft": "#edf4fa",
            "warn": "#9a5b12",
            "warn_soft": "#fff6e8",
            "warn_line": "#f2d3a0",
        }

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
            "version": tk.StringVar(value="1.0.0.0"),
            "display_version": tk.StringVar(value="1.0.0.0"),
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
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        self.checkbox_style_images = self.build_checkbox_style_images()
        style.configure("TFrame", background=self.colors["page"])
        style.configure("Surface.TFrame", background=self.colors["surface"])
        style.configure("PanelBody.TFrame", background=self.colors["surface"])
        style.configure("TLabel", background=self.colors["page"], foreground=self.colors["text"], font=("Segoe UI", 9))
        style.configure("Muted.TLabel", background=self.colors["page"], foreground=self.colors["muted"], font=("Segoe UI", 9))
        style.configure("Surface.TLabel", background=self.colors["surface"], foreground=self.colors["text"], font=("Segoe UI", 9))
        style.configure("PanelTitle.TLabel", background=self.colors["surface"], foreground=self.colors["text"], font=("Segoe UI Semibold", 10))
        style.configure("PanelNote.TLabel", background=self.colors["surface"], foreground=self.colors["muted"], font=("Segoe UI", 8))
        style.configure("Title.TLabel", background=self.colors["page"], foreground=self.colors["text"], font=("Segoe UI Semibold", 16))
        style.configure(
            "TButton",
            font=("Segoe UI Semibold", 9),
            padding=(14, 8),
            background=self.colors["surface"],
            foreground=self.colors["text"],
            borderwidth=1,
            focusthickness=0,
            relief="solid",
            focuscolor=self.colors["surface"],
            bordercolor=self.colors["border"],
            lightcolor=self.colors["border"],
            darkcolor=self.colors["border"],
        )
        style.map(
            "TButton",
            background=[
                ("active", self.colors["surface_alt"]),
                ("pressed", self.colors["surface_alt"]),
            ],
            foreground=[("disabled", self.colors["muted"])],
        )
        style.configure(
            "Accent.TButton",
            background=self.colors["accent"],
            foreground="#ffffff",
            borderwidth=1,
            focusthickness=0,
            focuscolor=self.colors["accent"],
            relief="solid",
            bordercolor=self.colors["accent"],
            lightcolor=self.colors["accent"],
            darkcolor=self.colors["accent"],
        )
        style.map(
            "Accent.TButton",
            background=[
                ("active", self.colors["accent_hover"]),
                ("pressed", self.colors["accent_hover"]),
            ],
            foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
        )
        style.configure(
            "Danger.TButton",
            background=self.colors["danger_soft"],
            foreground=self.colors["danger"],
            padding=(12, 8),
            borderwidth=1,
            focusthickness=0,
            focuscolor=self.colors["danger_line"],
            relief="solid",
            bordercolor=self.colors["danger_line"],
            lightcolor=self.colors["danger_line"],
            darkcolor=self.colors["danger_line"],
        )
        style.map(
            "Danger.TButton",
            background=[
                ("active", self.colors["danger"]),
                ("pressed", self.colors["danger"]),
            ],
            foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
        )
        style.element_create(
            "Web.Checkbutton.indicator",
            "image",
            self.checkbox_style_images["unchecked"],
            ("disabled", "selected", self.checkbox_style_images["checked_disabled"]),
            ("disabled", self.checkbox_style_images["unchecked_disabled"]),
            ("pressed", "selected", self.checkbox_style_images["checked_hover"]),
            ("active", "selected", self.checkbox_style_images["checked_hover"]),
            ("selected", self.checkbox_style_images["checked"]),
            ("pressed", self.checkbox_style_images["unchecked_hover"]),
            ("active", self.checkbox_style_images["unchecked_hover"]),
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
        style.configure("TCheckbutton", background=self.colors["surface"], foreground=self.colors["text"], font=("Segoe UI", 9), padding=(0, 2))
        style.map("TCheckbutton", foreground=[("disabled", self.colors["muted"])], background=[("active", self.colors["surface"])])
        style.configure(
            "TEntry",
            fieldbackground="#ffffff",
            background="#ffffff",
            foreground=self.colors["text"],
            padding=(12, 8, 12, 8),
            borderwidth=1,
            relief="flat",
            bordercolor=self.colors["border"],
            lightcolor=self.colors["border"],
            darkcolor=self.colors["border"],
        )
        style.map(
            "TEntry",
            fieldbackground=[
                ("disabled", self.colors["surface_alt"]),
                ("readonly", self.colors["surface_alt"]),
                ("focus", "#ffffff"),
            ],
            foreground=[
                ("disabled", self.colors["muted"]),
                ("readonly", self.colors["text"]),
            ],
            bordercolor=[
                ("focus", self.colors["accent_line"]),
            ],
            lightcolor=[
                ("focus", self.colors["accent_line"]),
            ],
            darkcolor=[
                ("focus", self.colors["accent_line"]),
            ],
        )
        style.configure(
            "Web.TCombobox",
            fieldbackground="#ffffff",
            background="#ffffff",
            foreground=self.colors["text"],
            padding=(12, 8, 12, 8),
            arrowsize=15,
            borderwidth=1,
            relief="flat",
            arrowcolor=self.colors["muted"],
            bordercolor=self.colors["border"],
            lightcolor=self.colors["border"],
            darkcolor=self.colors["border"],
            selectbackground="#ffffff",
            selectforeground=self.colors["text"],
        )
        style.map(
            "Web.TCombobox",
            fieldbackground=[
                ("disabled", self.colors["surface_alt"]),
                ("readonly", "#ffffff"),
                ("focus", "#ffffff"),
            ],
            background=[
                ("disabled", self.colors["surface_alt"]),
                ("readonly", "#ffffff"),
                ("focus", "#ffffff"),
            ],
            foreground=[("disabled", self.colors["muted"])],
            arrowcolor=[
                ("disabled", self.colors["muted"]),
                ("pressed", self.colors["text"]),
                ("active", self.colors["text"]),
            ],
            bordercolor=[
                ("focus", self.colors["accent_line"]),
                ("active", self.colors["accent_line"]),
            ],
            lightcolor=[
                ("focus", self.colors["accent_line"]),
                ("active", self.colors["accent_line"]),
            ],
            darkcolor=[
                ("focus", self.colors["accent_line"]),
                ("active", self.colors["accent_line"]),
            ],
        )
        style.configure(
            "TCombobox",
            fieldbackground="#ffffff",
            background="#ffffff",
            foreground=self.colors["text"],
            padding=(10, 6, 10, 6),
            arrowsize=11,
            borderwidth=1,
            relief="flat",
            arrowcolor=self.colors["muted"],
            bordercolor=self.colors["border"],
            lightcolor=self.colors["border"],
            darkcolor=self.colors["border"],
        )
        style.map(
            "TCombobox",
            fieldbackground=[
                ("disabled", self.colors["surface_alt"]),
                ("readonly", "#ffffff"),
                ("focus", "#ffffff"),
            ],
            background=[
                ("disabled", self.colors["surface_alt"]),
                ("readonly", "#ffffff"),
                ("focus", "#ffffff"),
            ],
            foreground=[("disabled", self.colors["muted"])],
            arrowcolor=[
                ("disabled", self.colors["muted"]),
                ("pressed", self.colors["text"]),
                ("active", self.colors["text"]),
            ],
            bordercolor=[
                ("focus", self.colors["accent_line"]),
                ("active", self.colors["accent_line"]),
            ],
            lightcolor=[
                ("focus", self.colors["accent_line"]),
                ("active", self.colors["accent_line"]),
            ],
            darkcolor=[
                ("focus", self.colors["accent_line"]),
                ("active", self.colors["accent_line"]),
            ],
        )
        style.configure(
            "Vertical.TScrollbar",
            background=self.colors["surface_alt"],
            troughcolor=self.colors["toolbar"],
            bordercolor=self.colors["border"],
            arrowcolor=self.colors["muted"],
            darkcolor=self.colors["surface_alt"],
            lightcolor=self.colors["surface_alt"],
            gripcount=0,
        )
        style.configure("TNotebook", background=self.colors["surface"], borderwidth=0, tabmargins=(0, 0, 0, 0))
        style.configure(
            "TNotebook.Tab",
            background=self.colors["page"],
            foreground=self.colors["muted"],
            padding=(16, 10),
            borderwidth=0,
        )
        style.map(
            "TNotebook.Tab",
            background=[
                ("selected", self.colors["surface"]),
                ("active", self.colors["surface_alt"]),
            ],
            foreground=[("selected", self.colors["text"]), ("active", self.colors["text"])],
        )

    def build_checkbox_style_images(self):
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
            "unchecked": make_checkbox("#ffffff", self.colors["border"]),
            "unchecked_hover": make_checkbox(self.colors["surface_alt"], self.colors["accent_line"]),
            "unchecked_disabled": make_checkbox(self.colors["surface_alt"], self.colors["border"]),
            "checked": make_checkbox(self.colors["accent"], self.colors["accent"], "#ffffff"),
            "checked_hover": make_checkbox(self.colors["accent_hover"], self.colors["accent_hover"], "#ffffff"),
            "checked_disabled": make_checkbox(self.colors["accent_line"], self.colors["accent_line"], "#ffffff"),
        }

    def create_combobox(self, parent, *, textvariable, values, width=None):
        kwargs = {
            "textvariable": textvariable,
            "values": values,
            "style": "Web.TCombobox",
            "state": "readonly",
        }
        if width is not None:
            kwargs["width"] = width
        return ttk.Combobox(parent, **kwargs)

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

        self.root.bind_all("<MouseWheel>", self.scroll_form, add="+")
        self.root.bind_all("<Button-4>", self.scroll_form, add="+")
        self.root.bind_all("<Button-5>", self.scroll_form, add="+")

        notebook_shell = self.create_panel(body, "Project Settings", "Build and preview the PortableApps project from one place.")
        notebook_shell.grid(row=0, column=0, sticky="nsew")
        notebook_shell.content.columnconfigure(0, weight=1)
        notebook_shell.content.rowconfigure(1, weight=1)

        tab_bar = tk.Frame(notebook_shell.content, bg=self.colors["toolbar"], highlightthickness=1, highlightbackground=self.colors["border"], bd=0)
        tab_bar.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        tabs_host = tk.Frame(tab_bar, bg=self.colors["toolbar"], highlightthickness=0, bd=0)
        tabs_host.pack(side="left", anchor="w", padx=8, pady=8)
        actions = tk.Frame(tab_bar, bg=self.colors["toolbar"], highlightthickness=0, bd=0)
        actions.pack(side="right", anchor="e", padx=8, pady=8)
        self.help_button = ttk.Button(actions, text="Help", style="Danger.TButton", command=self.open_help)
        self.help_button.pack(side="left")

        tab_content = ttk.Frame(notebook_shell.content, style="Surface.TFrame")
        tab_content.grid(row=1, column=0, sticky="nsew")
        tab_content.columnconfigure(0, weight=1)
        tab_content.rowconfigure(0, weight=1)

        appinfo_tab, appinfo_content = self.create_scrollable_tab(tab_content)
        launcher_tab, launcher_content = self.create_scrollable_tab(tab_content)
        installer_tab, installer_content = self.create_scrollable_tab(tab_content)
        registry_tab, registry_content = self.create_scrollable_tab(tab_content)
        icon_tab, icon_content = self.create_scrollable_tab(tab_content)
        templates_tab, templates_content = self.create_scrollable_tab(tab_content)
        preview_tab = ttk.Frame(tab_content, style="Surface.TFrame", padding=0)
        preview_tab.columnconfigure(0, weight=1)
        preview_tab.rowconfigure(0, weight=0)
        self.create_main_tab(tabs_host, "appinfo", "appinfo.ini", appinfo_tab)
        self.launcher_tab = self.create_main_tab(tabs_host, "launcher", "AppNamePortable.ini", launcher_tab)
        self.create_main_tab(tabs_host, "installer", "installer.ini", installer_tab)
        self.create_main_tab(tabs_host, "registry", "Registry", registry_tab)
        self.create_main_tab(tabs_host, "icon", "Icon", icon_tab)
        self.create_main_tab(tabs_host, "templates", "Splash", templates_tab)
        self.create_main_tab(tabs_host, "preview", "Preview", preview_tab)
        self.select_main_tab("appinfo")

        self.create_appinfo_editor(appinfo_content)

        self.create_launcher_editor(launcher_content)
        self.create_installer_editor(installer_content)

        self.create_registry_editor(registry_content)
        self.create_icon_editor(icon_content)
        self.create_template_editor(templates_content)

        preview_tab.rowconfigure(1, weight=1)
        preview_bar = tk.Frame(preview_tab, bg=self.colors["toolbar"], highlightthickness=1, highlightbackground=self.colors["border"], bd=0)
        preview_bar.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        preview_content = ttk.Frame(preview_tab, style="Surface.TFrame")
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
            text = self.create_preview_text(tab)
            text.grid(row=0, column=0, sticky="nsew")
            self.preview_texts[key] = text
            self.create_preview_tab(preview_bar, key, label, tab)
        self.select_preview_tab("folder")
        footer_shell = self.create_panel(self.root, "Project Paths")
        footer_shell.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 16))
        footer = footer_shell.content
        footer.columnconfigure(1, weight=1)
        footer.columnconfigure(4, weight=1)

        ttk.Label(footer, textvariable=self.status_var, style="PanelNote.TLabel").grid(row=0, column=0, columnspan=5, sticky="w", pady=(0, 10))
        ttk.Label(footer, textvariable=self.generator_status_var, style="PanelNote.TLabel").grid(row=0, column=5, columnspan=2, sticky="e", pady=(0, 10))
        ttk.Label(footer, text="Application EXE", style="Surface.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 8))
        ttk.Entry(footer, textvariable=self.vars["app_exe"]).grid(row=1, column=1, sticky="ew")
        ttk.Button(footer, text="Browse", command=self.choose_app_exe).grid(row=1, column=2, sticky="ew", padx=(8, 16))
        ttk.Label(footer, text="Output Folder", style="Surface.TLabel").grid(row=1, column=3, sticky="w", padx=(0, 8))
        ttk.Entry(footer, textvariable=self.vars["output_dir"]).grid(row=1, column=4, sticky="ew")
        ttk.Button(footer, text="Browse", command=self.choose_output_dir).grid(row=1, column=5, sticky="ew", padx=(8, 16))
        actions = ttk.Frame(footer, style="PanelBody.TFrame")
        actions.grid(row=1, column=6, sticky="e")
        self.validate_button = ttk.Button(actions, text="Validate", command=self.validate_current_project)
        self.validate_button.pack(side="left", padx=(0, 8))
        self.create_button = ttk.Button(actions, text="Create Project + EXE", style="Accent.TButton", command=self.create_project)
        self.create_button.pack(side="left")

    def create_preview_text(self, parent):
        return tk.Text(
            parent,
            height=18,
            wrap="none",
            bg="#fbfcfe",
            fg=self.colors["text"],
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            font=("Consolas", 9),
            padx=10,
            pady=10,
        )

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
        button.pack(side="left", padx=(8 if not self.preview_tab_buttons else 0, 6), pady=8)
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
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        content = ttk.Frame(canvas, style="Surface.TFrame", padding=16)
        content.columnconfigure(0, weight=1)
        window_id = canvas.create_window((0, 0), window=content, anchor="nw")

        content.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(window_id, width=event.width))

        canvas._scroll_canvas = canvas
        content._scroll_canvas = canvas
        canvas.bind("<Enter>", lambda _event, target=canvas: self.set_active_scroll_canvas(target), add="+")
        content.bind("<Enter>", lambda _event, target=canvas: self.set_active_scroll_canvas(target), add="+")
        canvas.bind("<Leave>", lambda _event, target=canvas: self.clear_active_scroll_canvas(target), add="+")
        content.bind("<Leave>", lambda _event, target=canvas: self.clear_active_scroll_canvas(target), add="+")

        return outer, content

    def create_panel(self, parent, title, note=""):
        outer = tk.Frame(parent, bg=self.colors["border"], highlightthickness=0, bd=0)
        inner = ttk.Frame(outer, style="Surface.TFrame", padding=(16, 14, 16, 16))
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        inner.columnconfigure(0, weight=1)
        ttk.Label(inner, text=title, style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w")
        content_row = 1
        if note:
            ttk.Label(inner, text=note, style="PanelNote.TLabel", wraplength=760).grid(row=1, column=0, sticky="ew", pady=(2, 12))
            content_row = 2
        inner.rowconfigure(content_row, weight=1)
        content = ttk.Frame(inner, style="PanelBody.TFrame")
        content.grid(row=content_row, column=0, sticky="nsew")
        content.columnconfigure(0, weight=1)
        outer.content = content
        return outer

    def create_text_editor(self, parent, height=4, background="#ffffff"):
        return tk.Text(
            parent,
            height=height,
            wrap="none",
            bg=background,
            fg=self.colors["text"],
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            font=("Consolas", 9),
            padx=8,
            pady=8,
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
                frame.columnconfigure(index, weight=1 if index % 2 else 0)
            row += 1
            return frame

        def add_entry(frame, field_row, column, label, key, value_span=1):
            ttk.Label(frame, text=label, style="Surface.TLabel").grid(
                row=field_row,
                column=column,
                sticky="w",
                padx=(0, 6),
                pady=4,
            )
            ttk.Entry(frame, textvariable=self.vars[key]).grid(
                row=field_row,
                column=column + 1,
                columnspan=value_span,
                sticky="ew",
                padx=(0, 8),
                pady=4,
            )

        def add_combo(frame, field_row, column, label, key, values, value_span=1):
            ttk.Label(frame, text=label, style="Surface.TLabel").grid(
                row=field_row,
                column=column,
                sticky="w",
                padx=(0, 6),
                pady=4,
            )
            self.create_combobox(frame, textvariable=self.vars[key], values=values).grid(
                row=field_row,
                column=column + 1,
                columnspan=value_span,
                sticky="ew",
                padx=(0, 8),
                pady=4,
            )

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
        self.import_registry_button = ttk.Button(
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
        for column in range(6):
            launch_group.columnconfigure(column, weight=1 if column % 2 else 0)

        ttk.Label(launch_group, text="ProgramExecutable", style="Surface.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.launch_program_executable_var = tk.StringVar()
        ttk.Entry(launch_group, textvariable=self.launch_program_executable_var, state="readonly").grid(row=0, column=1, columnspan=5, sticky="ew", pady=4)

        ttk.Label(launch_group, text="Arguments", style="Surface.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(launch_group, textvariable=self.vars["command_line"]).grid(row=1, column=1, columnspan=5, sticky="ew", pady=4)

        ttk.Label(launch_group, text="Working Dir", style="Surface.TLabel").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(launch_group, textvariable=self.vars["working_directory"]).grid(row=2, column=1, columnspan=3, sticky="ew", pady=4, padx=(0, 10))
        ttk.Label(launch_group, text="Close EXE", style="Surface.TLabel").grid(row=2, column=4, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(launch_group, textvariable=self.vars["close_exe"]).grid(row=2, column=5, sticky="ew", pady=4)

        ttk.Label(launch_group, text="Min OS", style="Surface.TLabel").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=4)
        self.create_combobox(launch_group, textvariable=self.vars["min_os"], values=os_values, width=16).grid(row=3, column=1, sticky="w", pady=4, padx=(0, 10))
        ttk.Label(launch_group, text="Max OS", style="Surface.TLabel").grid(row=3, column=2, sticky="w", padx=(0, 8), pady=4)
        self.create_combobox(launch_group, textvariable=self.vars["max_os"], values=os_values, width=16).grid(row=3, column=3, sticky="w", pady=4, padx=(0, 10))
        ttk.Label(launch_group, text="Run As Admin", style="Surface.TLabel").grid(row=3, column=4, sticky="w", padx=(0, 8), pady=4)
        self.create_combobox(launch_group, textvariable=self.vars["run_as_admin"], values=run_as_admin_values, width=16).grid(row=3, column=5, sticky="w", pady=4)

        ttk.Label(launch_group, text="Refresh Shell Icons", style="Surface.TLabel").grid(row=4, column=0, sticky="w", padx=(0, 8), pady=4)
        self.create_combobox(launch_group, textvariable=self.vars["refresh_shell_icons"], values=refresh_shell_values, width=16).grid(row=4, column=1, sticky="w", pady=4, padx=(0, 10))
        ttk.Label(launch_group, text="Supports UNC", style="Surface.TLabel").grid(row=4, column=2, sticky="w", padx=(0, 8), pady=4)
        self.create_combobox(launch_group, textvariable=self.vars["supports_unc"], values=unc_values, width=16).grid(row=4, column=3, sticky="w", pady=4)

        launch_checks = ttk.Frame(launch_group, style="PanelBody.TFrame")
        launch_checks.grid(row=5, column=0, columnspan=6, sticky="ew", pady=(10, 0))
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
        activate_row = ttk.Frame(activate_group, style="PanelBody.TFrame")
        activate_row.grid(row=0, column=0, sticky="w")
        ttk.Label(activate_row, text="Java", style="Surface.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        self.create_combobox(activate_row, textvariable=self.vars["activate_java"], values=java_values, width=18).grid(row=0, column=1, sticky="w", pady=4)
        ttk.Checkbutton(activate_group, text="Enable XML support", variable=self.vars["activate_xml"]).grid(row=1, column=0, sticky="w", pady=(10, 0))

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
        row_frame = ttk.Frame(check_running_group, style="PanelBody.TFrame")
        row_frame.grid(row=0, column=0, sticky="w")
        ttk.Label(row_frame, text="CloseEXE", style="Surface.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(row_frame, textvariable=self.vars["installer_close_exe"], width=28).grid(row=0, column=1, sticky="w", pady=4, padx=(0, 12))
        ttk.Label(row_frame, text="CloseName", style="Surface.TLabel").grid(row=0, column=2, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(row_frame, textvariable=self.vars["installer_close_name"], width=28).grid(row=0, column=3, sticky="w", pady=4)

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
        for column in range(6):
            optional_group.columnconfigure(column, weight=1 if column % 2 else 0)
        ttk.Checkbutton(
            optional_group,
            text="Enable optional components section",
            variable=self.vars["optional_components_enabled"],
        ).grid(row=0, column=0, columnspan=6, sticky="w", pady=(0, 10))
        ttk.Label(optional_group, text="Main Title", style="Surface.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(optional_group, textvariable=self.vars["main_section_title"]).grid(row=1, column=1, sticky="ew", pady=4, padx=(0, 10))
        ttk.Label(optional_group, text="Main Description", style="Surface.TLabel").grid(row=1, column=2, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(optional_group, textvariable=self.vars["main_section_description"]).grid(row=1, column=3, columnspan=3, sticky="ew", pady=4)
        ttk.Label(optional_group, text="Optional Title", style="Surface.TLabel").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(optional_group, textvariable=self.vars["optional_section_title"]).grid(row=2, column=1, sticky="ew", pady=4, padx=(0, 10))
        ttk.Label(optional_group, text="Optional Description", style="Surface.TLabel").grid(row=2, column=2, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(optional_group, textvariable=self.vars["optional_section_description"]).grid(row=2, column=3, columnspan=3, sticky="ew", pady=4)
        ttk.Label(optional_group, text="Selected InstallType", style="Surface.TLabel").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(optional_group, textvariable=self.vars["optional_section_selected_install_type"]).grid(row=3, column=1, sticky="ew", pady=4, padx=(0, 10))
        ttk.Label(optional_group, text="Not Selected InstallType", style="Surface.TLabel").grid(row=3, column=2, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(optional_group, textvariable=self.vars["optional_section_not_selected_install_type"]).grid(row=3, column=3, sticky="ew", pady=4, padx=(0, 10))
        ttk.Label(optional_group, text="Preselect If Non-English", style="Surface.TLabel").grid(row=3, column=4, sticky="w", padx=(0, 8), pady=4)
        self.create_combobox(optional_group, textvariable=self.vars["optional_section_preselected"], values=("", "true", "false")).grid(row=3, column=5, sticky="ew", pady=4)

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
        row_frame = ttk.Frame(icon_group, style="PanelBody.TFrame")
        row_frame.grid(row=0, column=0, sticky="w")

        ttk.Label(row_frame, text="Icon Index", style="Surface.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=4)
        ttk.Entry(row_frame, textvariable=self.vars["icon_index"], width=5).grid(row=0, column=1, sticky="w", pady=4, padx=(0, 12))

        ttk.Label(row_frame, text="Icon Override", style="Surface.TLabel").grid(row=0, column=2, sticky="w", padx=(0, 6), pady=4)
        ttk.Entry(row_frame, textvariable=self.vars["icon_source"], width=42).grid(row=0, column=3, sticky="w", pady=4)
        ttk.Button(row_frame, text="Browse", command=self.choose_icon).grid(row=0, column=4, sticky="w", padx=(8, 0), pady=4)

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
        ttk.Button(assets_toolbar, text="Open Assets Folder", command=self.open_template_folder).grid(row=0, column=0, sticky="w")

        for row_index, (relative_path, label, filetypes) in enumerate(TEMPLATE_ASSET_SPECS, start=1):
            self.add_template_asset_row(assets_group, row_index, relative_path, label, filetypes)

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
        ttk.Label(parent, text=label, style="Surface.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 6), pady=4)
        path_var = tk.StringVar(value=str(asset_path(relative_path)))
        self.template_asset_path_vars[relative_path] = path_var
        ttk.Entry(parent, textvariable=path_var, state="readonly").grid(row=row, column=1, sticky="ew", padx=(0, 6), pady=4)
        ttk.Button(
            parent,
            text="Open",
            command=lambda target=relative_path: self.open_template_asset(target),
        ).grid(row=row, column=2, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(
            parent,
            text="Replace",
            command=lambda target=relative_path, chooser=filetypes: self.replace_template_asset(target, chooser),
        ).grid(row=row, column=3, sticky="ew", pady=4)

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
            return

        splash_path = splash_asset_path()
        if splash_path.exists():
            try:
                with Image.open(splash_path) as splash_image:
                    preview = ImageOps.contain(splash_image.convert("RGBA"), (320, 180), Image.Resampling.LANCZOS)
                self.template_splash_image = ImageTk.PhotoImage(preview)
                self.template_splash_label.configure(image=self.template_splash_image, text="")
                self.template_splash_caption.configure(text=str(splash_path))
                return
            except OSError:
                pass

        self.template_splash_image = None
        self.template_splash_label.configure(image="", text="Splash preview unavailable")
        self.template_splash_caption.configure(text=str(splash_path))

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
        ttk.Button(toolbar, text="Additional Variable Help", command=self.open_variable_help_site).grid(
            row=1,
            column=0,
            sticky="w",
        )
        ttk.Button(toolbar, text="Additional Registry Help", command=self.open_registry_help_site).grid(
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
        ttk.Entry(parent, textvariable=self.vars[key]).grid(row=row, column=column + 1, columnspan=columnspan, sticky="ew", pady=2)

    def add_inline_combo(self, parent, row, column, label, key, values):
        ttk.Label(parent, text=label).grid(row=row, column=column, sticky="w", padx=(0, 6), pady=2)
        self.create_combobox(parent, textvariable=self.vars[key], values=values).grid(row=row, column=column + 1, sticky="ew", pady=2)

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
        previous_state = str(text.cget("state"))
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
        ttk.Entry(parent, textvariable=self.vars[key]).grid(row=row, column=1, columnspan=2, sticky="ew", padx=(0, 6), pady=5)
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
        ttk.Entry(parent, textvariable=self.vars[key]).grid(row=row, column=column + 1, columnspan=value_span, sticky="ew", padx=(0, 6), pady=4)

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
        text.configure(wrap="word", highlightthickness=0)
        text.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 10))
        text.insert("1.0", content)
        text.configure(state="disabled")

    def add_combo_row(self, parent, row, label, key, values):
        ttk.Label(parent, text=label, style="Surface.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 4), pady=5)
        self.create_combobox(parent, textvariable=self.vars[key], values=values).grid(row=row, column=1, columnspan=2, sticky="ew", padx=(0, 6), pady=5)
        return row + 1

    def add_path_row(self, parent, row, label, key, command, file_hint):
        ttk.Label(parent, text=label, style="Surface.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 4), pady=5)
        ttk.Entry(parent, textvariable=self.vars[key]).grid(row=row, column=1, sticky="ew", padx=(0, 6), pady=5)
        ttk.Button(parent, text="Browse", command=command).grid(row=row, column=2, sticky="ew", padx=(8, 0), pady=5)
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
        if not self.icon_preview_labels or self.icon_preview_caption is None:
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
        for icon_label, (_size_label, display_size) in zip(self.icon_preview_labels, ICON_PREVIEW_DISPLAY_SIZES):
            preview = image.resize((display_size, display_size), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(preview)
            self.icon_preview_images.append(photo)
            icon_label.configure(image=photo)
        self.icon_preview_caption.configure(text=caption)
        self.icon_preview_cache_key = cache_key

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
        ttk.Button(footer, text="Close", command=self.close_validation_popup).grid(row=0, column=1, sticky="e")

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
        folder_lines = [
            f"{project.portable_name}\\",
            f"  {project.portable_name}.exe  (build with PortableApps.com Launcher)",
            "  help.html",
            "  App\\",
            "    AppInfo\\",
            "      appinfo.ini",
            "      installer.ini  (optional)",
            "      appicon.ico",
            "      appicon_16.png",
            "      appicon_32.png",
            "      appicon_75.png",
            "      appicon_128.png",
            "      appicon_256.png",
            f"      source icon index: {project.icon_index}",
            "      Launcher\\",
            f"        {project.portable_name}.ini  (launcher settings)",
            "        Splash.jpg",
            f"    {project.package_name}\\",
            f"      {project.app_exe_name or 'YourApp.exe'}",
            "  Data\\",
            "    settings\\",
            "  Other\\",
            "    Help\\",
            "      Images\\",
            "        appicon_128.png",
            "    Source\\",
            "      Readme.txt",
        ]
        previews = {
            "folder": "\n".join(folder_lines),
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
            text.insert("1.0", content)
            text.configure(state="disabled")
        self.update_icon_preview()

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
    root = tk.Tk()
    PortableAppsLauncherMaker(root)
    root.mainloop()
