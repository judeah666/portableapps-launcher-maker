import html
import os
import re
import shutil
import struct
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path

from PIL import Image, ImageDraw, ImageOps


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
DEFAULT_FALLBACK_ICON = "icons/default_portable_icon.png"
SOFTWARE_LOGO = "icons/software_logo.png"
SOFTWARE_ICON = "icons/software_icon.ico"
SOFTWARE_ICON_PNG = "icons/software_icon_256.png"
DEFAULT_SPLASH_ASSET = "splash/default_splash.png"
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
    command_line: str = ""
    working_directory: str = ""
    close_exe: str = ""
    wait_for_other_instances: bool = False
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
    copy_app_files: bool = True
    wait_for_program: bool = True
    icon_source: str = ""
    icon_index: int = 0
    language: str = "Multilingual"
    trademarks: str = ""
    donate: str = ""
    install_type: str = ""
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
    dependency_requires_admin: str = "no"
    dependency_requires_portable_app: str = ""
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
    registry_enabled: bool = False
    registry_keys: str = ""
    registry_cleanup_if_empty: str = ""
    registry_cleanup_force: str = ""
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
    files_move: str = ""
    directories_move: str = ""

    @property
    def portable_name(self) -> str:
        return f"{self.package_name}Portable"

    @property
    def portable_display_name(self) -> str:
        return f"{self.app_name} Portable"

    @property
    def app_exe_name(self) -> str:
        return Path(self.app_exe).name

    @property
    def special_paths_plugins(self) -> str:
        return self.special_plugins

    @special_paths_plugins.setter
    def special_paths_plugins(self, value: str) -> None:
        self.special_plugins = value


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
    stem = Path(path).stem.replace("-", " ").replace("_", " ")
    return clean_display_name(stem)


def portableapps_launcher_candidates() -> list[Path]:
    candidates: list[Path] = []

    cwd = Path.cwd().resolve()
    drive_root = cwd.anchor or ""
    if drive_root:
        candidates.append(Path(drive_root) / "PortableApps" / "PortableApps.comLauncher" / "PortableApps.comLauncherGenerator.exe")

    executable_path = Path(sys.executable).resolve()
    for base in (cwd, executable_path.parent):
        current = base
        visited: set[Path] = set()
        while current not in visited:
            visited.add(current)
            anchor = current.anchor
            if anchor:
                candidates.append(Path(anchor) / "PortableApps" / "PortableApps.comLauncher" / "PortableApps.comLauncherGenerator.exe")
            portable_apps = current / "PortableApps" / "PortableApps.comLauncher" / "PortableApps.comLauncherGenerator.exe"
            candidates.append(portable_apps)
            parent = current.parent
            if parent == current:
                break
            current = parent

    env_path = os.environ.get("PORTABLEAPPS_LAUNCHER_GENERATOR")
    if env_path:
        candidates.append(Path(env_path))

    unique_candidates: list[Path] = []
    seen: set[str] = set()
    for candidate in candidates:
        key = str(candidate).casefold()
        if key not in seen:
            seen.add(key)
            unique_candidates.append(candidate)
    return unique_candidates


def find_portableapps_launcher() -> Path | None:
    for candidate in portableapps_launcher_candidates():
        if candidate.exists():
            return candidate
    return None


def app_base_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def asset_root() -> Path:
    base = app_base_path()
    direct_root = base / "assets"
    bundled_root = base / "app" / "assets"
    if direct_root.exists():
        return direct_root
    if bundled_root.exists():
        return bundled_root
    return direct_root


def default_portableapps_output_dir() -> str:
    cwd = Path.cwd().resolve()
    anchor = cwd.anchor or ""
    if anchor:
        return str(Path(anchor) / "PortableApps")
    return str(cwd / "PortableApps")


def asset_path(name: str) -> Path:
    return asset_root() / Path(name)


def help_template_path(*parts: str) -> Path:
    return asset_path(str(Path("help", *parts)))


def help_image_asset_path(name: str) -> Path:
    return help_template_path("Images", name)


def load_help_html_template() -> str:
    template_path = help_template_path("help.html")
    if template_path.exists():
        return template_path.read_text(encoding="utf-8")
    return "<html><body><h1>**App Name** Portable Help</h1></body></html>"


def launcher_template_asset_path(name: str) -> Path:
    return asset_path(str(Path("splash", name)))


def splash_asset_path() -> Path:
    return asset_path(DEFAULT_SPLASH_ASSET)


def open_folder_in_explorer(path: Path) -> None:
    try:
        os.startfile(str(path))
    except OSError:
        subprocess.Popen(["explorer", str(path)])


def resolve_project_tokens(value: str, project: LauncherProject) -> str:
    resolved = value or ""
    replacements = {
        "{app_name}": project.package_name or "MyApp",
        "{package_name}": project.package_name or "MyApp",
        "{portable_name}": project.portable_name or "MyAppPortable",
    }
    for token, token_value in replacements.items():
        resolved = resolved.replace(token, token_value)
    return resolved


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
            errors.append("Application EXE path does not exist.")
        elif app_exe.suffix.lower() != ".exe":
            errors.append("Application EXE must point to an .exe file.")
    if not project.output_dir.strip():
        errors.append("Output Folder is required.")
    if project.icon_source.strip() and not Path(project.icon_source).exists():
        errors.append("Icon Override path does not exist.")
    if project.version.strip() and not re.fullmatch(r"\d+\.\d+\.\d+\.\d+", project.version.strip()):
        errors.append("Package Version must use the PortableApps format: major.minor.patch.build")

    if not project.registry_enabled and (
        has_ini_lines(project.registry_keys)
        or has_ini_lines(project.registry_cleanup_if_empty)
        or has_ini_lines(project.registry_cleanup_force)
    ):
        errors.append("Enable Registry in the Registry tab before using RegistryKeys or cleanup sections.")
    return errors


def validate_ini_mapping_lines(text: str, section_name: str) -> list[str]:
    errors: list[str] = []
    for index, line in enumerate((text or "").splitlines(), start=1):
        cleaned = line.strip()
        if not cleaned:
            continue
        if "=" not in cleaned:
            errors.append(f"{section_name} line {index} must use key=value format.")
            continue
        key, value = cleaned.split("=", 1)
        if not key.strip() or not value.strip():
            errors.append(f"{section_name} line {index} must use key=value format.")
    return errors


def build_validation_items(project: LauncherProject, launcher_path: Path | None = None) -> list[ValidationItem]:
    mapped_errors = set(validate_project(project))
    items: list[ValidationItem] = []
    error_field_map = {
        "App Name is required.": "App Name",
        "Package ID is required.": "Package ID",
        "Application EXE is required.": "Application EXE",
        "Application EXE path does not exist.": "Application EXE",
        "Application EXE must point to an .exe file.": "Application EXE",
        "Output Folder is required.": "Output Folder",
        "Icon Override path does not exist.": "Icon Override",
        "Package Version must use the PortableApps format: major.minor.patch.build": "Package Version",
        "Enable Registry in the Registry tab before using RegistryKeys or cleanup sections.": "Registry",
    }
    for error in validate_project(project):
        items.append(ValidationItem("error", error_field_map.get(error, "Project"), error))

    for section_name, text in (
        ("FilesMove", project.files_move),
        ("DirectoriesMove", project.directories_move),
        ("RegistryKeys", project.registry_keys),
        ("RegistryCleanupIfEmpty", project.registry_cleanup_if_empty),
        ("RegistryCleanupForce", project.registry_cleanup_force),
        ("Languages", project.installer_languages),
        ("DirectoriesToPreserve", project.preserve_directories),
        ("DirectoriesToRemove", project.remove_directories),
        ("FilesToPreserve", project.preserve_files),
        ("FilesToRemove", project.remove_files),
        ("FileTypeIcons", project.file_type_icons),
    ):
        for error in validate_ini_mapping_lines(text, section_name):
            items.append(ValidationItem("warning", section_name, error))

    if launcher_path is None:
        launcher_path = find_portableapps_launcher()
    if launcher_path is None or not launcher_path.exists():
        items.append(ValidationItem("warning", "PortableApps Generator", "PortableApps.com Launcher Generator not found."))
    else:
        items.append(ValidationItem("ok", "PortableApps Generator", launcher_path.name))

    if project.app_name.strip() and "App Name is required." not in mapped_errors:
        items.append(ValidationItem("ok", "App Name", project.app_name.strip()))
    if project.package_name.strip() and "Package ID is required." not in mapped_errors:
        items.append(ValidationItem("ok", "Package ID", project.package_name.strip()))
    if project.output_dir.strip() and "Output Folder is required." not in mapped_errors:
        items.append(ValidationItem("ok", "Output Folder", project.output_dir.strip()))
    if project.version.strip() and "Package Version must use the PortableApps format: major.minor.patch.build" not in mapped_errors:
        items.append(ValidationItem("ok", "Package Version", project.version.strip()))

    if project.icon_source.strip() and "Icon Override path does not exist." not in mapped_errors:
        items.append(ValidationItem("ok", "Icon Override", Path(project.icon_source).name))

    if project.registry_enabled:
        items.append(ValidationItem("ok", "Registry", "Enabled"))
    elif has_ini_lines(project.registry_keys) or has_ini_lines(project.registry_cleanup_if_empty) or has_ini_lines(project.registry_cleanup_force):
        items.append(ValidationItem("warning", "Registry", "Registry sections have content but registry handling is disabled."))

    if has_ini_lines(project.files_move):
        items.append(ValidationItem("ok", "FilesMove", "Entries ready"))
    if has_ini_lines(project.directories_move):
        items.append(ValidationItem("ok", "DirectoriesMove", "Entries ready"))
    if project.optional_components_enabled:
        items.append(ValidationItem("ok", "OptionalComponents", "Enabled"))

    return items


def render_validation_report(items: list[ValidationItem]) -> tuple[str, str, str]:
    errors = [item for item in items if item.level == "error"]
    warnings = [item for item in items if item.level == "warning"]
    oks = [item for item in items if item.level == "ok"]

    if errors:
        title = "Validation Found Errors"
        tone = "error"
    elif warnings:
        title = "Validation Found Warnings"
        tone = "warning"
    else:
        title = "Validation Passed"
        tone = "ok"
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
    report = "\n".join(sections).strip()
    return title, report, tone


def build_appinfo_ini(project: LauncherProject) -> str:
    version = project.version or "1.0.0.0"
    display_version = project.display_version or version
    lines = [
        "[Format]",
        "Type=PortableApps.comFormat",
        "Version=3.9",
        "",
        "[Details]",
        f"Name={project.portable_display_name}",
        f"AppID={project.portable_name}",
        f"Publisher={project.publisher or 'Unknown Publisher'}",
        f"Homepage={project.homepage or 'https://portableapps.com/'}",
        f"Category={project.category or 'Utilities'}",
        f"Description={project.description or project.app_name + ' portable launcher'}",
        f"Language={project.language or 'Multilingual'}",
    ]
    if project.trademarks.strip():
        lines.append(f"Trademarks={project.trademarks.strip()}")
    if project.donate.strip():
        lines.append(f"Donate={project.donate.strip()}")
    if project.install_type.strip():
        lines.append(f"InstallType={project.install_type.strip()}")
    lines.extend(
        [
            "",
            "[License]",
            f"Shareable={bool_to_ini(project.license_shareable)}",
            f"OpenSource={bool_to_ini(project.license_open_source)}",
            f"Freeware={bool_to_ini(project.license_freeware)}",
            f"CommercialUse={bool_to_ini(project.license_commercial_use)}",
        ]
    )
    if project.license_eula_version.strip():
        lines.append(f"EULAVersion={project.license_eula_version.strip()}")
    lines.extend(
        [
            "",
            "[Version]",
            f"PackageVersion={version}",
            f"DisplayVersion={display_version}",
            "",
            "[SpecialPaths]",
            f"Plugins={project.special_plugins.strip() or 'NONE'}",
            "",
            "[Dependencies]",
            f"UsesGhostscript={project.dependency_uses_ghostscript.strip() or 'no'}",
            f"UsesJava={project.dependency_uses_java.strip() or 'no'}",
            f"UsesDotNetVersion={project.dependency_uses_dotnet_version.strip()}",
            f"Requires64bitOS={project.dependency_requires_64bit_os.strip() or 'no'}",
            f"RequiresAdmin={project.dependency_requires_admin.strip() or 'no'}",
        ]
    )
    if project.dependency_requires_portable_app.strip():
        lines.append(f"RequiresPortableApp={project.dependency_requires_portable_app.strip()}")

    control_lines = [
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
            control_lines.append(f"{key}={value.strip()}")
    if control_lines:
        lines.extend(["", "[Control]", *control_lines])

    association_lines = [
        f"FileTypes={project.association_file_types.strip()}",
        f"FileTypeCommandLine={project.association_file_type_command_line.strip()}",
        f"FileTypeCommandLine-extension={project.association_file_type_command_line_extension.strip()}",
        f"Protocols={project.association_protocols.strip()}",
        f"ProtocolCommandLine={project.association_protocol_command_line.strip()}",
        f"ProtocolCommandLine-protocol={project.association_protocol_command_line_protocol.strip()}",
        f"SendTo={bool_to_ini(project.association_send_to)}",
        f"SendToCommandLine={project.association_send_to_command_line.strip()}",
        f"Shell={bool_to_ini(project.association_shell)}",
        f"ShellCommand={project.association_shell_command.strip()}",
    ]
    non_default_association_lines = [
        line for line in association_lines if not line.endswith("=") and not line.endswith("=false")
    ]
    if non_default_association_lines:
        lines.extend(["", "[Associations]", *association_lines])

    file_type_icon_lines = clean_ini_lines(project.file_type_icons)
    if file_type_icon_lines:
        lines.extend(["", "[FileTypeIcons]", *file_type_icon_lines])

    return "\n".join(lines).strip() + "\n"


def build_launcher_ini(project: LauncherProject) -> str:
    package_dir = project.package_name or "YourApp"
    app_exe_name = project.app_exe_name or "YourApp.exe"
    working_directory = resolve_project_tokens(project.working_directory, project)
    lines = [
        "[Launch]",
        f"ProgramExecutable={package_dir}\\{app_exe_name}",
        f"CommandLineArguments={project.command_line}",
        f"WorkingDirectory={working_directory}",
    ]
    if project.close_exe:
        lines.append(f"CloseEXE={project.close_exe}")
    lines.append(f"WaitForProgram={bool_to_ini(project.wait_for_program)}")
    lines.append(f"WaitForOtherInstances={bool_to_ini(project.wait_for_other_instances)}")
    if project.min_os:
        lines.append(f"MinOS={project.min_os}")
    if project.max_os:
        lines.append(f"MaxOS={project.max_os}")
    if project.run_as_admin:
        lines.append(f"RunAsAdmin={project.run_as_admin}")
    if project.refresh_shell_icons:
        lines.append(f"RefreshShellIcons={project.refresh_shell_icons}")
    if project.hide_command_line_window:
        lines.append("HideCommandLineWindow=true")
    if project.no_spaces_in_path:
        lines.append("NoSpacesInPath=true")
    if project.supports_unc:
        lines.append(f"SupportsUNC={project.supports_unc}")

    lines.extend(["", "[Activate]", f"Registry={bool_to_ini(project.registry_enabled)}"])
    if project.activate_java:
        lines.append(f"Java={project.activate_java}")
    if project.activate_xml:
        lines.append("XML=true")

    if project.live_mode_copy_app or project.live_mode_copy_data:
        lines.extend(["", "[LiveMode]"])
        if project.live_mode_copy_app:
            lines.append("CopyApp=true")
        if project.live_mode_copy_data:
            lines.append("CopyData=true")

    file_move_lines = clean_ini_lines(project.files_move)
    if file_move_lines:
        lines.extend(["", "[FilesMove]", *file_move_lines])

    directory_move_lines = clean_ini_lines(project.directories_move)
    if directory_move_lines:
        lines.extend(["", "[DirectoriesMove]", *directory_move_lines])

    if project.registry_enabled:
        for section_name, text in (
            ("RegistryKeys", project.registry_keys),
            ("RegistryCleanupIfEmpty", project.registry_cleanup_if_empty),
            ("RegistryCleanupForce", project.registry_cleanup_force),
        ):
            section_lines = clean_ini_lines(text)
            if section_lines:
                lines.extend(["", f"[{section_name}]", *section_lines])

    return "\n".join(lines).strip() + "\n"


def build_installer_ini(project: LauncherProject) -> str:
    lines: list[str] = []

    if project.installer_close_exe or project.installer_close_name:
        lines.extend(["[CheckRunning]"])
        if project.installer_close_exe:
            lines.append(f"CloseEXE={project.installer_close_exe}")
        if project.installer_close_name:
            lines.append(f"CloseName={project.installer_close_name}")
        lines.append("")

    if project.include_installer_source:
        lines.extend(["[Source]", "IncludeInstallerSource=true", ""])

    if project.remove_app_directory or project.remove_data_directory or project.remove_other_directory:
        lines.extend(["[MainDirectories]"])
        if project.remove_app_directory:
            lines.append("RemoveAppDirectory=true")
        if project.remove_data_directory:
            lines.append("RemoveDataDirectory=true")
        if project.remove_other_directory:
            lines.append("RemoveOtherDirectory=true")
        lines.append("")

    if project.optional_components_enabled:
        lines.extend(["[OptionalComponents]"])
        lines.append("OptionalComponents=true")
        for key, value in (
            ("MainSectionTitle", project.main_section_title),
            ("MainSectionDescription", project.main_section_description),
            ("OptionalSectionTitle", project.optional_section_title),
            ("OptionalSectionDescription", project.optional_section_description),
            ("OptionalSectionSelectedInstallType", project.optional_section_selected_install_type),
            ("OptionalSectionNotSelectedInstallType", project.optional_section_not_selected_install_type),
            ("OptionalSectionPreSelectedIfNonEnglish", project.optional_section_preselected),
        ):
            if value.strip():
                lines.append(f"{key}={value.strip()}")
        lines.append("")

    language_lines = clean_ini_lines(project.installer_languages)
    if language_lines:
        lines.append("[Languages]")
        lines.extend(language_lines)
        lines.append("")

    preserve_dir_lines = clean_ini_lines(project.preserve_directories)
    if preserve_dir_lines:
        lines.append("[DirectoriesToPreserve]")
        lines.extend(preserve_dir_lines)
        lines.append("")

    remove_dir_lines = clean_ini_lines(project.remove_directories)
    if remove_dir_lines:
        lines.append("[DirectoriesToRemove]")
        lines.extend(remove_dir_lines)
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
        raise ValueError("RVA outside section table.")

    return sections, rva_to_offset, rva_to_offset(resource_rva)


def parse_resource_name(data: bytes, resource_base: int, value: int):
    if value & 0x80000000:
        offset = resource_base + (value & 0x7FFFFFFF)
        length = read_uint16(data, offset)
        name = data[offset + 2:offset + 2 + length * 2].decode("utf-16le")
        return name
    return value


def parse_resource_directory(data: bytes, resource_base: int, directory_offset: int):
    absolute_offset = resource_base + directory_offset
    named_count = read_uint16(data, absolute_offset + 12)
    id_count = read_uint16(data, absolute_offset + 14)
    entries_offset = absolute_offset + 16
    entries = []
    for index in range(named_count + id_count):
        entry_offset = entries_offset + index * 8
        name = parse_resource_name(data, resource_base, read_uint32(data, entry_offset))
        target = read_uint32(data, entry_offset + 4)
        entries.append((name, target))
    return entries


def collect_resource_data(data: bytes, resource_base: int, rva_to_offset, resource_type: int):
    type_entries = parse_resource_directory(data, resource_base, 0)
    for name, target in type_entries:
        if name != resource_type or not (target & 0x80000000):
            continue
        name_directory = target & 0x7FFFFFFF
        collected = []
        for resource_name, name_target in parse_resource_directory(data, resource_base, name_directory):
            if not (name_target & 0x80000000):
                continue
            language_directory = name_target & 0x7FFFFFFF
            for language_name, language_target in parse_resource_directory(data, resource_base, language_directory):
                if language_target & 0x80000000:
                    continue
                data_entry = resource_base + language_target
                data_rva = read_uint32(data, data_entry)
                size = read_uint32(data, data_entry + 4)
                collected.append((resource_name, language_name, rva_to_offset(data_rva), size))
        return collected
    return []


def extract_icon_group_from_exe(source_exe: Path, destination_ico: Path, icon_index: int = 0) -> bool:
    try:
        data = source_exe.read_bytes()
        _sections, rva_to_offset, resource_base = get_pe_sections_and_resource_offset(data)
        group_entries = collect_resource_data(data, resource_base, rva_to_offset, 14)
        icon_entries = collect_resource_data(data, resource_base, rva_to_offset, 3)
        if not group_entries or not icon_entries:
            return False

        icon_payloads = {}
        for resource_name, language_name, offset, size in icon_entries:
            icon_payloads[(resource_name, language_name)] = data[offset:offset + size]
            icon_payloads[(resource_name, None)] = data[offset:offset + size]

        icon_index = max(0, min(int(icon_index or 0), len(group_entries) - 1))
        _group_name, _group_language, group_offset, group_size = group_entries[icon_index]
        group_data = data[group_offset:group_offset + group_size]
        if len(group_data) < 6:
            return False

        reserved = read_uint16(group_data, 0)
        icon_type = read_uint16(group_data, 2)
        count = read_uint16(group_data, 4)
        if reserved != 0 or icon_type != 1 or count <= 0:
            return False

        icon_dir = bytearray()
        icon_dir.extend(struct.pack("<HHH", 0, 1, count))
        payload_chunks = []
        payload_offset = 6 + count * 16
        for entry_index in range(count):
            entry_offset = 6 + entry_index * 14
            entry = group_data[entry_offset:entry_offset + 14]
            if len(entry) < 14:
                return False
            width, height, color_count, reserved_byte, planes, bit_count, bytes_in_res, resource_id = struct.unpack(
                "<BBBBHHIH", entry
            )
            payload = icon_payloads.get((resource_id, _group_language)) or icon_payloads.get((resource_id, None))
            if payload is None:
                return False
            icon_dir.extend(struct.pack("<BBBBHHII", width, height, color_count, reserved_byte, planes, bit_count, len(payload), payload_offset))
            payload_chunks.append(payload)
            payload_offset += len(payload)

        with destination_ico.open("wb") as handle:
            handle.write(icon_dir)
            for payload in payload_chunks:
                handle.write(payload)
        return destination_ico.exists() and destination_ico.stat().st_size > 0
    except (OSError, ValueError, struct.error):
        return False


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
            "if ($icon -eq $null) { exit 2 }; "
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
    app_dir = project_root / "App"
    appinfo_dir = app_dir / "AppInfo"
    launcher_dir = appinfo_dir / "Launcher"
    package_dir = app_dir / project.package_name
    data_dir = project_root / "Data"
    settings_dir = data_dir / "settings"
    other_help_dir = project_root / "Other" / "Help"
    help_images_dir = other_help_dir / "Images"
    other_source_dir = project_root / "Other" / "Source"

    for path in (launcher_dir, package_dir, settings_dir, help_images_dir, other_source_dir):
        ensure_empty_or_create(path)

    if project.copy_app_files:
        copy_application_folder(app_exe, package_dir, project_root=project_root)
    else:
        shutil.copy2(app_exe, package_dir / app_exe.name)

    create_portableapps_icons(project, appinfo_dir, app_exe)

    (appinfo_dir / "appinfo.ini").write_text(build_appinfo_ini(project), encoding="utf-8")
    (launcher_dir / f"{project.portable_name}.ini").write_text(build_launcher_ini(project), encoding="utf-8")
    installer_text = build_installer_ini(project)
    if installer_text:
        (appinfo_dir / "installer.ini").write_text(installer_text + "\n", encoding="utf-8")
    else:
        installer_path = appinfo_dir / "installer.ini"
        if installer_path.exists():
            installer_path.unlink()

    create_help_images(help_images_dir, appinfo_dir)
    create_launcher_template_assets(launcher_dir)

    (project_root / "help.html").write_text(build_help_html(project), encoding="utf-8")
    (other_source_dir / "Readme.txt").write_text(build_readme(project), encoding="utf-8")
    return project_root
