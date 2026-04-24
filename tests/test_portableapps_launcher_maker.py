import unittest
import uuid
from pathlib import Path
import tkinter as tk

from PIL import Image

from app.portableapps_launcher_maker import (
    LAUNCHER_TEMPLATE_FILENAMES,
    LauncherProject,
    HELP_IMAGE_FILENAMES,
    PortableAppsLauncherMaker,
    build_appinfo_ini,
    build_help_html,
    build_installer_ini,
    build_launcher_ini,
    clean_identifier,
    create_launcher_project,
    detect_app_name_from_exe,
    extract_embedded_icon,
    make_fallback_icon,
)


class PortableAppsLauncherMakerTests(unittest.TestCase):
    def test_clean_identifier_uses_pascal_case_words(self) -> None:
        self.assertEqual(clean_identifier("movie nfo-editor!"), "MovieNfoEditor")

    def test_detect_app_name_from_exe_humanizes_stem(self) -> None:
        self.assertEqual(detect_app_name_from_exe(r"C:\Tools\my-app.exe"), "my app")

    def test_apply_selected_app_exe_updates_previous_auto_detected_values(self) -> None:
        root = tk.Tk()
        root.withdraw()
        try:
            app = PortableAppsLauncherMaker(root)
            app.apply_selected_app_exe(r"C:\Apps\FirstTool.exe")
            self.assertEqual("FirstTool.exe", Path(app.vars["app_exe"].get()).name)
            self.assertEqual("FirstTool", app.vars["app_name"].get())
            self.assertEqual("FirstTool", app.vars["package_name"].get())
            self.assertEqual("FirstTool portable launcher", app.vars["description"].get())
            self.assertEqual("FirstToolPortable.exe", app.vars["control_start"].get())

            app.apply_selected_app_exe(r"C:\Apps\SecondTool.exe")
            self.assertEqual("SecondTool", app.vars["app_name"].get())
            self.assertEqual("SecondTool", app.vars["package_name"].get())
            self.assertEqual("SecondTool portable launcher", app.vars["description"].get())
            self.assertEqual("SecondToolPortable.exe", app.vars["control_start"].get())
        finally:
            root.destroy()

    def test_apply_selected_app_exe_preserves_customized_values(self) -> None:
        root = tk.Tk()
        root.withdraw()
        try:
            app = PortableAppsLauncherMaker(root)
            app.apply_selected_app_exe(r"C:\Apps\FirstTool.exe")
            app.vars["app_name"].set("Custom Name")
            app.vars["package_name"].set("CustomPackage")
            app.vars["description"].set("Custom description")
            app.vars["control_start"].set("CustomPortable.exe")

            app.apply_selected_app_exe(r"C:\Apps\SecondTool.exe")
            self.assertEqual("Custom Name", app.vars["app_name"].get())
            self.assertEqual("CustomPackage", app.vars["package_name"].get())
            self.assertEqual("Custom description", app.vars["description"].get())
            self.assertEqual("CustomPortable.exe", app.vars["control_start"].get())
        finally:
            root.destroy()

    def test_make_fallback_icon_returns_image(self) -> None:
        image = make_fallback_icon("Sample App")
        self.assertEqual(image.size, (256, 256))

    def test_builds_appinfo_launcher_and_installer_ini(self) -> None:
        project = LauncherProject(
            app_name="Sample App",
            package_name="SampleApp",
            publisher="Sample Publisher",
            homepage="https://example.com",
            category="Utilities",
            description="Sample description",
            version="2.3.4.5",
            display_version="2.3 Release 5",
            app_exe=r"C:\Apps\Sample.exe",
            output_dir=r"C:\Out",
            command_line="--portable",
            wait_for_other_instances=False,
            min_os="XP",
            max_os="7",
            run_as_admin="try",
            refresh_shell_icons="both",
            hide_command_line_window=True,
            no_spaces_in_path=True,
            supports_unc="warn",
            activate_java="find",
            activate_xml=True,
            live_mode_copy_app=True,
            live_mode_copy_data=True,
            files_move=r"settings\config.ini=%PAL:AppDir%\SampleApp",
            directories_move=r"settings=%APPDATA%\SampleApp",
            license_shareable=False,
            license_open_source=True,
            license_freeware=False,
            license_commercial_use=False,
            special_plugins="NONE",
            dependency_uses_ghostscript="optional",
            dependency_uses_java="no",
            dependency_uses_dotnet_version="4.8",
            dependency_requires_64bit_os="yes",
            dependency_requires_portable_app="CommonFiles\\Java",
            dependency_requires_admin="no",
            control_icons="1",
            control_extract_icon=r"App\SampleApp\Sample.exe",
            control_base_app_id=r"%BASELAUNCHERPATH%\App\SampleApp\Sample.exe",
            control_exit_exe=r"App\SampleApp\Sample.exe",
            control_exit_parameters="-exit",
            association_file_types="sample,smp",
            association_file_type_command_line='"%1"',
            association_protocols="sample",
            association_protocol_command_line="--url=%1",
            association_send_to=True,
            association_send_to_command_line='-multiplefiles "%1"',
            association_shell=True,
            association_shell_command="/idlist,%I,%L",
            file_type_icons="sample=app\nsmp=custom",
            installer_close_exe="Sample.exe",
            installer_close_name="Sample App",
            include_installer_source=True,
            remove_app_directory=True,
            remove_data_directory=False,
            optional_components_enabled=True,
            main_section_title="Sample App Portable [Required]",
            optional_section_title="Additional Languages",
            optional_section_selected_install_type="Multilingual",
            optional_section_preselected="true",
            installer_languages="ENGLISH=true\nGERMAN=true",
            preserve_directories=r"PreserveDirectory1=App\SampleApp\plugins",
            remove_files=r"RemoveFile1=App\SampleApp\*.lang",
        )

        appinfo = build_appinfo_ini(project)
        launcher = build_launcher_ini(project)
        installer = build_installer_ini(project)

        self.assertIn("Name=Sample App Portable", appinfo)
        self.assertIn("AppID=SampleAppPortable", appinfo)
        self.assertIn("Start=SampleAppPortable.exe", appinfo)
        self.assertIn("PackageVersion=2.3.4.5", appinfo)
        self.assertIn("DisplayVersion=2.3 Release 5", appinfo)
        self.assertIn("Shareable=false", appinfo)
        self.assertIn("OpenSource=true", appinfo)
        self.assertIn("Freeware=false", appinfo)
        self.assertIn("CommercialUse=false", appinfo)
        self.assertIn("[SpecialPaths]", appinfo)
        self.assertIn("Plugins=NONE", appinfo)
        self.assertIn("[Dependencies]", appinfo)
        self.assertIn("UsesGhostscript=optional", appinfo)
        self.assertIn("UsesDotNetVersion=4.8", appinfo)
        self.assertIn("Requires64bitOS=yes", appinfo)
        self.assertIn("RequiresPortableApp=CommonFiles\\Java", appinfo)
        self.assertIn("[Control]", appinfo)
        self.assertIn(r"ExtractIcon=App\SampleApp\Sample.exe", appinfo)
        self.assertIn(r"BaseAppID=%BASELAUNCHERPATH%\App\SampleApp\Sample.exe", appinfo)
        self.assertIn("ExitParameters=-exit", appinfo)
        self.assertIn("[Associations]", appinfo)
        self.assertIn("FileTypes=sample,smp", appinfo)
        self.assertIn('FileTypeCommandLine="%1"', appinfo)
        self.assertIn("Protocols=sample", appinfo)
        self.assertIn("SendTo=true", appinfo)
        self.assertIn("Shell=true", appinfo)
        self.assertIn("[FileTypeIcons]", appinfo)
        self.assertIn("sample=app", appinfo)
        self.assertIn("smp=custom", appinfo)
        self.assertIn(r"ProgramExecutable=SampleApp\Sample.exe", launcher)
        self.assertIn("CommandLineArguments=--portable", launcher)
        self.assertIn("WaitForOtherInstances=false", launcher)
        self.assertIn("MinOS=XP", launcher)
        self.assertIn("MaxOS=7", launcher)
        self.assertIn("RunAsAdmin=try", launcher)
        self.assertIn("RefreshShellIcons=both", launcher)
        self.assertIn("HideCommandLineWindow=true", launcher)
        self.assertIn("NoSpacesInPath=true", launcher)
        self.assertIn("SupportsUNC=warn", launcher)
        self.assertIn("Java=find", launcher)
        self.assertIn("XML=true", launcher)
        self.assertIn("[LiveMode]", launcher)
        self.assertIn("CopyApp=true", launcher)
        self.assertIn("CopyData=true", launcher)
        self.assertIn("[FilesMove]", launcher)
        self.assertIn(r"settings\config.ini=%PAL:AppDir%\SampleApp", launcher)
        self.assertIn("[DirectoriesMove]", launcher)
        self.assertIn(r"settings=%APPDATA%\SampleApp", launcher)
        self.assertIn("[CheckRunning]", installer)
        self.assertIn("CloseEXE=Sample.exe", installer)
        self.assertIn("[Source]", installer)
        self.assertIn("IncludeInstallerSource=true", installer)
        self.assertIn("[MainDirectories]", installer)
        self.assertIn("RemoveAppDirectory=true", installer)
        self.assertIn("[OptionalComponents]", installer)
        self.assertIn("OptionalComponents=true", installer)
        self.assertIn("OptionalSectionSelectedInstallType=Multilingual", installer)
        self.assertIn("[Languages]", installer)
        self.assertIn("ENGLISH=true", installer)
        self.assertIn("[DirectoriesToPreserve]", installer)
        self.assertIn(r"PreserveDirectory1=App\SampleApp\plugins", installer)
        self.assertIn("[FilesToRemove]", installer)
        self.assertIn(r"RemoveFile1=App\SampleApp\*.lang", installer)

    def test_omits_installer_ini_when_empty(self) -> None:
        project = LauncherProject(
            app_name="Sample App",
            package_name="SampleApp",
            publisher="Sample Publisher",
            homepage="https://example.com",
            category="Utilities",
            description="Sample description",
            version="1.0.0.0",
            display_version="1.0.0",
            app_exe=r"C:\Apps\Sample.exe",
            output_dir=r"C:\Out",
        )

        installer = build_installer_ini(project)

        self.assertEqual("", installer)

    def test_build_help_html_references_help_image(self) -> None:
        project = LauncherProject(
            app_name="Sample App",
            package_name="SampleApp",
            publisher="Sample Publisher",
            homepage="https://example.com",
            category="Utilities",
            description="Sample description",
            version="1.0.0.0",
            display_version="1.0.0",
            app_exe=r"C:\Apps\Sample.exe",
            output_dir=r"C:\Out",
        )

        help_html = build_help_html(project)

        self.assertIn("<title>Sample App Portable Help</title>", help_html)
        self.assertIn("<h1 class=\"hastagline\">Sample App Portable Help</h1>", help_html)
        self.assertIn("Learn more about Sample App...", help_html)
        self.assertIn("Go to the Sample App Portable Homepage", help_html)

    def test_omits_empty_associations_and_file_type_icons_sections(self) -> None:
        project = LauncherProject(
            app_name="Sample App",
            package_name="SampleApp",
            publisher="Sample Publisher",
            homepage="https://example.com",
            category="Utilities",
            description="Sample description",
            version="1.0.0.0",
            display_version="1.0.0",
            app_exe=r"C:\Apps\Sample.exe",
            output_dir=r"C:\Out",
        )

        appinfo = build_appinfo_ini(project)

        self.assertNotIn("[Associations]", appinfo)
        self.assertNotIn("[FileTypeIcons]", appinfo)

    def test_keeps_associations_section_when_flags_are_enabled(self) -> None:
        project = LauncherProject(
            app_name="Sample App",
            package_name="SampleApp",
            publisher="Sample Publisher",
            homepage="https://example.com",
            category="Utilities",
            description="Sample description",
            version="1.0.0.0",
            display_version="1.0.0",
            app_exe=r"C:\Apps\Sample.exe",
            output_dir=r"C:\Out",
            association_send_to=True,
        )

        appinfo = build_appinfo_ini(project)

        self.assertIn("[Associations]", appinfo)
        self.assertIn("SendTo=true", appinfo)

    def test_create_launcher_project_writes_portableapps_structure(self) -> None:
        temp_path = Path(__file__).parent / "tmp" / f"portableapps-maker-{uuid.uuid4().hex}"
        temp_path.mkdir(parents=True, exist_ok=True)

        exe = temp_path / "Sample.exe"
        exe.write_bytes(b"fake exe")
        sibling_file = temp_path / "helper.dll"
        sibling_file.write_bytes(b"fake dll")
        sibling_dir = temp_path / "plugins"
        sibling_dir.mkdir()
        sibling_plugin = sibling_dir / "plugin.dat"
        sibling_plugin.write_bytes(b"fake plugin")
        icon = temp_path / "Sample.ico"
        Image.new("RGBA", (256, 256), "#216b52").save(icon, sizes=[(16, 16), (32, 32), (48, 48), (256, 256)])
        launcher = temp_path / "Launcher.exe"
        launcher.write_bytes(b"fake launcher")

        project = LauncherProject(
            app_name="Sample App",
            package_name="SampleApp",
            publisher="Sample Publisher",
            homepage="https://example.com",
            category="Utilities",
            description="Sample description",
            version="2.3.4.5",
            display_version="2.3 Release 5",
            app_exe=str(exe),
            output_dir=str(temp_path / "out"),
            icon_source=str(icon),
        )

        project_root = create_launcher_project(project)

        self.assertFalse((project_root / "SampleAppPortable.exe").exists())
        self.assertTrue((project_root / "App" / "SampleApp" / "Sample.exe").exists())
        self.assertTrue((project_root / "App" / "SampleApp" / "helper.dll").exists())
        self.assertTrue((project_root / "App" / "SampleApp" / "plugins" / "plugin.dat").exists())
        self.assertTrue((project_root / "App" / "AppInfo" / "appicon.ico").exists())
        self.assertTrue((project_root / "App" / "AppInfo" / "appicon_16.png").exists())
        self.assertTrue((project_root / "App" / "AppInfo" / "appicon_32.png").exists())
        self.assertTrue((project_root / "App" / "AppInfo" / "appicon_75.png").exists())
        self.assertTrue((project_root / "App" / "AppInfo" / "appicon_128.png").exists())
        self.assertTrue((project_root / "App" / "AppInfo" / "appicon_256.png").exists())
        self.assertTrue((project_root / "App" / "AppInfo" / "appinfo.ini").exists())
        self.assertFalse((project_root / "App" / "AppInfo" / "installer.ini").exists())
        self.assertTrue((project_root / "App" / "AppInfo" / "Launcher" / "SampleAppPortable.ini").exists())
        self.assertTrue((project_root / "App" / "AppInfo" / "Launcher" / "Splash.jpg").exists())
        self.assertTrue((project_root / "help.html").exists())
        self.assertTrue((project_root / "Other" / "Help" / "Images" / "appicon_128.png").exists())
        self.assertTrue((project_root / "Other" / "Help" / "Images" / "Favicon.ico").exists())
        self.assertTrue((project_root / "Other" / "Help" / "Images" / "Help_Background_Header.png").exists())
        self.assertTrue((project_root / "Other" / "Help" / "Images" / "Help_Background_Footer.png").exists())
        self.assertTrue((project_root / "Other" / "Help" / "Images" / "Help_Logo_Top.png").exists())
        self.assertTrue((project_root / "Other" / "Help" / "Images" / "Donation_Button.png").exists())
        self.assertTrue((project_root / "Data" / "settings").is_dir())
        self.assertTrue((project_root / "Other" / "Source" / "Readme.txt").exists())
        expected_help_html = (Path(__file__).parents[1] / "app" / "help_template" / "help.html").read_text(encoding="utf-8")
        self.assertIn("Sample App Portable Help", (project_root / "help.html").read_text(encoding="utf-8"))
        self.assertIn("[DESCRIPTION OF APP FUNCTION HERE]", expected_help_html)
        for filename in HELP_IMAGE_FILENAMES:
            expected = (Path(__file__).parents[1] / "app" / "help_template" / "Images" / filename).read_bytes()
            actual = (project_root / "Other" / "Help" / "Images" / filename).read_bytes()
            self.assertEqual(actual, expected)
        for filename in LAUNCHER_TEMPLATE_FILENAMES:
            expected = (Path(__file__).parents[1] / "app" / "help_template" / "Launcher" / filename).read_bytes()
            actual = (project_root / "App" / "AppInfo" / "Launcher" / filename).read_bytes()
            self.assertEqual(actual, expected)

        with Image.open(project_root / "App" / "AppInfo" / "appicon_75.png") as generated:
            self.assertEqual(generated.size, (75, 75))

    def test_create_launcher_project_writes_installer_ini_when_configured(self) -> None:
        temp_path = Path(__file__).parent / "tmp" / f"portableapps-installer-{uuid.uuid4().hex}"
        temp_path.mkdir(parents=True, exist_ok=True)

        exe = temp_path / "Sample.exe"
        exe.write_bytes(b"fake exe")
        icon = temp_path / "Sample.ico"
        Image.new("RGBA", (256, 256), "#216b52").save(icon, sizes=[(16, 16), (32, 32), (48, 48), (256, 256)])

        project = LauncherProject(
            app_name="Sample App",
            package_name="SampleApp",
            publisher="Sample Publisher",
            homepage="https://example.com",
            category="Utilities",
            description="Sample description",
            version="1.0.0.0",
            display_version="1.0.0",
            app_exe=str(exe),
            output_dir=str(temp_path / "out"),
            icon_source=str(icon),
            include_installer_source=True,
        )

        project_root = create_launcher_project(project)

        installer_ini = project_root / "App" / "AppInfo" / "installer.ini"
        self.assertTrue(installer_ini.exists())
        self.assertIn("IncludeInstallerSource=true", installer_ini.read_text(encoding="utf-8"))

    def test_extract_embedded_icon_accepts_icon_index(self) -> None:
        source_exe = Path(__file__).parents[1] / "dist" / "PortableAppsLauncherMaker" / "PortableAppsLauncherMaker.exe"
        if not source_exe.exists():
            self.skipTest("release executable is not built")

        temp_path = Path(__file__).parent / "tmp" / f"embedded-icon-{uuid.uuid4().hex}"
        temp_path.mkdir(parents=True, exist_ok=True)
        destination = temp_path / "icon.ico"

        self.assertTrue(extract_embedded_icon(source_exe, destination, icon_index=0))
        self.assertTrue(destination.exists())
        self.assertGreater(destination.stat().st_size, 0)


if __name__ == "__main__":
    unittest.main()
