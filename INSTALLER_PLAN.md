# Windows installer plan — childRegistry

Target repository: https://github.com/nchotenin-tech/childRegistry

## Release pipeline

Source code → GitHub Actions Windows runner (Python 3.12 x64) → tests → PyInstaller onedir/windowed → Inno Setup → ChildRegistry-Setup-x.y.z.exe → GitHub Release.

End users install the Setup executable and use a desktop/start-menu shortcut. They do not install Python, pip or Git. First release should target Windows 10/11 x64 and be tested on a clean machine without Python. Start with manual version releases; do not implement silent automatic updates initially.

## Required changes before packaging

- Separate bundled resources from writable user data. The current application uses ROOT for template discovery; verify that frozen builds can find templateRegistry.xlsx inside the bundle.
- Ship a sanitized template with no student records, no cached student values in Entry, and no personal information in hidden sheets, comments or other cells. Do not ship the current populated workbook unchanged.
- Create the initial writable database on first launch only, in a user-owned folder such as Documents/ChildRegistry. Never replace an existing database on launch, upgrade or uninstall.
- Keep data-source selection and persistence using QSettings. The installer must not capture the developer's chosen database path.
- Keep database-specific .settings.json and backups next to the user's workbook. Exclude user data from the installer and Git repository.
- Make all tests self-contained: several current tests read studentRegistry.xlsx and assume the sample cohort. Move sample generation to fixtures using temporary workbooks before running CI; do not solve this by committing the active database.
- Pin and record tested runtime/build dependencies for reproducible releases.

## Repository contents

Include oral_registry/, run.py, tests/, requirements files, documentation, a sanitized resource template, packaging/ChildRegistry.spec, packaging/installer.iss and .github/workflows/windows-release.yml.

Exclude studentRegistry.xlsx, all operational Excel workbooks, *.settings.json, backups/, .venv/, .test-*/, tmp/, .pytest_cache/, build/, dist/, installer output and extracted files containing personal data. Review git status and staged files explicitly before every first push. Do not use git add . against the current workspace until exclusions and sanitization are verified.

## Installer decisions

- Use PyInstaller onedir; bundle that directory in Inno Setup. Build on Windows.
- Install per Windows user into LocalAppData/Programs/ChildRegistry; PrivilegesRequired=lowest.
- Use a permanent AppId across releases and a visible AppVersion.
- Add Start-menu and optional desktop shortcuts; support launching after installation.
- Upgrades replace application binaries only. Close the running application before replacing files.
- Uninstall removes application files, not workbooks, settings or backups containing user data.
- Consider executable and installer code signing for distribution; unsigned software can show Windows reputation warnings.

## GitHub Actions and releases

Use a workflow triggered by workflow_dispatch initially; after verification, allow version tags such as v1.0.0. Run tests before packaging. Upload a build artifact for review, then attach the approved installer and SHA-256 checksum to a GitHub Release. Limit release-write permissions to the release job. Store signing credentials in secrets, never in the repository.

If the repository is private, users need access to download its release assets. Confirm the intended distribution audience before publishing.

## Acceptance checks

1. Install and launch on a clean Windows machine without Python or development dependencies.
2. All Thai labels, icons, charts and workbook mappings display correctly.
3. First launch creates or selects a writable database and remembers its location.
4. Save, reopen, edit, delete with confirmation, backup and restore work.
5. Installing a newer version retains the database, settings and backups.
6. Uninstall does not remove user data.
7. Installer contains no current student records or developer-specific paths.
8. Verify installer version and checksum against the GitHub release.

## Current status

Packaging implemented with a sanitized template, first-run database creation, self-contained tests, a windowed executable smoke test, and Inno Setup. Requires Windows 10 build 19041 or later x64. Incompatible Poppler ICU DLLs are excluded; Qt uses Windows ICU. A separate clean-machine deployment check remains recommended.

References:
- https://pyinstaller.org/en/stable/operating-mode.html
- https://jrsoftware.org/ishelp/topic_setup_privilegesrequired.htm
- https://docs.github.com/en/repositories/releasing-projects-on-github/about-releases
