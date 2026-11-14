# Application Packaging Automation V2

## Launch
Run `Start-PackagingAutomation.bat` on a Windows machine with Windows PowerShell 5.1 and the required packaging tools/templates.

## V2 changes
- Added a Configuration tab with repeatable Additional Installers and configuration-action command generation.
- Added direct editable command areas for all six PSADT lifecycle phases.
- Configuration actions are inserted into the selected phase editor for review before building.
- Added registry-based EXE uninstall command generation and pre-install older-version removal action.
- Added Refresh / New Package control.
- Removed the visible PID Document and QA tabs while retaining automatic PID generation during Build & Output.
- Preserved the existing package, template settings, build, run log and Non-Automated IntuneWin workflows.

## Templates
The `Templates` folder is intentionally empty in this archive, matching the supplied V1 baseline. Add your existing PSADT templates, PID document template, Silent Switch Finder module, PowerRun and IntuneWinAppUtil files as required by your environment.

## Important test notes
This package was updated in a non-Windows environment, so the WinForms UI and PSADT build workflow could not be executed here. Test first with a non-production application and verify generated commands against the PSADT template/version you use. Review registry-based uninstall commands before deployment, especially where a quiet uninstall string is not registered.


## Test Fix 2 notes
Main installer selection now populates available metadata after selection. For additional installers, select files located under the configured Source Folder so they can be included in the package. Configuration preview aggregates added actions; phase editors remain the final review/edit locations. Registry uninstall detection and MSI MST behavior must be tested on Windows with the appropriate templates/toolkit present.

## Follow-up Fixes (2026-10-23)
- Clear stale MSI/EXE metadata before populating metadata for a newly selected EXE.
- Prefer exact installed-app name/version matches during uninstall detection and avoid silently choosing a different version.
- Populate detected EXE uninstall path and arguments for review; keep the generated registry-based uninstall action in the Uninstallation editor.
- Extend portable path conversion to common Public Documents paths.
- Preview saved configuration actions plus the current selection, even if the current command matches a saved action.
- Prevent adding the same generated command repeatedly to the same phase.
- Improve firewall-rule input parsing and validate port values/ranges; support program paths or ports, direction, protocol, and Allow/Block.

The Windows PowerShell parser, WinForms UI, MSI COM access, Silent Switch Finder integration, and package build must still be validated on a Windows test machine.
### Additional Installer uninstall configuration
Uninstallation is optional for each Additional Installer. The tool attempts registry detection when an installer is selected. If no uninstall string is found, it displays a status message and allows the installer to be saved without uninstall details. Add uninstall command and parameters only when that component requires uninstallation.


### TestFix7 - Additional Installer EXE/MSI uninstall detection
Uninstall detection supports both EXE and MSI installers. MSI detection prioritizes ProductCode; EXE detection searches uninstall registry entries using application metadata and filename-derived names. If an EXE has not been installed on the current machine, a registry uninstall string may not exist, so manual entry remains available when uninstall is required. Uninstall remains optional.


## TestFix13
- Preserves user-edited application metadata after initial installer metadata detection.
- Browse for Remove file disables shortcut dereferencing so .lnk paths are selected literally.
