# CCTV IP Toolkit

[![Latest Release](https://img.shields.io/github/v/release/theLostPing/cctv-ip-toolkit?style=flat-square&color=0969da&label=version)](https://github.com/theLostPing/cctv-ip-toolkit/releases/latest)
[![Code Signing](https://img.shields.io/badge/Code%20Signing-IN%20PROGRESS-f59e0b?style=flat-square)](https://cctv.thelostping.net/#transparency)
[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20me%20a%20coffee-%E2%98%95-FFDD00?style=flat-square&labelColor=1a1a2e)](https://buymeacoffee.com/thelostping)

[**📥 Download the latest release**](https://github.com/theLostPing/cctv-ip-toolkit/releases/latest) — installer or bare exe, both ship with every tagged release.

Docs, install notes, and the SmartScreen walkthrough live at [cctv.thelostping.net](https://cctv.thelostping.net) (no binary hosted there — GitHub is the source of truth for downloads).

Windows GUI toolkit for field techs programming Axis, Bosch, and Hanwha/Wisenet IP cameras.

> ⚠ **CODE SIGNING IS IN PROGRESS.** [SignPath Foundation](https://signpath.io/foundation) application submitted (free code-signing program for OSS projects). Signed builds ship as soon as the application is approved. Until then, the `.exe` isn't Microsoft-recognized and Windows SmartScreen will warn on first run — if that's a dealbreaker, compile it yourself from this repo.

## Two ways to install

| File on the release page | Use it for |
|---|---|
| `CCTVIPToolkit-Setup-vX.Y.Z.exe` | **Recommended.** Inno Setup installer: Start Menu entry, optional Desktop shortcut, uninstaller in Add/Remove Programs, one-click in-app updates. |
| `CCTVIPToolkit.exe` | Bare exe. Runs from anywhere — USB stick, Downloads folder, wherever. No install, no Start Menu entry, manual updates. |

Both are built from the same source per release.

## Features

- ARP + LLDP network discovery (know which switch port each camera is on)
- Built-in password list manager, tries known creds in bulk
- Step-by-step programming wizard with live checklist and firmware capture
- Additional-users bulk add during initial programming
- Cross-subnet programming without juggling network adapters
- Firmware version / model / serial / MAC capture to CSV / XLSX
- Ping + auth verify loop after every change
- Snapshot grabber for proof-of-install bundles
- DPI-aware UI; dialogs clamp to the active monitor on multi-display rigs
- Remembers per-job preferences in an INI file
- One-click in-app update from the GitHub release (4.4.0+)

## Where your data lives

Two folders, on purpose: the program is disposable, your data isn't.

| | Path | What | Why |
|---|---|---|---|
| **Config** | `%APPDATA%\CCTVIPToolkit\` | `passwords.json`, `cameras.json`, `settings.ini`, `additional_users.json` | Private per-user profile; follows you across reinstalls and machine moves. Never shipped inside the .exe. |
| **Exports** | `%USERPROFILE%\Documents\CCTV Toolkit\` *(default, user-configurable)* | `programmed_cameras.csv`, `ping_results.csv`, `found_passwords.csv`, `screenshots/`, `triplett/` | Visible, browseable, zippable for the site report. Change it from **File → Settings → Export Folder** to point at a per-site folder if you want each job kept separate. |

Upgrades — bare exe drop-in OR installer over-install — never touch either folder. Your password list and camera history survive every update.

Quick jumps: **File → Open Export Folder** / **File → Open Config Folder**.

If you used v4.0, a one-time migration on first launch copies your old `./data/` folder into the two new locations and leaves the original untouched as a backup.

## Supported brands

| Brand | Protocol(s) |
|---|---|
| Axis | VAPIX + `param.cgi` fallback for older firmware |
| Bosch | RCP+ with service-account auth |
| Hanwha / Wisenet | SUNAPI |

Auto-detected via MAC OUI + HTTP fingerprint. Manual override available when auto-detect guesses wrong.

## Build from source

Windows with Python 3.10 or newer:

```
pip install requests pillow openpyxl pyinstaller
pyinstaller --onefile --icon=app.ico --name=CCTVIPToolkit cctv_toolkit.py
```

Or use the included `build.bat` — handles dependency install, PyInstaller bundling, and (4.4.0+) auto-installs Inno Setup via winget if missing, then compiles `installer.iss` to produce both:

```
dist/CCTVIPToolkit.exe                    (bare exe)
dist/CCTVIPToolkit-Setup-vX.Y.Z.exe       (installer)
```

The installer source is `installer.iss` in the repo root. Stable AppId GUID across all versions means future installers always upgrade prior installs in place.

## Transparency

- No telemetry, analytics, or phone-home code.
- No camera stream capture, recording, or upload.
- No credentials transmitted anywhere except the specific camera you're programming.
- **Updates are user-initiated.** The app can check the GitHub releases API at startup (silent) or via Help → Check for Updates. If a newer version is published, you get a dialog with release notes and an explicit *Install vX.Y.Z now* button. Nothing downloads or installs without your click. The "Remind me later" button suppresses the nag for that specific version. Network destinations: only `api.github.com` and `objects.githubusercontent.com` (the asset CDN).

Network access is otherwise needed to talk to cameras on the LAN (VAPIX, RCP+, SUNAPI, ICMP, ARP) and to query switches via LLDP/CDP for port discovery. That's it.

## Feedback

Bug reports and feature requests: [cctv.thelostping.net/feedback](https://cctv.thelostping.net/feedback) (goes to the author as email) or open a GitHub issue here.

## License / reuse

Created by **Brian Preston**. Free to use, modify, audit, redistribute. Attribution appreciated.

If this saved you a service call or an afternoon, [buy me a coffee](https://buymeacoffee.com/thelostping) ☕
