# CCTV IP Toolkit

[![Latest Release](https://img.shields.io/github/v/release/theLostPing/cctv-ip-toolkit?style=flat-square&color=0969da&label=Download%20EXE)](https://github.com/theLostPing/cctv-ip-toolkit/releases/latest)
[![Code Signing](https://img.shields.io/badge/Code%20Signing-IN%20PROGRESS-f59e0b?style=flat-square)](https://cctv.thelostping.net/#transparency)
[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20me%20a%20coffee-%E2%98%95-FFDD00?style=flat-square&labelColor=1a1a2e)](https://buymeacoffee.com/thelostping)

> **⬇ Download the EXE from the [latest GitHub release](https://github.com/theLostPing/cctv-ip-toolkit/releases/latest).** Docs + install notes at [cctv.thelostping.net](https://cctv.thelostping.net). This repo is the source mirror — published so every binary can be audited or rebuilt from scratch.

Windows GUI toolkit for field techs programming Axis, Bosch, and Hanwha/Wisenet IP cameras.

> ⚠ **CODE SIGNING IS IN PROGRESS.** Azure Trusted Signing application submitted, identity verification underway. Signed builds ship as soon as Microsoft completes the review. Until then, the `.exe` isn't Microsoft-code-signed and Windows SmartScreen will warn on first run — if that's a dealbreaker, compile it yourself from this repo.

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
pyinstaller --onefile --icon=app.ico --name=CCTVIPToolkit axis_toolkit_v3.py
```

Or use the included `build.bat` — it handles dependency install + PyInstaller bundling and drops `CCTVIPToolkit.exe` into `dist/`.

## Transparency

- No telemetry, analytics, or phone-home code.
- No camera stream capture, recording, or upload.
- No credentials transmitted anywhere except the specific camera you're programming.
- No auto-updates that fetch remote code. Updates are manual.

Network access is needed to talk to cameras on the LAN (VAPIX, RCP+, SUNAPI, ICMP, ARP) and to query switches via LLDP/CDP for port discovery. That's it.

## Feedback

Bug reports and feature requests: [cctv.thelostping.net/feedback](https://cctv.thelostping.net/feedback) (goes to the author as email) or open a GitHub issue here.

## License / reuse

Created by **Brian Preston**. Free to use, modify, audit, redistribute. Attribution appreciated.

If this saved you a service call or an afternoon, [buy me a coffee](https://buymeacoffee.com/thelostping) &#9749;
