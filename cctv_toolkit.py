#!/usr/bin/env python3
"""
CCTV IP Toolkit v3.0 - Windows GUI Application
Created by Brian Preston

Features:
- Integrated camera list editor (no external files needed)
- Built-in password list manager
- Step-by-step wizards with validation
- Settings file to remember preferences
- "Don't show again" options for warnings
"""

import sys
# Enable per-monitor DPI awareness BEFORE Tkinter is imported. Without this,
# Windows scales the whole UI up as a bitmap (fuzzy + cut-off widgets) when
# the user runs at 125%/150% display scaling. With it, Tk renders at the
# correct logical pixel size and ttk widgets grow to accommodate text.
if sys.platform == 'win32':
    try:
        import ctypes
        # Try per-monitor DPI v2 first (Win10 1703+)
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()  # Win Vista+ fallback
        except Exception:
            pass

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog, simpledialog
import threading
import time
import requests
from requests.auth import HTTPDigestAuth
import os
import csv
import json
from datetime import datetime
import queue
import re
import socket
import struct
from io import BytesIO
import configparser
import ftplib
import shutil
import urllib.request
from pathlib import Path

try:
    from PIL import Image, ImageTk, ImageDraw, ImageFont
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ============================================================================
# CONFIGURATION
# ============================================================================
APP_VERSION = "4.2.8"
GITHUB_LATEST_API = "https://api.github.com/repos/theLostPing/cctv-ip-toolkit/releases/latest"
GITHUB_RELEASES_PAGE = "https://github.com/theLostPing/cctv-ip-toolkit/releases/latest"
# In-app upgrade link routes through the fieldtoolkit.com tracker so upgrades
# show up in the same download analytics as the website. The tracker 302s
# straight to the GitHub release asset.
FIELDTOOLKIT_DOWNLOAD_URL = "https://fieldtoolkit.com/download.php?asset=CCTVIPToolkit.exe"

# v4.1 split: config (private, survives upgrades) vs. exports (visible, user-configurable).
#   CONFIG_DIR -> %APPDATA%\CCTVIPToolkit\       (passwords, cameras, settings)
#   EXPORT_DIR -> %USERPROFILE%\Documents\CCTV Toolkit\  (CSVs, screenshots, FTP pulls)
# Both are created on first launch; legacy ./data/ next to the .exe is auto-migrated once.
APP_NAME = "CCTVIPToolkit"

def _default_config_dir() -> Path:
    if sys.platform == 'win32':
        root = os.environ.get('APPDATA') or str(Path.home() / 'AppData' / 'Roaming')
    elif sys.platform == 'darwin':
        root = str(Path.home() / 'Library' / 'Application Support')
    else:
        root = os.environ.get('XDG_CONFIG_HOME') or str(Path.home() / '.config')
    return Path(root) / APP_NAME

def _default_export_dir() -> Path:
    # Documents folder if it resolves; otherwise home.
    docs = Path.home() / 'Documents'
    if not docs.exists():
        docs = Path.home()
    return docs / 'CCTV Toolkit'

CONFIG_DIR = _default_config_dir()
# EXPORT_DIR is mutable — settings may point it elsewhere after load. Initial value is the default;
# SettingsManager rebinds it via set_export_dir() during startup.
EXPORT_DIR = _default_export_dir()

# Config files (private, per-user)
SETTINGS_FILE         = str(CONFIG_DIR / "settings.ini")
CAMERAS_FILE          = str(CONFIG_DIR / "cameras.json")
PASSWORDS_FILE        = str(CONFIG_DIR / "passwords.json")
ADDITIONAL_USERS_FILE = str(CONFIG_DIR / "additional_users.json")

# Export files (visible, user-configurable). These are rebuilt whenever EXPORT_DIR changes
# via _rebind_export_paths() so the rest of the code can keep using the module-level names.
IMAGES_DIR = TRIPLETT_DIR = OUTPUT_CSV = PING_RESULTS = SUCCESSFUL_PASSWORDS = ""

def _rebind_export_paths():
    global IMAGES_DIR, TRIPLETT_DIR, OUTPUT_CSV, PING_RESULTS, SUCCESSFUL_PASSWORDS
    IMAGES_DIR           = str(EXPORT_DIR / "screenshots")
    TRIPLETT_DIR         = str(EXPORT_DIR / "triplett")
    OUTPUT_CSV           = str(EXPORT_DIR / "programmed_cameras.csv")
    PING_RESULTS         = str(EXPORT_DIR / "ping_results.csv")
    SUCCESSFUL_PASSWORDS = str(EXPORT_DIR / "found_passwords.csv")

_rebind_export_paths()

def _exe_dir() -> Path:
    # Directory of the running .exe (PyInstaller) or the .py script.
    return Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent

_MIGRATION_NOTE: list = []  # filled by _migrate_legacy_data(); surfaced as a one-time dialog

def _migrate_legacy_data():
    """One-time: if old ./data/ lives next to the .exe, split it into CONFIG_DIR and EXPORT_DIR.
    Copies (not moves) so the original stays as a belt-and-suspenders backup."""
    legacy = _exe_dir() / 'data'
    if not legacy.is_dir():
        return
    marker = CONFIG_DIR / '.migrated_from_data'
    if marker.exists():
        return
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    config_names = {'settings.ini', 'cameras.json', 'passwords.json', 'additional_users.json'}
    export_file_map = {
        'programmed_cameras.csv': 'programmed_cameras.csv',
        'ping_results.csv':       'ping_results.csv',
        'found_passwords.csv':    'found_passwords.csv',
    }
    export_dir_map = {
        'images':   'screenshots',
        'triplett': 'triplett',
    }
    copied = []
    for item in legacy.iterdir():
        try:
            if item.is_file() and item.name in config_names:
                dst = CONFIG_DIR / item.name
                if not dst.exists():
                    shutil.copy2(item, dst)
                    copied.append(('config', item.name))
            elif item.is_file() and item.name in export_file_map:
                dst = EXPORT_DIR / export_file_map[item.name]
                if not dst.exists():
                    shutil.copy2(item, dst)
                    copied.append(('export', item.name))
            elif item.is_dir() and item.name in export_dir_map:
                dst = EXPORT_DIR / export_dir_map[item.name]
                if not dst.exists():
                    shutil.copytree(item, dst)
                    copied.append(('export', item.name + '/'))
        except Exception as e:
            copied.append(('error', f'{item.name}: {e}'))
    # Record that we migrated, even if nothing copied (folder existed but was empty).
    try:
        marker.write_text(f"migrated at {datetime.now().isoformat()}\n")
    except Exception:
        pass
    if copied:
        _MIGRATION_NOTE.append((str(legacy), str(CONFIG_DIR), str(EXPORT_DIR), copied))

# Create dirs + run migration before any data manager touches disk.
CONFIG_DIR.mkdir(parents=True, exist_ok=True)
EXPORT_DIR.mkdir(parents=True, exist_ok=True)
_migrate_legacy_data()
TIMEOUT = 5

# Bosch camera constants
BOSCH_OUIS = [b'\x00\x07\x5f']
BOSCH_DEFAULT_IP = '192.168.0.1'
BOSCH_DEFAULT_USER = 'service'
RCP_CMD = {
    'ip':            0x007c,
    'subnet':        0x007d,
    'gateway':       0x007f,
    'dhcp':          0x00af,
    'mac':           0x00bc,
    'hw_ver':        0x002e,
    'sw_ver':        0x002f,
    'unit_name':     0x0024,
    'factory_reset': 0x09a0,
    'reboot':        0x0811,
    'password':      0x028b,  # P_STRING, num: 1=LIVE, 2=USER, 3=SERVICE
}

# Hanwha/Wisenet camera constants
HANWHA_OUIS = [b'\x00\x09\x18', b'\x00\x16\x6c', b'\x00\x09\x12', b'\xc4\xf1\xd1', b'\x9c\xdc\x71']
HANWHA_DEFAULT_IP = '192.168.1.100'
HANWHA_DEFAULT_USER = 'admin'
HANWHA_DEFAULT_PASSWORD = 'admin'  # Factory default before initial setup
HANWHA_LOCKOUT_COOLDOWN = 35  # seconds after 490 lockout

# Default settings
DEFAULT_SETTINGS = {
    'general': {
        'factory_ip': '192.168.0.90',
        'bosch_factory_ip': '192.168.0.1',
        'username': 'root',
        'android_ip': '',
        'first_run': 'true',
        'brand': 'axis',
        'interface_index': '',
        'export_dir': '',  # blank = use default Documents/CCTV Toolkit
        'last_seen_version': '',  # for "what's new" popup on first launch of a new version
        'last_dismissed_version': '',  # suppresses nag when user chose "remind me later"
    },
    'warnings': {
        'show_incomplete_camera_warning': 'true',
        'show_batch_test_explanation': 'true',
        'show_programming_intro': 'true',
        'show_ip_update_intro': 'true',
        'show_hash_warning': 'true'
    }
}

OUTPUT_CSV_HEADER = ['CameraName', 'IPAddress', 'SerialNumber', 'MACAddress', 'Model', 'Firmware', 'BuildingReportsLabel', 'Timestamp']


def _ensure_output_csv_header():
    """Create programmed_cameras.csv with current header, or migrate older headers
    in-place. Migrations supported (newest-first):
      - 7-col (no BuildingReportsLabel) → add blank cell before Timestamp
      - 6-col (no Firmware, no BuildingReportsLabel) → add both blanks"""
    if not os.path.exists(OUTPUT_CSV):
        os.makedirs(os.path.dirname(OUTPUT_CSV) or '.', exist_ok=True)
        with open(OUTPUT_CSV, 'w', newline='') as f:
            csv.writer(f).writerow(OUTPUT_CSV_HEADER)
        return
    try:
        with open(OUTPUT_CSV, 'r', newline='') as f:
            rows = list(csv.reader(f))
    except Exception:
        return
    if not rows:
        with open(OUTPUT_CSV, 'w', newline='') as f:
            csv.writer(f).writerow(OUTPUT_CSV_HEADER)
        return
    header = rows[0]
    if header == OUTPUT_CSV_HEADER:
        return
    # Migrate
    new_rows = [OUTPUT_CSV_HEADER]
    for row in rows[1:]:
        if len(row) == 7 and 'Firmware' in header:
            # 7-col: CameraName,IPAddress,SerialNumber,MACAddress,Model,Firmware,Timestamp
            new_rows.append(row[:6] + [''] + row[6:])
        elif len(row) == 6:
            # 6-col: CameraName,IPAddress,SerialNumber,MACAddress,Model,Timestamp
            new_rows.append(row[:5] + ['', ''] + row[5:])
        else:
            new_rows.append(row)
    with open(OUTPUT_CSV, 'w', newline='') as f:
        csv.writer(f).writerows(new_rows)


# ============================================================================
# SETTINGS MANAGER
# ============================================================================
class SettingsManager:
    def __init__(self):
        self.config = configparser.ConfigParser()
        self.load()
        self.apply_export_dir()

    def load(self):
        if os.path.exists(SETTINGS_FILE):
            self.config.read(SETTINGS_FILE)
        # Ensure all default sections/keys exist
        for section, values in DEFAULT_SETTINGS.items():
            if section not in self.config:
                self.config[section] = {}
            for key, val in values.items():
                if key not in self.config[section]:
                    self.config[section][key] = val
        self.save()

    def save(self):
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        with open(SETTINGS_FILE, 'w') as f:
            self.config.write(f)

    def apply_export_dir(self):
        """Point EXPORT_DIR at the user's configured folder (or the default) and rebind the
        module-level export paths so the rest of the code picks it up without further work."""
        global EXPORT_DIR
        configured = (self.config.get('general', 'export_dir', fallback='') or '').strip()
        target = Path(configured) if configured else _default_export_dir()
        try:
            target.mkdir(parents=True, exist_ok=True)
        except Exception:
            target = _default_export_dir()
            target.mkdir(parents=True, exist_ok=True)
        EXPORT_DIR = target
        _rebind_export_paths()
    
    def get(self, section, key):
        return self.config.get(section, key, fallback=DEFAULT_SETTINGS.get(section, {}).get(key, ''))
    
    def get_bool(self, section, key):
        return self.config.getboolean(section, key, fallback=True)
    
    def set(self, section, key, value):
        if section not in self.config:
            self.config[section] = {}
        self.config[section][key] = str(value)
        self.save()


# ============================================================================
# DATA MANAGERS
# ============================================================================
class CameraDataManager:
    """Manages camera list - stored as JSON internally"""
    def __init__(self):
        self.cameras = []
        self.load()
    
    def load(self):
        if os.path.exists(CAMERAS_FILE):
            try:
                with open(CAMERAS_FILE, 'r') as f:
                    self.cameras = json.load(f)
            except:
                self.cameras = []
        # Auto-clean duplicates on load
        removed = self.dedup_camera_list()
        if removed:
            self.save()
        return self.cameras
    
    def save(self):
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        with open(CAMERAS_FILE, 'w') as f:
            json.dump(self.cameras, f, indent=2)
    
    def add(self, camera):
        self.cameras.append(camera)
        self.save()
    
    def upsert(self, camera):
        """Add or update camera. Match by serial first, then MAC, then IP fallback.
        When updating, merges new data INTO existing — doesn't wipe user edits.
        Also removes any duplicate entries sharing the same serial or MAC."""
        def merge_into(existing, new_data):
            """Merge new_data into existing, only filling empty/placeholder fields"""
            for key, val in new_data.items():
                old_val = existing.get(key, '')
                if val and (not old_val or old_val == '(Auth Required)'):
                    existing[key] = val
            return existing
        
        def normalize_mac(mac):
            return mac.upper().replace(':', '').replace('-', '').strip() if mac else ''
        
        def dedup_after_merge(keep_index):
            """Remove any OTHER entries that share the same serial or MAC as the kept entry."""
            kept = self.cameras[keep_index]
            kept_serial = kept.get('serial', '')
            kept_mac = normalize_mac(kept.get('mac', ''))
            to_remove = []
            for i, cam in enumerate(self.cameras):
                if i == keep_index:
                    continue
                if kept_serial and cam.get('serial') == kept_serial:
                    to_remove.append(i)
                elif kept_mac and normalize_mac(cam.get('mac', '')) == kept_mac:
                    to_remove.append(i)
            for i in sorted(to_remove, reverse=True):
                del self.cameras[i]
            return len(to_remove)
        
        # 1. Match by serial (hardware identity)
        cam_serial = camera.get('serial')
        if cam_serial:
            for i, existing in enumerate(self.cameras):
                if existing.get('serial') == cam_serial:
                    merge_into(existing, camera)
                    dedup_after_merge(i)
                    return 'updated'
        
        # 2. Match by MAC if no serial match
        cam_mac = normalize_mac(camera.get('mac', ''))
        if cam_mac:
            for i, existing in enumerate(self.cameras):
                if normalize_mac(existing.get('mac', '')) == cam_mac:
                    merge_into(existing, camera)
                    dedup_after_merge(i)
                    return 'updated'
        
        # 3. Match by IP as last resort (only if existing has no serial/mac)
        cam_ip = camera.get('ip')
        if cam_ip:
            for i, existing in enumerate(self.cameras):
                if existing.get('ip') == cam_ip:
                    if not existing.get('serial') and not existing.get('mac'):
                        merge_into(existing, camera)
                        return 'updated'
        
        self.cameras.append(camera)
        # Clean up stale entries: if this new camera has identity (serial/MAC),
        # remove any existing entries at the same IP (or new_ip) with NO identity — 
        # those are stale auth-required ghosts of the same physical camera
        new_idx = len(self.cameras) - 1
        new_serial = camera.get('serial', '')
        new_mac = normalize_mac(camera.get('mac', ''))
        if new_serial or new_mac:
            new_ips = set()
            for key in ('ip', 'new_ip'):
                val = camera.get(key, '')
                if val:
                    new_ips.add(val)
            if new_ips:
                to_remove = []
                for i, existing in enumerate(self.cameras):
                    if i == new_idx:
                        continue
                    if not existing.get('serial') and not normalize_mac(existing.get('mac', '')):
                        existing_ips = set()
                        for key in ('ip', 'new_ip'):
                            val = existing.get(key, '')
                            if val:
                                existing_ips.add(val)
                        if new_ips & existing_ips:
                            to_remove.append(i)
                for i in sorted(to_remove, reverse=True):
                    if i < new_idx:
                        new_idx -= 1
                    del self.cameras[i]
        return 'added'
    
    def update(self, index, camera):
        if 0 <= index < len(self.cameras):
            self.cameras[index] = camera
            self.save()
    
    def delete(self, index):
        if 0 <= index < len(self.cameras):
            del self.cameras[index]
            self.save()
    
    def clear(self):
        self.cameras = []
        self.save()
    
    def get_all(self):
        return self.cameras
    
    def dedup_camera_list(self):
        """Remove duplicate entries sharing the same serial or MAC.
        When duplicates found, keep the entry with the most data (most non-empty fields),
        but merge any missing fields from removed entries first.
        Returns number of entries removed."""
        def normalize_mac(mac):
            return mac.upper().replace(':', '').replace('-', '').strip() if mac else ''
        
        def richness(cam):
            """Count non-empty meaningful fields — richer entry wins"""
            score = 0
            for key in ('name', 'hostname', 'ip', 'model', 'serial', 'mac', 'gateway', 'subnet', 'number'):
                val = cam.get(key, '')
                if val and val != '(Auth Required)':
                    score += 1
            return score
        
        def merge_into(target, source):
            """Fill empty fields in target from source"""
            for key, val in source.items():
                old_val = target.get(key, '')
                if val and (not old_val or old_val == '(Auth Required)'):
                    target[key] = val
        
        # Group by serial
        serial_groups = {}
        for i, cam in enumerate(self.cameras):
            s = cam.get('serial', '')
            if s:
                serial_groups.setdefault(s, []).append(i)
        
        # Group by MAC
        mac_groups = {}
        for i, cam in enumerate(self.cameras):
            m = normalize_mac(cam.get('mac', ''))
            if m:
                mac_groups.setdefault(m, []).append(i)
        
        to_remove = set()
        
        # For each group with duplicates, keep richest entry, merge others into it
        for group in list(serial_groups.values()) + list(mac_groups.values()):
            if len(group) <= 1:
                continue
            best = max(group, key=lambda i: richness(self.cameras[i]))
            for i in group:
                if i != best:
                    # Merge data from duplicate into the keeper before removing
                    merge_into(self.cameras[best], self.cameras[i])
                    to_remove.add(i)
        
        for i in sorted(to_remove, reverse=True):
            del self.cameras[i]
        
        return len(to_remove)
    
    def get_valid_for_programming(self):
        """Returns cameras with all required fields for programming"""
        return [c for c in self.cameras if c.get('ip') and c.get('gateway') and c.get('subnet') and not c.get('processed')]
    
    def get_valid_for_basic_ops(self):
        """Returns cameras with at least name and IP"""
        return [c for c in self.cameras if c.get('ip')]
    
    def mark_processed(self, index):
        if 0 <= index < len(self.cameras):
            self.cameras[index]['processed'] = True
            self.cameras[index]['status'] = 'done'
            self.save()
    
    def mark_failed(self, index, reason=''):
        if 0 <= index < len(self.cameras):
            self.cameras[index]['processed'] = False
            self.cameras[index]['status'] = 'failed'
            self.cameras[index]['fail_reason'] = reason
            self.save()
    
    def import_from_csv(self, filepath):
        """Import from CSV/TXT file"""
        imported = 0
        with open(filepath, 'r') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#'):
                    parts = line.split(',')
                    if len(parts) >= 2:
                        cam = {
                            'name': parts[0].strip(),
                            'ip': parts[1].strip(),
                            'gateway': parts[2].strip() if len(parts) > 2 else '',
                            'subnet': parts[3].strip() if len(parts) > 3 else '',
                            'model': parts[4].strip() if len(parts) > 4 else '',
                            'new_ip': parts[5].strip() if len(parts) > 5 else '',
                            'processed': False
                        }
                        self.cameras.append(cam)
                        imported += 1
        self.save()
        return imported
    
    def export_to_csv(self, filepath):
        """Export to CSV file"""
        with open(filepath, 'w', newline='') as f:
            w = csv.writer(f)
            w.writerow(['Camera Name', 'IP Address', 'Gateway', 'Subnet', 'Model', 'New IP', 'Processed'])
            for cam in self.cameras:
                w.writerow([cam.get('name',''), cam.get('ip',''), cam.get('gateway',''), 
                           cam.get('subnet',''), cam.get('model',''), cam.get('new_ip',''),
                           'Yes' if cam.get('processed') else 'No'])


class PasswordDataManager:
    """Manages password list - stored as JSON"""
    def __init__(self):
        self.passwords = []
        self.load()
    
    def load(self):
        if os.path.exists(PASSWORDS_FILE):
            try:
                with open(PASSWORDS_FILE, 'r') as f:
                    self.passwords = json.load(f)
            except:
                self.passwords = []
        return self.passwords
    
    def save(self):
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        with open(PASSWORDS_FILE, 'w') as f:
            json.dump(self.passwords, f, indent=2)
    
    def add(self, password):
        if password and password not in self.passwords:
            self.passwords.append(password)
            self.save()
    
    def delete(self, index):
        if 0 <= index < len(self.passwords):
            del self.passwords[index]
            self.save()
    
    def clear(self):
        self.passwords = []
        self.save()
    
    def get_all(self):
        return self.passwords


class AdditionalUsersDataManager:
    """Manages additional camera users - stored as JSON list of {username, password, role}"""
    # ROLES — Axis user role options.
    # 'Administrator', 'Operator', 'Viewer' = added to BOTH VAPIX and ONVIF
    #   databases. Standard for users who need web-UI access AND VMS pickup.
    # 'ONVIF-only Operator' (v4.3 #10c) = added ONLY to the ONVIF database,
    #   not VAPIX. For VMS service accounts (Milestone, Genetec, exacqVision)
    #   that don't need web-UI access. add_user() detects this role and
    #   skips the VAPIX pwdgrp.cgi POST.
    ROLES = ['Administrator', 'Operator', 'Viewer', 'ONVIF-only Operator']

    def __init__(self):
        self.users = []
        self.load()

    def load(self):
        if os.path.exists(ADDITIONAL_USERS_FILE):
            try:
                with open(ADDITIONAL_USERS_FILE, 'r') as f:
                    self.users = json.load(f)
            except:
                self.users = []
        return self.users

    def save(self):
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        with open(ADDITIONAL_USERS_FILE, 'w') as f:
            json.dump(self.users, f, indent=2)

    def add(self, username, password, role='Operator'):
        if not username:
            return False
        # Don't add duplicate usernames
        for u in self.users:
            if u['username'].lower() == username.lower():
                return False
        self.users.append({'username': username, 'password': password, 'role': role})
        self.save()
        return True

    def delete(self, index):
        if 0 <= index < len(self.users):
            del self.users[index]
            self.save()

    def clear(self):
        self.users = []
        self.save()

    def get_all(self):
        return self.users


# ============================================================================
# AXIS CAMERA DISCOVERY (UDP Port 19540)
# ============================================================================
class AxisDiscovery:
    """
    Discovers Axis cameras on the network using the proprietary
    Axis discovery protocol (same as AXIS IP Utility).
    Sends a broadcast on UDP port 19540 and cameras respond with their info.
    """
    DISCOVERY_PORT = 19540
    DISCOVERY_MAGIC = b'\x00\x01\x00\x00'  # Axis discovery request
    
    @staticmethod
    def get_local_ips():
        """Get all local IP addresses to broadcast from"""
        local_ips = []
        try:
            # Get all network interfaces
            hostname = socket.gethostname()
            local_ips = socket.gethostbyname_ex(hostname)[2]
        except:
            pass
        # Filter out localhost
        local_ips = [ip for ip in local_ips if not ip.startswith('127.')]
        if not local_ips:
            local_ips = ['0.0.0.0']
        return local_ips
    
    @staticmethod
    def discover(timeout=5, callback=None):
        """
        Discover Axis cameras on the network.
        Returns list of dicts with camera info.
        callback(camera_dict) is called for each camera found (optional).
        """
        cameras = []
        seen_macs = set()
        
        try:
            # Create UDP socket
            sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            sock.settimeout(0.5)  # Short timeout for recv
            sock.bind(('', 0))  # Bind to any available port
            
            # Send discovery broadcast
            broadcast_addr = ('255.255.255.255', AxisDiscovery.DISCOVERY_PORT)
            
            # Axis discovery packet - simple broadcast
            # The actual Axis protocol uses a specific packet format
            discovery_packet = AxisDiscovery.DISCOVERY_MAGIC + b'\x00' * 60
            
            sock.sendto(discovery_packet, broadcast_addr)
            
            # Also try sending to common subnet broadcasts
            for local_ip in AxisDiscovery.get_local_ips():
                try:
                    parts = local_ip.split('.')
                    subnet_broadcast = f"{parts[0]}.{parts[1]}.{parts[2]}.255"
                    sock.sendto(discovery_packet, (subnet_broadcast, AxisDiscovery.DISCOVERY_PORT))
                except:
                    pass
            
            # Listen for responses
            end_time = datetime.now().timestamp() + timeout
            while datetime.now().timestamp() < end_time:
                try:
                    data, addr = sock.recvfrom(1024)
                    if len(data) > 20:
                        camera = AxisDiscovery.parse_response(data, addr[0])
                        if camera and camera.get('mac') not in seen_macs:
                            seen_macs.add(camera.get('mac'))
                            cameras.append(camera)
                            if callback:
                                callback(camera)
                except socket.timeout:
                    continue
                except Exception as e:
                    continue
            
            sock.close()
        except Exception as e:
            print(f"Discovery error: {e}")
        
        return cameras
    
    @staticmethod
    def parse_response(data, source_ip):
        """Parse Axis discovery response packet"""
        try:
            # Axis response format varies by firmware, but generally contains:
            # - MAC address (6 bytes, often at offset 4-9)
            # - IP address (4 bytes)
            # - Model name (string)
            # - Serial number (string)
            
            camera = {
                'ip': source_ip,
                'mac': '',
                'model': '',
                'serial': '',
                'name': ''
            }
            
            # Try to extract MAC address (look for 6 consecutive bytes that look like MAC)
            for i in range(len(data) - 5):
                # Check if this could be a MAC (non-zero, not all same)
                mac_bytes = data[i:i+6]
                if mac_bytes[0:3] == b'\x00\x40\x8c' or mac_bytes[0:3] == b'\xac\xcc\x8e' or mac_bytes[0:3] == b'\xb8\xa4\x4f':
                    # Common Axis MAC prefixes
                    camera['mac'] = ':'.join(f'{b:02X}' for b in mac_bytes)
                    break
            
            # Try to extract readable strings (model, serial)
            # Look for null-terminated strings in the packet
            text_parts = []
            current = b''
            for byte in data:
                if 32 <= byte <= 126:  # Printable ASCII
                    current += bytes([byte])
                else:
                    if len(current) >= 3:
                        text_parts.append(current.decode('ascii', errors='ignore'))
                    current = b''
            if len(current) >= 3:
                text_parts.append(current.decode('ascii', errors='ignore'))
            
            # Try to identify model and serial from text parts
            for part in text_parts:
                part = part.strip()
                if not part:
                    continue
                # Axis models often start with letters like P, M, Q, etc.
                if re.match(r'^[PMQVFA]\d{4}', part) or 'AXIS' in part.upper():
                    camera['model'] = part
                # Serial numbers are often MAC-like or all caps/numbers
                elif re.match(r'^[A-F0-9]{12}$', part) or re.match(r'^ACCC[A-F0-9]{8}$', part):
                    camera['serial'] = part
                    if not camera['mac']:
                        # Convert serial to MAC format
                        camera['mac'] = ':'.join(part[i:i+2] for i in range(0, 12, 2))
            
            # Name = axis-serial (default hostname) or IP fallback
            if camera['serial']:
                camera['name'] = f"axis-{camera['serial'].lower()}"
            else:
                camera['name'] = source_ip
            
            return camera
        except Exception as e:
            return {'ip': source_ip, 'mac': '', 'model': 'Unknown', 'serial': '', 'name': source_ip}


# ============================================================================
# AXIS CAMERA DISCOVERY via DHCP SNOOPING
# ============================================================================
class AxisDHCPDiscovery:
    """
    Discovers Axis cameras by listening for DHCP DISCOVER broadcasts.
    Firmware 12.0+ cameras with no DHCP server broadcast DHCPDISCOVER
    every few seconds with hostname, model, MAC embedded in DHCP options.
    This is the fastest discovery method — purely passive, Layer 2 broadcast,
    no routing or multicast required. Works on any subnet.
    """
    DHCP_SERVER_PORT = 67
    DHCP_CLIENT_PORT = 68
    DHCP_MAGIC = b'\x63\x82\x53\x63'

    # Known Axis MAC OUI prefixes
    AXIS_OUIS = [b'\x00\x40\x8c', b'\xac\xcc\x8e', b'\xb8\xa4\x4f']

    @staticmethod
    def discover(timeout=5, callback=None):
        """Listen for DHCP DISCOVER broadcasts from Axis cameras.
        Returns list of camera dicts with mac, model, serial, hostname."""
        cameras = []
        seen_macs = set()

        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
            sock.bind(('', AxisDHCPDiscovery.DHCP_SERVER_PORT))
            sock.settimeout(0.5)
        except OSError:
            # Port 67 in use (DHCP server running) — can't snoop
            return cameras

        end_time = time.time() + timeout
        while time.time() < end_time:
            try:
                data, addr = sock.recvfrom(2048)
                camera = AxisDHCPDiscovery._parse_dhcp(data)
                if camera:
                    key = camera['mac']
                    if key not in seen_macs:
                        seen_macs.add(key)
                        cameras.append(camera)
                        if callback:
                            callback(camera)
            except socket.timeout:
                continue
            except Exception:
                continue

        sock.close()
        return cameras

    @staticmethod
    def _parse_dhcp(data):
        """Parse a DHCP packet and extract Axis camera info.
        Returns camera dict or None if not an Axis camera."""
        if len(data) < 240:
            return None

        op = data[0]
        if op != 1:  # Not a BOOTP request
            return None

        hlen = data[2]
        if hlen != 6:
            return None

        # Client MAC address at bytes 28-33
        mac_bytes = data[28:34]
        mac = ':'.join(f'{b:02X}' for b in mac_bytes)

        # Check if it's an Axis or Bosch MAC
        is_axis_mac = any(mac_bytes[:3] == oui for oui in AxisDHCPDiscovery.AXIS_OUIS)
        is_bosch_mac = any(mac_bytes[:3] == oui for oui in BOSCH_OUIS)

        # Check for DHCP magic cookie at offset 236
        if data[236:240] != AxisDHCPDiscovery.DHCP_MAGIC:
            return None

        # Parse DHCP options
        hostname = ''
        vendor = ''
        msg_type = 0
        i = 240
        while i < len(data):
            opt = data[i]
            if opt == 255:  # End
                break
            if opt == 0:  # Pad
                i += 1
                continue
            if i + 1 >= len(data):
                break
            olen = data[i + 1]
            if i + 2 + olen > len(data):
                break
            oval = data[i + 2:i + 2 + olen]
            if opt == 53:   # Message Type
                msg_type = oval[0]
            elif opt == 12:  # Hostname
                hostname = oval.decode('ascii', errors='ignore')
            elif opt == 60:  # Vendor Class Identifier
                vendor = oval.decode('ascii', errors='ignore')
            i += 2 + olen

        # Only care about DISCOVER (type 1) from Axis or Bosch cameras
        if msg_type != 1:
            return None
        if not is_axis_mac and not is_bosch_mac and 'axis' not in hostname.lower() and 'AXIS' not in vendor:
            return None

        # Determine brand
        brand = 'bosch' if is_bosch_mac else 'axis'

        # Parse vendor string: "AXIS,Dome Camera,P3268-LV,12.3.56"
        model = ''
        serial = ''
        if vendor and brand == 'axis':
            parts = vendor.split(',')
            if len(parts) >= 3:
                model = f"AXIS {parts[2].strip()}"

        # Derive serial from MAC (Axis serials = MAC without colons)
        serial = mac.replace(':', '')

        # Derive link-local IP from MAC (RFC 3927 — camera picks one, but we
        # know it from mDNS or can compute a guess). For now, leave IP empty
        # since DHCP DISCOVER comes from 0.0.0.0.
        # The actual IP will be found via mDNS A record after route is added.

        name_prefix = 'bosch' if brand == 'bosch' else 'axis'
        return {
            'ip': '',  # Camera has no routable IP yet — we'll find it
            'mac': mac,
            'model': model,
            'serial': serial,
            'hostname': hostname,
            'name': hostname or f"{name_prefix}-{serial.lower()}",
            'vendor': vendor,
            'brand': brand,
            '_source': 'dhcp',
        }


# ============================================================================
# BOSCH RCP-OVER-HTTP HELPER
# ============================================================================
class BoschRCP:
    """Static helper for Bosch RCP (Remote Control Protocol) over HTTP.
    Bosch cameras expose /rcp.xml for reading and writing configuration."""

    @staticmethod
    def rcp_read(ip, cmd_hex, rcp_type='P_STRING', auth=None, timeout=3):
        """Read a value from a Bosch camera via RCP-over-HTTP.
        rcp_type: P_STRING, T_DWORD, T_OCTET
        Returns parsed value string, or None on error."""
        try:
            params = {
                'command': f'0x{cmd_hex:04x}',
                'type': rcp_type,
                'direction': 'READ',
                'num': '1',
            }
            kwargs = {'timeout': timeout}
            if auth:
                kwargs['auth'] = HTTPDigestAuth(auth[0], auth[1])
            r = requests.get(f'http://{ip}/rcp.xml', params=params, **kwargs)
            if r.status_code != 200:
                return None
            text = r.text
            # Check for error
            err_m = re.search(r'<err>(0x[0-9a-fA-F]+)</err>', text)
            if err_m:
                return None
            # Parse based on type
            if rcp_type == 'P_STRING':
                m = re.search(r'<str>([^<]*)</str>', text)
                return m.group(1).strip() if m else None
            elif rcp_type == 'T_DWORD':
                m = re.search(r'<dec>(\d+)</dec>', text)
                return int(m.group(1)) if m else None
            elif rcp_type == 'T_OCTET':
                m = re.search(r'<str>([^<]*)</str>', text)
                return m.group(1).strip() if m else None
            return None
        except Exception:
            return None

    @staticmethod
    def rcp_write(ip, cmd_hex, rcp_type, payload, auth, timeout=5, num=1):
        """Write a value to a Bosch camera via RCP-over-HTTP.
        auth: (username, password) tuple. num: RCP instance number. Returns True on success."""
        try:
            params = {
                'command': f'0x{cmd_hex:04x}',
                'type': rcp_type,
                'direction': 'WRITE',
                'num': str(num),
                'payload': str(payload),
            }
            r = requests.get(f'http://{ip}/rcp.xml', params=params,
                             auth=HTTPDigestAuth(auth[0], auth[1]), timeout=timeout)
            if r.status_code != 200:
                return False
            if re.search(r'<err>', r.text):
                return False
            return True
        except Exception:
            return False

    @staticmethod
    def get_device_info(ip, timeout=3):
        """Get Bosch camera model/firmware/hardware from /config.js.
        Returns dict with 'model', 'firmware', 'hardware' or None.
        Older Bosch cameras may not have CTN (model) but will have HI/SW/Unit."""
        try:
            r = requests.get(f'http://{ip}/config.js', timeout=timeout)
            if r.status_code != 200:
                return None
            info = {}

            for line in r.text.split('\n'):
                line = line.strip().rstrip(';')
                if line.startswith('var CTN'):
                    m = re.search(r'"([^"]+)"', line)
                    if m:
                        info['model'] = m.group(1)
                elif line.startswith('var SW'):
                    m = re.search(r"['\"]([^'\"]+)['\"]", line)
                    if m:
                        info['firmware'] = m.group(1)
                elif line.startswith('var HI'):
                    m = re.search(r"['\"]([^'\"]+)['\"]", line)
                    if m:
                        info['hardware'] = m.group(1)
                elif line.startswith('var Unit'):
                    m = re.search(r"['\"]([^'\"]+)['\"]", line)
                    if m:
                        info['unit'] = m.group(1)
            # Detect as Bosch if we got any Bosch-specific fields
            if info:
                if not info.get('model'):
                    # Older cameras: use Unit or hardware as model fallback
                    info['model'] = info.get('unit', info.get('hardware', 'Bosch Camera'))
                return info
            return None
        except Exception:
            return None

    @staticmethod
    def get_network_config(ip, timeout=3):
        """Read network config from Bosch camera via RCP (no auth required).
        Returns dict with ip, subnet, gateway, dhcp, mac or empty dict."""
        config = {}
        val = BoschRCP.rcp_read(ip, RCP_CMD['ip'], 'P_STRING', timeout=timeout)
        if val:
            config['ip'] = val
        val = BoschRCP.rcp_read(ip, RCP_CMD['subnet'], 'P_STRING', timeout=timeout)
        if val:
            config['subnet'] = val
        val = BoschRCP.rcp_read(ip, RCP_CMD['gateway'], 'P_STRING', timeout=timeout)
        if val:
            config['gateway'] = val
        val = BoschRCP.rcp_read(ip, RCP_CMD['dhcp'], 'T_DWORD', timeout=timeout)
        if val is not None:
            config['dhcp'] = 'Yes' if val == 1 else 'No'
        val = BoschRCP.rcp_read(ip, RCP_CMD['mac'], 'T_OCTET', timeout=timeout)
        if val:
            # MAC comes as "00 07 5f 9c 9e 75 " — normalize to colon-separated
            parts = val.split()
            if len(parts) >= 6:
                config['mac'] = ':'.join(p.upper() for p in parts[:6])
        return config


# ============================================================================
# LOCKOUT ERROR (Hanwha 490 response)
# ============================================================================
class LockoutError(Exception):
    """Raised when a Hanwha camera returns HTTP 490 (lockout)."""
    pass


# ============================================================================
# CAMERA PROTOCOL ABSTRACTION (ABC)
# ============================================================================
from abc import ABC, abstractmethod


class CameraProtocol(ABC):
    """Abstract base class for brand-specific camera operations.
    Each brand implements this interface so all operations can be
    brand-agnostic — just call self.protocol.method()."""

    # Subclasses set these
    BRAND_NAME = ''
    BRAND_KEY = ''
    DEFAULT_USER = ''
    DEFAULT_PASSWORD = ''
    FACTORY_IP = ''
    MAC_OUIS = []

    @abstractmethod
    def create_initial_user(self, ip, password):
        """Set password on a factory-default camera. Returns True/False."""

    @abstractmethod
    def set_network(self, ip, password, new_ip, subnet, gateway):
        """Set static IP, subnet, gateway and disable DHCP. Returns True/False."""

    @abstractmethod
    def set_hostname(self, ip, password, hostname):
        """Set network hostname. Returns True/False."""

    @abstractmethod
    def reboot(self, ip, password):
        """Reboot camera. Returns True/False."""

    @abstractmethod
    def get_serial(self, ip, password):
        """Get serial number (requires auth). Returns string or 'UNKNOWN'."""

    @abstractmethod
    def get_model_noauth(self, ip):
        """Query model without authentication. Returns string or None."""

    def probe_unrestricted(self, ip):
        """Single no-auth probe returning model + serial + MAC + firmware +
        hardware in one dict. Brands that don't have an unauthenticated
        rich-info endpoint return what they can. AxisProtocol overrides with
        the basicdeviceinfo.cgi getAllUnrestrictedProperties call which works
        on factory AND password-locked cameras.
        Returns dict with keys: model, model_full, serial, mac, firmware,
        hardware, brand. Any/all may be None on failure."""
        out = {'model': None, 'model_full': None, 'serial': None, 'mac': None,
               'firmware': None, 'hardware': None, 'brand': None}
        try:
            out['model'] = self.get_model_noauth(ip)
        except Exception:
            pass
        return out

    def get_firmware(self, ip, password):
        """Get firmware version. Returns string or 'UNKNOWN'.
        Default implementation returns 'UNKNOWN' — subclasses should override."""
        return 'UNKNOWN'

    @abstractmethod
    def get_image(self, ip, username, password):
        """Get JPEG snapshot bytes. Returns bytes or None."""

    @abstractmethod
    def test_password(self, ip, username, password):
        """Test if credentials work. Returns True/False. May raise LockoutError for Hanwha."""

    @abstractmethod
    def change_password(self, ip, username, old_pwd, new_pwd):
        """Change password. Returns True/False."""

    @abstractmethod
    def set_dhcp(self, ip, password, enable=True):
        """Enable or disable DHCP. Returns True/False."""

    @abstractmethod
    def factory_reset(self, ip, password):
        """Factory reset camera. Returns True/False."""

    def add_user(self, ip, admin_password, username, user_password, role='Operator'):
        """Add an additional user account to the camera. Returns True/False.
        Default implementation returns False (not supported)."""
        return False

    @abstractmethod
    def get_programming_steps(self, cam, password, options=None):
        """Return ordered list of (description, callable) steps for programming a factory-default camera.
        Each callable returns True on success, False on failure."""

    def get_discovery_info(self, ip, timeout=2):
        """Try to identify this brand's camera at the given IP without auth.
        Returns a dict with camera info, or None if not this brand."""
        return None


# ============================================================================
# AXIS PROTOCOL
# ============================================================================
class AxisProtocol(CameraProtocol):
    # Axis is the cleanest of the three brands by a mile. VAPIX has been
    # backwards-compatible since like firmware 5.x (2013-ish) — the same
    # pwdgrp.cgi / param.cgi calls below still work on 11.x. Easy money.
    BRAND_NAME = 'Axis'
    BRAND_KEY = 'axis'
    DEFAULT_USER = 'root'
    DEFAULT_PASSWORD = ''  # factory cameras are unauthenticated until you set the root pwd
    FACTORY_IP = '192.168.0.90'
    # OUIs cover the bulk of what shows up on jobs. The b8:a4:4f range is the
    # newer ARTPEC chips (M30/M32 series and up); 00:40:8c is older P-series.
    MAC_OUIS = [b'\x00\x40\x8c', b'\xac\xcc\x8e', b'\xb8\xa4\x4f']

    def create_initial_user(self, ip, password):
        # On a factory Axis, /pwdgrp.cgi is wide open until the first root pwd
        # gets set — that's by design, it's how their out-of-box flow works.
        # First call locks down the camera and turns on auth for everything else.
        try:
            r = requests.get(f"http://{ip}/axis-cgi/pwdgrp.cgi",
                params={"action": "add", "user": "root", "pwd": password,
                        "grp": "root", "sgrp": "admin:operator:viewer:ptz"},
                timeout=TIMEOUT)
            if r.status_code != 200:
                return False
        except:
            # Bare except because anything that explodes here (DNS, refused,
            # cert, you name it) means the camera isn't reachable or isn't
            # actually factory. Either way: not our problem.
            return False

        # The ONVIF user is a SEPARATE database from the VAPIX user. Same name,
        # same password, different table inside the camera. If you only create
        # the VAPIX one, ONVIF clients (Milestone, Genetec, exacqVision) can't
        # log in even though the web UI works fine. Found that out the hard way
        # on a Chase job in 2014. Now I always create both up-front.
        # 2026-05-02 v4.3 #10: this ONVIF "root" user is REQUIRED for set_network
        # SOAP. After programming, the wizard calls delete_onvif_user(static_ip)
        # to remove it (unless the operator opted to keep it). Net effect:
        # camera ends with VAPIX root only, no leftover ONVIF account.
        try:
            soap = (f'<?xml version="1.0"?>'
                    f'<Envelope xmlns="http://www.w3.org/2003/05/soap-envelope"><Header/><Body>'
                    f'<CreateUsers xmlns="http://www.onvif.org/ver10/device/wsdl" '
                    f'xmlns:tt="http://www.onvif.org/ver10/schema">'
                    f'<User><tt:Username>root</tt:Username>'
                    f'<tt:Password>{password}</tt:Password>'
                    f'<tt:UserLevel>Administrator</tt:UserLevel>'
                    f'</User></CreateUsers></Body></Envelope>')
            requests.post(f"http://{ip}/vapix/services", data=soap,
                headers={"Content-Type": "application/soap+xml"},
                auth=HTTPDigestAuth("root", password), timeout=TIMEOUT)
        except:
            # Don't fail the whole programming step if this one barfs — older
            # firmware (pre 6.50) sometimes 500s on this call but the user
            # still ends up created. Web UI confirms it. Just move on.
            pass
        return True

    def add_onvif_user(self, ip, password, new_username, new_password, level='Operator'):
        """v4.3 #10b — create a NAMED ONVIF user (e.g. for VMS service account)
        AFTER programming completes. Used in the rename pattern: the wizard
        deletes the transient ONVIF root, then calls this to create the
        operator-specified named user. Auths via VAPIX root (same admin creds
        the rest of the wizard uses). level: 'Administrator' / 'Operator' /
        'User'. Returns True on success."""
        try:
            soap = (f'<?xml version="1.0"?>'
                    f'<Envelope xmlns="http://www.w3.org/2003/05/soap-envelope"><Header/><Body>'
                    f'<CreateUsers xmlns="http://www.onvif.org/ver10/device/wsdl" '
                    f'xmlns:tt="http://www.onvif.org/ver10/schema">'
                    f'<User><tt:Username>{new_username}</tt:Username>'
                    f'<tt:Password>{new_password}</tt:Password>'
                    f'<tt:UserLevel>{level}</tt:UserLevel>'
                    f'</User></CreateUsers></Body></Envelope>')
            r = requests.post(f"http://{ip}/vapix/services", data=soap,
                headers={"Content-Type": "application/soap+xml"},
                auth=HTTPDigestAuth("root", password), timeout=TIMEOUT)
            return r.status_code in (200, 204)
        except Exception:
            return False

    def verify_camera_state(self, ip, password):
        """v4.3 #12 — Confirm Programming. Audit one camera at one IP, return
        a dict of read-only checks. Caller compares against the expected
        camera entry to decide pass/fail.
        Checks:
          - reachable      : probe_unrestricted got something back
          - mac/model/firmware : from no-auth probe (Axis serial == MAC)
          - auth_ok        : root+password gets 200 from param.cgi (Network)
          - dhcp_off       : Network.eth0.BootProto == 'static' (most installs
                             on a camera VLAN need DHCP off; on, only safe
                             with MAC reservations)
          - actual_ip / actual_gateway / actual_subnet : what the camera says
                             its config is, for diff against expected entry
        Returns a dict with all fields; missing/unread fields are None or False."""
        out = {
            'reachable': False, 'mac': None, 'model': None, 'firmware': None,
            'auth_ok': False, 'dhcp_off': None,
            'actual_ip': None, 'actual_gateway': None, 'actual_subnet': None,
        }
        # Quick no-auth probe — also covers reachability + identity
        probe = self.probe_unrestricted(ip)
        if probe.get('mac') or probe.get('model'):
            out['reachable'] = True
            out['mac'] = probe.get('mac')
            out['model'] = probe.get('model')
            out['firmware'] = probe.get('firmware')
        if not out['reachable']:
            return out
        # Auth + network config
        try:
            r = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                params={"action": "list", "group": "Network"},
                auth=HTTPDigestAuth("root", password), timeout=TIMEOUT)
            if r.status_code == 200:
                out['auth_ok'] = True
                for line in r.text.split('\n'):
                    line = line.strip()
                    if not line or '=' not in line:
                        continue
                    if 'Network.eth0.IPAddress=' in line or line.startswith('Network.IPAddress='):
                        out['actual_ip'] = line.split('=', 1)[1].strip()
                    elif 'Network.eth0.DefaultRouter=' in line:
                        out['actual_gateway'] = line.split('=', 1)[1].strip()
                    elif 'Network.eth0.SubnetMask=' in line:
                        out['actual_subnet'] = line.split('=', 1)[1].strip()
                    elif 'Network.eth0.BootProto=' in line or line.startswith('Network.BootProto='):
                        proto = line.split('=', 1)[1].strip().lower()
                        out['dhcp_off'] = (proto != 'dhcp')
        except Exception:
            pass
        return out

    def delete_onvif_user(self, ip, password, username='root'):
        """v4.3 #10 — security cleanup after programming completes. The ONVIF
        user that create_initial_user added is a transient TOOL needed for
        set_network's ONVIF SOAP call; once set_network is done it's just a
        leftover account. Delete via ONVIF DeleteUsers SOAP. Idempotent —
        if the user is already gone, returns True (no-op).
        Returns True on success or already-absent, False on a real failure
        (auth fail, network unreachable, SOAP fault). Non-fatal at the wizard
        level — wizard logs a warning if False but keeps going."""
        try:
            soap = (f'<?xml version="1.0"?>'
                    f'<Envelope xmlns="http://www.w3.org/2003/05/soap-envelope"><Header/><Body>'
                    f'<DeleteUsers xmlns="http://www.onvif.org/ver10/device/wsdl">'
                    f'<Username>{username}</Username>'
                    f'</DeleteUsers></Body></Envelope>')
            r = requests.post(f"http://{ip}/vapix/services", data=soap,
                headers={"Content-Type": "application/soap+xml"},
                auth=HTTPDigestAuth("root", password), timeout=TIMEOUT)
            if r.status_code in (200, 204):
                return True
            # 4xx with "user not found" SOAP fault = already deleted; treat as ok
            if 'NoSuchUser' in r.text or 'not found' in r.text.lower():
                return True
            return False
        except Exception:
            return False

    def set_network(self, ip, password, new_ip, subnet, gateway):
        # Two paths: ONVIF SOAP first because it's atomic — gateway/IP/DHCP
        # all flip together so the camera doesn't end up in a half-configured
        # state if the connection drops. VAPIX param.cgi is the fallback for
        # older firmware (anything before like 7.10 or so) where the ONVIF
        # SetNetworkInterfaces call returns a Fault.
        auth = HTTPDigestAuth("root", password)
        cidr = sum(bin(int(x)).count('1') for x in subnet.split('.')) if subnet else 24

        # --- ONVIF path ---
        onvif_ok = False
        # Gateway has to go in its own SOAP call. Tried merging it into the
        # SetNetworkInterfaces envelope once — Axis silently ignored it. This
        # is the supported pattern in the ONVIF Network Configuration spec.
        gw_soap = (f'<?xml version="1.0"?>'
                   f'<Envelope xmlns="http://www.w3.org/2003/05/soap-envelope"><Header/><Body>'
                   f'<SetNetworkDefaultGateway xmlns="http://www.onvif.org/ver10/device/wsdl">'
                   f'<IPv4Address>{gateway}</IPv4Address>'
                   f'</SetNetworkDefaultGateway></Body></Envelope>')
        try:
            requests.post(f"http://{ip}/vapix/services", data=gw_soap,
                headers={"Content-Type": "application/soap+xml"},
                auth=auth, timeout=TIMEOUT)
        except:
            pass  # gw failure isn't fatal — IP can still be set, we'll just retry gw

        # Now the static IP + DHCP off in one shot. PrefixLength is CIDR notation
        # (the /24 in 192.168.1.0/24), Axis expects an integer here.
        ip_soap = (f'<?xml version="1.0"?>'
                   f'<Envelope xmlns="http://www.w3.org/2003/05/soap-envelope" '
                   f'xmlns:tds="http://www.onvif.org/ver10/device/wsdl" '
                   f'xmlns:tt="http://www.onvif.org/ver10/schema"><Header/><Body>'
                   f'<tds:SetNetworkInterfaces>'
                   f'<tds:InterfaceToken>eth0</tds:InterfaceToken>'
                   f'<tds:NetworkInterface><tt:Enabled>true</tt:Enabled>'
                   f'<tt:IPv4><tt:Enabled>true</tt:Enabled>'
                   f'<tt:Manual><tt:Address>{new_ip}</tt:Address>'
                   f'<tt:PrefixLength>{cidr}</tt:PrefixLength></tt:Manual>'
                   f'<tt:DHCP>false</tt:DHCP></tt:IPv4></tds:NetworkInterface>'
                   f'</tds:SetNetworkInterfaces></Body></Envelope>')
        try:
            r = requests.post(f"http://{ip}/vapix/services", data=ip_soap,
                headers={"Content-Type": "application/soap+xml"},
                auth=auth, timeout=TIMEOUT)
            # Axis returns 200 even on a failed config change, so the body has
            # to be checked for <Fault>. Found this on a P3245 that wouldn't
            # stick a static IP and was lying about it — a 200 with a Fault
            # block. Why they don't return 4xx in that case I do not know.
            if r.status_code == 200 and 'Fault' not in r.text:
                onvif_ok = True
        except:
            pass

        if onvif_ok:
            return True

        # --- VAPIX fallback ---
        # Order matters here. If you set IP first then DHCP, the camera goes
        # offline mid-config and you can't finish. DHCP off first, gateway,
        # subnet, IP last — that order has been bulletproof since at least
        # the M-series firmware in 2012.
        try:
            requests.get(f"http://{ip}/axis-cgi/param.cgi",
                params={"action": "update", "root.Network.IPv4.DHCP": "no"},
                auth=auth, timeout=TIMEOUT)
            requests.get(f"http://{ip}/axis-cgi/param.cgi",
                params={"action": "update", "root.Network.DefaultRouter": gateway},
                auth=auth, timeout=TIMEOUT)
            requests.get(f"http://{ip}/axis-cgi/param.cgi",
                params={"action": "update", "root.Network.SubnetMask": subnet},
                auth=auth, timeout=TIMEOUT)
            # Setting IPAddress is the call that yanks the rug — the response
            # may never come back because the kernel has switched interfaces
            # by the time it tries to ack. Hence the ConnectionError below
            # being treated as success.
            r = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                params={"action": "update", "root.Network.IPAddress": new_ip},
                auth=auth, timeout=TIMEOUT)
            return r.status_code == 200
        except requests.exceptions.ConnectionError:
            # Connection died mid-call — that's actually GOOD on this one,
            # means the IP change took effect. Return True; verification step
            # higher up will confirm the camera is reachable on the new IP.
            return True
        except:
            return False

    def set_hostname(self, ip, password, hostname):
        # 63-char DNS limit, anything not alphanum or dash gets sanitized to a dash.
        # Cameras themselves are more permissive than RFC 952 but switches and
        # NVRs aren't — Genetec in particular barfs on underscores in hostnames.
        clean = re.sub(r'[^a-zA-Z0-9-]', '-', hostname).strip('-')[:63]
        try:
            r = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                params={"action": "update", "root.Network.HostName": clean},
                auth=HTTPDigestAuth("root", password), timeout=TIMEOUT)
            # Also stamp the Bonjour name — most NVRs that auto-populate camera
            # inventories pull this for the friendly label, not HostName. AXIS
            # Camera Station and Milestone do, anyway.
            try:
                requests.get(f"http://{ip}/axis-cgi/param.cgi",
                    params={"action": "update", "root.Network.Bonjour.FriendlyName": clean},
                    auth=HTTPDigestAuth("root", password), timeout=TIMEOUT)
            except:
                pass  # Bonjour is best-effort; some MFA-locked-down builds disable it
            return r.status_code == 200 and 'OK' in r.text
        except:
            return False

    def reboot(self, ip, password):
        # Note: cameras already reboot themselves after a network change, so this
        # is mostly for the "user clicked Reboot" path. Hitting it twice in a
        # row is harmless — Axis just queues the second request and ignores it.
        try:
            r = requests.get(f"http://{ip}/axis-cgi/restart.cgi",
                auth=HTTPDigestAuth("root", password), timeout=TIMEOUT)
            return r.status_code == 200
        except:
            return False

    def set_dhcp(self, ip, password, enable=True):
        # Used for the "I need to reset this back to DHCP for staging" workflow.
        # Camera will renew immediately and probably end up at a different IP,
        # so the caller is responsible for re-discovery.
        val = 'yes' if enable else 'no'
        try:
            r = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                params={"action": "update", "root.Network.IPv4.DHCP": val},
                auth=HTTPDigestAuth("root", password), timeout=TIMEOUT)
            return r.status_code == 200 and 'OK' in r.text
        except:
            return False

    def get_serial(self, ip, password):
        # Three retries because freshly-booted cameras (especially after a
        # factory reset) sometimes 503 the first time you hit this. Web UI
        # is up but the param subsystem isn't ready. Sleep a sec and try again.
        for attempt in range(3):
            try:
                # param.cgi path — works on every Axis I've ever touched. If this
                # call fails it's almost always a wrong-password thing, not a
                # missing-feature thing.
                r = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                    params={"action": "list", "group": "Properties.System.SerialNumber"},
                    auth=HTTPDigestAuth("root", password), timeout=TIMEOUT)
                if r.status_code == 200:
                    for line in r.text.split('\n'):
                        if 'SerialNumber=' in line:
                            return line.split('=')[1].strip()
            except:
                pass
            # Backup path: basicdeviceinfo.cgi — newer (5.50+) but cleaner JSON.
            # Some really old M-series only have param.cgi, which is why this
            # is the second choice not the first.
            try:
                r = requests.post(f"http://{ip}/axis-cgi/basicdeviceinfo.cgi",
                    json={"apiVersion": "1.0", "method": "getAllProperties"},
                    auth=HTTPDigestAuth("root", password), timeout=TIMEOUT)
                if r.status_code == 200:
                    data = r.json()
                    if 'data' in data and 'propertyList' in data['data']:
                        serial = data['data']['propertyList'].get('SerialNumber')
                        if serial:
                            return serial
            except:
                pass
            time.sleep(1)
        # Returning the literal string "UNKNOWN" not None so it survives a
        # CSV export without breaking the column. Caller flags these for
        # manual entry on the report.
        return "UNKNOWN"

    def get_model_noauth(self, ip):
        # Backward-compatible thin wrapper around probe_unrestricted — returns
        # just the ProdNbr string. New code should call probe_unrestricted
        # directly to get model + serial + MAC + firmware in one call.
        m = self.probe_unrestricted(ip).get('model')
        if m:
            return m
        # Fallback for pre-7.x firmware — basicdeviceinfo.cgi doesn't exist there
        try:
            r = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                params={"action": "list", "group": "Brand.ProdNbr"},
                timeout=TIMEOUT)
            if r.status_code == 200:
                for line in r.text.split('\n'):
                    if 'ProdNbr=' in line:
                        return line.split('=')[1].strip()
        except:
            pass
        return None

    def probe_unrestricted(self, ip):
        """Single no-auth probe via basicdeviceinfo.cgi getAllUnrestrictedProperties.
        Works on factory AND password-locked cameras (the *Unrestricted* variant
        doesn't require auth even when root is set; plain getAllProperties returns
        401 on configured cams). Returns model + serial + MAC + firmware + hardware
        in one round trip. For Axis the SerialNumber IS the MAC (12-char hex).
        Used by the wizard's verify_model step so MAC and model are both captured
        together, eliminating the "expected model, got ?" path when ARP pin missed
        AND the legacy single-method getModel fell through."""
        out = {'model': None, 'model_full': None, 'serial': None, 'mac': None,
               'firmware': None, 'hardware': None, 'brand': None}
        try:
            r = requests.post(f"http://{ip}/axis-cgi/basicdeviceinfo.cgi",
                json={"apiVersion": "1.0", "method": "getAllUnrestrictedProperties"},
                timeout=TIMEOUT)
            if r.status_code == 200:
                props = (r.json().get('data') or {}).get('propertyList') or {}
                out['model'] = (props.get('ProdNbr') or '').strip() or None
                out['model_full'] = (props.get('ProdFullName') or '').strip() or None
                out['serial'] = (props.get('SerialNumber') or '').strip() or None
                out['firmware'] = (props.get('Version') or '').strip() or None
                out['hardware'] = (props.get('HardwareID') or '').strip() or None
                out['brand'] = (props.get('Brand') or '').strip() or None
                if out['serial'] and len(out['serial']) == 12 and all(c in '0123456789ABCDEFabcdef' for c in out['serial']):
                    s = out['serial'].upper()
                    out['mac'] = ':'.join(s[i:i+2] for i in range(0, 12, 2))
        except Exception:
            pass
        return out

    def get_firmware(self, ip, password):
        # Three-tier fallback because firmware version is THE thing customers
        # ask about and the most-supported endpoint changed twice between
        # firmware 5.x, 6.x, and 9.x. Don't ask me why this is so hard. Below
        # in order of preference: no-auth JSON, auth JSON, auth param.cgi.

        # No-auth basicdeviceinfo via getAllUnrestrictedProperties — works on
        # factory AND configured cameras (the *Unrestricted* variant doesn't
        # require auth even when root is set; plain getAllProperties returns 401).
        try:
            r = requests.post(f"http://{ip}/axis-cgi/basicdeviceinfo.cgi",
                json={"apiVersion": "1.0", "method": "getAllUnrestrictedProperties"},
                timeout=TIMEOUT)
            if r.status_code == 200:
                data = r.json()
                if 'data' in data and 'propertyList' in data['data']:
                    fw = data['data']['propertyList'].get('Version')
                    if fw:
                        return fw
        except:
            pass

        # Auth basicdeviceinfo — once root is set, no-auth gets 401. Same call,
        # different auth posture.
        if password:
            try:
                r = requests.post(f"http://{ip}/axis-cgi/basicdeviceinfo.cgi",
                    json={"apiVersion": "1.0", "method": "getAllProperties"},
                    auth=HTTPDigestAuth("root", password), timeout=TIMEOUT)
                if r.status_code == 200:
                    data = r.json()
                    if 'data' in data and 'propertyList' in data['data']:
                        fw = data['data']['propertyList'].get('Version')
                        if fw:
                            return fw
            except:
                pass

        # Last resort: param.cgi. Older M10/M11 cameras only have this.
        try:
            r = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                params={"action": "list", "group": "Properties.Firmware.Version"},
                auth=HTTPDigestAuth("root", password) if password else None,
                timeout=TIMEOUT)
            if r.status_code == 200:
                for line in r.text.split('\n'):
                    if 'Version=' in line:
                        # split with maxsplit=1 because firmware strings can
                        # contain '=' (build metadata) — losing it would lie.
                        return line.split('=', 1)[1].strip()
        except:
            pass
        return 'UNKNOWN'

    def get_image(self, ip, username, password):
        # Pulls a single JPEG snapshot for the verification panel. Resolution
        # is whatever the camera's main stream is set to — we don't try to
        # downscale here, the GUI handles fitting. Bandwidth on this is
        # noticeable when you've got 30 of these running in parallel during
        # discovery; that's why TIMEOUT is short.
        try:
            r = requests.get(f"http://{ip}/axis-cgi/jpg/image.cgi",
                auth=HTTPDigestAuth(username, password), timeout=TIMEOUT)
            if r.status_code == 200:
                return r.content
        except:
            pass
        return None

    def test_password(self, ip, username, password):
        # Cheapest authenticated call I can think of — Brand group is tiny,
        # always present, and 200/Brand-in-text is a definitive yes/no.
        # Don't replace this with a get_serial test, that's like 4x slower.
        try:
            r = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                params={"action": "list", "group": "Brand"},
                auth=HTTPDigestAuth(username, password), timeout=TIMEOUT)
            return r.status_code == 200 and 'Brand' in r.text
        except:
            return False

    def change_password(self, ip, username, old_pwd, new_pwd):
        # Note: changing root's password while logged in as root works. The
        # session doesn't get invalidated mid-call. (Bosch is NOT this nice.)
        try:
            return requests.get(f"http://{ip}/axis-cgi/pwdgrp.cgi",
                params={"action": "update", "user": username, "pwd": new_pwd},
                auth=HTTPDigestAuth(username, old_pwd), timeout=TIMEOUT).status_code == 200
        except:
            return False

    def add_user(self, ip, admin_password, username, user_password, role='Operator'):
        # sgrp is Axis-speak for "secondary groups". You stack the privileges
        # by colon-separating them. Operator implies viewer + ptz, Viewer is
        # view-only. Customer reqs almost always end up with Operator for the
        # NVR account and Viewer for VMS-of-last-resort.
        role_groups = {
            'Administrator': 'admin:operator:viewer:ptz',
            'Operator': 'operator:viewer:ptz',
            'Viewer': 'viewer',
        }
        sgrp = role_groups.get(role, 'operator:viewer:ptz')

        # v4.3 #10c — 'ONVIF-only Operator' role skips the VAPIX side entirely.
        # The user only exists in the ONVIF user database, so they can hit the
        # camera via ONVIF (Milestone/Genetec/exacqVision) but cannot log into
        # the web UI. Useful for VMS service accounts where you want zero
        # human-facing creds.
        onvif_only = (role == 'ONVIF-only Operator')

        if not onvif_only:
            # VAPIX user first (this is the auth gate), ONVIF after for VMS hookup.
            try:
                r = requests.get(f"http://{ip}/axis-cgi/pwdgrp.cgi",
                    params={"action": "add", "user": username, "pwd": user_password,
                            "grp": "users", "sgrp": sgrp},
                    auth=HTTPDigestAuth("root", admin_password), timeout=TIMEOUT)
                if r.status_code != 200:
                    return False
            except:
                return False

        # ONVIF level only has Administrator / Operator / User — Axis maps
        # Viewer to "User". 'ONVIF-only Operator' maps to 'Operator'. Same
        # caveat as create_initial_user about needing both records or VMS-side
        # login fails.
        if role in ('Administrator', 'Operator', 'ONVIF-only Operator'):
            onvif_level = 'Administrator' if role == 'Administrator' else 'Operator'
        else:
            onvif_level = 'User'
        try:
            soap = (f'<?xml version="1.0"?>'
                    f'<Envelope xmlns="http://www.w3.org/2003/05/soap-envelope"><Header/><Body>'
                    f'<CreateUsers xmlns="http://www.onvif.org/ver10/device/wsdl" '
                    f'xmlns:tt="http://www.onvif.org/ver10/schema">'
                    f'<User><tt:Username>{username}</tt:Username>'
                    f'<tt:Password>{user_password}</tt:Password>'
                    f'<tt:UserLevel>{onvif_level}</tt:UserLevel>'
                    f'</User></CreateUsers></Body></Envelope>')
            r = requests.post(f"http://{ip}/vapix/services", data=soap,
                headers={"Content-Type": "application/soap+xml"},
                auth=HTTPDigestAuth("root", admin_password), timeout=TIMEOUT)
            # For ONVIF-only role, ONVIF write must succeed (it's the ONLY user
            # creation). For dual-write roles, ONVIF is best-effort (VAPIX is
            # already gold-pathed).
            if onvif_only and r.status_code not in (200, 204):
                return False
        except:
            if onvif_only:
                return False  # ONVIF write failed AND it's our only user creation path
            pass  # dual-write: ONVIF best-effort
        return True

    def factory_reset(self, ip, password):
        # THREE different reset endpoints because Axis has changed their mind
        # twice. Walk them in newest-first order; older firmware just won't
        # know about the newer ones.
        #
        # Hard reset wipes EVERYTHING — IP, password, certs, the lot. There's
        # also a "soft" mode that preserves network config but in 17 years of
        # this I've never wanted that. If a camera needs resetting, I want it
        # blank.

        # firmwaremanagement.cgi — current API (firmware 9.x+).
        try:
            r = requests.post(f"http://{ip}/axis-cgi/firmwaremanagement.cgi",
                json={"apiVersion": "1.0", "method": "factoryDefault",
                      "params": {"mode": "hard"}},
                auth=HTTPDigestAuth("root", password), timeout=10)
            if r.status_code == 200:
                try:
                    data = r.json()
                    if 'error' not in data:
                        return True
                except:
                    # Some firmware doesn't actually return JSON despite the
                    # API contract saying it does. If body has no 'error'
                    # string, call it good.
                    if 'error' not in r.text.lower():
                        return True
        except requests.exceptions.Timeout:
            # Camera reset and dropped the TCP connection mid-response. That's
            # the success path for a destructive call like this.
            return True
        except:
            pass

        # Legacy hardfactorydefault.cgi — Axis 6.x and earlier.
        try:
            r = requests.get(f"http://{ip}/axis-cgi/hardfactorydefault.cgi",
                auth=HTTPDigestAuth("root", password), timeout=10)
            if r.status_code == 200:
                return True
        except requests.exceptions.Timeout:
            return True  # see above — connection drop = win
        except:
            pass

        # param.cgi route — really old (5.x). Last ditch.
        try:
            r = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                params={"action": "update", "System.HardFactoryDefault": "yes"},
                auth=HTTPDigestAuth("root", password), timeout=10)
            if r.status_code == 200:
                return True
        except requests.exceptions.Timeout:
            return True
        except:
            pass
        return False

    def get_programming_steps(self, cam, password, options=None):
        # Returns the sequenced list of (label, callable) tuples that the
        # programmer worker chews through. Order is mandatory: user FIRST,
        # because everything after needs auth; network SECOND, because the
        # IP change is the moment the camera "moves" off its factory
        # 192.168.0.90 (or DHCP) onto the project subnet.
        ip = cam.get('_program_ip', cam['ip'])  # factory/link-local IP we found it on
        static_ip = cam['ip']
        gateway = cam['gateway']
        subnet = cam['subnet']
        set_hostname = options.get('set_hostname', False) if options else False

        steps = [
            ("Creating system user + ONVIF user",
             lambda: self.create_initial_user(ip, password)),
            ("Setting gateway + IP + disabling DHCP",
             lambda: self.set_network(ip, password, static_ip, subnet, gateway)),
        ]
        if set_hostname:
            # Hostname format I've used for a decade: <position>-<brand>-<serial>.
            # That sorts cleanly in any NVR camera list and makes the truck
            # paperwork match the hostname when somebody pulls a stream months
            # later trying to figure out which camera went down.
            cam_number = cam.get('number', '1')
            serial = cam.get('serial', 'unknown')
            hostname = f"{cam_number}-axis-{serial.lower()}"
            steps.append(("Setting hostname",
                lambda: self.set_hostname(static_ip, password, hostname)))
        return steps

    def get_discovery_info(self, ip, timeout=2):
        # Used by the network sweeper to figure out who's home at a given IP.
        # JSON path is the cleanest fingerprint. 401 means "Axis camera, but
        # configured" — that's still useful info (we know it's there, just
        # need a password to do anything else).
        try:
            r = requests.post(f"http://{ip}/axis-cgi/basicdeviceinfo.cgi",
                json={"apiVersion": "1.0", "method": "getAllProperties"},
                timeout=timeout)
            if r.status_code == 200:
                data = r.json()
                if 'data' in data and 'propertyList' in data['data']:
                    props = data['data']['propertyList']
                    return {
                        'ip': ip,
                        'model': props.get('ProdFullName', props.get('ProdShortName', '')),
                        'serial': props.get('SerialNumber', ''),
                        'brand': 'axis',
                    }
            elif r.status_code == 401:
                return {'ip': ip, 'model': '(Auth Required)', 'brand': 'axis'}
        except:
            pass
        # Last-ditch: tickle the snapshot endpoint. Even with no creds, factory
        # cameras serve image.cgi. A 401 here is gold — it means there IS an
        # Axis on the IP, just with auth on. (Used to use this as the primary
        # check before basicdeviceinfo existed.)
        try:
            r = requests.get(f"http://{ip}/axis-cgi/jpg/image.cgi", timeout=1)
            if r.status_code == 401:
                return {'ip': ip, 'model': '(Auth Required)', 'brand': 'axis'}
        except:
            pass
        return None


# ============================================================================
# BOSCH PROTOCOL
# ============================================================================
# RCP+ is Bosch's binary control protocol — it's older than the cameras most of
# us are still installing, predates ONVIF, and almost zero documentation exists
# online. Reverse-engineered from packet captures of the Configuration Manager
# tool (Bosch's own provisioning software). The good news: once you have the
# command numbers and type tags right, it's deterministic. The bad news: every
# value gets typed (P_STRING, T_DWORD, T_OCTET, F_FLAG) and the wrong type just
# returns "OK" with no actual write happening. Don't trust the ack — verify.
#
# Also: Bosch has THREE password levels (service=admin, user=ops, live=viewer)
# and the customer almost never knows that. They ask "what's the password" and
# you have to ask which one. So we just set all three to the same value.
class BoschProtocol(CameraProtocol):
    BRAND_NAME = 'Bosch'
    BRAND_KEY = 'bosch'
    DEFAULT_USER = 'service'
    DEFAULT_PASSWORD = 'service'    # Bosch ships with service/service. Yes really.
    FACTORY_IP = '192.168.0.1'      # which is also the gateway IP on most home
                                    # networks — fun for tech support calls
    MAC_OUIS = [b'\x00\x07\x5f']    # Robert Bosch GmbH OUI; basically all their cams

    def create_initial_user(self, ip, password):
        # Change all three password tiers in one shot. Walking high->low so if
        # the service write fails (the only one that's REALLY bad), we bail
        # before partially-resetting the lower-priv ones. Return False only on
        # service failure — failing on user/live is annoying but not blocking.
        auth = (BOSCH_DEFAULT_USER, 'service')
        ok = True
        for num, name in [(3, 'service'), (2, 'user'), (1, 'live')]:
            if not BoschRCP.rcp_write(ip, RCP_CMD['password'], 'P_STRING',
                                       password, auth, num=num):
                if num == 3:
                    ok = False
        return ok

    def set_network(self, ip, password, new_ip, subnet, gateway):
        # Same ordering principle as Axis: do everything that DOESN'T move the
        # camera first, then yank the rug at the end. Gateway, subnet, DHCP off,
        # then IP last. RCP is connectionless on UDP so the IP change doesn't
        # tear down a TCP session — but the camera still drops off the L2 with
        # the old ARP, so callers should re-resolve.
        auth = (BOSCH_DEFAULT_USER, password)
        ok = True
        if gateway:
            if not BoschRCP.rcp_write(ip, RCP_CMD['gateway'], 'P_STRING', gateway, auth):
                ok = False
        if subnet:
            if not BoschRCP.rcp_write(ip, RCP_CMD['subnet'], 'P_STRING', subnet, auth):
                ok = False
        # DHCP off — note T_DWORD typing here, the dhcp command doesn't take a
        # P_STRING despite "0"/"1" looking like strings. Tried that. RCP acks
        # it but does nothing.
        BoschRCP.rcp_write(ip, RCP_CMD['dhcp'], 'T_DWORD', '0', auth)
        # IP last — see above for why
        if not BoschRCP.rcp_write(ip, RCP_CMD['ip'], 'P_STRING', new_ip, auth):
            ok = False
        return ok

    def set_hostname(self, ip, password, hostname):
        # Bosch calls hostname "unit name" because RCP predates DNS-everywhere
        # conventions. It DOES end up being the camera's announced hostname on
        # the network though, so set it the same way you would a real hostname.
        auth = (BOSCH_DEFAULT_USER, password)
        return BoschRCP.rcp_write(ip, RCP_CMD['unit_name'], 'P_STRING', hostname, auth) or False

    def reboot(self, ip, password):
        # Bosch took two firmware revs to settle on the right TYPE for the
        # reboot RCP command. F_FLAG was the original (5.x firmware), T_DWORD
        # is current. Try both. If neither sticks, fall back to the HTTP /reset
        # endpoint which always works. This thing was 4 hours of head-scratching.
        auth = (BOSCH_DEFAULT_USER, password)
        if BoschRCP.rcp_write(ip, RCP_CMD['reboot'], 'F_FLAG', '1', auth, timeout=10):
            return True
        if BoschRCP.rcp_write(ip, RCP_CMD['reboot'], 'T_DWORD', '1', auth, timeout=10):
            return True
        try:
            r = requests.get(f'http://{ip}/reset', timeout=5)
            if r.status_code == 200:
                return True
        except requests.exceptions.ConnectionError:
            # Camera dropped while booting — call it good
            return True
        except:
            pass
        return False

    def set_dhcp(self, ip, password, enable=True):
        # Same T_DWORD typing gotcha as set_network — do not pass strings here.
        auth = (BOSCH_DEFAULT_USER, password)
        payload = '1' if enable else '0'
        return BoschRCP.rcp_write(ip, RCP_CMD['dhcp'], 'T_DWORD', payload, auth) or False

    def get_serial(self, ip, password):
        # Bosch doesn't expose a serial number over RCP the way Axis does
        # (Properties.System.SerialNumber). Closest we get is the MAC, which
        # is also engraved on the box. We dehyphenate it and pretend that's
        # the serial — it's unique per camera and that's what matters for
        # the report. T_OCTET is byte-level, comes back as space-separated hex.
        mac = BoschRCP.rcp_read(ip, RCP_CMD['mac'], 'T_OCTET')
        if mac:
            parts = mac.split()
            if len(parts) >= 6:
                return ''.join(p.upper() for p in parts[:6])
        return 'UNKNOWN'

    def get_model_noauth(self, ip):
        # /config.js is a public endpoint Bosch ships with the web UI. No auth.
        # That's where Configuration Manager pulls the model string from too,
        # which is how I found it.
        info = BoschRCP.get_device_info(ip, timeout=2)
        if info and info.get('model'):
            return info['model']
        return None

    def get_firmware(self, ip, password):
        # Same /config.js path as above — model and firmware live in the same
        # blob. password is unused here on purpose (Bosch coughs it up to
        # everyone), kept in the signature to match the abstract interface.
        info = BoschRCP.get_device_info(ip, timeout=3)
        if info and info.get('firmware'):
            return info['firmware']
        return 'UNKNOWN'

    def get_image(self, ip, username, password):
        # Two attempts: with auth, then without. Bosch snapshots are usually
        # auth-required but some firmware/SKU combos (notably the FLEXIDOME
        # IP starlight 7000 line at one point) shipped with snap.jpg open by
        # default. Try authed first, that's the right way; fall back so we
        # still get a thumbnail for the verification panel.
        #
        # JpegSize=M is "medium" — about 640x480 in most camera configs.
        # Bandwidth-friendly for the discovery wizard. Caller can re-pull L
        # if they want something better for the report.
        try:
            r = requests.get(f'http://{ip}/snap.jpg?JpegSize=M',
                             auth=HTTPDigestAuth(username, password), timeout=TIMEOUT)
            # >1000 byte check — some cams return a 200 with a tiny stub image
            # ("no signal" placeholder) when the encoder isn't ready yet.
            # Real snapshots are always over a kilobyte.
            if r.status_code == 200 and len(r.content) > 1000:
                return r.content
        except:
            pass
        try:
            r = requests.get(f'http://{ip}/snap.jpg?JpegSize=M', timeout=TIMEOUT)
            if r.status_code == 200 and len(r.content) > 1000:
                return r.content
        except:
            pass
        return None

    def test_password(self, ip, username, password):
        # Reading unit_name is cheap and won't lock the account on a wrong pw.
        # Don't use anything that writes — Bosch's lockout logic is aggressive
        # if you fat-finger the service password too many times.
        result = BoschRCP.rcp_read(ip, RCP_CMD['unit_name'], 'P_STRING',
                                   auth=(username, password))
        return result is not None

    def change_password(self, ip, username, old_pwd, new_pwd):
        # Same three-tier walk as create_initial_user. Use auth = (service, old)
        # because the service-level account is the only one allowed to write
        # password fields. Doing this from a non-service login is a 401.
        auth = (BOSCH_DEFAULT_USER, old_pwd)
        ok = True
        for num, name in [(3, 'service'), (2, 'user'), (1, 'live')]:
            if not BoschRCP.rcp_write(ip, RCP_CMD['password'], 'P_STRING',
                                       new_pwd, auth, num=num):
                if num == 3:
                    ok = False
        return ok

    def factory_reset(self, ip, password):
        # Two paths. RCP factory_reset first — it's the documented one and
        # respects the 'preserve network' flag if the camera supports it
        # (firmware-dependent, FLEXIDOME 7000 series and up). HTTP /reset is
        # a sledgehammer that wipes everything; that's our fallback.
        auth = (BOSCH_DEFAULT_USER, password)
        if BoschRCP.rcp_write(ip, RCP_CMD['factory_reset'], 'T_DWORD', '1', auth, timeout=10):
            return True
        try:
            r = requests.get(f'http://{ip}/reset', timeout=5)
            if r.status_code == 200:
                return True
        except requests.exceptions.ConnectionError:
            return True  # boot dropped the connection -> success
        except:
            pass
        return False

    def get_programming_steps(self, cam, password, options=None):
        # Bosch needs a reboot after network change — the new IP doesn't fully
        # commit to flash without one. Found this when staged cameras would
        # forget their static IP after a power cycle on site. Add the reboot
        # step explicitly and tail the programming with it.
        ip = cam.get('_program_ip', cam['ip'])
        static_ip = cam['ip']
        gateway = cam['gateway']
        subnet = cam['subnet']
        set_hostname = options.get('set_hostname', False) if options else False

        steps = [
            ("Setting password via RCP",
             lambda: self.create_initial_user(ip, password)),
            ("Setting network via RCP",
             lambda: self.set_network(ip, password, static_ip, subnet, gateway)),
            ("Rebooting camera",
             lambda: self.reboot(ip, password)),
        ]
        if set_hostname:
            cam_number = cam.get('number', '1')
            serial = cam.get('serial', 'unknown')
            hostname = f"{cam_number}-bosch-{serial.lower()}"
            steps.append(("Setting unit name",
                lambda: self.set_hostname(static_ip, password, hostname)))
        return steps

    def get_discovery_info(self, ip, timeout=2):
        # Returns enough metadata for the discovery dialog to pre-populate a
        # camera row with model, MAC-as-serial, current network config. The
        # network read matters because we want to TELL the user what the
        # camera currently has set, not just our defaults.
        info = BoschRCP.get_device_info(ip, timeout=timeout)
        if info:
            net = BoschRCP.get_network_config(ip, timeout=timeout)
            cam = {
                'ip': ip,
                'model': info.get('model', ''),
                'serial': '',
                'brand': 'bosch',
            }
            if net.get('mac'):
                cam['mac'] = net['mac']
                # Use bare MAC as serial to keep it stable across reboots.
                # Customer paperwork already expects a 12-char hex string.
                cam['serial'] = net['mac'].replace(':', '')
            if net.get('subnet'):
                cam['subnet'] = net['subnet']
            # 0.0.0.0 gateway means "no gateway set" not "literal 0.0.0.0",
            # so swallow that as a missing value rather than report it.
            if net.get('gateway') and net['gateway'] != '0.0.0.0':
                cam['gateway'] = net['gateway']
            if net.get('dhcp'):
                cam['dhcp'] = net['dhcp']
            return cam
        return None


# ============================================================================
# HANWHA/WISENET PROTOCOL
# ============================================================================
class HanwhaProtocol(CameraProtocol):
    BRAND_NAME = 'Hanwha'
    BRAND_KEY = 'hanwha'
    DEFAULT_USER = 'admin'
    DEFAULT_PASSWORD = 'admin'
    FACTORY_IP = '192.168.1.100'
    MAC_OUIS = [b'\x00\x09\x18', b'\x00\x16\x6c', b'\x00\x09\x12', b'\xc4\xf1\xd1', b'\x9c\xdc\x71']
    LOCKOUT_COOLDOWN = HANWHA_LOCKOUT_COOLDOWN

    def _stw_get(self, ip, path, auth=None, timeout=None):
        # GET wrapper. Hanwha is digest auth (HTTPDigestAuth) — DO NOT use
        # basic. Tried that for an hour once when a tcpdump showed plaintext
        # auth flying back; turned out the camera was sending a 401 with a
        # WWW-Authenticate: Digest the requests lib was happily ignoring
        # because I'd handed it HTTPBasicAuth.
        if timeout is None:
            timeout = TIMEOUT
        kwargs = {'timeout': timeout}
        if auth:
            kwargs['auth'] = HTTPDigestAuth(auth[0], auth[1])
        return requests.get(f"http://{ip}{path}", **kwargs)

    def _stw_post(self, ip, path, data=None, auth=None, timeout=None):
        # POST wrapper. STW-CGI uses form-urlencoded data, NOT JSON.
        # The newer Hanwha cameras have a JSON-y endpoint behind /stw-cgi/json
        # but the form path is older and supported on every camera I've ever
        # touched, including the 4MP QND series from 2017. Backwards compat
        # wins here.
        if timeout is None:
            timeout = TIMEOUT
        kwargs = {'timeout': timeout}
        if auth:
            kwargs['auth'] = HTTPDigestAuth(auth[0], auth[1])
        if data:
            kwargs['data'] = data
        return requests.post(f"http://{ip}{path}", **kwargs)

    def create_initial_user(self, ip, password):
        # Hanwha factory cameras don't have a default password — they REQUIRE
        # you to set one before anything else works (similar to Axis 7.10+).
        # Password rules are 8-15 chars and 3+ character types (upper/lower/
        # digit/special). If the customer hands you a 6-char password, this
        # call will 200 OK and silently not save it. Validate before sending.
        try:
            r = self._stw_post(ip,
                '/stw-cgi/user.cgi?msubmenu=admin&action=set',
                data={'NewPassword': password, 'ID': 'admin'},
                timeout=10)
            return r.status_code == 200
        except:
            return False

    def set_network(self, ip, password, new_ip, subnet, gateway):
        # Hanwha is the easiest of the three for network config — one POST
        # sets all four fields atomically. The Type='Static' is what flips
        # DHCP off as a side effect; you don't have to set DHCP separately.
        # Wisenet labels the field 'Type' (not 'IPType') in the API even
        # though the web UI calls it "IP Type". Took me a minute to find that.
        try:
            r = self._stw_post(ip,
                '/stw-cgi/network.cgi?msubmenu=ethernet&action=set',
                data={
                    'Type': 'Static',
                    'IPAddress': new_ip,
                    'SubnetMask': subnet,
                    'DefaultGateway': gateway,
                },
                auth=('admin', password),
                timeout=10)
            return r.status_code == 200
        except:
            return False

    def set_hostname(self, ip, password, hostname):
        # Same DNS-clean rule as the others. Hanwha calls it "HostName" on the
        # general system page, doesn't expose a Bonjour/mDNS field separately.
        clean = re.sub(r'[^a-zA-Z0-9-]', '-', hostname).strip('-')[:63]
        try:
            r = self._stw_post(ip,
                '/stw-cgi/system.cgi?msubmenu=general&action=set',
                data={'HostName': clean},
                auth=('admin', password))
            return r.status_code == 200
        except:
            return False

    def reboot(self, ip, password):
        # The reboot endpoint immediately drops the connection — POST returns
        # ConnectionError or Timeout, both of which mean "successfully rebooted".
        # If you get a 200 back something is weird (camera queued the reboot
        # for later? rare).
        try:
            r = self._stw_post(ip,
                '/stw-cgi/system.cgi?msubmenu=reboot&action=execute',
                auth=('admin', password), timeout=10)
            return r.status_code == 200
        except requests.exceptions.ConnectionError:
            return True  # camera bounced before sending the response
        except requests.exceptions.Timeout:
            return True
        except:
            return False

    def set_dhcp(self, ip, password, enable=True):
        # Just flip the Type back to DHCP. Same endpoint as set_network but
        # we don't include the IP/subnet/gateway fields — camera will pull
        # those from DHCP after the next renew.
        try:
            mode = 'DHCP' if enable else 'Static'
            r = self._stw_post(ip,
                '/stw-cgi/network.cgi?msubmenu=ethernet&action=set',
                data={'Type': mode},
                auth=('admin', password))
            return r.status_code == 200
        except:
            return False

    def get_serial(self, ip, password):
        # The deviceinfo endpoint returns a flat key=value text block (NOT
        # JSON despite what you'd expect from an API named STW-CGI). Walk it
        # twice: first looking for explicit "SerialNumber=" / "Serial=" keys,
        # then a fuzzy match on anything containing "serial" — different
        # firmware revisions use different exact key names. Standard Wisenet
        # spelling has been moving target since at least the QND-7080R era.
        try:
            r = self._stw_get(ip,
                '/stw-cgi/system.cgi?msubmenu=deviceinfo&action=view',
                auth=('admin', password))
            if r.status_code == 200:
                for line in r.text.split('\n'):
                    if 'SerialNumber=' in line or 'Serial=' in line:
                        return line.split('=', 1)[1].strip()
                # Fuzzy fallback for variant key names
                for line in r.text.split('\n'):
                    line = line.strip()
                    if '=' in line:
                        key, val = line.split('=', 1)
                        if 'serial' in key.lower():
                            return val.strip()
        except:
            pass
        return 'UNKNOWN'

    def get_model_noauth(self, ip):
        # Most Hanwha firmware leaves deviceinfo open without auth (just the
        # model field, not the serial). Used during discovery to fingerprint
        # without needing a password. Newer locked-down builds (post 2.20)
        # might 401 this; that's why discovery has the 401 fallback path.
        try:
            r = self._stw_get(ip,
                '/stw-cgi/system.cgi?msubmenu=deviceinfo&action=view',
                timeout=3)
            if r.status_code == 200:
                for line in r.text.split('\n'):
                    line = line.strip()
                    if '=' in line:
                        key, val = line.split('=', 1)
                        if 'model' in key.lower():
                            return val.strip()
        except:
            pass
        return None

    def get_firmware(self, ip, password):
        # Two-pass: authed first (most reliable, gets the full firmware string
        # including patch level), then no-auth (some firmware leaks the version
        # to anonymous probes). Walk the response for any key that smells like
        # a firmware version — the field name has changed between Wisenet
        # firmware lines (FwVersion, FirmwareVersion, AppFwVersion, etc.).
        for auth in [('admin', password) if password else None, None]:
            try:
                r = self._stw_get(ip,
                    '/stw-cgi/system.cgi?msubmenu=deviceinfo&action=view',
                    auth=auth, timeout=3)
                if r.status_code == 200:
                    for line in r.text.split('\n'):
                        line = line.strip()
                        if '=' in line:
                            key, val = line.split('=', 1)
                            kl = key.lower()
                            if 'firmware' in kl or 'fwversion' in kl or kl.endswith('version'):
                                v = val.strip()
                                if v:
                                    return v
            except:
                pass
        return 'UNKNOWN'

    def get_image(self, ip, username, password):
        # Hanwha snapshot path. Same >1000-byte sanity check as Bosch — if a
        # camera returns a 200 with a tiny image it's usually the placeholder
        # because the encoder is still warming up post-reboot.
        try:
            r = self._stw_get(ip,
                '/stw-cgi/image.cgi?msubmenu=snapshot&action=view',
                auth=(username, password))
            if r.status_code == 200 and len(r.content) > 1000:
                return r.content
        except:
            pass
        return None

    def test_password(self, ip, username, password):
        # Hanwha lockout is the worst of the three brands. After ~5 wrong
        # passwords the camera returns HTTP 490 (a non-standard code Hanwha
        # invented for "you've been bad") and the account is unusable for
        # 30 minutes. Even with the right password. There's no way to clear
        # the lockout short of a factory reset or waiting it out.
        #
        # We propagate LockoutError up the stack so the GUI can disable the
        # row and stop hammering. Not catching this and treating it as a
        # generic auth fail will keep the lockout timer rolling forever.
        try:
            r = self._stw_get(ip,
                '/stw-cgi/system.cgi?msubmenu=deviceinfo&action=view',
                auth=(username, password))
            if r.status_code == 200:
                return True
            elif r.status_code == 490:
                raise LockoutError(f"Camera {ip} is locked out (HTTP 490)")
            elif r.status_code == 401:
                return False
            return False
        except LockoutError:
            raise
        except (requests.exceptions.ConnectionError, requests.exceptions.Timeout):
            # Hanwha sometimes goes silent during lockout instead of returning
            # 490 — same effective behavior, treat it the same way so we
            # don't keep retrying and extending the cooldown.
            raise LockoutError(f"Camera {ip} not responding (likely locked out)")
        except:
            return False

    def change_password(self, ip, username, old_pwd, new_pwd):
        # Note Hanwha REQUIRES OldPassword in the same call as NewPassword —
        # it's a CSRF guard, not just for verification. Sending only NewPassword
        # silently no-ops. Found out the hard way during a customer pw rotation
        # where every camera "succeeded" but kept the old password.
        try:
            r = self._stw_post(ip,
                '/stw-cgi/user.cgi?msubmenu=admin&action=set',
                data={'OldPassword': old_pwd, 'NewPassword': new_pwd, 'ID': 'admin'},
                auth=('admin', old_pwd))
            return r.status_code == 200
        except:
            return False

    def add_user(self, ip, admin_password, username, user_password, role='Operator'):
        # Hanwha role names: admin / manager / user. Manager == Axis Operator,
        # User == Axis Viewer. Sticking with our standard Administrator/Operator/
        # Viewer terminology in the GUI and remapping here.
        role_map = {
            'Administrator': 'admin',
            'Operator': 'manager',
            'Viewer': 'user',
        }
        group = role_map.get(role, 'manager')
        try:
            r = self._stw_post(ip,
                '/stw-cgi/user.cgi?msubmenu=adduser&action=set',
                data={'UserID': username, 'Password': user_password, 'UserGroup': group},
                auth=('admin', admin_password))
            return r.status_code == 200
        except:
            return False

    def factory_reset(self, ip, password):
        # No mode flag on Hanwha — factory reset is always a full hard reset.
        # Network config, password, certs, schedules, all gone. The camera comes
        # back up factory-default and waiting for the initial admin password.
        try:
            r = self._stw_post(ip,
                '/stw-cgi/system.cgi?msubmenu=factory&action=execute',
                auth=('admin', password), timeout=10)
            return r.status_code == 200
        except requests.exceptions.ConnectionError:
            return True  # camera dropped during reset, that's the win path
        except requests.exceptions.Timeout:
            return True
        except:
            return False

    def get_programming_steps(self, cam, password, options=None):
        # Reboot at the end is mandatory on Hanwha — the network change
        # technically takes effect immediately but a freshly-programmed camera
        # behaves weirdly until it's been power-cycled (RTSP stream stutters,
        # ONVIF discovery skips it). Single-rev older firmware especially.
        # Easier to just reboot and skip the troubleshooting.
        ip = cam.get('_program_ip', cam['ip'])
        static_ip = cam['ip']
        gateway = cam['gateway']
        subnet = cam['subnet']
        set_hostname = options.get('set_hostname', False) if options else False

        steps = [
            ("Setting admin password",
             lambda: self.create_initial_user(ip, password)),
            ("Setting network (IP, subnet, gateway, static)",
             lambda: self.set_network(ip, password, static_ip, subnet, gateway)),
            ("Rebooting camera",
             lambda: self.reboot(ip, password)),
        ]
        if set_hostname:
            cam_number = cam.get('number', '1')
            serial = cam.get('serial', 'unknown')
            hostname = f"{cam_number}-hanwha-{serial.lower()}"
            steps.append(("Setting hostname",
                lambda: self.set_hostname(static_ip, password, hostname)))
        return steps

    def get_discovery_info(self, ip, timeout=2):
        # 200 OR 401 both mean "Hanwha is here." 401 means it's locked down,
        # which is still useful info for the discovery dialog. Anything else
        # (timeout, 404, connection refused) and we return None so the
        # multi-protocol scanner can try the next brand.
        try:
            r = requests.get(f"http://{ip}/stw-cgi/system.cgi?msubmenu=deviceinfo&action=view",
                             timeout=timeout)
            if r.status_code in (200, 401):
                cam = {'ip': ip, 'brand': 'hanwha'}
                if r.status_code == 200:
                    for line in r.text.split('\n'):
                        line = line.strip()
                        if '=' in line:
                            key, val = line.split('=', 1)
                            if 'model' in key.lower():
                                cam['model'] = val.strip()
                            elif 'serial' in key.lower():
                                cam['serial'] = val.strip()
                    cam.setdefault('model', 'Hanwha Camera')
                else:
                    cam['model'] = '(Auth Required)'
                return cam
        except:
            pass
        return None


# Vendor protocol registry. Order doesn't matter for lookup but does matter
# for the multi-brand discovery sweep — that walks this dict in declaration
# order and tries each brand's discovery probe in turn. Axis first because
# their discovery is the cheapest (basicdeviceinfo no-auth), Hanwha last
# because of the lockout risk on a wrong probe.
PROTOCOLS = {
    'axis': AxisProtocol,
    'bosch': BoschProtocol,
    'hanwha': HanwhaProtocol,
}


# ============================================================================
# AXIS CAMERA DISCOVERY via mDNS (Multicast DNS / Bonjour)
# ============================================================================
class AxisMDNSDiscovery:
    """
    Discovers Axis cameras using mDNS (Multicast DNS / Bonjour).
    This is the same method used by AXIS IP Utility.
    Works with link-local (169.254.x.x) addresses on firmware 12.0+.
    Returns model, serial, MAC, IP without authentication.
    """
    MDNS_ADDR = '224.0.0.251'
    MDNS_PORT = 5353
    
    # Service types to query (Axis cameras respond to these)
    SERVICE_TYPES = [
        '_axis-video._tcp.local',
        '_vapix-http._tcp.local', 
        '_vapix-https._tcp.local',
    ]
    
    @staticmethod
    def build_mdns_query(service_name):
        """Build an mDNS PTR query packet for the given service"""
        # DNS header: ID=0, Flags=0 (standard query), QDCOUNT=1
        header = b'\x00\x00'  # Transaction ID
        header += b'\x00\x00'  # Flags: standard query
        header += b'\x00\x01'  # Questions: 1
        header += b'\x00\x00'  # Answer RRs: 0
        header += b'\x00\x00'  # Authority RRs: 0
        header += b'\x00\x00'  # Additional RRs: 0
        
        # Build the question section
        # Service name like "_axis-video._tcp.local" becomes labels
        question = b''
        for part in service_name.split('.'):
            question += bytes([len(part)]) + part.encode('utf-8')
        question += b'\x00'  # Null terminator
        question += b'\x00\x0c'  # Type: PTR (12)
        question += b'\x00\x01'  # Class: IN (1)
        
        return header + question
    
    @staticmethod
    def parse_dns_name(data, offset):
        """Parse a DNS name from packet data, handling compression"""
        parts = []
        original_offset = offset
        jumped = False
        max_jumps = 10  # Prevent infinite loops
        jumps = 0
        
        while True:
            if offset >= len(data):
                break
            length = data[offset]
            
            if length == 0:
                offset += 1
                break
            elif (length & 0xc0) == 0xc0:
                # Compression pointer
                if offset + 1 >= len(data):
                    break
                pointer = ((length & 0x3f) << 8) | data[offset + 1]
                if not jumped:
                    original_offset = offset + 2
                jumped = True
                offset = pointer
                jumps += 1
                if jumps > max_jumps:
                    break
            else:
                offset += 1
                if offset + length > len(data):
                    break
                parts.append(data[offset:offset + length].decode('utf-8', errors='ignore'))
                offset += length
        
        name = '.'.join(parts)
        return name, (original_offset if jumped else offset)
    
    @staticmethod
    def parse_mdns_response(data, source_ip, source_mac=None):
        """Parse an mDNS response packet and extract camera info"""
        try:
            if len(data) < 12:
                return None
            
            # Parse DNS header
            flags = (data[2] << 8) | data[3]
            is_response = (flags & 0x8000) != 0
            if not is_response:
                return None
            
            qdcount = (data[4] << 8) | data[5]
            ancount = (data[6] << 8) | data[7]
            nscount = (data[8] << 8) | data[9]
            arcount = (data[10] << 8) | data[11]
            
            camera = {
                'ip': source_ip,
                'mac': source_mac or '',
                'model': '',
                'serial': '',
                'name': '',
                'hostname': '',
                'ipv6': '',
            }
            
            offset = 12
            
            # Skip questions
            for _ in range(qdcount):
                _, offset = AxisMDNSDiscovery.parse_dns_name(data, offset)
                offset += 4  # Skip QTYPE and QCLASS
            
            # Parse answers and additional records
            for _ in range(ancount + nscount + arcount):
                if offset >= len(data):
                    break
                    
                name, offset = AxisMDNSDiscovery.parse_dns_name(data, offset)
                
                if offset + 10 > len(data):
                    break
                
                rtype = (data[offset] << 8) | data[offset + 1]
                rclass = (data[offset + 2] << 8) | data[offset + 3]
                ttl = (data[offset + 4] << 24) | (data[offset + 5] << 16) | (data[offset + 6] << 8) | data[offset + 7]
                rdlength = (data[offset + 8] << 8) | data[offset + 9]
                offset += 10
                
                if offset + rdlength > len(data):
                    break
                
                rdata = data[offset:offset + rdlength]
                offset += rdlength
                
                # PTR record (12) - contains service instance name with model
                if rtype == 12:
                    ptr_name, _ = AxisMDNSDiscovery.parse_dns_name(data, offset - rdlength)
                    # Extract model from PTR like "AXIS P3268-LV._axis-video._tcp.local"
                    if 'AXIS' in ptr_name.upper():
                        parts = ptr_name.split('._')
                        if parts:
                            model_part = parts[0]
                            # Handle "AXIS P3268-LV - B8A44F8BF3BB" format
                            if ' - ' in model_part:
                                model_part = model_part.split(' - ')[0]
                            camera['model'] = model_part.strip()
                
                # A record (1) - IPv4 address
                elif rtype == 1 and rdlength == 4:
                    ip = f"{rdata[0]}.{rdata[1]}.{rdata[2]}.{rdata[3]}"
                    camera['ip'] = ip
                
                # AAAA record (28) - IPv6 address  
                elif rtype == 28 and rdlength == 16:
                    ipv6_parts = [f"{rdata[i]:02x}{rdata[i+1]:02x}" for i in range(0, 16, 2)]
                    camera['ipv6'] = ':'.join(ipv6_parts)
                
                # TXT record (16) - contains serial number and other info
                elif rtype == 16:
                    txt_offset = 0
                    while txt_offset < len(rdata):
                        txt_len = rdata[txt_offset]
                        txt_offset += 1
                        if txt_offset + txt_len > len(rdata):
                            break
                        txt = rdata[txt_offset:txt_offset + txt_len].decode('utf-8', errors='ignore')
                        txt_offset += txt_len
                        
                        # Parse key=value pairs
                        if '=' in txt:
                            key, value = txt.split('=', 1)
                            if key.lower() == 'sn':
                                camera['serial'] = value.upper()
                                # Derive MAC from serial (Axis serials are MAC addresses)
                                if len(value) == 12 and not camera['mac']:
                                    camera['mac'] = ':'.join(value[i:i+2].upper() for i in range(0, 12, 2))
                
                # SRV record (33) - contains hostname
                elif rtype == 33 and rdlength >= 6:
                    # SRV format: priority(2) + weight(2) + port(2) + target
                    srv_target, _ = AxisMDNSDiscovery.parse_dns_name(data, offset - rdlength + 6)
                    if srv_target:
                        camera['hostname'] = srv_target.replace('.local', '')
            
            # Generate name from serial or hostname
            if camera['serial']:
                camera['name'] = f"axis-{camera['serial'].lower()}"
            elif camera['hostname']:
                camera['name'] = camera['hostname']
            else:
                camera['name'] = source_ip
            
            # Only return if we got meaningful data
            if camera['model'] or camera['serial']:
                return camera
            return None
            
        except Exception as e:
            return None
    
    @staticmethod
    def discover(timeout=5, callback=None):
        """
        Discover Axis cameras using mDNS.
        Returns list of dicts with camera info.
        callback(camera_dict) is called for each camera found (optional).
        
        Tries zeroconf library first (handles Windows quirks), falls back to manual.
        """
        # Try using zeroconf library if available (handles all edge cases)
        try:
            return AxisMDNSDiscovery._discover_zeroconf(timeout, callback)
        except ImportError:
            pass
        except Exception:
            pass
        
        # Fall back to manual mDNS implementation
        return AxisMDNSDiscovery._discover_manual(timeout, callback)
    
    @staticmethod
    def _discover_zeroconf(timeout=5, callback=None):
        """Discover using zeroconf library (pip install zeroconf)"""
        from zeroconf import Zeroconf, ServiceBrowser, ServiceListener
        
        cameras = []
        seen = set()
        
        class AxisListener(ServiceListener):
            def add_service(self, zc, type_, name):
                info = zc.get_service_info(type_, name)
                if info:
                    camera = {
                        'ip': '',
                        'mac': '',
                        'model': '',
                        'serial': '',
                        'name': '',
                        'hostname': info.server.rstrip('.').replace('.local', '') if info.server else '',
                    }
                    
                    # Get IP addresses
                    if info.addresses:
                        import socket
                        camera['ip'] = socket.inet_ntoa(info.addresses[0])
                    
                    # Parse service name for model (e.g., "AXIS P3268-LV._axis-video._tcp.local")
                    if 'AXIS' in name.upper():
                        parts = name.split('._')
                        if parts:
                            model_part = parts[0]
                            if ' - ' in model_part:
                                model_part = model_part.split(' - ')[0]
                            camera['model'] = model_part.strip()
                    
                    # Parse TXT records for serial
                    if info.properties:
                        for key, val in info.properties.items():
                            if isinstance(key, bytes):
                                key = key.decode('utf-8', errors='ignore')
                            if isinstance(val, bytes):
                                val = val.decode('utf-8', errors='ignore')
                            if key.lower() == 'sn':
                                camera['serial'] = val.upper()
                                if len(val) == 12:
                                    camera['mac'] = ':'.join(val[i:i+2].upper() for i in range(0, 12, 2))
                    
                    # Generate name
                    if camera['serial']:
                        camera['name'] = f"axis-{camera['serial'].lower()}"
                    elif camera['hostname']:
                        camera['name'] = camera['hostname']
                    else:
                        camera['name'] = camera['ip']
                    
                    # Dedupe and add
                    key = camera.get('mac') or camera.get('ip')
                    if key and key not in seen and (camera['model'] or camera['serial']):
                        seen.add(key)
                        cameras.append(camera)
                        if callback:
                            callback(camera)
            
            def remove_service(self, zc, type_, name):
                pass
            
            def update_service(self, zc, type_, name):
                pass
        
        zc = Zeroconf()
        listener = AxisListener()
        browsers = []
        
        for service in AxisMDNSDiscovery.SERVICE_TYPES:
            browsers.append(ServiceBrowser(zc, service, listener))
        
        time.sleep(timeout)
        
        for browser in browsers:
            browser.cancel()
        zc.close()
        
        return cameras
    
    @staticmethod
    def _discover_manual(timeout=5, callback=None):
        """Manual mDNS discovery (fallback when zeroconf not available)"""
        cameras = []
        seen = set()  # Track by MAC or IP to avoid duplicates
        
        try:
            # Create UDP socket for mDNS
            sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM, socket.IPPROTO_UDP)
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            
            # On Windows, try SO_EXCLUSIVEADDRUSE = False to allow port sharing
            try:
                sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
            except:
                pass
            
            # Set multicast TTL to 255 (required for mDNS)
            sock.setsockopt(socket.IPPROTO_IP, socket.IP_MULTICAST_TTL, 255)
            
            # Enable multicast loopback (to see our own queries, helps debugging)
            sock.setsockopt(socket.IPPROTO_IP, socket.IP_MULTICAST_LOOP, 1)
            
            # Try to bind to mDNS port - may fail if already in use
            bound = False
            try:
                sock.bind(('', AxisMDNSDiscovery.MDNS_PORT))
                bound = True
            except OSError:
                # Port 5353 in use (e.g., Bonjour service) - bind to any port
                sock.bind(('', 0))
            
            # Join multicast group on all interfaces
            try:
                mreq = socket.inet_aton(AxisMDNSDiscovery.MDNS_ADDR) + socket.inet_aton('0.0.0.0')
                sock.setsockopt(socket.IPPROTO_IP, socket.IP_ADD_MEMBERSHIP, mreq)
            except Exception as e:
                pass  # May fail on some systems, continue anyway
            
            sock.settimeout(0.5)
            
            # Send queries for each service type
            for service in AxisMDNSDiscovery.SERVICE_TYPES:
                try:
                    query = AxisMDNSDiscovery.build_mdns_query(service)
                    sock.sendto(query, (AxisMDNSDiscovery.MDNS_ADDR, AxisMDNSDiscovery.MDNS_PORT))
                except Exception as e:
                    pass
            
            # Listen for responses
            end_time = time.time() + timeout
            queries_sent = 1
            while time.time() < end_time:
                try:
                    data, addr = sock.recvfrom(4096)
                    source_ip = addr[0]
                    
                    camera = AxisMDNSDiscovery.parse_mdns_response(data, source_ip)
                    if camera:
                        # Dedupe by MAC or IP
                        key = camera.get('mac') or camera.get('ip')
                        if key and key not in seen:
                            seen.add(key)
                            cameras.append(camera)
                            if callback:
                                callback(camera)
                                
                except socket.timeout:
                    # Send more queries periodically
                    if queries_sent < 3 and time.time() < end_time - 1:
                        for service in AxisMDNSDiscovery.SERVICE_TYPES:
                            try:
                                query = AxisMDNSDiscovery.build_mdns_query(service)
                                sock.sendto(query, (AxisMDNSDiscovery.MDNS_ADDR, AxisMDNSDiscovery.MDNS_PORT))
                            except:
                                pass
                        queries_sent += 1
                except Exception:
                    continue
            
            sock.close()
        except Exception as e:
            print(f"mDNS Discovery error: {e}")
        
        return cameras


# ============================================================================
# SMART DATA ANALYZER
# ============================================================================
class SmartDataAnalyzer:
    """Analyzes data columns to guess what they contain"""
    
    @staticmethod
    def is_ip_address(value):
        """Check if value looks like an IP address"""
        pattern = r'^(\d{1,3}\.){3}\d{1,3}$'
        if re.match(pattern, value):
            parts = value.split('.')
            return all(0 <= int(p) <= 255 for p in parts)
        return False
    
    @staticmethod
    def is_subnet_mask(value):
        """Check if value looks like a subnet mask (255.x.x.x)"""
        if not SmartDataAnalyzer.is_ip_address(value):
            return False
        return value.startswith('255.')
    
    @staticmethod
    def is_likely_gateway(value):
        """Check if value looks like a gateway (ends in .1, .254, etc.)"""
        if not SmartDataAnalyzer.is_ip_address(value):
            return False
        last_octet = int(value.split('.')[-1])
        return last_octet in [1, 2, 254]
    
    @staticmethod
    def is_mac_address(value):
        """Check if value looks like a MAC address"""
        value = value.upper().strip()
        if re.match(r'^([0-9A-F]{2}:){5}[0-9A-F]{2}$', value):
            return True
        if re.match(r'^([0-9A-F]{2}-){5}[0-9A-F]{2}$', value):
            return True
        if re.match(r'^[0-9A-F]{12}$', value):
            return True
        return False
    
    @staticmethod
    def is_model(value):
        """Check if value looks like an Axis model"""
        value = value.upper().strip()
        if re.match(r'^[PMQVFAT]\d{4}', value):
            return True
        if 'AXIS' in value:
            return True
        return False
    
    @staticmethod
    def find_repeated_ips(values):
        """Find IPs that repeat - likely gateways"""
        ip_counts = {}
        for v in values:
            v = v.strip()
            if SmartDataAnalyzer.is_ip_address(v):
                ip_counts[v] = ip_counts.get(v, 0) + 1
        
        # If any IP appears more than once, it's likely a gateway
        repeated = [ip for ip, count in ip_counts.items() if count > 1]
        return repeated
    
    @staticmethod
    def guess_column_type(values, all_columns_data=None):
        """Guess what type of data a column contains"""
        non_empty = [v.strip() for v in values if v.strip()]
        if not non_empty:
            return 'unknown', 0
        
        total = len(non_empty)
        
        # Check for subnet first (most specific: 255.x.x.x)
        subnet_count = sum(1 for v in non_empty if SmartDataAnalyzer.is_subnet_mask(v))
        if subnet_count > total * 0.5:
            return 'subnet', subnet_count / total
        
        # Check for repeated IPs (strong gateway indicator)
        repeated_ips = SmartDataAnalyzer.find_repeated_ips(non_empty)
        if repeated_ips:
            repeated_count = sum(1 for v in non_empty if v.strip() in repeated_ips)
            if repeated_count > total * 0.5:
                return 'gateway', repeated_count / total
        
        # Check for gateway pattern (ends in .1, .254)
        gateway_count = sum(1 for v in non_empty if SmartDataAnalyzer.is_likely_gateway(v))
        if gateway_count > total * 0.7:
            return 'gateway', gateway_count / total
        
        # Check for MAC address
        mac_count = sum(1 for v in non_empty if SmartDataAnalyzer.is_mac_address(v))
        if mac_count > total * 0.5:
            return 'mac', mac_count / total
        
        # Check for regular IP (not subnet, not gateway)
        ip_count = sum(1 for v in non_empty if SmartDataAnalyzer.is_ip_address(v))
        if ip_count > total * 0.7:
            if subnet_count == 0 and gateway_count < total * 0.3:
                return 'ip', ip_count / total
        
        # Check for model (Axis model patterns)
        model_count = sum(1 for v in non_empty if SmartDataAnalyzer.is_model(v))
        if model_count > total * 0.3:
            return 'model', model_count / total
        
        # Check for small integers (1-3 digits, likely row numbers, phase, DI, port)
        small_int_count = sum(1 for v in non_empty if re.match(r'^\d{1,3}$', v))
        if small_int_count > total * 0.5:
            return 'unknown', small_int_count / total
        
        # Check for switch-like names (contain "sw", "switch", "acc", etc.)
        switch_count = sum(1 for v in non_empty if re.search(
            r'(sw\d|switch|acc[- ]|dist[- ]|core[- ]|mdf|idf)', v, re.IGNORECASE))
        if switch_count > total * 0.3:
            return 'unknown', switch_count / total
        
        # Check for rack-like names (contain "rack", "room", "zone", "closet", "mdf")
        rack_count = sum(1 for v in non_empty if re.search(
            r'(rack|room|zone|closet|mdf|idf|cabinet|row)', v, re.IGNORECASE))
        if rack_count > total * 0.3:
            return 'unknown', rack_count / total
        
        # Check for port patterns (Gi1/0/1, Fa0/1, ge-0/0/1, etc.)
        port_count = sum(1 for v in non_empty if re.match(
            r'^(Gi|Fa|Te|Eth|ge-|xe-|et-)', v, re.IGNORECASE))
        if port_count > total * 0.3:
            return 'unknown', port_count / total
        
        # Default to name - long text with mixed chars is most likely camera name
        return 'name', 0.5
    
    # Header keywords to field mapping
    HEADER_MAP = {
        'name': 'name', 'camera': 'name', 'camera name': 'name', 'camera_name': 'name',
        'cam': 'name', 'description': 'name',
        'ip': 'ip', 'ip address': 'ip', 'ip_address': 'ip', 'ipaddress': 'ip', 'address': 'ip',
        'gateway': 'gateway', 'default gateway': 'gateway', 'router': 'gateway', 
        'broadcast/router': 'gateway', 'broadcast router': 'gateway',
        'broadcast/router/gateway': 'gateway', 'broadcast router gateway': 'gateway',
        'default router': 'gateway', 'defaultrouter': 'gateway',
        'subnet': 'subnet', 'subnet mask': 'subnet', 'subnetmask': 'subnet', 'mask': 'subnet',
        'netmask': 'subnet', 'subnet_mask': 'subnet',
        'model': 'model', 'camera model': 'model', 'type': 'model',
        'mac': 'mac', 'mac address': 'mac', 'mac_address': 'mac', 'macaddress': 'mac',
        'serial': 'serial', 'serial number': 'serial', 'serialnumber': 'serial',
        'rack': 'rack', 'rack location': 'rack', 'rack_location': 'rack',
        'switch': 'switch', 'switch name': 'switch', 'switch_name': 'switch',
        'port': 'port', 'switch port': 'port', 'switch_port': 'port', 'switchport': 'port',
        'new ip': 'new_ip', 'new_ip': 'new_ip', 'target ip': 'new_ip',
        # Common spreadsheet columns to skip
        'number': 'unknown', 'num': 'unknown', '#': 'unknown',
        'phase': 'unknown', 'phase/rom': 'unknown', 'phase rom': 'unknown', 'rom': 'unknown',
        'di': 'unknown', 'di#': 'unknown', 'di #': 'unknown',
        'vlan': 'unknown', 'notes': 'unknown', 'comments': 'unknown', 'location': 'unknown',
    }
    
    @staticmethod
    def detect_header_row(row):
        """Check if a row looks like a header. Returns field mapping dict or None."""
        matches = {}
        for col_idx, val in enumerate(row):
            clean = val.strip().strip('#').strip().lower()
            # Remove common punctuation and normalize whitespace
            clean = re.sub(r'[:/\n\r]', ' ', clean).strip()
            clean = re.sub(r'\s+', ' ', clean)
            # Try exact match first
            if clean in SmartDataAnalyzer.HEADER_MAP:
                matches[col_idx] = SmartDataAnalyzer.HEADER_MAP[clean]
                continue
            # Try partial/keyword match
            for keyword, field in SmartDataAnalyzer.HEADER_MAP.items():
                if keyword in clean or clean in keyword:
                    matches[col_idx] = field
                    break
            # If still no match, check individual words
            if col_idx not in matches:
                words = clean.split()
                for word in words:
                    if word in SmartDataAnalyzer.HEADER_MAP:
                        matches[col_idx] = SmartDataAnalyzer.HEADER_MAP[word]
                        break
        
        # If we matched 2+ columns including IP, it's a header row
        if len(matches) >= 2 and 'ip' in matches.values():
            return matches
        return None
    
    @staticmethod
    def analyze_data(rows):
        """
        Analyze rows of data and return column mapping.
        Returns dict: {column_index: {'type': str, 'confidence': float, 'sample': str}}
        """
        if not rows:
            return {}
        
        # Check first row for headers
        header_map = SmartDataAnalyzer.detect_header_row(rows[0])
        if header_map:
            # First row is headers - use header mapping and skip header row
            rows = rows[1:]  # NOTE: modifies local reference only
            column_types = {}
            for col_idx in range(max(len(rows[0]) if rows else 0, len(header_map))):
                field = header_map.get(col_idx, 'unknown')
                # Skip non-camera fields
                if field in ('rack', 'switch', 'port', 'unknown'):
                    field = 'unknown'
                sample = rows[0][col_idx] if rows and col_idx < len(rows[0]) else ''
                column_types[col_idx] = {
                    'type': field,
                    'confidence': 1.0,
                    'sample': sample
                }
            return column_types
        
        num_cols = max(len(row) for row in rows)
        
        # First pass: collect all column data
        all_columns = {}
        for col_idx in range(num_cols):
            all_columns[col_idx] = [row[col_idx] if col_idx < len(row) else '' for row in rows]
        
        # Second pass: analyze each column with awareness of other columns
        column_types = {}
        assigned_types = set()
        
        for col_idx in range(num_cols):
            values = all_columns[col_idx]
            col_type, confidence = SmartDataAnalyzer.guess_column_type(values, all_columns)
            
            # Skip non-camera fields
            if col_type in ('rack', 'switch', 'port'):
                col_type = 'unknown'
            
            # Avoid duplicate assignments (prefer first match)
            if col_type in assigned_types and col_type in ['ip', 'gateway', 'subnet']:
                # Check if this might be new_ip
                if col_type == 'ip':
                    col_type = 'new_ip'
                else:
                    col_type = 'unknown'
            
            if col_type != 'unknown':
                assigned_types.add(col_type)
            
            column_types[col_idx] = {
                'type': col_type,
                'confidence': confidence,
                'sample': values[0] if values else ''
            }
        
        # Promote first small-integer 'unknown' column to 'number' (camera ID for hostname)
        if 'number' not in assigned_types:
            for col_idx in range(num_cols):
                if column_types[col_idx]['type'] == 'unknown':
                    values = [v.strip() for v in all_columns[col_idx] if v.strip()]
                    int_count = sum(1 for v in values if re.match(r'^\d+$', v))
                    if values and int_count > len(values) * 0.5:
                        column_types[col_idx]['type'] = 'number'
                        break  # Only the first one
        
        return column_types


# ============================================================================
# SMART IMPORT DIALOG
# ============================================================================
class SmartImportDialog(tk.Toplevel):
    """Dialog for importing data with smart column detection"""
    
    def __init__(self, parent, initial_data=None):
        super().__init__(parent)
        self.title("Smart Import")
        self.result = None
        self.transient(parent)
        self.grab_set()
        
        # Center on PARENT monitor (not monitor 1) — multi-monitor fix 2026-04-30.
        # Use _center_on_parent which uses parent.winfo_rootx/y to track which
        # screen the parent is on. winfo_screenwidth() is the PRIMARY screen,
        # not the screen that contains parent.
        _center_on_parent(self, parent, 1400, 800)
        
        self.rows = []
        self.column_mappings = {}
        
        frame = ttk.Frame(self, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Instructions
        ttk.Label(frame, text="Smart Import - Paste or Load Camera Data", 
                 font=('Helvetica', 14, 'bold')).pack(anchor=tk.W)
        ttk.Label(frame, text="Paste data below or load from file. Columns are auto-detected.", 
                 font=('Helvetica', 10), foreground='gray').pack(anchor=tk.W, pady=(0, 10))
        
        # Text area for pasting
        paste_frame = ttk.LabelFrame(frame, text="Paste Data Here (CSV, tab-separated, or any format)", padding="5")
        paste_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.paste_text = scrolledtext.ScrolledText(paste_frame, font=('Courier', 10), height=10)
        self.paste_text.pack(fill=tk.BOTH, expand=True)
        
        # Right-click context menu for paste area
        self.context_menu = tk.Menu(self.paste_text, tearoff=0)
        self.context_menu.add_command(label="Paste", command=self._paste)
        self.context_menu.add_command(label="Select All", command=lambda: self.paste_text.tag_add('sel', '1.0', 'end'))
        self.context_menu.add_command(label="Clear", command=lambda: self.paste_text.delete('1.0', tk.END))
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Paste & Analyze", command=self._paste_and_analyze)
        self.paste_text.bind("<Button-3>", self._show_context_menu)
        
        if initial_data:
            self.paste_text.insert('1.0', initial_data)
        
        # Buttons for paste area
        paste_btn_frame = ttk.Frame(paste_frame)
        paste_btn_frame.pack(fill=tk.X, pady=(5, 0))
        ttk.Button(paste_btn_frame, text="📁 Load from File", command=self.load_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(paste_btn_frame, text="🔍 Analyze Data", command=self.analyze_data).pack(side=tk.LEFT, padx=2)
        ttk.Button(paste_btn_frame, text="📋 Paste from Clipboard", command=self.paste_clipboard).pack(side=tk.LEFT, padx=2)
        
        # Column mapping area - scrollable for wide data
        mapping_outer = ttk.LabelFrame(frame, text="Column Mapping (click to change)", padding="10")
        mapping_outer.pack(fill=tk.X, pady=(0, 10))
        
        mapping_canvas = tk.Canvas(mapping_outer, height=90)
        mapping_scroll = ttk.Scrollbar(mapping_outer, orient='horizontal', command=mapping_canvas.xview)
        self.mapping_container = ttk.Frame(mapping_canvas)
        
        self.mapping_container.bind("<Configure>", lambda e: mapping_canvas.configure(scrollregion=mapping_canvas.bbox("all")))
        mapping_canvas.create_window((0, 0), window=self.mapping_container, anchor='nw')
        mapping_canvas.configure(xscrollcommand=mapping_scroll.set)
        
        mapping_canvas.pack(fill=tk.X, expand=True)
        mapping_scroll.pack(fill=tk.X)
        
        self.mapping_label = ttk.Label(mapping_outer, text="Paste data above and click 'Analyze Data'", 
                                       foreground='gray')
        self.mapping_label.pack()
        
        # Preview area
        preview_frame = ttk.LabelFrame(frame, text="Preview (first 5 rows)", padding="5")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        columns = ('number', 'name', 'ip', 'gateway', 'subnet', 'model')
        self.preview_tree = ttk.Treeview(preview_frame, columns=columns, show='headings', height=5)
        for col in columns:
            self.preview_tree.heading(col, text=col.title() if col != 'number' else '#')
            self.preview_tree.column(col, width=50 if col == 'number' else 120)
        self.preview_tree.pack(fill=tk.BOTH, expand=True)
        
        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X)
        ttk.Button(btn_frame, text="✓ Import Cameras", command=self.do_import).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.cancel).pack(side=tk.RIGHT)
        
        self.wait_window(self)
    
    def _show_context_menu(self, event):
        self.context_menu.tk_popup(event.x_root, event.y_root)
    
    def _paste(self):
        try:
            clipboard = self.clipboard_get()
            self.paste_text.insert(tk.INSERT, clipboard)
        except:
            pass
    
    def _paste_and_analyze(self):
        try:
            clipboard = self.clipboard_get()
            self.paste_text.delete('1.0', tk.END)
            self.paste_text.insert('1.0', clipboard)
            self.analyze_data()
        except:
            pass
    
    def paste_clipboard(self):
        try:
            clipboard = self.clipboard_get()
            self.paste_text.delete('1.0', tk.END)
            self.paste_text.insert('1.0', clipboard)
            self.analyze_data()
        except:
            messagebox.showwarning("Clipboard", "Nothing in clipboard or unable to paste")
    
    def load_file(self):
        filepath = filedialog.askopenfilename(
            title="Load Data File",
            filetypes=[("All Supported", "*.csv *.txt *.tsv"), ("CSV Files", "*.csv"), 
                      ("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        if filepath:
            try:
                with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                    self.paste_text.delete('1.0', tk.END)
                    self.paste_text.insert('1.0', f.read())
                self.analyze_data()
            except Exception as e:
                messagebox.showerror("Error", f"Could not load file: {e}")
    
    def parse_data(self):
        """Parse the pasted data into rows"""
        text = self.paste_text.get('1.0', tk.END).strip()
        if not text:
            return []
        
        rows = []
        for line in text.split('\n'):
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            
            # Try different delimiters
            if '\t' in line:
                parts = line.split('\t')
            elif ',' in line:
                parts = line.split(',')
            elif ';' in line:
                parts = line.split(';')
            else:
                parts = line.split()
            
            parts = [p.strip() for p in parts]
            if parts:
                rows.append(parts)
        
        return rows
    
    def analyze_data(self):
        """Analyze the pasted data and show column mappings"""
        self.rows = self.parse_data()
        
        if not self.rows:
            messagebox.showinfo("No Data", "No valid data found. Please paste some camera data.")
            return
        
        # Check if first row is a header and strip it
        header_map = SmartDataAnalyzer.detect_header_row(self.rows[0])
        if header_map:
            self.rows = self.rows[1:]
            if not self.rows:
                messagebox.showinfo("No Data", "Only found a header row, no camera data.")
                return
        
        # Analyze columns
        analysis = SmartDataAnalyzer.analyze_data(self.rows if not header_map else self.rows)
        # If header was detected, use that mapping instead
        if header_map:
            for col_idx, field in header_map.items():
                if field in ('rack', 'switch', 'port', 'unknown'):
                    field = 'unknown'
                sample = self.rows[0][col_idx] if col_idx < len(self.rows[0]) else ''
                analysis[col_idx] = {'type': field, 'confidence': 1.0, 'sample': sample}
        
        # Clear existing mapping widgets
        for widget in self.mapping_container.winfo_children():
            widget.destroy()
        self.mapping_label.pack_forget()
        
        # Create mapping dropdowns
        field_options = ['(skip)', 'number', 'name', 'ip', 'gateway', 'subnet', 'model', 'mac', 'serial', 'new_ip']
        self.column_vars = []
        
        num_cols = max(len(row) for row in self.rows)
        
        for col_idx in range(min(num_cols, 20)):  # Support up to 20 columns
            col_frame = ttk.Frame(self.mapping_container)
            col_frame.pack(side=tk.LEFT, padx=5, pady=5)
            
            # Sample value
            sample = self.rows[0][col_idx] if col_idx < len(self.rows[0]) else ''
            if len(sample) > 15:
                sample = sample[:15] + '...'
            ttk.Label(col_frame, text=f"Col {col_idx + 1}", font=('Helvetica', 9, 'bold')).pack()
            ttk.Label(col_frame, text=f'"{sample}"', font=('Courier', 8), foreground='gray').pack()
            
            # Dropdown
            var = tk.StringVar()
            guess = analysis.get(col_idx, {}).get('type', 'name')
            if guess == 'unknown':
                guess = '(skip)'
            var.set(guess)
            self.column_vars.append(var)
            
            combo = ttk.Combobox(col_frame, textvariable=var, values=field_options, width=10, state='readonly')
            combo.pack()
            combo.bind('<<ComboboxSelected>>', lambda e: self.update_preview())
        
        # Show confidence note
        ttk.Label(self.mapping_container, 
                 text=f"\n✓ Analyzed {len(self.rows)} rows", 
                 foreground='green').pack(side=tk.LEFT, padx=20)
        
        self.update_preview()
    
    def update_preview(self):
        """Update the preview based on current mappings"""
        self.preview_tree.delete(*self.preview_tree.get_children())
        
        if not self.rows or not hasattr(self, 'column_vars'):
            return
        
        # Build mapping
        self.column_mappings = {}
        for idx, var in enumerate(self.column_vars):
            field = var.get()
            if field != '(skip)':
                self.column_mappings[field] = idx
        
        # Show preview
        for row in self.rows[:5]:
            preview_row = []
            for field in ['number', 'name', 'ip', 'gateway', 'subnet', 'model']:
                if field in self.column_mappings:
                    col_idx = self.column_mappings[field]
                    value = row[col_idx] if col_idx < len(row) else ''
                else:
                    value = ''
                preview_row.append(value)
            self.preview_tree.insert('', 'end', values=preview_row)
    
    def do_import(self):
        """Import the data with current mappings"""
        if not self.rows:
            messagebox.showwarning("No Data", "No data to import")
            return
        
        if 'ip' not in self.column_mappings:
            messagebox.showwarning("Missing IP", "You must map at least an IP Address column")
            return
        
        cameras = []
        for i, row in enumerate(self.rows, start=1):
            cam = {'processed': False}
            
            for field, col_idx in self.column_mappings.items():
                if col_idx < len(row):
                    value = row[col_idx].strip()
                    if value:
                        cam[field] = value
            
            # Default number to sequential if not mapped
            if 'number' not in cam:
                cam['number'] = str(i)
            
            # Name logic:
            # - If CSV gave a name, use it
            # - If CSV gave a number and serial exists, name = {number}-axis-{serial}
            # - If CSV gave a number but no serial, name = {number}
            # - Otherwise name = IP
            if 'name' not in cam:
                if cam.get('serial'):
                    cam['name'] = f"{cam['number']}-axis-{cam['serial'].lower()}"
                elif cam.get('number'):
                    cam['name'] = cam['number']
                elif cam.get('ip'):
                    cam['name'] = cam['ip']
            
            if cam.get('ip'):
                cameras.append(cam)
        
        self.result = cameras
        self.destroy()
    
    def cancel(self):
        self.result = None
        self.destroy()


# ============================================================================
# DISCOVERY RESULTS DIALOG
# ============================================================================
class DiscoveryResultsDialog(tk.Toplevel):
    """Dialog to show discovered cameras and select which to add"""
    
    def __init__(self, parent, cameras, settings=None):
        super().__init__(parent)
        self.title("Discovered Cameras")
        self.result = None
        self.transient(parent)
        self.grab_set()
        _center_on_parent(self, parent, 950, 600)
        self.settings = settings
        
        self.cameras = cameras
        
        frame = ttk.Frame(self, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header = ttk.Frame(frame)
        header.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(header, text=f"Found {len(cameras)} Axis Camera(s)", 
                 font=('Helvetica', 14, 'bold')).pack(side=tk.LEFT)
        
        # Option to get more details with password
        detail_frame = ttk.Frame(header)
        detail_frame.pack(side=tk.RIGHT)
        ttk.Label(detail_frame, text="Password (optional):").pack(side=tk.LEFT, padx=(0, 5))
        self.password_var = tk.StringVar()
        pwd_entry = ttk.Entry(detail_frame, textvariable=self.password_var, show="*", width=15)
        pwd_entry.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(detail_frame, text="🔍 Get Details", command=self.fetch_details).pack(side=tk.LEFT)
        ToolTip(pwd_entry, "Enter password to fetch DHCP status, gateway, subnet from cameras")
        
        ttk.Label(frame, text="Click checkboxes to select • Double-click name to edit • Enter password above for more details", 
                 foreground='gray').pack(anchor=tk.W, pady=(0, 10))
        
        # Treeview container
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Treeview with checkboxes
        columns = ('select', 'ip', 'mac', 'model', 'serial', 'dhcp', 'gateway', 'name')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
        
        self.tree.heading('select', text='Add')
        self.tree.heading('ip', text='IP Address')
        self.tree.heading('mac', text='MAC Address')
        self.tree.heading('model', text='Model')
        self.tree.heading('serial', text='Serial')
        self.tree.heading('dhcp', text='DHCP')
        self.tree.heading('gateway', text='Gateway')
        self.tree.heading('name', text='Name')
        
        self.tree.column('select', width=40, anchor='center')
        self.tree.column('ip', width=110)
        self.tree.column('mac', width=130)
        self.tree.column('model', width=100)
        self.tree.column('serial', width=120)
        self.tree.column('dhcp', width=50, anchor='center')
        self.tree.column('gateway', width=100)
        self.tree.column('name', width=130)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Populate tree
        self.check_vars = {}
        for i, cam in enumerate(cameras):
            self.check_vars[i] = True  # All selected by default
            self.tree.insert('', 'end', iid=str(i), values=(
                '☑' if self.check_vars[i] else '☐',
                cam.get('ip', ''),
                cam.get('mac', ''),
                cam.get('model', ''),
                cam.get('serial', ''),
                cam.get('dhcp', '?'),
                cam.get('gateway', ''),
                cam.get('name', '')
            ))
        
        self.tree.bind('<Button-1>', self.on_click)
        self.tree.bind('<Double-1>', self.on_double_click)
        
        # Buttons at bottom
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(btn_frame, text="☑ Select All", command=self.select_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="☐ Select None", command=self.select_none).pack(side=tk.LEFT, padx=2)
        
        # Big add button
        add_btn = tk.Button(btn_frame, text="✓ Add Selected to Camera List", 
                           command=self.add_selected, bg='#4CAF50', fg='white',
                           font=('Helvetica', 10, 'bold'), padx=20)
        add_btn.pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(btn_frame, text="Cancel", command=self.cancel).pack(side=tk.RIGHT, padx=2)
        
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.wait_window(self)
    
    def fetch_details(self):
        """Fetch ALL details from cameras using password"""
        password = self.password_var.get()
        if not password:
            messagebox.showinfo("Password Required", "Enter a password to fetch camera details")
            return
        
        for i, cam in enumerate(self.cameras):
            ip = cam.get('ip', '')
            if not ip:
                continue
            
            try:
                auth = HTTPDigestAuth("root", password)
                
                # Get model via basicdeviceinfo (try with auth)
                try:
                    r = requests.post(f"http://{ip}/axis-cgi/basicdeviceinfo.cgi",
                        json={"apiVersion": "1.0", "method": "getAllProperties"},
                        auth=auth, timeout=3)
                    if r.status_code == 200:
                        data = r.json()
                        if 'data' in data and 'propertyList' in data['data']:
                            props = data['data']['propertyList']
                            cam['model'] = props.get('ProdFullName', props.get('ProdShortName', cam.get('model', '')))
                            if not cam.get('serial'):
                                cam['serial'] = props.get('SerialNumber', '')
                except:
                    pass
                
                # Fallback model from Brand.ProdFullName
                if not cam.get('model') or cam['model'] == '(Auth Required)':
                    try:
                        r = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                            params={"action": "list", "group": "Brand"},
                            auth=auth, timeout=3)
                        if r.status_code == 200:
                            for line in r.text.split('\n'):
                                if 'ProdFullName=' in line:
                                    cam['model'] = line.split('=', 1)[1].strip()
                    except:
                        pass
                
                # Get network info (gateway, subnet, DHCP)
                r = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                    params={"action": "list", "group": "Network"},
                    auth=auth, timeout=3)
                
                if r.status_code == 200:
                    for line in r.text.split('\n'):
                        line = line.strip()
                        if '=' not in line:
                            continue
                        if 'Network.eth0.DefaultRouter=' in line:
                            cam['gateway'] = line.split('=', 1)[1].strip()
                        elif 'Network.eth0.SubnetMask=' in line:
                            cam['subnet'] = line.split('=', 1)[1].strip()
                        elif 'Network.eth0.BootProto=' in line or 'Network.BootProto=' in line:
                            proto = line.split('=', 1)[1].strip().lower()
                            cam['dhcp'] = 'Yes' if proto == 'dhcp' else 'No'
                
                # Get serial if still missing
                if not cam.get('serial'):
                    try:
                        r2 = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                            params={"action": "list", "group": "Properties.System.SerialNumber"},
                            auth=auth, timeout=3)
                        if r2.status_code == 200:
                            for line in r2.text.split('\n'):
                                if 'SerialNumber=' in line:
                                    cam['serial'] = line.split('=', 1)[1].strip()
                    except:
                        pass
                
                # Derive MAC from serial
                if cam.get('serial') and len(cam['serial']) == 12 and not cam.get('mac'):
                    cam['mac'] = ':'.join(cam['serial'][j:j+2] for j in range(0, 12, 2))
                
                # Update tree row
                values = list(self.tree.item(str(i), 'values'))
                values[2] = cam.get('mac', '')       # MAC
                values[3] = cam.get('model', '')      # Model
                values[4] = cam.get('serial', '')     # Serial
                values[5] = cam.get('dhcp', '?')      # DHCP
                values[6] = cam.get('gateway', '')    # Gateway
                self.tree.item(str(i), values=values)
                
            except Exception as e:
                pass
        
        messagebox.showinfo("Done", "Finished fetching camera details")
    
    def on_click(self, event):
        """Toggle checkbox on click"""
        region = self.tree.identify_region(event.x, event.y)
        if region == 'cell':
            col = self.tree.identify_column(event.x)
            if col == '#1':  # Select column
                item = self.tree.identify_row(event.y)
                if item:
                    idx = int(item)
                    self.check_vars[idx] = not self.check_vars[idx]
                    values = list(self.tree.item(item, 'values'))
                    values[0] = '☑' if self.check_vars[idx] else '☐'
                    self.tree.item(item, values=values)
    
    def on_double_click(self, event):
        """Edit camera name on double-click"""
        col = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        if item and col == '#8':  # Name column
            idx = int(item)
            cam = self.cameras[idx]
            
            new_name = simpledialog.askstring("Edit Name", 
                f"Camera name for {cam.get('ip', '')}:",
                initialvalue=cam.get('name', ''),
                parent=self)
            if new_name:
                self.cameras[idx]['name'] = new_name
                values = list(self.tree.item(item, 'values'))
                values[7] = new_name
                self.tree.item(item, values=values)
    
    def select_all(self):
        for i in range(len(self.cameras)):
            self.check_vars[i] = True
            values = list(self.tree.item(str(i), 'values'))
            values[0] = '☑'
            self.tree.item(str(i), values=values)
    
    def select_none(self):
        for i in range(len(self.cameras)):
            self.check_vars[i] = False
            values = list(self.tree.item(str(i), 'values'))
            values[0] = '☐'
            self.tree.item(str(i), values=values)
    
    def add_selected(self):
        self.result = [self.cameras[i] for i in range(len(self.cameras)) if self.check_vars.get(i)]
        if not self.result:
            messagebox.showinfo("None Selected", "Please select at least one camera to add")
            return
        self.destroy()
    
    def cancel(self):
        self.result = None
        self.destroy()


# ============================================================================
# HELPER DIALOGS
# ============================================================================
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.show)
        self.widget.bind("<Leave>", self.hide)
        
    def show(self, event=None):
        x = self.widget.winfo_rootx() + 25
        y = self.widget.winfo_rooty() + 25
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        label = ttk.Label(self.tooltip, text=self.text, background="#ffffe0", 
                         relief="solid", borderwidth=1, padding=(5, 2))
        label.pack()
        
    def hide(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None


def _center_on_parent(dialog, parent, width, height):
    """Center a dialog on its parent window. Works across multiple monitors.
    The width/height args act as MINIMUM sizes — if the dialog's content
    requests more space (likely under DPI scaling), the dialog grows so
    nothing is cut off. Also clamps to screen bounds and sets minsize so
    the user can't drag it smaller than the content."""
    parent.update_idletasks()
    dialog.update_idletasks()
    req_w = dialog.winfo_reqwidth()
    req_h = dialog.winfo_reqheight()
    # Use the larger of (caller-requested, content-required) so widgets fit
    width = max(width, req_w) if width > 0 else req_w
    height = max(height, req_h) if height > 0 else req_h
    # Clamp to screen so dialog isn't bigger than the display
    sw = dialog.winfo_screenwidth()
    sh = dialog.winfo_screenheight()
    width = min(width, sw - 40)
    height = min(height, sh - 80)
    # Center on parent — use parent's actual position (works on any monitor)
    px = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
    py = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
    px = max(0, min(px, sw - width))
    py = max(0, min(py, sh - height))
    dialog.geometry(f"{width}x{height}+{px}+{py}")
    # Keep dialog at-or-above the content's minimum size during user resize
    dialog.minsize(width, height)


class PasswordDialog(tk.Toplevel):
    def __init__(self, parent, title="Enter Password", prompt="Password:"):
        super().__init__(parent)
        self.title(title)
        self.result = None
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)

        frame = ttk.Frame(self, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=prompt, font=('Helvetica', 12), wraplength=520, justify=tk.LEFT).pack(anchor=tk.W, fill=tk.X)

        pwd_frame = ttk.Frame(frame)
        pwd_frame.pack(fill=tk.X, pady=(8, 14))
        self.password_var = tk.StringVar()
        self.show_password = tk.BooleanVar(value=False)
        self.pwd_entry = ttk.Entry(pwd_frame, textvariable=self.password_var, show="*", width=30, font=('Helvetica', 12))
        self.pwd_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=2)
        ttk.Checkbutton(pwd_frame, text="Show", variable=self.show_password,
                       command=self.toggle_show).pack(side=tk.LEFT, padx=(8, 0))

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(btn_frame, text="OK", command=self.ok, width=10).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(btn_frame, text="Cancel", command=self.cancel, width=10).pack(side=tk.RIGHT)
        # Center AFTER widgets are packed so winfo_reqwidth includes content
        _center_on_parent(self, parent, 480, 220)
        
        self.pwd_entry.bind("<Return>", lambda e: self.ok())
        self.bind("<Escape>", lambda e: self.cancel())
        self.after(10, self._set_focus)
        self.wait_window(self)
    
    def _set_focus(self):
        self.pwd_entry.focus_force()
        
    def toggle_show(self):
        self.pwd_entry.config(show="" if self.show_password.get() else "*")
    
    def ok(self):
        self.result = self.password_var.get()
        self.destroy()
    
    def cancel(self):
        self.result = None
        self.destroy()


class WarningDialog(tk.Toplevel):
    """Warning dialog with 'Don't show again' option"""
    def __init__(self, parent, title, message, setting_key=None, settings_manager=None):
        super().__init__(parent)
        self.title(title)
        self.result = False
        self.transient(parent)
        self.grab_set()
        _center_on_parent(self, parent, 550, 430)
        self.setting_key = setting_key
        self.settings_manager = settings_manager
        
        frame = ttk.Frame(self, padding="25")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Warning icon and message
        msg_frame = ttk.Frame(frame)
        msg_frame.pack(fill=tk.X, pady=(0, 20))
        ttk.Label(msg_frame, text="⚠️", font=('Helvetica', 36)).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Label(msg_frame, text=message, font=('Helvetica', 11), wraplength=420, justify=tk.LEFT).pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Don't show again checkbox
        self.dont_show_var = tk.BooleanVar(value=False)
        if setting_key:
            ttk.Checkbutton(frame, text="Don't show this message again", variable=self.dont_show_var).pack(pady=(0, 15))
        
        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.pack()
        ttk.Button(btn_frame, text="Continue", command=self.on_continue).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.on_cancel).pack(side=tk.LEFT, padx=5)
        
        self.wait_window(self)
    
    def on_continue(self):
        if self.setting_key and self.dont_show_var.get() and self.settings_manager:
            self.settings_manager.set('warnings', self.setting_key, 'false')
        self.result = True
        self.destroy()
    
    def on_cancel(self):
        self.result = False
        self.destroy()


class ContinueDialog(tk.Toplevel):
    """Dialog shown between camera programming with preview"""
    def __init__(self, parent, message, next_camera=None, next_model=None, image_data=None):
        super().__init__(parent)
        self.title("Camera Complete")
        self.result = False
        self.transient(parent)
        self.grab_set()
        self.preview_image = None

        frame = ttk.Frame(self, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)

        # Top section with message and image side by side
        top_frame = ttk.Frame(frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))

        info_frame = ttk.Frame(top_frame)
        info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Label(info_frame, text="✓ " + message, font=('Helvetica', 12, 'bold'), foreground='green').pack(anchor=tk.W)

        if image_data and HAS_PIL:
            try:
                img = Image.open(BytesIO(image_data))
                img.thumbnail((240, 180), Image.Resampling.LANCZOS)
                self.preview_image = ImageTk.PhotoImage(img)
                preview_frame = ttk.LabelFrame(top_frame, text="Preview", padding="3")
                preview_frame.pack(side=tk.RIGHT, padx=(10, 0))
                ttk.Label(preview_frame, image=self.preview_image).pack()
            except: pass

        if next_camera:
            ttk.Separator(frame, orient='horizontal').pack(fill=tk.X, pady=5)
            ttk.Label(frame, text="NEXT CAMERA:", font=('Helvetica', 11, 'bold')).pack()
            ttk.Label(frame, text=next_camera, font=('Courier', 22, 'bold')).pack(pady=(3, 3))
            if next_model:
                ttk.Label(frame, text=next_model, font=('Courier', 16, 'bold')).pack(pady=(0, 5))

        ttk.Label(frame, text="Connect next camera and press Continue", font=('Helvetica', 10)).pack(pady=(8, 10))

        btn_frame = ttk.Frame(frame)
        btn_frame.pack()
        self.continue_btn = ttk.Button(btn_frame, text="Continue (Space/Enter)", command=self.on_continue)
        self.continue_btn.pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Stop", command=self.on_cancel).pack(side=tk.LEFT, padx=10)

        self.continue_btn.focus_set()
        # Auto-size to content then center on parent
        _center_on_parent(self, parent, 0, 0)
        self.bind("<space>", lambda e: self.on_continue())
        self.bind("<Return>", lambda e: self.on_continue())
        self.bind("<Escape>", lambda e: self.on_cancel())
        self.wait_window(self)
    
    def on_continue(self):
        self.result = True
        self.destroy()
    
    def on_cancel(self):
        self.result = False
        self.destroy()


# ============================================================================
# PROGRAM OPTIONS DIALOG
# ============================================================================
class ProgramOptionsDialog(tk.Toplevel):
    """Dialog to set factory IP and programming options before starting"""

    @staticmethod
    def _get_network_interfaces():
        """Get list of network interfaces with IP addresses for the dropdown."""
        import subprocess
        interfaces = []
        try:
            cmd = ("Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | "
                   "Where-Object { $_.IPAddress -ne '127.0.0.1' -and "
                   "-not $_.IPAddress.StartsWith('169.254.') } | "
                   "Select-Object IPAddress,InterfaceIndex,InterfaceAlias | "
                   "ConvertTo-Json -Compress")
            result = subprocess.run(
                ['powershell', '-NoProfile', '-Command', cmd],
                capture_output=True, text=True, timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode == 0 and result.stdout.strip():
                entries = json.loads(result.stdout)
                if isinstance(entries, dict):
                    entries = [entries]
                for e in entries:
                    ip = e.get('IPAddress', '')
                    idx = e.get('InterfaceIndex', '')
                    alias = e.get('InterfaceAlias', '')
                    if ip:
                        interfaces.append({
                            'ip': ip, 'index': idx,
                            'label': f"{alias} ({ip})",
                        })
        except:
            pass
        return interfaces

    def __init__(self, parent, factory_ip='192.168.0.90', additional_users_count=0):
        super().__init__(parent)
        self.title("Programming Options")
        self.result = None
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)

        frame = ttk.Frame(self, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)

        # Network interface selector
        self._interfaces = self._get_network_interfaces()
        if self._interfaces:
            ttk.Label(frame, text="Programming Interface:", font=('Helvetica', 10, 'bold')).grid(
                row=0, column=0, columnspan=2, sticky='w', pady=(0, 5))
            iface_labels = [i['label'] for i in self._interfaces]
            iface_labels.insert(0, "Auto-detect (default)")
            self.iface_var = tk.StringVar(value=iface_labels[0])
            iface_combo = ttk.Combobox(frame, textvariable=self.iface_var,
                                       values=iface_labels, state='readonly', width=40)
            iface_combo.grid(row=1, column=0, columnspan=2, sticky='w', padx=(10, 0), pady=(0, 5))
            ttk.Label(frame, text="Select which NIC is connected to the cameras",
                     foreground='gray', font=('Helvetica', 8)).grid(
                row=2, column=0, columnspan=2, sticky='w', padx=(10, 0))
            sep_row = 3
        else:
            self.iface_var = tk.StringVar(value='')
            sep_row = 0

        # Separator before discovery
        ttk.Separator(frame, orient='horizontal').grid(
            row=sep_row, column=0, columnspan=2, sticky='ew', pady=8)

        # Discovery method section
        base_row = sep_row + 1
        ttk.Label(frame, text="Camera Discovery Method:", font=('Helvetica', 10, 'bold')).grid(
            row=base_row, column=0, columnspan=2, sticky='w', pady=(0, 5))

        self.discovery_var = tk.StringVar(value='both')
        r = base_row + 1  # running row counter

        # DHCP/mDNS only (for firmware 12.0+ link-local cameras)
        ttk.Radiobutton(frame, text="DHCP/mDNS only (firmware 12.0+ link-local)",
            variable=self.discovery_var, value='mdns',
            command=self._update_ip_state).grid(row=r, column=0, columnspan=2, sticky='w', padx=(10, 0))
        r += 1

        # Factory IP only (legacy)
        ttk.Radiobutton(frame, text="Factory IP only",
            variable=self.discovery_var, value='factory',
            command=self._update_ip_state).grid(row=r, column=0, columnspan=2, sticky='w', padx=(10, 0))
        r += 1

        # Both (recommended)
        ttk.Radiobutton(frame, text="Both DHCP/mDNS + Factory IP (recommended)",
            variable=self.discovery_var, value='both',
            command=self._update_ip_state).grid(row=r, column=0, columnspan=2, sticky='w', padx=(10, 0))
        r += 1

        # Factory IP entry
        self.ip_label = ttk.Label(frame, text="Factory Default IP:", font=('Helvetica', 10))
        self.ip_label.grid(row=r, column=0, sticky='w', pady=(10, 5))
        self.ip_entry = ttk.Entry(frame, width=20)
        self.ip_entry.insert(0, factory_ip)
        self.ip_entry.grid(row=r, column=1, sticky='w', pady=(10, 5), padx=(10, 0))
        r += 1

        # Separator
        ttk.Separator(frame, orient='horizontal').grid(
            row=r, column=0, columnspan=2, sticky='ew', pady=10)
        r += 1

        # Hostname checkbox
        self.hostname_var = tk.BooleanVar(value=False)
        self.hostname_check = ttk.Checkbutton(frame,
            text="Change network hostname",
            variable=self.hostname_var)
        self.hostname_check.grid(row=r, column=0, columnspan=2, sticky='w', pady=2)
        r += 1
        ttk.Label(frame, text="Sets hostname to <number>-<brand>-<serial> (lowercase)",
                 foreground='gray', font=('Helvetica', 8)).grid(
            row=r, column=0, columnspan=2, sticky='w', padx=(22, 0))
        r += 1

        # Additional users checkbox
        self.additional_users_var = tk.BooleanVar(value=False)
        user_label = f"Create additional users ({additional_users_count} defined)"
        if additional_users_count == 0:
            user_label = "Create additional users (none defined)"
        self.additional_users_check = ttk.Checkbutton(frame,
            text=user_label,
            variable=self.additional_users_var)
        self.additional_users_check.grid(row=r, column=0, columnspan=2, sticky='w', pady=2)
        if additional_users_count == 0:
            self.additional_users_check.configure(state='disabled')
        r += 1
        ttk.Label(frame, text="Add users in the Passwords tab before programming",
                 foreground='gray', font=('Helvetica', 8)).grid(
            row=r, column=0, columnspan=2, sticky='w', padx=(22, 0))
        r += 1

        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=r, column=0, columnspan=2, pady=(15, 0))
        ttk.Button(btn_frame, text="Start Programming", command=self.ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.cancel).pack(side=tk.LEFT, padx=5)
        
        self._update_ip_state()
        self.ip_entry.focus_set()
        self.bind("<Return>", lambda e: self.ok())
        self.bind("<Escape>", lambda e: self.cancel())
        # Center on parent after content is laid out (auto-size: width=0, height=0)
        _center_on_parent(self, parent, 0, 0)
        self.wait_window(self)
    
    def _update_ip_state(self):
        """Enable/disable factory IP entry based on discovery method"""
        if self.discovery_var.get() == 'mdns':
            self.ip_entry.configure(state='disabled')
            self.ip_label.configure(foreground='gray')
        else:
            self.ip_entry.configure(state='normal')
            self.ip_label.configure(foreground='black')
    
    def ok(self):
        discovery = self.discovery_var.get()
        ip = self.ip_entry.get().strip()
        
        # Factory IP required unless mDNS-only mode
        if discovery != 'mdns' and not ip:
            messagebox.showwarning("Required", "Factory IP is required for this mode", parent=self)
            return
        
        # Resolve selected interface
        selected_iface = None
        iface_selection = self.iface_var.get()
        if iface_selection and iface_selection != 'Auto-detect (default)':
            for iface in self._interfaces:
                if iface['label'] == iface_selection:
                    selected_iface = iface
                    break

        self.result = {
            'factory_ip': ip if discovery != 'mdns' else None,
            'discovery_mode': discovery,  # 'mdns', 'factory', or 'both'
            'set_hostname': self.hostname_var.get(),
            'add_additional_users': self.additional_users_var.get(),
            'interface': selected_iface,  # {'ip': ..., 'index': ..., 'label': ...} or None
        }
        self.destroy()
    
    def cancel(self):
        self.destroy()


# ============================================================================
# PROGRAM WIZARD DIALOG (step-by-step setup)
# ============================================================================
class ProgramWizardDialog(tk.Toplevel):
    """Step-by-step wizard for programming new cameras.
    Replaces the all-in-one ProgramOptionsDialog with one decision per screen.
    Returns a result dict with password + all options, or None on cancel."""

    def __init__(self, parent, brand_name, factory_ip='192.168.0.90',
                 camera_count=0, additional_users_count=0):
        super().__init__(parent)
        self.title(f"Program {brand_name} Cameras — Setup Wizard")
        self.result = None
        self.brand_name = brand_name
        self.camera_count = camera_count
        self.additional_users_count = additional_users_count
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)

        # Variables (shared across steps)
        self.password_var = tk.StringVar()
        self.password_confirm_var = tk.StringVar()
        self.discovery_var = tk.StringVar(value='both')
        self.factory_ip_var = tk.StringVar(value=factory_ip)
        self.hostname_var = tk.BooleanVar(value=False)
        self.additional_users_var = tk.BooleanVar(value=False)
        # v4.3 #10: by default the wizard creates an ONVIF root user (required
        # by set_network's ONVIF SOAP) AND deletes it after programming so the
        # camera ends with VAPIX root only. Operator can opt to keep it (e.g.
        # customer requires ONVIF account for VMS handoff).
        self.keep_onvif_user_var = tk.BooleanVar(value=False)
        # v4.3 #10b — optional custom ONVIF user/password. If both are filled
        # AND keep_onvif_user is checked, the wizard deletes the transient
        # ONVIF root and creates this named user instead (rename pattern).
        # If empty, the kept ONVIF user remains as 'root'/(VAPIX password).
        self.onvif_username_var = tk.StringVar(value='')
        self.onvif_password_var = tk.StringVar(value='')
        # v4.3 #13 — Building Reports sticker numbering. If enabled, each
        # successfully-programmed camera gets the next sequential sticker
        # number assigned (8-digit per Brian's rolls), starting from the
        # operator-supplied seed. Saved into the OUTPUT_CSV's
        # BuildingReportsLabel column. Operator peels stickers off the
        # spool in order as cameras complete.
        self.add_br_stickers_var = tk.BooleanVar(value=False)
        self.br_first_label_var = tk.StringVar(value='')
        # v4.3 — "Factory default before programming" workflow. For the case
        # where Brian gets a USED camera (already configured for a previous
        # site, has a known root password) and wants to wipe + reprogram for
        # the new install. If checked + existing pwd is filled, the wizard
        # fires factory_reset(camera_ip, existing_pwd) right after MAC-
        # freshness gate, waits for the camera to come back into factory
        # state, THEN proceeds with create_initial_user + set_network.
        # Also addresses the eager-grab problem (operator couldn't physically
        # hold the factory-reset button before wizard started writing).
        self.factory_first_var = tk.BooleanVar(value=False)
        self.existing_root_pwd_var = tk.StringVar(value='')
        self.iface_var = tk.StringVar(value='Auto-detect (default)')
        self._interfaces = ProgramOptionsDialog._get_network_interfaces()

        # Layout: header bar + step area + nav bar
        outer = ttk.Frame(self, padding=(0, 0, 0, 0))
        outer.pack(fill=tk.BOTH, expand=True)

        # Header
        header = tk.Frame(outer, bg='#4CAF50', padx=20, pady=14)
        header.pack(fill=tk.X)
        self.header_title = tk.Label(header, text='', bg='#4CAF50', fg='white',
                                     font=('Helvetica', 16, 'bold'), anchor='w')
        self.header_title.pack(fill=tk.X)
        self.header_subtitle = tk.Label(header, text='', bg='#4CAF50', fg='white',
                                        font=('Helvetica', 10), anchor='w')
        self.header_subtitle.pack(fill=tk.X)

        # Step content area
        self.body = ttk.Frame(outer, padding=20)
        self.body.pack(fill=tk.BOTH, expand=True)

        # Build all step frames (only shown one at a time)
        self.steps = []
        self._build_step_welcome()
        self._build_step_password()
        self._build_step_discovery()
        self._build_step_extras()
        self._build_step_review()

        # Nav bar
        nav = ttk.Frame(outer, padding=(20, 10, 20, 15))
        nav.pack(fill=tk.X)
        self.back_btn = ttk.Button(nav, text='← Back', command=self.go_back, width=12)
        self.back_btn.pack(side=tk.LEFT)
        ttk.Button(nav, text='Cancel', command=self.cancel, width=12).pack(side=tk.LEFT, padx=(8, 0))
        self.next_btn = tk.Button(nav, text='Next →', command=self.go_next,
                                  bg='#4CAF50', fg='white', font=('Helvetica', 10, 'bold'),
                                  padx=14, pady=6, relief=tk.RAISED, cursor='hand2', width=18)
        self.next_btn.pack(side=tk.RIGHT)
        self.step_label = ttk.Label(nav, text='', font=('Helvetica', 9), foreground='gray')
        self.step_label.pack(side=tk.RIGHT, padx=(0, 12))

        self.current_step = 0
        self.show_step(0)

        self.bind("<Escape>", lambda e: self.cancel())
        # Min size 720x600 — bumped from 640x460 because Step 2's content
        # (network dropdown + 3 radio explanations + factory IP field) was
        # taller than 460 and clipped the Next button. Brian flagged 2026-04-30
        # via screenshot. _center_on_parent treats these as MINIMUMS so it
        # grows further if content needs more.
        _center_on_parent(self, parent, 720, 600)
        self.wait_window(self)

    # ------------------------------------------------------------------
    # Step builders
    # ------------------------------------------------------------------
    def _new_step(self, title, subtitle):
        f = ttk.Frame(self.body)
        self.steps.append({'frame': f, 'title': title, 'subtitle': subtitle})
        return f

    def _build_step_welcome(self):
        f = self._new_step("Welcome",
                           f"You're about to program {self.camera_count} {self.brand_name} camera(s).")
        msg = (
            "This wizard walks you through programming brand-new cameras one at a time.\n\n"
            "What to expect:\n"
            "  •  You'll set a password and a few options on the next screens.\n"
            "  •  Then you'll be told when to plug in each camera.\n"
            "  •  Programming each camera takes about 1–2 minutes.\n"
            "  •  A live checklist will show exactly what's happening.\n\n"
            "Before you continue, make sure:\n"
            "  ✓  You have your camera list ready in the Camera List tab.\n"
            "  ✓  Your laptop is plugged into the camera switch.\n"
            "  ✓  PoE is powering the cameras.\n\n"
            "When you're ready, click  Next →"
        )
        ttk.Label(f, text=msg, justify=tk.LEFT, font=('Helvetica', 10)).pack(
            anchor='w', pady=(10, 0))

    def _build_step_password(self):
        f = self._new_step("Step 1 of 4 — Set Password",
                           "This password will be set on every camera you program.")
        ttk.Label(f, text="Password:", font=('Helvetica', 11, 'bold')).pack(
            anchor='w', pady=(20, 4))
        e1 = ttk.Entry(f, textvariable=self.password_var, show='•', width=40, font=('Helvetica', 11))
        e1.pack(anchor='w', ipady=4)
        self._pwd_entry = e1

        ttk.Label(f, text="Confirm password:", font=('Helvetica', 11, 'bold')).pack(
            anchor='w', pady=(15, 4))
        e2 = ttk.Entry(f, textvariable=self.password_confirm_var, show='•', width=40, font=('Helvetica', 11))
        e2.pack(anchor='w', ipady=4)

        ttk.Label(f,
                  text="\nTip: Use a strong password — this is the camera's admin login.",
                  foreground='gray', font=('Helvetica', 9)).pack(anchor='w', pady=(10, 0))

    def _build_step_discovery(self):
        f = self._new_step("Step 2 of 4 — How to find cameras",
                           "Pick how the toolkit should discover the camera when you plug it in.")

        if self._interfaces:
            ttk.Label(f, text="Network interface (which port the cameras are on):",
                      font=('Helvetica', 10, 'bold')).pack(anchor='w', pady=(15, 4))
            iface_labels = ['Auto-detect (default)'] + [i['label'] for i in self._interfaces]
            ttk.Combobox(f, textvariable=self.iface_var, values=iface_labels,
                         state='readonly', width=50).pack(anchor='w')

        ttk.Label(f, text="Discovery method:", font=('Helvetica', 10, 'bold')).pack(
            anchor='w', pady=(18, 4))

        rb_frame = ttk.Frame(f)
        rb_frame.pack(fill=tk.X)

        ttk.Radiobutton(rb_frame, text="Both — DHCP/mDNS + Factory IP  (recommended)",
                        variable=self.discovery_var, value='both',
                        command=self._update_factory_state).pack(anchor='w', pady=2)
        ttk.Label(rb_frame, text="    Works for both old and new (firmware 12+) cameras.",
                  foreground='gray', font=('Helvetica', 9)).pack(anchor='w')

        ttk.Radiobutton(rb_frame, text="DHCP/mDNS only",
                        variable=self.discovery_var, value='mdns',
                        command=self._update_factory_state).pack(anchor='w', pady=(10, 2))
        ttk.Label(rb_frame, text="    For firmware 12.0+ cameras with link-local 169.254.x.x addresses.",
                  foreground='gray', font=('Helvetica', 9)).pack(anchor='w')

        ttk.Radiobutton(rb_frame, text="Factory IP only",
                        variable=self.discovery_var, value='factory',
                        command=self._update_factory_state).pack(anchor='w', pady=(10, 2))
        ttk.Label(rb_frame, text="    Older cameras with a fixed factory IP address.",
                  foreground='gray', font=('Helvetica', 9)).pack(anchor='w')

        ip_row = ttk.Frame(f)
        ip_row.pack(fill=tk.X, pady=(18, 0))
        self.ip_label = ttk.Label(ip_row, text="Factory default IP:", font=('Helvetica', 10, 'bold'))
        self.ip_label.pack(side=tk.LEFT)
        self.ip_entry = ttk.Entry(ip_row, textvariable=self.factory_ip_var, width=20, font=('Helvetica', 10))
        self.ip_entry.pack(side=tk.LEFT, padx=(10, 0), ipady=3)

    def _build_step_extras(self):
        f = self._new_step("Step 3 of 4 — Extras",
                           "Optional things to do during programming.")

        ttk.Checkbutton(f, text="Set network hostname automatically",
                        variable=self.hostname_var).pack(anchor='w', pady=(20, 2))
        ttk.Label(f, text="    Sets hostname to <number>-<brand>-<serial> on each camera.",
                  foreground='gray', font=('Helvetica', 9)).pack(anchor='w')

        ttk.Checkbutton(f, text="Keep ONVIF user after programming (don't delete)",
                        variable=self.keep_onvif_user_var).pack(anchor='w', pady=(15, 2))
        ttk.Label(f, text="    Default: ONVIF user is deleted after set_network completes,\n"
                          "    leaving VAPIX root only. Check this if your customer/VMS\n"
                          "    requires the ONVIF user to remain.",
                  foreground='gray', font=('Helvetica', 9)).pack(anchor='w')

        # ONVIF custom-creds row — only meaningful if Keep is checked
        onvif_creds = ttk.Frame(f)
        onvif_creds.pack(fill=tk.X, pady=(4, 2))
        ttk.Label(onvif_creds, text="    Custom ONVIF user (optional):",
                  font=('Helvetica', 9)).pack(side=tk.LEFT)
        ttk.Entry(onvif_creds, textvariable=self.onvif_username_var, width=14).pack(side=tk.LEFT, padx=(4, 8))
        ttk.Label(onvif_creds, text="Password:", font=('Helvetica', 9)).pack(side=tk.LEFT)
        ttk.Entry(onvif_creds, textvariable=self.onvif_password_var, width=14).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Label(f, text="    Both filled → wizard deletes ONVIF root and creates this user "
                          "after programming.\n"
                          "    Empty → if Keep is checked, ONVIF root remains as-is. Ignored if "
                          "Keep is unchecked.",
                  foreground='gray', font=('Helvetica', 9)).pack(anchor='w')

        # Building Reports stickers (v4.3 #13)
        ttk.Checkbutton(f, text="Add Building Reports stickers (sequential 8-digit labels)",
                        variable=self.add_br_stickers_var).pack(anchor='w', pady=(15, 2))
        br_row = ttk.Frame(f)
        br_row.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(br_row, text="    First sticker number on roll:",
                  font=('Helvetica', 9)).pack(side=tk.LEFT)
        ttk.Entry(br_row, textvariable=self.br_first_label_var, width=14).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Label(f, text="    Each programmed camera gets the next sequential number assigned.\n"
                          "    Saved in the BuildingReportsLabel column of the success CSV.\n"
                          "    Peel stickers off the spool in order as cameras complete.",
                  foreground='gray', font=('Helvetica', 9)).pack(anchor='w')

        # Factory-default-before-program (v4.3 — for re-using cameras from prior installs)
        ttk.Checkbutton(f, text="Factory default before programming (for used cameras)",
                        variable=self.factory_first_var).pack(anchor='w', pady=(15, 2))
        fact_row = ttk.Frame(f)
        fact_row.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(fact_row, text="    Existing root password:",
                  font=('Helvetica', 9)).pack(side=tk.LEFT)
        ttk.Entry(fact_row, textvariable=self.existing_root_pwd_var, show='*', width=20).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Label(f, text="    Wipes the camera (using the existing password) before applying\n"
                          "    new programming. For cameras you're reprovisioning from a\n"
                          "    previous site. Also gives you time to physically hold the\n"
                          "    factory-reset button if needed — programming pauses for the wipe.",
                  foreground='gray', font=('Helvetica', 9)).pack(anchor='w')

        if self.additional_users_count > 0:
            label = f"Create additional user accounts ({self.additional_users_count} defined)"
            state = 'normal'
        else:
            label = "Create additional user accounts (none defined)"
            state = 'disabled'
        cb = ttk.Checkbutton(f, text=label, variable=self.additional_users_var)
        cb.pack(anchor='w', pady=(15, 2))
        cb.configure(state=state)
        ttk.Label(f,
                  text="    Add users in the Passwords tab before programming if you need this.",
                  foreground='gray', font=('Helvetica', 9)).pack(anchor='w')

        ttk.Label(f, text="\nNo extras needed? Just click  Next →",
                  foreground='gray', font=('Helvetica', 10)).pack(anchor='w', pady=(20, 0))

    def _build_step_review(self):
        f = self._new_step("Step 4 of 4 — Review",
                           "Last check before you start programming.")
        self._review_text = tk.Text(f, height=14, width=70, font=('Consolas', 10),
                                    relief=tk.SUNKEN, borderwidth=1, wrap=tk.WORD,
                                    bg='#f8f8f8')
        self._review_text.pack(fill=tk.BOTH, expand=True, pady=(15, 0))
        self._review_text.configure(state='disabled')

    def _populate_review(self):
        iface = self.iface_var.get()
        mode = self.discovery_var.get()
        mode_name = {'both': 'Both DHCP/mDNS + Factory IP',
                     'mdns': 'DHCP/mDNS only',
                     'factory': 'Factory IP only'}.get(mode, mode)
        lines = [
            f"Brand                : {self.brand_name}",
            f"Cameras to program   : {self.camera_count}",
            "",
            f"Password             : {'•' * len(self.password_var.get())}",
            f"Network interface    : {iface}",
            f"Discovery method     : {mode_name}",
        ]
        if mode != 'mdns':
            lines.append(f"Factory IP           : {self.factory_ip_var.get()}")
        lines.append("")
        lines.append(f"Set hostname         : {'YES' if self.hostname_var.get() else 'no'}")
        lines.append(f"Add extra users      : {'YES' if self.additional_users_var.get() else 'no'}")
        lines.append("")
        lines.append("Click  Start Programming  to begin.")
        lines.append("You'll be prompted to plug in each camera one at a time.")

        self._review_text.configure(state='normal')
        self._review_text.delete('1.0', tk.END)
        self._review_text.insert('1.0', '\n'.join(lines))
        self._review_text.configure(state='disabled')

    # ------------------------------------------------------------------
    # Navigation
    # ------------------------------------------------------------------
    def show_step(self, idx):
        for s in self.steps:
            s['frame'].pack_forget()
        self.current_step = idx
        step = self.steps[idx]
        step['frame'].pack(fill=tk.BOTH, expand=True)
        self.header_title.config(text=step['title'])
        self.header_subtitle.config(text=step['subtitle'])
        self.step_label.config(text=f"Step {idx + 1} of {len(self.steps)}")
        self.back_btn.config(state='normal' if idx > 0 else 'disabled')
        if idx == len(self.steps) - 1:
            self.next_btn.config(text='✓ Start Programming')
            self._populate_review()
        else:
            self.next_btn.config(text='Next →')
        # Focus first input on the password step
        if idx == 1:
            self.after(50, self._pwd_entry.focus_set)

    def go_back(self):
        if self.current_step > 0:
            self.show_step(self.current_step - 1)

    def go_next(self):
        # Validate current step
        idx = self.current_step
        if idx == 1:  # password
            pwd = self.password_var.get()
            if not pwd:
                messagebox.showwarning("Required", "Password is required.", parent=self)
                return
            if pwd != self.password_confirm_var.get():
                messagebox.showerror("Mismatch", "Passwords don't match!", parent=self)
                return
        elif idx == 2:  # discovery
            mode = self.discovery_var.get()
            if mode != 'mdns' and not self.factory_ip_var.get().strip():
                messagebox.showwarning("Required",
                                       "Factory IP is required for this mode.", parent=self)
                return

        if idx == len(self.steps) - 1:
            self.finish()
        else:
            self.show_step(idx + 1)

    def _update_factory_state(self):
        if self.discovery_var.get() == 'mdns':
            self.ip_entry.configure(state='disabled')
            self.ip_label.configure(foreground='gray')
        else:
            self.ip_entry.configure(state='normal')
            self.ip_label.configure(foreground='black')

    def finish(self):
        # Resolve interface selection
        selected_iface = None
        sel = self.iface_var.get()
        if sel and sel != 'Auto-detect (default)':
            for iface in self._interfaces:
                if iface['label'] == sel:
                    selected_iface = iface
                    break
        mode = self.discovery_var.get()
        self.result = {
            'password': self.password_var.get(),
            'factory_ip': self.factory_ip_var.get().strip() if mode != 'mdns' else None,
            'discovery_mode': mode,
            'set_hostname': self.hostname_var.get(),
            'add_additional_users': self.additional_users_var.get(),
            'keep_onvif_user': self.keep_onvif_user_var.get(),
            'onvif_username': self.onvif_username_var.get().strip(),
            'onvif_password': self.onvif_password_var.get().strip(),
            'add_br_stickers': self.add_br_stickers_var.get(),
            'br_first_label': self.br_first_label_var.get().strip(),
            'factory_first': self.factory_first_var.get(),
            'existing_root_pwd': self.existing_root_pwd_var.get(),
            'interface': selected_iface,
        }
        self.destroy()

    def cancel(self):
        self.result = None
        self.destroy()


# ============================================================================
# BRAND SELECTION DIALOG
# ============================================================================
class BrandSelectionDialog(tk.Toplevel):
    """Modal dialog to select camera brand. Shown on first run or when switching."""
    def __init__(self, parent, current_brand='axis'):
        super().__init__(parent)
        self.title("Select Camera Brand")
        self.result = None
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)

        _center_on_parent(self, parent, 500, 350)

        frame = ttk.Frame(self, padding="25")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Select Camera Brand",
                 font=('Helvetica', 16, 'bold')).pack(pady=(0, 5))
        ttk.Label(frame, text="All operations will use this brand's protocol.\nSwitch brands anytime from the brand bar.",
                 font=('Helvetica', 10), foreground='gray', justify=tk.CENTER).pack(pady=(0, 15))

        self.brand_var = tk.StringVar(value=current_brand)

        brands = [
            ('axis', 'Axis', 'VAPIX/ONVIF  •  Factory IP: 192.168.0.90  •  User: root'),
            ('bosch', 'Bosch', 'RCP-over-HTTP  •  Factory IP: 192.168.0.1  •  User: service'),
            ('hanwha', 'Hanwha / Wisenet', 'STW-CGI/Sunapi  •  Factory IP: 192.168.1.100  •  User: admin'),
        ]

        for key, name, desc in brands:
            btn_frame = ttk.Frame(frame)
            btn_frame.pack(fill=tk.X, pady=4)
            rb = ttk.Radiobutton(btn_frame, text=name, variable=self.brand_var, value=key,
                                style='Toolbutton')
            rb.pack(side=tk.LEFT, padx=(10, 0))
            ttk.Label(btn_frame, text=desc, foreground='gray',
                     font=('Helvetica', 9)).pack(side=tk.LEFT, padx=(15, 0))

        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=(20, 0))
        ok_btn = tk.Button(btn_frame, text="Select Brand", command=self.ok,
                          bg='#4CAF50', fg='white', font=('Helvetica', 11, 'bold'),
                          padx=20, pady=5)
        ok_btn.pack(side=tk.LEFT, padx=5)

        self.bind("<Return>", lambda e: self.ok())
        self.protocol("WM_DELETE_WINDOW", self.ok)
        self.wait_window(self)

    def ok(self):
        self.result = self.brand_var.get()
        self.destroy()


# ============================================================================
# CAMERA EDITOR DIALOG
# ============================================================================
class CameraEditorDialog(tk.Toplevel):
    """Dialog for adding/editing a single camera"""
    def __init__(self, parent, camera=None, settings=None):
        super().__init__(parent)
        self.title("Add Camera" if camera is None else "Edit Camera")
        self.result = None
        self.camera = camera
        self.settings = settings
        self.transient(parent)
        self.grab_set()
        _center_on_parent(self, parent, 550, 500)

        frame = ttk.Frame(self, padding="25")
        frame.pack(fill=tk.BOTH, expand=True)

        self.entries = {}
        row = 0

        # IP Address — always editable (user needs to set target IP for programming)
        ttk.Label(frame, text="IP Address:").grid(row=row, column=0, sticky='w', pady=3)
        ip_val = camera.get('ip', '') if camera else ''
        ip_entry = ttk.Entry(frame, width=30)
        ip_entry.grid(row=row, column=1, sticky='ew', pady=3, padx=(10, 0))
        if ip_val:
            ip_entry.insert(0, ip_val)
        self.entries['ip'] = ip_entry
        ip_entry.bind('<FocusOut>', lambda e: self._auto_fill_gateway())
        row += 1

        # Camera Name
        ttk.Label(frame, text="Camera Name:").grid(row=row, column=0, sticky='w', pady=3)
        name_entry = ttk.Entry(frame, width=30)
        name_entry.grid(row=row, column=1, sticky='ew', pady=3, padx=(10, 0))
        if camera and camera.get('name'):
            name_entry.insert(0, camera['name'])
        self.entries['name'] = name_entry
        row += 1

        # Gateway — auto-filled from IP (.1)
        ttk.Label(frame, text="Gateway:").grid(row=row, column=0, sticky='w', pady=3)
        gw_entry = ttk.Entry(frame, width=30)
        gw_entry.grid(row=row, column=1, sticky='ew', pady=3, padx=(10, 0))
        if camera and camera.get('gateway'):
            gw_entry.insert(0, camera['gateway'])
        elif ip_val and not ip_val.startswith('169.254.'):
            parts = ip_val.split('.')
            if len(parts) == 4:
                gw_entry.insert(0, f"{parts[0]}.{parts[1]}.{parts[2]}.1")
        self.entries['gateway'] = gw_entry
        row += 1

        # Subnet Mask — with quick-select buttons
        ttk.Label(frame, text="Subnet Mask:").grid(row=row, column=0, sticky='w', pady=3)
        subnet_frame = ttk.Frame(frame)
        subnet_frame.grid(row=row, column=1, sticky='ew', pady=3, padx=(10, 0))
        subnet_entry = ttk.Entry(subnet_frame, width=18)
        subnet_entry.pack(side=tk.LEFT)
        if camera and camera.get('subnet'):
            subnet_entry.insert(0, camera['subnet'])
        else:
            subnet_entry.insert(0, '255.255.255.0')
        self.entries['subnet'] = subnet_entry
        def set_subnet(mask):
            subnet_entry.delete(0, tk.END)
            subnet_entry.insert(0, mask)
        ttk.Button(subnet_frame, text="/24", width=4,
                   command=lambda: set_subnet('255.255.255.0')).pack(side=tk.LEFT, padx=2)
        ttk.Button(subnet_frame, text="/16", width=4,
                   command=lambda: set_subnet('255.255.0.0')).pack(side=tk.LEFT, padx=2)
        ttk.Button(subnet_frame, text="/8", width=4,
                   command=lambda: set_subnet('255.0.0.0')).pack(side=tk.LEFT, padx=2)
        row += 1

        # Remaining fields
        for label, key in [("New IP (for updates):", "new_ip"),
                           ("Hostname (optional):", "hostname"),
                           ("Model (optional):", "model")]:
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky='w', pady=3)
            entry = ttk.Entry(frame, width=30)
            entry.grid(row=row, column=1, sticky='ew', pady=3, padx=(10, 0))
            if camera and camera.get(key):
                entry.insert(0, camera[key])
            self.entries[key] = entry
            row += 1
        
        # DHCP checkbox after the form fields
        next_row = row
        self.dhcp_var = tk.BooleanVar(value=False)
        if camera and camera.get('dhcp', '').lower() == 'yes':
            self.dhcp_var.set(True)
        
        dhcp_frame = ttk.Frame(frame)
        dhcp_frame.grid(row=next_row, column=0, columnspan=2, sticky='w', pady=(8, 0))
        self.dhcp_check = ttk.Checkbutton(dhcp_frame, text="DHCP Enabled", variable=self.dhcp_var)
        self.dhcp_check.pack(side=tk.LEFT)
        ttk.Label(dhcp_frame, text="(uncheck to set static IP during programming)", 
                 foreground='gray', font=('Helvetica', 8)).pack(side=tk.LEFT, padx=(10, 0))
        
        frame.columnconfigure(1, weight=1)
        
        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=next_row + 1, column=0, columnspan=2, pady=(20, 0))
        ttk.Button(btn_frame, text="Save", command=self.save).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.cancel).pack(side=tk.LEFT, padx=5)
        
        self.entries['name'].focus_set()
        self.bind("<Return>", lambda e: self.save())
        self.bind("<Escape>", lambda e: self.cancel())
        self.wait_window(self)
    
    def _auto_fill_gateway(self):
        """Auto-fill gateway from IP address (replace last octet with .1) if gateway is empty."""
        try:
            ip = self.entries['ip'].get().strip()
            gw = self.entries['gateway'].get().strip()
            if ip and (not gw or gw == '0.0.0.0') and not ip.startswith('169.254.'):
                parts = ip.split('.')
                if len(parts) == 4 and all(p.isdigit() for p in parts):
                    self.entries['gateway'].delete(0, 'end')
                    self.entries['gateway'].insert(0, f"{parts[0]}.{parts[1]}.{parts[2]}.1")
        except:
            pass

    def validate_ip(self, ip):
        if not ip: return True  # Empty is OK for optional fields
        pattern = r'^(\d{1,3}\.){3}\d{1,3}$'
        if not re.match(pattern, ip): return False
        parts = ip.split('.')
        return all(0 <= int(p) <= 255 for p in parts)
    
    def save(self):
        # Validate
        name = self.entries['name'].get().strip()
        ip = self.entries['ip'].get().strip()
        
        if not name:
            messagebox.showwarning("Required", "Camera Name is required")
            self.entries['name'].focus_set()
            return
        
        if not ip:
            messagebox.showwarning("Required", "IP Address is required")
            return
        
        if not self.validate_ip(ip):
            messagebox.showwarning("Invalid", "IP Address format is invalid")
            return
        
        gateway = self.entries['gateway'].get().strip()
        if gateway and not self.validate_ip(gateway):
            messagebox.showwarning("Invalid", "Gateway format is invalid")
            self.entries['gateway'].focus_set()
            return
        
        # Track what user actually changed
        pending = []
        if self.camera:
            new_hostname = self.entries['hostname'].get().strip()
            if new_hostname and new_hostname != self.camera.get('hostname', ''):
                pending.append('hostname')
            new_dhcp = 'Yes' if self.dhcp_var.get() else 'No'
            if new_dhcp != self.camera.get('dhcp', 'No'):
                pending.append('dhcp')
            new_ip_val = self.entries['new_ip'].get().strip()
            if new_ip_val:
                pending.append('ip')
        
        self.result = {
            'name': name,
            'hostname': self.entries['hostname'].get().strip(),
            'ip': ip,
            'gateway': gateway,
            'subnet': self.entries['subnet'].get().strip(),
            'model': self.entries['model'].get().strip(),
            'new_ip': self.entries['new_ip'].get().strip(),
            'dhcp': 'Yes' if self.dhcp_var.get() else 'No',
            'serial': self.camera.get('serial', '') if self.camera else '',
            'mac': self.camera.get('mac', '') if self.camera else '',
            'brand': self.camera.get('brand', 'axis') if self.camera else 'axis',
            'pending': pending,
            'processed': False
        }
        self.destroy()
    
    def cancel(self):
        self.result = None
        self.destroy()


# ============================================================================
# LLDP SWITCH PORT DISCOVERY
# ============================================================================

class LldpDiscoveryDialog(tk.Toplevel):
    """Listens for LLDP frames to identify the connected switch and port.
    Uses pktmon (built-in Windows 10/11) for packet capture."""

    CAPTURE_TIMEOUT = 35  # LLDP default interval is 30s

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Identify Switch Port (LLDP)")
        self.transient(parent)
        self.grab_set()
        self.resizable(True, True)
        self._thread = None
        self._cancel = threading.Event()
        self._temp_files = []
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        frame = ttk.Frame(self, padding=15)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="LLDP Switch Port Discovery",
                  font=('Helvetica', 14, 'bold')).pack(anchor='w')
        ttk.Label(frame, text="Listens for LLDP advertisements from connected switches.\n"
                  "Most managed switches broadcast every 30 seconds.",
                  foreground='gray', font=('Helvetica', 9)).pack(anchor='w', pady=(2, 10))

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(0, 8))
        self.listen_btn = ttk.Button(btn_frame, text="Start Listening",
                                      command=self._start)
        self.listen_btn.pack(side=tk.LEFT)
        self.stop_btn = ttk.Button(btn_frame, text="Stop",
                                    command=self._stop, state='disabled')
        self.stop_btn.pack(side=tk.LEFT, padx=(10, 0))

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(frame, textvariable=self.status_var,
                  font=('Helvetica', 10)).pack(anchor='w')
        self.progress = ttk.Progressbar(frame, mode='determinate',
                                         maximum=self.CAPTURE_TIMEOUT)
        self.progress.pack(fill=tk.X, pady=(4, 10))

        results_lf = ttk.LabelFrame(frame, text="Results", padding=8)
        results_lf.pack(fill=tk.BOTH, expand=True)
        self.results_text = tk.Text(results_lf, height=14, width=65,
                                     state='disabled', font=('Consolas', 10),
                                     wrap='word')
        scroll = ttk.Scrollbar(results_lf, orient='vertical',
                                command=self.results_text.yview)
        self.results_text.configure(yscrollcommand=scroll.set)
        self.results_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        bottom = ttk.Frame(frame)
        bottom.pack(fill=tk.X, pady=(8, 0))
        self.copy_btn = ttk.Button(bottom, text="Copy to Clipboard",
                                    command=self._copy, state='disabled')
        self.copy_btn.pack(side=tk.RIGHT)
        ttk.Button(bottom, text="Close",
                   command=self._on_close).pack(side=tk.RIGHT, padx=(0, 8))

        # Use _center_on_parent for multi-monitor correctness + content-aware sizing
        _center_on_parent(self, parent, 580, 500)

    def _start(self):
        self._cancel.clear()
        self.listen_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.copy_btn.config(state='disabled')
        self.progress['value'] = 0
        self._set_results("")
        self.status_var.set("Starting packet capture...")
        self._thread = threading.Thread(target=self._capture_worker, daemon=True)
        self._thread.start()
        self._tick(0)

    def _tick(self, elapsed):
        if self._cancel.is_set() or self._thread is None or not self._thread.is_alive():
            return
        self.progress['value'] = elapsed
        remaining = self.CAPTURE_TIMEOUT - elapsed
        self.status_var.set(f"Listening for LLDP frames... {remaining}s remaining")
        if elapsed < self.CAPTURE_TIMEOUT:
            self.after(1000, self._tick, elapsed + 1)

    def _stop(self):
        self._cancel.set()
        self.stop_btn.config(state='disabled')
        self.status_var.set("Stopping capture...")

    def _on_close(self):
        self._cancel.set()
        self.after(500, self.destroy)

    def _set_results(self, text):
        self.results_text.config(state='normal')
        self.results_text.delete('1.0', tk.END)
        if text:
            self.results_text.insert('1.0', text)
        self.results_text.config(state='disabled')

    def _copy(self):
        text = self.results_text.get('1.0', tk.END).strip()
        if text:
            self.clipboard_clear()
            self.clipboard_append(text)
            self.status_var.set("Copied to clipboard!")

    def _capture_worker(self):
        """Background thread: capture LLDP via pktmon, parse results."""
        import subprocess as sp
        import tempfile

        etl = tempfile.mktemp(suffix='.etl')
        pcap = tempfile.mktemp(suffix='.pcapng')
        self._temp_files = [etl, pcap]
        kw = dict(capture_output=True, timeout=10,
                  creationflags=0x08000000)  # CREATE_NO_WINDOW

        try:
            # Clean slate
            sp.run(['pktmon', 'stop'], **kw)
            sp.run(['pktmon', 'filter', 'remove'], **kw)

            # Filter for LLDP EtherType 0x88CC.
            # Modern pktmon (Win10 2004+ / Win11) uses -d for ethertype/datalink;
            # older versions used -t. Try new syntax first, fall back to old.
            r = sp.run(['pktmon', 'filter', 'add', 'LLDP', '-d', '0x88CC'], **kw)
            if r.returncode != 0:
                # Fallback to legacy syntax for pre-2004 pktmon
                r = sp.run(['pktmon', 'filter', 'add', 'LLDP', '-t', '0x88CC'], **kw)
            if r.returncode != 0:
                self.after(0, self._done, None,
                           "pktmon filter failed:\n" +
                           r.stderr.decode(errors='replace') +
                           "\n(Need Windows 10 2004+ or Windows 11, with admin rights.)")
                return

            # Start capture
            r = sp.run(['pktmon', 'start', '-c', '--pkt-size', '512',
                         '-f', etl], **kw)
            if r.returncode != 0:
                self.after(0, self._done, None,
                           "pktmon start failed:\n" +
                           r.stderr.decode(errors='replace'))
                return

            # Wait for frames or cancellation
            t0 = time.time()
            while time.time() - t0 < self.CAPTURE_TIMEOUT:
                if self._cancel.is_set():
                    break
                time.sleep(0.5)

            sp.run(['pktmon', 'stop'], **kw)
            time.sleep(0.3)

            # Convert ETL -> pcapng
            sp.run(['pktmon', 'etl2pcap', etl, '-o', pcap], **kw)

            if not os.path.exists(pcap) or os.path.getsize(pcap) < 100:
                self.after(0, self._done, [], None)
                return

            # Parse
            packets = self._read_pcapng(pcap)
            results = []
            seen = set()
            for pkt in packets:
                info = self._parse_lldp(pkt)
                if info:
                    key = (info['switch_name'], info['port_id'])
                    if key not in seen:
                        seen.add(key)
                        results.append(info)

            self.after(0, self._done, results, None)

        except Exception as e:
            self.after(0, self._done, None, f"Error: {e}")
        finally:
            try:
                sp.run(['pktmon', 'filter', 'remove'], capture_output=True,
                       timeout=5, creationflags=0x08000000)
            except:
                pass
            for f in self._temp_files:
                try:
                    os.remove(f)
                except:
                    pass

    def _done(self, results, error):
        """Main-thread callback when capture finishes."""
        self.listen_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        self.progress['value'] = self.CAPTURE_TIMEOUT

        if error:
            self.status_var.set("Capture failed.")
            self._set_results(error)
            return

        if not results:
            self.status_var.set("No LLDP frames received.")
            self._set_results(
                "No LLDP advertisements detected.\n\n"
                "Possible causes:\n"
                "  - LLDP is not enabled on the switch\n"
                "  - Not connected via wired Ethernet\n"
                "  - Switch interval > 35s — try again\n"
                "  - Some unmanaged switches don't support LLDP")
            return

        self.status_var.set(f"Found {len(results)} switch port(s)!")
        self.copy_btn.config(state='normal')

        lines = []
        for i, r in enumerate(results):
            if i > 0:
                lines.append("\n" + "\u2500" * 48)
            lines.append(f"Switch Name : {r['switch_name'] or '(not advertised)'}")
            lines.append(f"Port ID     : {r['port_id'] or '(not advertised)'}")
            if r['port_desc']:
                lines.append(f"Port Desc   : {r['port_desc']}")
            if r['chassis_id']:
                lines.append(f"Chassis ID  : {r['chassis_id']}")
            lines.append(f"Switch MAC  : {r['switch_mac']}")
            if r['mgmt_ip']:
                lines.append(f"Mgmt IP     : {r['mgmt_ip']}")
            if r['vlan']:
                lines.append(f"VLAN        : {r['vlan']}")
            if r['system_desc']:
                # Truncate long sysDescr to keep it readable
                desc = r['system_desc']
                if len(desc) > 120:
                    desc = desc[:117] + '...'
                lines.append(f"System      : {desc}")

        self._set_results('\n'.join(lines))

    @staticmethod
    def _read_pcapng(filepath):
        """Extract raw packet data from a pcapng file."""
        packets = []
        with open(filepath, 'rb') as f:
            data = f.read()
        offset = 0
        while offset + 12 <= len(data):
            btype = struct.unpack('<I', data[offset:offset + 4])[0]
            blen = struct.unpack('<I', data[offset + 4:offset + 8])[0]
            if blen < 12 or offset + blen > len(data):
                break
            if btype == 0x00000006:  # Enhanced Packet Block
                if offset + 28 <= len(data):
                    cap_len = struct.unpack('<I', data[offset + 20:offset + 24])[0]
                    end = offset + 28 + cap_len
                    if end <= len(data):
                        packets.append(data[offset + 28:end])
            offset += blen
        return packets

    @staticmethod
    def _parse_lldp(frame):
        """Parse LLDP TLVs from a raw Ethernet frame."""
        if len(frame) < 16:
            return None

        # Standard Ethernet: dst(6) src(6) ethertype(2)
        ethertype = struct.unpack('!H', frame[12:14])[0]
        payload_off = 14

        # Skip 802.1Q tag
        if ethertype == 0x8100 and len(frame) >= 18:
            ethertype = struct.unpack('!H', frame[16:18])[0]
            payload_off = 18

        # pktmon may prepend metadata — scan for 0x88CC in first 64 bytes
        if ethertype != 0x88CC:
            for i in range(12, min(64, len(frame) - 1)):
                if frame[i] == 0x88 and frame[i + 1] == 0xCC:
                    payload_off = i + 2
                    ethertype = 0x88CC
                    break
        if ethertype != 0x88CC:
            return None

        # Source MAC is 6 bytes before ethertype
        mac_start = payload_off - 8  # src is 6 bytes before the 2-byte ethertype
        if mac_start < 0:
            mac_start = 6
        src_mac = frame[mac_start:mac_start + 6]

        r = {'switch_mac': ':'.join(f'{b:02X}' for b in src_mac),
             'switch_name': '', 'port_id': '', 'port_desc': '',
             'system_desc': '', 'mgmt_ip': '', 'chassis_id': '',
             'ttl': 0, 'vlan': ''}

        off = payload_off
        while off + 2 <= len(frame):
            hdr = struct.unpack('!H', frame[off:off + 2])[0]
            ttype = (hdr >> 9) & 0x7F
            tlen = hdr & 0x01FF
            off += 2
            if ttype == 0:
                break
            if off + tlen > len(frame):
                break
            val = frame[off:off + tlen]
            off += tlen

            if ttype == 1 and len(val) > 1:        # Chassis ID
                sub = val[0]
                if sub == 4 and len(val) >= 7:
                    r['chassis_id'] = ':'.join(f'{b:02X}' for b in val[1:7])
                elif sub == 5 and len(val) >= 6 and val[1] == 1:
                    r['chassis_id'] = socket.inet_ntoa(val[2:6])
                else:
                    r['chassis_id'] = val[1:].decode('utf-8', errors='replace').rstrip('\x00')

            elif ttype == 2 and len(val) > 1:       # Port ID
                sub = val[0]
                if sub == 3 and len(val) >= 7:
                    r['port_id'] = ':'.join(f'{b:02X}' for b in val[1:7])
                else:
                    r['port_id'] = val[1:].decode('utf-8', errors='replace').rstrip('\x00')

            elif ttype == 3 and len(val) >= 2:      # TTL
                r['ttl'] = struct.unpack('!H', val[:2])[0]

            elif ttype == 4:                         # Port Description
                r['port_desc'] = val.decode('utf-8', errors='replace').rstrip('\x00')

            elif ttype == 5:                         # System Name
                r['switch_name'] = val.decode('utf-8', errors='replace').rstrip('\x00')

            elif ttype == 6:                         # System Description
                r['system_desc'] = val.decode('utf-8', errors='replace').rstrip('\x00')

            elif ttype == 8 and len(val) >= 6:       # Management Address
                if val[0] >= 5 and val[1] == 1:
                    r['mgmt_ip'] = socket.inet_ntoa(val[2:6])

            elif ttype == 127 and len(val) >= 6:     # Org-specific
                oui, sub = val[0:3], val[3]
                if oui == b'\x00\x80\xc2' and sub == 1:       # Port VLAN ID
                    r['vlan'] = str(struct.unpack('!H', val[4:6])[0])
                elif oui == b'\x00\x80\xc2' and sub == 3 and len(val) >= 7:  # VLAN Name
                    vid = struct.unpack('!H', val[4:6])[0]
                    nlen = val[6]
                    vname = val[7:7 + nlen].decode('utf-8', errors='replace') if nlen else ''
                    r['vlan'] = f"{vid}" + (f" ({vname})" if vname else '')

        return r


# ============================================================================
# MAIN APPLICATION
# ============================================================================
class CCTVToolkitApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"CCTV IP Toolkit v{APP_VERSION} - Brian Preston")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Start maximized
        try:
            self.root.state('zoomed')
        except tk.TclError:
            # Fallback: get screen size and set geometry
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            self.root.geometry(f"{sw}x{sh}+0+0")
        
        # Try to set icon (app icon, not the personal logo)
        try:
            import sys
            if getattr(sys, 'frozen', False):
                icon_path = os.path.join(sys._MEIPASS, 'app.ico')
            else:
                icon_path = 'app.ico'
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except: pass
        
        # Initialize managers
        self.settings = SettingsManager()
        self.camera_data = CameraDataManager()
        self.password_data = PasswordDataManager()
        self.additional_users_data = AdditionalUsersDataManager()

        # Initialize protocol from saved brand
        saved_brand = self.settings.get('general', 'brand')
        if saved_brand not in PROTOCOLS:
            saved_brand = 'axis'
        self.protocol = PROTOCOLS[saved_brand]()

        self.log_queue = queue.Queue()
        self.cancel_flag = False
        self.preview_image = None
        self.startup_scan_complete = False
        self._scan_running = False
        self._periodic_scan_id = None
        self._post_op_scan_id = None
        self._countdown_tick_id = None
        
        # Create UI
        self.create_menu()
        self.create_main_ui()

        # One-time notice if we migrated a legacy ./data/ folder on startup
        if _MIGRATION_NOTE:
            legacy, cfg, exp, copied = _MIGRATION_NOTE[0]
            n_cfg = sum(1 for kind, _ in copied if kind == 'config')
            n_exp = sum(1 for kind, _ in copied if kind == 'export')
            self.root.after(400, lambda: messagebox.showinfo(
                "Data folders moved",
                "Your existing data folder has been split so upgrades don't wipe it out:\n\n"
                f"  Config ({n_cfg} file(s)): {cfg}\n"
                f"    • passwords, camera list, settings\n\n"
                f"  Exports ({n_exp} item(s)): {exp}\n"
                f"    • CSVs, screenshots, FTP pulls\n\n"
                f"Your original folder is untouched at:\n  {legacy}\n"
                "You can delete it after confirming the new folders look right."
            ))

        # Check for first run
        if self.settings.get_bool('general', 'first_run'):
            self.show_welcome()
            # Show brand selection on first run
            dialog = BrandSelectionDialog(self.root, self.protocol.BRAND_KEY)
            if dialog.result:
                self.protocol = PROTOCOLS[dialog.result]()
                self.brand_var.set(dialog.result)
                self.settings.set('general', 'brand', dialog.result)
                self.factory_ip_label.config(
                    text=f"Factory IP: {self.protocol.FACTORY_IP}  |  User: {self.protocol.DEFAULT_USER}")
            self.settings.set('general', 'first_run', 'false')
        
        self.process_log_queue()

        # Start background network scan
        self.root.after(500, self.background_scan)

        # First-launch-of-this-version: show What's New popup once.
        last_seen = self.settings.get('general', 'last_seen_version') or ''
        if last_seen != APP_VERSION:
            # Delay so main UI is visible first, then fire the what's-new popup.
            self.root.after(1200, lambda: self.show_whats_new(APP_VERSION))
            self.settings.set('general', 'last_seen_version', APP_VERSION)

        # Silent update check on every launch. Runs off the UI thread so a slow
        # network call can't block startup; if a newer version exists and the user
        # hasn't dismissed that exact version, the dialog fires on the main thread.
        threading.Thread(
            target=lambda: self.root.after(2500, lambda: self.check_for_update(silent=True)),
            daemon=True,
        ).start()
    
    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Discover Cameras on Network...", command=self.discover_cameras)
        file_menu.add_command(label="Smart Import/Paste...", command=self.smart_import)
        file_menu.add_command(label="Export Cameras to CSV...", command=self.export_cameras)
        file_menu.add_separator()
        file_menu.add_command(label="Open Export Folder", command=self.open_export_folder)
        file_menu.add_command(label="Open Config Folder", command=self.open_config_folder)
        file_menu.add_separator()
        file_menu.add_command(label="Settings", command=self.show_settings)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Identify Switch Port (LLDP)...",
                               command=self.show_lldp_discovery)

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="Quick Start Guide", command=self.show_quick_start)
        help_menu.add_command(label="What's New", command=self.show_whats_new)
        help_menu.add_command(label="Check for Updates...", command=lambda: self.check_for_update(silent=False))
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_separator()
        help_menu.add_command(label="Buy Me A Coffee ☕", command=lambda: __import__('webbrowser').open('https://buymeacoffee.com/thelostping'))
        help_menu.add_command(label="Report Issues", command=lambda: __import__('webbrowser').open('mailto:axisprogrammer@thelostping.net'))
    
    def create_main_ui(self):
        # Brand selection bar (always visible above tabs)
        self.brand_bar = ttk.Frame(self.root)
        self.brand_bar.pack(fill=tk.X, padx=10, pady=(5, 0))

        ttk.Label(self.brand_bar, text="BRAND:", font=('Helvetica', 10, 'bold')).pack(side=tk.LEFT)

        self.brand_var = tk.StringVar(value=self.protocol.BRAND_KEY)
        for key, name in [('axis', 'Axis'), ('bosch', 'Bosch'), ('hanwha', 'Hanwha')]:
            rb = ttk.Radiobutton(self.brand_bar, text=name, variable=self.brand_var,
                                value=key, command=self._on_brand_change)
            rb.pack(side=tk.LEFT, padx=(10, 0))

        ttk.Separator(self.brand_bar, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=15)

        self.factory_ip_label = ttk.Label(self.brand_bar,
            text=f"Factory IP: {self.protocol.FACTORY_IP}  |  User: {self.protocol.DEFAULT_USER}",
            font=('Courier', 9), foreground='#666666')
        self.factory_ip_label.pack(side=tk.LEFT)

        # Main container with notebook (tabs)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tab 0: Camera List
        self.cameras_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.cameras_tab, text="📋 Camera List")
        self.create_cameras_tab()
        
        # Tab 1: Discovered Cameras
        self.discovered_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.discovered_tab, text="📡 Discovered")
        self.create_discovered_tab()
        
        # Tab 2: Password List
        self.passwords_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.passwords_tab, text="🔑 Passwords")
        self.create_passwords_tab()
        
        # Tab 3: Operations
        self.operations_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.operations_tab, text="⚡ Operations")
        self.create_operations_tab()
        
        # Tab 4: Programming Status (live checklist for new wizard)
        self.status_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.status_tab, text="🟢 Programming Status")
        self.create_status_tab()

        # Tab 5: Log/Results
        self.log_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.log_tab, text="📊 Log & Results")
        self.create_log_tab()

    def _on_brand_change(self):
        """Handle brand radio button change."""
        new_brand = self.brand_var.get()
        if new_brand not in PROTOCOLS:
            return
        self.protocol = PROTOCOLS[new_brand]()
        self.settings.set('general', 'brand', new_brand)
        self.factory_ip_label.config(
            text=f"Factory IP: {self.protocol.FACTORY_IP}  |  User: {self.protocol.DEFAULT_USER}")
        # Clear discovered list (different brand = different cameras)
        self.discovered_cameras = []
        self.refresh_discovered_list()
        self.log(f"Switched to {self.protocol.BRAND_NAME} protocol")
        # Restart background scan for new brand
        self.root.after(1000, lambda: self.background_scan(force=True, quiet=False))

    def create_cameras_tab(self):
        """Camera list editor tab"""
        self.cameras_frame = ttk.Frame(self.cameras_tab, padding="10")
        self.cameras_frame.pack(fill=tk.BOTH, expand=True)
        frame = self.cameras_frame
        
        # Header with instructions
        header = ttk.Frame(frame)
        header.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(header, text="Camera List", font=('Helvetica', 16, 'bold')).pack(side=tk.LEFT)
        ttk.Label(header, text="  •  Add cameras here for programming, pinging, and other operations", 
                 font=('Helvetica', 10), foreground='gray').pack(side=tk.LEFT, padx=(10, 0))
        
        # Toolbar
        toolbar = ttk.Frame(frame)
        toolbar.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(toolbar, text="➕ Add Camera", command=self.add_camera).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="✏️ Edit", command=self.edit_camera).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="🗑️ Delete", command=self.delete_camera).pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=10)
        ttk.Button(toolbar, text="📋 Smart Import/Paste", command=self.smart_import).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="📤 Export CSV", command=self.export_cameras).pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=10)
        ttk.Button(toolbar, text="🔄 Reset Status", command=self.reset_status).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="🗑️ Clear Done", command=self.clear_processed).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="🗑️ Clear All", command=self.clear_all_cameras).pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=10)
        ttk.Button(toolbar, text="⚠️ Factory Default", command=self.factory_default_camera).pack(side=tk.LEFT, padx=2)
        
        # Camera list treeview
        columns = ('name', 'hostname', 'ip', 'mac', 'gateway', 'subnet', 'model', 'new_ip', 'status')
        self.camera_tree = ttk.Treeview(frame, columns=columns, show='headings', height=15, selectmode='extended')
        
        self.camera_tree.heading('name', text='Camera Name')
        self.camera_tree.heading('hostname', text='Hostname')
        self.camera_tree.heading('ip', text='IP Address')
        self.camera_tree.heading('mac', text='MAC Address')
        self.camera_tree.heading('gateway', text='Gateway')
        self.camera_tree.heading('subnet', text='Subnet')
        self.camera_tree.heading('model', text='Model')
        self.camera_tree.heading('new_ip', text='New IP')
        self.camera_tree.heading('status', text='Status')
        
        self.camera_tree.column('name', width=130)
        self.camera_tree.column('hostname', width=130)
        self.camera_tree.column('ip', width=110)
        self.camera_tree.column('mac', width=130)
        self.camera_tree.column('gateway', width=100)
        self.camera_tree.column('subnet', width=100)
        self.camera_tree.column('model', width=100)
        self.camera_tree.column('new_ip', width=110)
        self.camera_tree.column('status', width=80)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.camera_tree.yview)
        self.camera_tree.configure(yscrollcommand=scrollbar.set)
        
        self.camera_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Double-click to edit
        self.camera_tree.bind('<Double-1>', lambda e: self.edit_camera())
        # Keyboard shortcuts
        self.camera_tree.bind('<Delete>', lambda e: self.delete_camera())
        self.camera_tree.bind('<Return>', lambda e: self.edit_camera())
        self.camera_tree.bind('<Control-a>', lambda e: self.camera_tree.selection_set(self.camera_tree.get_children()))
        
        # Status bar
        self.camera_status = tk.StringVar(value="0 cameras")
        ttk.Label(frame, textvariable=self.camera_status, font=('Helvetica', 10)).pack(anchor=tk.W, pady=(5, 0))
        
        # Refresh list
        self.refresh_camera_list()
    
    def create_discovered_tab(self):
        """Discovered cameras tab - shows what's on the network"""
        frame = ttk.Frame(self.discovered_tab, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header = ttk.Frame(frame)
        header.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(header, text="Discovered Cameras", font=('Helvetica', 16, 'bold')).pack(side=tk.LEFT)
        ttk.Label(header, text="  •  Cameras found on the network (not your programming list)", 
                 font=('Helvetica', 10), foreground='gray').pack(side=tk.LEFT, padx=(10, 0))
        
        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(0, 5))
        self.rescan_btn = ttk.Button(btn_frame, text="🔄 Rescan Network", command=self.rescan_network)
        self.rescan_btn.pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="📋 Copy Selected to Camera List", command=self.copy_discovered_to_list).pack(side=tk.LEFT, padx=2)
        self.copy_new_btn = ttk.Button(btn_frame, text="📋 Copy All New to Camera List", command=self.copy_new_discovered_to_list)
        self.copy_new_btn.pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="🗑 Clear", command=self.clear_discovered).pack(side=tk.LEFT, padx=2)
        
        # Password bar placeholder for discovered tab
        self.discovered_password_bar = tk.Frame(frame)
        self.discovered_password_bar.pack(fill=tk.X)
        
        # Treeview
        columns = ('ip', 'hostname', 'model', 'mac', 'gateway', 'subnet', 'dhcp', 'on_list')
        self.discovered_tree = ttk.Treeview(frame, columns=columns, show='headings', selectmode='extended')
        
        col_widths = {'ip': 120, 'hostname': 150, 'model': 140, 'mac': 140, 'gateway': 110, 'subnet': 100, 'dhcp': 50, 'on_list': 200}
        col_labels = {'ip': 'IP', 'hostname': 'HOSTNAME', 'model': 'MODEL', 'mac': 'MAC', 
                      'gateway': 'GATEWAY', 'subnet': 'SUBNET', 'dhcp': 'DHCP', 'on_list': 'ON LIST'}
        for col in columns:
            self.discovered_tree.heading(col, text=col_labels.get(col, col.upper()))
            self.discovered_tree.column(col, width=col_widths.get(col, 100))
        
        # Tag for cameras already on list (gray out)
        self.discovered_tree.tag_configure('on_list', foreground='#999999')
        self.discovered_tree.tag_configure('new_cam', foreground='#000000')
        self.discovered_tree.tag_configure('mismatch', foreground='#CC6600')
        self.discovered_tree.tag_configure('linklocal', foreground='#0066CC', font=('Helvetica', 9, 'bold'))
        self.discovered_tree.tag_configure('bosch', foreground='#006633')
        self.discovered_tree.tag_configure('bosch_new', foreground='#006633', font=('Helvetica', 9, 'bold'))
        
        scroll = ttk.Scrollbar(frame, orient='vertical', command=self.discovered_tree.yview)
        self.discovered_tree.configure(yscrollcommand=scroll.set)
        self.discovered_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        scroll.pack(fill=tk.Y, side=tk.RIGHT)
        # Keyboard shortcuts for discovered tree
        self.discovered_tree.bind('<Return>', lambda e: self.copy_discovered_to_list())
        self.discovered_tree.bind('<Control-a>', lambda e: self.discovered_tree.selection_set(self.discovered_tree.get_children()))
        
        # Status bar with countdown
        status_frame = ttk.Frame(frame)
        status_frame.pack(fill=tk.X, pady=(5, 0))
        self.discovered_status = tk.StringVar(value="Startup scan running...")
        ttk.Label(status_frame, textvariable=self.discovered_status, font=('Helvetica', 10)).pack(side=tk.LEFT)
        self.rescan_countdown_var = tk.StringVar(value="")
        ttk.Label(status_frame, textvariable=self.rescan_countdown_var, 
                 font=('Helvetica', 9), foreground='#888888').pack(side=tk.RIGHT)
        self._rescan_seconds_left = 0
        
        # Store discovered cameras separately
        self.discovered_cameras = []
    
    def _get_camera_list_macs(self):
        """Get dict of MAC → camera data for duplicate/mismatch checking"""
        mac_map = {}
        for cam in self.camera_data.cameras:
            mac = cam.get('mac', '').upper().replace(':', '').replace('-', '').strip()
            if mac:
                mac_map[mac] = cam
        return mac_map
    
    def _check_discovered_vs_list(self, discovered_cam, list_cam):
        """Compare discovered camera against Camera List entry, return list of mismatches"""
        mismatches = []
        placeholder = '(Auth Required)'
        
        # Compare IP
        disc_ip = discovered_cam.get('ip', '').strip()
        list_ip = list_cam.get('ip', '').strip()
        if disc_ip and list_ip and disc_ip != list_ip:
            mismatches.append(f"IP: {disc_ip} ≠ {list_ip}")
        
        # Compare hostname — flag if discovered has one but list doesn't
        disc_host = discovered_cam.get('hostname', '').strip().lower()
        list_host = list_cam.get('hostname', '').strip().lower()
        if disc_host and list_host and disc_host != list_host:
            mismatches.append(f"hostname differs")
        elif disc_host and not list_host:
            mismatches.append("hostname missing")
        
        # Compare model — flag placeholder
        disc_model = discovered_cam.get('model', '').strip()
        list_model = list_cam.get('model', '').strip()
        if list_model == placeholder and disc_model and disc_model != placeholder:
            mismatches.append("model outdated")
        elif disc_model and list_model and disc_model != list_model and list_model != placeholder:
            mismatches.append("model differs")
        
        # Compare gateway
        disc_gw = discovered_cam.get('gateway', '').strip()
        list_gw = list_cam.get('gateway', '').strip()
        if disc_gw and list_gw and disc_gw != list_gw:
            mismatches.append("gateway")
        
        # Compare subnet
        disc_sub = discovered_cam.get('subnet', '').strip()
        list_sub = list_cam.get('subnet', '').strip()
        if disc_sub and list_sub and disc_sub != list_sub:
            mismatches.append("subnet")
        
        return mismatches
    
    def refresh_discovered_list(self):
        """Refresh the discovered cameras treeview"""
        self.discovered_tree.delete(*self.discovered_tree.get_children())
        mac_map = self._get_camera_list_macs()
        # Also build IP and serial maps for matching auth-required cameras
        ip_map = {}
        serial_map = {}
        for cam in self.camera_data.cameras:
            ip = cam.get('ip', '').strip()
            if ip:
                ip_map[ip] = cam
            serial = cam.get('serial', '')
            if serial:
                serial_map[serial] = cam
        
        new_count = 0
        mismatch_count = 0
        linklocal_count = 0
        for cam in self.discovered_cameras:
            # Check MAC against camera list first, then serial, then IP
            cam_mac = cam.get('mac', '').upper().replace(':', '').replace('-', '').strip()
            cam_serial = cam.get('serial', '')
            cam_ip = cam.get('ip', '').strip()
            is_linklocal = cam_ip.startswith('169.254.')
            
            if is_linklocal:
                linklocal_count += 1
            
            list_cam = None
            if cam_mac:
                list_cam = mac_map.get(cam_mac)
            if not list_cam and cam_serial:
                list_cam = serial_map.get(cam_serial)
            if not list_cam and cam_ip and not is_linklocal:
                # Don't match link-local IPs since they're temporary
                list_cam = ip_map.get(cam_ip)
            
            if list_cam:
                mismatches = self._check_discovered_vs_list(cam, list_cam)
                if mismatches:
                    on_list = f"⚠ Changed: {', '.join(mismatches[:2])}"
                    tag = 'mismatch'
                    mismatch_count += 1
                else:
                    on_list = '✓ Match'
                    tag = 'on_list'
            else:
                if is_linklocal:
                    on_list = '🔧 UNPROGRAMMED'  # Link-local = factory default
                    tag = 'linklocal'
                else:
                    on_list = ''
                    tag = 'new_cam'
                new_count += 1
            
            # Show link-local marker in IP column
            display_ip = cam_ip
            if is_linklocal:
                display_ip = f"{cam_ip} [LL]"

            display_model = cam.get('model', '')

            self.discovered_tree.insert('', 'end', values=(
                display_ip,
                cam.get('hostname', ''),
                display_model,
                cam.get('mac', ''),
                cam.get('gateway', ''),
                cam.get('subnet', ''),
                cam.get('dhcp', ''),
                on_list
            ), tags=(tag,))
        total = len(self.discovered_cameras)
        on = total - new_count - mismatch_count
        parts = [f"{total} discovered"]
        if linklocal_count:
            parts.append(f"{linklocal_count} link-local")
        if new_count:
            parts.append(f"{new_count} new")
        if mismatch_count:
            parts.append(f"{mismatch_count} changed")
        if on:
            parts.append(f"{on} match")
        self.discovered_status.set("  •  ".join(parts))
        # Grey out Copy All New button if nothing new
        if new_count == 0:
            self.copy_new_btn.state(['disabled'])
        else:
            self.copy_new_btn.state(['!disabled'])
        self._update_tab_counts()
    
    def copy_discovered_to_list(self):
        """Copy selected discovered cameras to the main camera list"""
        selected = self.discovered_tree.selection()
        if not selected:
            messagebox.showinfo("No Selection", "Select cameras to copy to the Camera List.")
            return
        
        added = 0
        for item in selected:
            idx = self.discovered_tree.index(item)
            if idx >= len(self.discovered_cameras):
                continue
            src = self.discovered_cameras[idx]
            
            cam = {
                'name': src.get('hostname', src.get('ip', '')),
                'ip': src.get('ip', ''),
                'hostname': src.get('hostname', ''),
                'model': src.get('model', '') if src.get('model', '') != '(Auth Required)' else '',
                'serial': src.get('serial', ''),
                'mac': src.get('mac', ''),
                'gateway': src.get('gateway', ''),
                'subnet': src.get('subnet', ''),
                'dhcp': src.get('dhcp', ''),
                'brand': src.get('brand', 'axis'),
                'processed': False
            }
            result = self.camera_data.upsert(cam)
            if result == 'added':
                added += 1
        
        # Dedup pass
        removed = self.camera_data.dedup_camera_list()
        
        if added or removed:
            self.camera_data.save()
            self.refresh_camera_list()
            self.refresh_discovered_list()
            msg = f"Copied {added} camera(s) from Discovered to Camera List"
            if removed:
                msg += f", removed {removed} duplicate(s)"
            self.log(msg)
            messagebox.showinfo("Copied", msg)
    
    def copy_new_discovered_to_list(self):
        """Copy only cameras NOT already on the Camera List (by MAC, serial, or IP)"""
        existing_macs = {}
        existing_serials = {}
        existing_ips = {}
        for cam in self.camera_data.cameras:
            mac = cam.get('mac', '').upper().replace(':', '').replace('-', '').strip()
            if mac:
                existing_macs[mac] = cam
            serial = cam.get('serial', '')
            if serial:
                existing_serials[serial] = cam
            ip = cam.get('ip', '').strip()
            if ip:
                existing_ips[ip] = cam
        
        added = 0
        skipped = 0
        for src in self.discovered_cameras:
            cam_mac = src.get('mac', '').upper().replace(':', '').replace('-', '').strip()
            cam_serial = src.get('serial', '')
            cam_ip = src.get('ip', '').strip()
            
            # Skip if MAC, serial, or IP already on list
            if cam_mac and cam_mac in existing_macs:
                skipped += 1
                continue
            if cam_serial and cam_serial in existing_serials:
                skipped += 1
                continue
            if cam_ip and cam_ip in existing_ips:
                skipped += 1
                continue
            
            cam = {
                'name': src.get('hostname', src.get('ip', '')),
                'ip': src.get('ip', ''),
                'hostname': src.get('hostname', ''),
                'model': src.get('model', '') if src.get('model', '') != '(Auth Required)' else '',
                'serial': src.get('serial', ''),
                'mac': src.get('mac', ''),
                'gateway': src.get('gateway', ''),
                'subnet': src.get('subnet', ''),
                'dhcp': src.get('dhcp', ''),
                'brand': src.get('brand', 'axis'),
                'processed': False
            }
            result = self.camera_data.upsert(cam)
            if result == 'added':
                added += 1
                # Update maps so next iteration sees it
                if cam_mac:
                    existing_macs[cam_mac] = cam
                if cam_serial:
                    existing_serials[cam_serial] = cam
                if cam_ip:
                    existing_ips[cam_ip] = cam
        
        # Full dedup pass — after adds/merges, remove entries that now share a MAC/serial
        # with a richer entry (catches stale auth-required ghosts)
        removed = self.camera_data.dedup_camera_list()
        
        if added or removed:
            self.camera_data.save()
            self.refresh_camera_list()
            self.refresh_discovered_list()
            msg = f"Copied {added} new camera(s) to Camera List (skipped {skipped} already on list)"
            if removed:
                msg += f", removed {removed} duplicate(s)"
            self.log(msg)
            messagebox.showinfo("Copied", msg)
        else:
            messagebox.showinfo("Nothing New", f"All discovered cameras are already on the Camera List.")
    
    def clear_discovered(self):
        """Clear the discovered cameras list"""
        self.discovered_cameras = []
        self.refresh_discovered_list()
    
    def rescan_network(self):
        """Manual rescan from Discovered tab"""
        self.discovered_cameras = []
        self.refresh_discovered_list()
        self.rescan_btn.configure(text="⏳ Scanning...", state='disabled')
        self.discovered_status.set("Scanning network...")
        self._rescan_seconds_left = 0
        self.rescan_countdown_var.set("Scanning...")
        self.background_scan(force=True)
    
    def _reset_rescan_btn(self):
        """Re-enable the Rescan button after scan completes"""
        try:
            self.rescan_btn.configure(text="🔄 Rescan Network", state='normal')
        except:
            pass

    def create_passwords_tab(self):
        """Password list editor tab + additional users section"""
        # Top/bottom split: passwords on top, additional users on bottom
        outer_split = ttk.PanedWindow(self.passwords_tab, orient=tk.VERTICAL)
        outer_split.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # ---- TOP: Password List ----
        frame = ttk.Frame(outer_split)
        outer_split.add(frame, weight=3)

        # Header
        header = ttk.Frame(frame)
        header.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(header, text="Password List", font=('Helvetica', 16, 'bold')).pack(side=tk.LEFT)
        ttk.Label(header, text="  •  Used for batch password testing to find unknown camera passwords",
                 font=('Helvetica', 10), foreground='gray').pack(side=tk.LEFT, padx=(10, 0))

        # Split view: list on left, add on right
        split = ttk.PanedWindow(frame, orient=tk.HORIZONTAL)
        split.pack(fill=tk.BOTH, expand=True)
        
        # Left: password list
        left_frame = ttk.Frame(split)
        split.add(left_frame, weight=2)
        
        toolbar = ttk.Frame(left_frame)
        toolbar.pack(fill=tk.X, pady=(0, 5))
        
        # Show/hide toggle
        self.passwords_visible = tk.BooleanVar(value=False)
        self.show_hide_btn = ttk.Button(toolbar, text="👁 Show", command=self.toggle_password_visibility)
        self.show_hide_btn.pack(side=tk.LEFT, padx=2)
        
        ttk.Button(toolbar, text="🗑️ Delete Selected", command=self.delete_password).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="🗑️ Clear All", command=self.clear_all_passwords).pack(side=tk.LEFT, padx=2)
        
        self.password_listbox = tk.Listbox(left_frame, font=('Courier', 11), height=15)
        pwd_scroll = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.password_listbox.yview)
        self.password_listbox.configure(yscrollcommand=pwd_scroll.set)
        self.password_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        pwd_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Right: add password
        right_frame = ttk.LabelFrame(split, text="Add Password", padding="15")
        split.add(right_frame, weight=1)
        
        # Single entry
        ttk.Label(right_frame, text="Enter a password to add:", font=('Helvetica', 10)).pack(anchor=tk.W)
        self.new_password_var = tk.StringVar()
        pwd_entry = ttk.Entry(right_frame, textvariable=self.new_password_var, font=('Helvetica', 12), width=25)
        pwd_entry.pack(fill=tk.X, pady=(5, 10))
        pwd_entry.bind('<Return>', lambda e: self.add_password())
        
        ttk.Button(right_frame, text="➕ Add Password", command=self.add_password).pack()
        
        ttk.Separator(right_frame, orient='horizontal').pack(fill=tk.X, pady=15)
        
        # Mass entry
        ttk.Label(right_frame, text="Bulk add (one per line):", font=('Helvetica', 10, 'bold')).pack(anchor=tk.W)
        self.mass_password_text = scrolledtext.ScrolledText(right_frame, font=('Courier', 10), height=6, width=25)
        self.mass_password_text.pack(fill=tk.X, pady=(5, 5))
        
        # Right-click menu for mass entry
        mass_menu = tk.Menu(self.mass_password_text, tearoff=0)
        mass_menu.add_command(label="Paste", command=lambda: self.mass_password_text.event_generate('<<Paste>>'))
        mass_menu.add_command(label="Clear", command=lambda: self.mass_password_text.delete('1.0', tk.END))
        self.mass_password_text.bind("<Button-3>", lambda e: mass_menu.tk_popup(e.x_root, e.y_root))
        
        ttk.Button(right_frame, text="➕ Add All", command=self.mass_add_passwords).pack(pady=(0, 10))
        
        ttk.Separator(right_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        # Common defaults
        ttk.Label(right_frame, text="Common defaults:", font=('Helvetica', 10, 'bold')).pack(anchor=tk.W)
        common = ["pass", "admin", "root", "password", "123456", "camera", "axis"]
        for pwd in common:
            btn = ttk.Button(right_frame, text=f"Add '{pwd}'", 
                           command=lambda p=pwd: self.add_password_quick(p))
            btn.pack(anchor=tk.W, pady=2)
        
        # Status
        self.password_status = tk.StringVar(value="0 passwords")
        ttk.Label(left_frame, textvariable=self.password_status, font=('Helvetica', 10)).pack(anchor=tk.W, pady=(5, 0))
        
        self.refresh_password_list()

        # ---- BOTTOM: Additional Users ----
        users_frame = ttk.LabelFrame(outer_split, text="Additional Users  •  Created on each camera during programming",
                                     padding="10")
        outer_split.add(users_frame, weight=2)

        # Users treeview
        users_top = ttk.Frame(users_frame)
        users_top.pack(fill=tk.BOTH, expand=True)

        cols = ('username', 'password', 'role')
        self.users_tree = ttk.Treeview(users_top, columns=cols, show='headings', height=5)
        self.users_tree.heading('username', text='Username')
        self.users_tree.heading('password', text='Password')
        self.users_tree.heading('role', text='Role')
        # anchor='center' on data columns matches the default-centered headings
        # — without this, headers center but data left-aligns (Brian flagged
        # 2026-05-02: visual mismatch, alignment bug)
        self.users_tree.column('username', width=150, anchor='center')
        self.users_tree.column('password', width=150, anchor='center')
        self.users_tree.column('role', width=120, anchor='center')
        users_scroll = ttk.Scrollbar(users_top, orient=tk.VERTICAL, command=self.users_tree.yview)
        self.users_tree.configure(yscrollcommand=users_scroll.set)
        self.users_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        users_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Add/remove controls
        users_controls = ttk.Frame(users_frame)
        users_controls.pack(fill=tk.X, pady=(8, 0))

        ttk.Label(users_controls, text="Username:").pack(side=tk.LEFT, padx=(0, 3))
        self.new_user_name_var = tk.StringVar()
        ttk.Entry(users_controls, textvariable=self.new_user_name_var, width=14).pack(side=tk.LEFT, padx=(0, 8))

        ttk.Label(users_controls, text="Password:").pack(side=tk.LEFT, padx=(0, 3))
        self.new_user_pwd_var = tk.StringVar()
        ttk.Entry(users_controls, textvariable=self.new_user_pwd_var, width=14).pack(side=tk.LEFT, padx=(0, 8))

        ttk.Label(users_controls, text="Role:").pack(side=tk.LEFT, padx=(0, 3))
        self.new_user_role_var = tk.StringVar(value='Operator')
        role_combo = ttk.Combobox(users_controls, textvariable=self.new_user_role_var,
                                  values=AdditionalUsersDataManager.ROLES, state='readonly', width=12)
        role_combo.pack(side=tk.LEFT, padx=(0, 8))

        ttk.Button(users_controls, text="Add User", command=self.add_additional_user).pack(side=tk.LEFT, padx=3)
        ttk.Button(users_controls, text="Delete Selected", command=self.delete_additional_user).pack(side=tk.LEFT, padx=3)
        ttk.Button(users_controls, text="Clear All", command=self.clear_additional_users).pack(side=tk.LEFT, padx=3)

        self.additional_users_status = tk.StringVar(value="0 additional users")
        ttk.Label(users_controls, textvariable=self.additional_users_status,
                 font=('Helvetica', 9), foreground='gray').pack(side=tk.RIGHT)

        self.refresh_additional_users_list()

    def create_operations_tab(self):
        """Operations tab with big buttons and wizards"""
        frame = ttk.Frame(self.operations_tab, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        ttk.Label(frame, text="Operations", font=('Helvetica', 20, 'bold')).pack(pady=(0, 20))
        
        # Grid of operation buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.BOTH, expand=True)
        
        operations = [
            ("🔧 Program New Cameras", "Step-by-step wizard with live\nchecklist (recommended)",
             self.start_program_wizard, "#4CAF50"),
            ("✅ Confirm Programming", "Audit each camera in the list:\nIP / auth / DHCP-off match expected",
             self.start_confirm_wizard, "#00ACC1"),
            ("♻️ Factory Default", "Wipe a camera back to factory state\n(prompts for IP + existing password)",
             self.start_factory_default_wizard, "#E91E63"),
            ("🔄 Update Cameras", "Push changes (IP, hostname, DHCP)\nfrom Camera List to cameras",
             self.start_update_wizard, "#2196F3"),
            ("📷 Capture Images", "Download snapshot images from\nall cameras in the list",
             self.start_capture_wizard, "#9C27B0"),
            ("📡 Ping Test", "Check connectivity to all cameras\nand export results to CSV",
             self.start_ping_wizard, "#FF9800"),
            ("✓ Validate Password", "Test ONE password against\nALL cameras in the list",
             self.start_validate_wizard, "#607D8B"),
            ("🔑 Change Passwords", "Change password on all cameras\n(requires current password)",
             self.start_change_password_wizard, "#F44336"),
            ("🔍 Batch Password Test", "Test MULTIPLE passwords to find\nunknown camera credentials",
             self.start_batch_test_wizard, "#795548"),
            ("🛠️ Classic Programmer", "Original combined-options dialog\n(fallback if new wizard misbehaves)",
             self.start_program_wizard_classic, "#9E9E9E"),
        ]
        
        for i, (text, desc, cmd, color) in enumerate(operations):
            row, col = divmod(i, 3)
            
            btn_container = ttk.Frame(btn_frame)
            btn_container.grid(row=row, column=col, padx=10, pady=10, sticky='nsew')
            
            btn = tk.Button(btn_container, text=text, font=('Helvetica', 12, 'bold'),
                          width=25, height=2, command=cmd, bg='#f0f0f0', relief=tk.RAISED)
            btn.pack(pady=(0, 5))
            
            ttk.Label(btn_container, text=desc, font=('Helvetica', 9), 
                     foreground='gray', justify=tk.CENTER).pack()
        
        for i in range(3):
            btn_frame.columnconfigure(i, weight=1)
        
        # Cancel button and status
        bottom_frame = ttk.Frame(frame)
        bottom_frame.pack(fill=tk.X, pady=(20, 0))
        
        self.cancel_btn = ttk.Button(bottom_frame, text="⏹️ Cancel Operation", 
                                    command=self.cancel_operation, state='disabled')
        self.cancel_btn.pack(side=tk.LEFT)
        
        # Current operation display
        display_frame = ttk.LabelFrame(bottom_frame, text="Current Operation", padding="10")
        display_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(20, 0))
        
        self.current_camera_label = ttk.Label(display_frame, text="----", font=('Courier', 24, 'bold'))
        self.current_camera_label.pack(side=tk.LEFT)
        
        self.current_status_label = ttk.Label(display_frame, text="Ready", font=('Helvetica', 12))
        self.current_status_label.pack(side=tk.LEFT, padx=(20, 0))
        
        # Preview
        self.preview_frame = ttk.LabelFrame(bottom_frame, text="Preview", padding="5")
        self.preview_label = ttk.Label(self.preview_frame, text="No image")
        self.preview_label.pack()
        
        # ---- Triplett FTP Section ----
        # DORMANT in v4.1 — Triplett Android integration is coming in a future release.
        # The UI is hidden; the implementation (send/retrieve/csv builders, see
        # triplett_send_cameras / _send_passwords / _send_file / _retrieve / _build_triplett_csv_lines /
        # _get_triplett_address / _ftp_push methods below) stays in the file and is
        # re-exposed by flipping TRIPLETT_UI_ENABLED to True.
        TRIPLETT_UI_ENABLED = False
        if TRIPLETT_UI_ENABLED:
            triplett_frame = ttk.LabelFrame(frame, text="📱 Send to Triplett", padding="10")
            triplett_frame.pack(fill=tk.X, pady=(15, 0))

            ttk.Label(triplett_frame,
                     text="Start FTP on the Triplett, then type the address shown on its screen.",
                     foreground='gray', font=('Helvetica', 9)).pack(anchor='w')

            addr_row = ttk.Frame(triplett_frame)
            addr_row.pack(fill=tk.X, pady=(8, 8))
            ttk.Label(addr_row, text="Triplett address:", font=('Helvetica', 10, 'bold')).pack(side=tk.LEFT)
            self.triplett_addr_var = tk.StringVar()
            triplett_entry = ttk.Entry(addr_row, textvariable=self.triplett_addr_var, width=28, font=('Courier', 11))
            triplett_entry.pack(side=tk.LEFT, padx=(10, 0))
            ttk.Label(addr_row, text="e.g. ftp://<device-ip>:2121",
                     foreground='gray', font=('Helvetica', 9)).pack(side=tk.LEFT, padx=(10, 0))

            send_row = ttk.Frame(triplett_frame)
            send_row.pack(fill=tk.X)
            self.triplett_cam_btn = tk.Button(send_row, text="📋 Send Camera List",
                                             command=self.triplett_send_cameras,
                                             bg='#4CAF50', fg='white', font=('Helvetica', 10, 'bold'),
                                             padx=12, pady=4, cursor='hand2')
            self.triplett_cam_btn.pack(side=tk.LEFT, padx=(0, 8))
            self.triplett_pwd_btn = tk.Button(send_row, text="🔑 Send Passwords",
                                             command=self.triplett_send_passwords,
                                             bg='#607D8B', fg='white', font=('Helvetica', 10, 'bold'),
                                             padx=12, pady=4, cursor='hand2')
            self.triplett_pwd_btn.pack(side=tk.LEFT, padx=(0, 8))
            self.triplett_file_btn = tk.Button(send_row, text="📁 Send File...",
                                              command=self.triplett_send_file,
                                              bg='#795548', fg='white', font=('Helvetica', 10, 'bold'),
                                              padx=12, pady=4, cursor='hand2')
            self.triplett_file_btn.pack(side=tk.LEFT)

            recv_row = ttk.Frame(triplett_frame)
            recv_row.pack(fill=tk.X, pady=(8, 0))
            self.triplett_recv_btn = tk.Button(recv_row, text="📥 Retrieve from Triplett",
                                               command=self.triplett_retrieve,
                                               bg='#1976D2', fg='white', font=('Helvetica', 10, 'bold'),
                                               padx=12, pady=4, cursor='hand2')
            self.triplett_recv_btn.pack(side=tk.LEFT, padx=(0, 8))
            ttk.Label(recv_row, text="Downloads images, logs, passwords & results → <export folder>/triplett/",
                     foreground='gray', font=('Helvetica', 9)).pack(side=tk.LEFT)
        else:
            # Placeholder so the user sees the feature is on the roadmap.
            placeholder = ttk.LabelFrame(frame, text="📱 Triplett Android Integration — Coming Soon",
                                         padding="10")
            placeholder.pack(fill=tk.X, pady=(15, 0))
            ttk.Label(placeholder,
                      text="Send camera lists, passwords, and files to the Triplett handheld over FTP, "
                           "and pull back images / logs / results when you're done. Re-enabled in a future build.",
                      foreground='gray', font=('Helvetica', 9), wraplength=720, justify=tk.LEFT).pack(anchor='w')

    # ========================================================================
    # PROGRAMMING STATUS TAB (live checklist for new wizard)
    # ========================================================================
    PROG_STEPS = [
        ('discover',     '1. Detect camera on the network'),
        ('pin',          '2. Lock onto camera (ARP pin)'),
        ('verify_model', '3. Verify camera model'),
        ('firmware',     '4. Read firmware version'),
        ('auth',         '5. Set password / create user'),
        ('extra_users',  '6. Create additional users'),
        ('hostname',     '7. Set hostname'),
        ('network',      '8. Apply IP / network settings'),
        ('verify_online','9. Wait for camera at new IP'),
        ('capture',      '10. Capture serial / MAC / image'),
    ]

    def create_status_tab(self):
        """Live programming status: banner, checklist, log."""
        frame = ttk.Frame(self.status_tab, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)

        # ---- Big banner ----
        self._status_banner_frame = tk.Frame(frame, bg='#9E9E9E', padx=20, pady=18)
        self._status_banner_frame.pack(fill=tk.X)
        self._status_banner_label = tk.Label(self._status_banner_frame,
            text="READY", bg='#9E9E9E', fg='white',
            font=('Helvetica', 28, 'bold'))
        self._status_banner_label.pack()
        self._status_banner_sub = tk.Label(self._status_banner_frame,
            text="Start a programming run from the Operations tab.",
            bg='#9E9E9E', fg='white', font=('Helvetica', 11))
        self._status_banner_sub.pack(pady=(4, 0))

        # ---- Camera + progress row ----
        info_row = ttk.Frame(frame)
        info_row.pack(fill=tk.X, pady=(15, 5))
        self._status_camera_label = ttk.Label(info_row,
            text="—", font=('Helvetica', 14, 'bold'))
        self._status_camera_label.pack(side=tk.LEFT)
        self._status_progress_label = ttk.Label(info_row,
            text="", font=('Helvetica', 11), foreground='#555')
        self._status_progress_label.pack(side=tk.RIGHT)

        # ---- Two columns: checklist on left, preview on right ----
        body = ttk.Frame(frame)
        body.pack(fill=tk.BOTH, expand=True, pady=(5, 5))

        # Checklist column
        check_frame = ttk.LabelFrame(body, text="Programming Steps", padding=10)
        check_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._status_step_widgets = {}
        for step_id, label in self.PROG_STEPS:
            row = ttk.Frame(check_frame)
            row.pack(fill=tk.X, pady=2)
            icon = tk.Label(row, text='○', font=('Helvetica', 14),
                            fg='#999', width=2)
            icon.pack(side=tk.LEFT)
            text = tk.Label(row, text=label, font=('Helvetica', 11),
                            fg='#666', anchor='w')
            text.pack(side=tk.LEFT, fill=tk.X, expand=True)
            detail = tk.Label(row, text='', font=('Helvetica', 9),
                              fg='#888', anchor='e')
            detail.pack(side=tk.RIGHT)
            self._status_step_widgets[step_id] = {
                'icon': icon, 'text': text, 'detail': detail,
            }

        # Preview column
        preview_col = ttk.Frame(body)
        preview_col.pack(side=tk.RIGHT, fill=tk.Y, padx=(15, 0))
        prev_box = ttk.LabelFrame(preview_col, text="Camera Preview", padding=8)
        prev_box.pack(fill=tk.BOTH, expand=False)
        self._status_preview_label = ttk.Label(prev_box, text="(no image yet)",
                                               width=40, anchor='center')
        self._status_preview_label.pack(padx=4, pady=4)

        # ---- Bottom: cancel + log toggle ----
        bottom = ttk.Frame(frame)
        bottom.pack(fill=tk.X, pady=(10, 0))
        self._status_cancel_btn = ttk.Button(bottom, text="⏹️ Cancel",
                                             command=self.cancel_operation,
                                             state='disabled')
        self._status_cancel_btn.pack(side=tk.LEFT)

        self._status_log_visible = tk.BooleanVar(value=True)
        ttk.Checkbutton(bottom, text="Show detailed log",
                        variable=self._status_log_visible,
                        command=self._toggle_status_log).pack(side=tk.LEFT, padx=(15, 0))

        # ---- Detailed log (collapsible) ----
        self._status_log_container = ttk.LabelFrame(frame, text="Detailed Log", padding=5)
        self._status_log_container.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self._status_log_text = scrolledtext.ScrolledText(
            self._status_log_container, font=('Courier', 9), height=8)
        self._status_log_text.pack(fill=tk.BOTH, expand=True)

        # Tally
        self._status_ok_count = 0
        self._status_fail_count = 0

    def _toggle_status_log(self):
        if self._status_log_visible.get():
            self._status_log_container.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        else:
            self._status_log_container.pack_forget()

    # ---------- Status view API (called from worker thread via root.after) ----
    def status_set_banner(self, text, subtitle='', color='#9E9E9E'):
        self._status_banner_frame.config(bg=color)
        self._status_banner_label.config(text=text, bg=color)
        self._status_banner_sub.config(text=subtitle, bg=color)

    def status_set_camera(self, name='—', detail=''):
        self._status_camera_label.config(text=name)
        self._status_progress_label.config(text=detail)

    def status_reset_steps(self, used_steps=None):
        """Reset all step icons. If used_steps given, dim unused ones."""
        for step_id, w in self._status_step_widgets.items():
            if used_steps is not None and step_id not in used_steps:
                w['icon'].config(text='—', fg='#ccc')
                w['text'].config(fg='#bbb')
            else:
                w['icon'].config(text='○', fg='#999')
                w['text'].config(fg='#666')
            w['detail'].config(text='')

    def status_set_step(self, step_id, state, detail=''):
        """state: 'pending', 'active', 'ok', 'fail', 'skip'."""
        w = self._status_step_widgets.get(step_id)
        if not w:
            return
        if state == 'active':
            w['icon'].config(text='⏳', fg='#FF9800')
            w['text'].config(fg='#000', font=('Helvetica', 11, 'bold'))
        elif state == 'ok':
            w['icon'].config(text='✓', fg='#4CAF50')
            w['text'].config(fg='#444', font=('Helvetica', 11))
        elif state == 'fail':
            w['icon'].config(text='✗', fg='#F44336')
            w['text'].config(fg='#444', font=('Helvetica', 11))
        elif state == 'skip':
            w['icon'].config(text='—', fg='#bbb')
            w['text'].config(fg='#bbb', font=('Helvetica', 11))
        else:  # pending
            w['icon'].config(text='○', fg='#999')
            w['text'].config(fg='#666', font=('Helvetica', 11))
        if detail:
            w['detail'].config(text=detail)

    def status_log(self, msg):
        """Append a line to the embedded detailed log + the main log tab + the
        wizard-run log file (if a run is active — see _open_wizard_run_log)."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        try:
            self._status_log_text.insert(tk.END, f"[{timestamp}] {msg}\n")
            self._status_log_text.see(tk.END)
        except Exception:
            pass
        # Auto-persist to disk for the active wizard run (#11)
        fh = getattr(self, '_wizard_run_log_fh', None)
        if fh:
            try:
                fh.write(f"[{timestamp}] {msg}\n")
                fh.flush()
            except Exception:
                pass
        self.log(msg, _persist_to_file=False)  # main log tab only — file already written

    def _open_wizard_run_log(self):
        """Auto-open EXPORT_DIR/wizard_logs/wizard_run_<timestamp>.log for the
        duration of a wizard run. Every status_log/log line is also written to
        disk so a crashed/interrupted run still leaves a complete log artifact.
        Brian's CMH 2026-04-30 ask (v4.3 list item k / #11): the manual Save
        Log button is fine but operators forget to click it, and a crash mid-
        run loses everything. This makes persistence automatic.
        Returns path string or None on failure."""
        try:
            from datetime import datetime as _dt
            ts = _dt.now().strftime('%Y%m%d_%H%M%S')
            log_dir = EXPORT_DIR / 'wizard_logs'
            log_dir.mkdir(parents=True, exist_ok=True)
            path = log_dir / f'wizard_run_{ts}.log'
            self._wizard_run_log_fh = open(path, 'w', encoding='utf-8', buffering=1)
            self._wizard_run_log_fh.write(f"=== Wizard run started {_dt.now().isoformat()} ===\n")
            self._wizard_run_log_fh.flush()
            return str(path)
        except Exception:
            self._wizard_run_log_fh = None
            return None

    def _close_wizard_run_log(self):
        fh = getattr(self, '_wizard_run_log_fh', None)
        if fh:
            try:
                from datetime import datetime as _dt
                fh.write(f"=== Wizard run ended {_dt.now().isoformat()} ===\n")
                fh.close()
            except Exception:
                pass
            self._wizard_run_log_fh = None

    def status_set_preview(self, image_data=None):
        if not HAS_PIL:
            return
        if image_data is None:
            self._status_preview_label.config(image='', text='(no image yet)')
            self._status_preview_image = None
            return
        try:
            img = Image.open(BytesIO(image_data))
            img.thumbnail((280, 200), Image.Resampling.LANCZOS)
            self._status_preview_image = ImageTk.PhotoImage(img)
            self._status_preview_label.config(image=self._status_preview_image, text='')
        except Exception:
            self._status_preview_label.config(image='', text='(preview error)')

    def status_enable_cancel(self, enable):
        self._status_cancel_btn.config(state='normal' if enable else 'disabled')

    def create_log_tab(self):
        """Log and results tab"""
        frame = ttk.Frame(self.log_tab, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header = ttk.Frame(frame)
        header.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(header, text="Operation Log", font=('Helvetica', 16, 'bold')).pack(side=tk.LEFT)
        ttk.Button(header, text="Clear Log", command=self.clear_log).pack(side=tk.RIGHT)
        ttk.Button(header, text="Save Log...", command=self.save_log).pack(side=tk.RIGHT, padx=(0, 5))
        self.log_cancel_btn = ttk.Button(header, text="⏹️ Cancel Operation", 
                                         command=self.cancel_operation, state='disabled')
        self.log_cancel_btn.pack(side=tk.RIGHT, padx=(0, 10))
        
        # Log text
        self.log_text = scrolledtext.ScrolledText(frame, font=('Courier', 10), height=25)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Right-click context menu for log
        log_menu = tk.Menu(self.log_text, tearoff=0)
        log_menu.add_command(label="Copy", command=lambda: self.log_copy_selection())
        log_menu.add_command(label="Copy All", command=lambda: self.log_copy_all())
        log_menu.add_separator()
        log_menu.add_command(label="Select All", command=lambda: self.log_text.tag_add('sel', '1.0', 'end'))
        log_menu.add_separator()
        log_menu.add_command(label="Clear Log", command=self.clear_log)
        self.log_text.bind('<Button-3>', lambda e: log_menu.tk_popup(e.x_root, e.y_root))
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(frame, textvariable=self.status_var, font=('Helvetica', 10), 
                 relief=tk.SUNKEN, anchor=tk.W).pack(fill=tk.X, pady=(10, 0))
    
    # ========================================================================
    # CAMERA LIST MANAGEMENT
    # ========================================================================
    def refresh_camera_list(self):
        self.camera_tree.delete(*self.camera_tree.get_children())
        cameras = self.camera_data.get_all()
        valid = 0
        for i, cam in enumerate(cameras):
            if cam.get('status') == 'failed':
                reason = cam.get('fail_reason', 'unknown')
                status = f"✗ Failed: {reason}" if reason else "✗ Failed"
            elif cam.get('processed'):
                status = "✓ Done"
            elif cam.get('ip') and cam.get('gateway') and cam.get('subnet'):
                status = "Ready"
                valid += 1
            else:
                missing = []
                if not cam.get('ip'):
                    missing.append("IP")
                if not cam.get('gateway'):
                    missing.append("GW")
                if not cam.get('subnet'):
                    missing.append("Subnet")
                status = f"⚠ Need {', '.join(missing)}"
            display_model = cam.get('model', '')
            pass  # Brand is global — no per-camera prefix needed
            self.camera_tree.insert('', 'end', iid=str(i), values=(
                cam.get('name', ''), cam.get('hostname', ''), cam.get('ip', ''), cam.get('mac', ''),
                cam.get('gateway', ''), cam.get('subnet', ''), display_model,
                cam.get('new_ip', ''), status
            ))
        self.camera_status.set(f"{len(cameras)} cameras total, {valid} ready for programming")
        self._update_tab_counts()
    
    def _update_tab_counts(self):
        """Update tab labels with counts"""
        cam_count = len(self.camera_data.get_all())
        self.notebook.tab(self.cameras_tab, text=f"📋 Camera List ({cam_count})")
        if hasattr(self, 'discovered_tab'):
            disc_count = len(getattr(self, 'discovered_cameras', []))
            self.notebook.tab(self.discovered_tab, text=f"📡 Discovered ({disc_count})")
    
    def add_camera(self):
        dialog = CameraEditorDialog(self.root, settings=self.settings)
        if dialog.result:
            self.camera_data.add(dialog.result)
            self.refresh_camera_list()
            self.log(f"Added camera: {dialog.result['name']}")
    
    def edit_camera(self):
        selected = self.camera_tree.selection()
        if not selected:
            messagebox.showinfo("Select Camera", "Please select a camera to edit")
            return
        idx = int(selected[0])
        cameras = self.camera_data.get_all()
        if idx < len(cameras):
            dialog = CameraEditorDialog(self.root, camera=cameras[idx], settings=self.settings)
            if dialog.result:
                self.camera_data.update(idx, dialog.result)
                self.refresh_camera_list()
                self.log(f"Updated camera: {dialog.result['name']}")
    
    def delete_camera(self):
        selected = self.camera_tree.selection()
        if not selected:
            messagebox.showinfo("Select Camera", "Please select a camera to delete")
            return
        count = len(selected)
        label = "camera" if count == 1 else f"{count} cameras"
        if messagebox.askyesno("Confirm Delete", f"Delete {label}?"):
            # Delete in reverse order so indexes don't shift
            for item in sorted(selected, key=lambda x: int(x), reverse=True):
                self.camera_data.delete(int(item))
            self.refresh_camera_list()
            self.log(f"Deleted {count} camera(s)")
    
    def clear_all_cameras(self):
        if messagebox.askyesno("Confirm", "Delete ALL cameras from the list?"):
            self.camera_data.clear()
            self.refresh_camera_list()
            self.log("Cleared all cameras")
    
    def reset_status(self):
        """Reset all processed/failed flags back to Ready"""
        cameras = self.camera_data.get_all()
        count = sum(1 for cam in cameras if cam.get('processed') or cam.get('status') == 'failed')
        if count == 0:
            messagebox.showinfo("Nothing to Reset", 
                "No cameras are marked as done or failed.\n\n"
                "All cameras already show 'Ready' status.")
            return
        for cam in cameras:
            cam['processed'] = False
            cam.pop('status', None)
            cam.pop('fail_reason', None)
        self.camera_data.save()
        self.refresh_camera_list()
        self.log(f"Reset status on {count} camera(s) back to Ready")
    
    def clear_processed(self):
        """Delete cameras marked as Done from the list"""
        cameras = self.camera_data.get_all()
        done = [cam for cam in cameras if cam.get('processed')]
        if not done:
            messagebox.showinfo("Nothing to Clear", 
                "No cameras are marked as done.")
            return
        if messagebox.askyesno("Confirm", 
            f"Remove {len(done)} completed camera(s) from the list?\n\n"
            "This deletes them — use 'Reset Status' if you\n"
            "just want to reprogram them."):
            self.camera_data.cameras = [cam for cam in cameras if not cam.get('processed')]
            self.camera_data.save()
            self.refresh_camera_list()
            self.log(f"Removed {len(done)} completed camera(s) from list")
    
    def factory_default_camera(self):
        """Factory default a SINGLE selected camera — requires typing YES"""
        selected = self.camera_tree.selection()
        
        if not selected:
            messagebox.showinfo("No Selection", "Select ONE camera to factory default.")
            return
        
        if len(selected) > 1:
            messagebox.showwarning("One Camera Only", 
                "Factory default operates on ONE camera at a time.\n\n"
                "Please select exactly one camera.")
            return
        
        idx = int(selected[0])
        cam = self.camera_data.cameras[idx]
        cam_name = cam.get('name', cam.get('ip', '?'))
        cam_ip = cam.get('ip', '')
        
        if not cam_ip:
            messagebox.showwarning("No IP", "This camera has no IP address.")
            return
        
        # Scary confirmation dialog — must type YES
        confirm = simpledialog.askstring("⚠️ FACTORY DEFAULT ⚠️",
            f"You are about to FACTORY DEFAULT:\n\n"
            f"  Camera: {cam_name}\n"
            f"  IP: {cam_ip}\n\n"
            f"This will ERASE ALL SETTINGS on the camera.\n"
            f"The camera will reboot to factory defaults.\n"
            f"This cannot be undone.\n\n"
            f"Type YES to confirm:",
            parent=self.root)
        
        if not confirm or confirm.strip().upper() != 'YES':
            self.log(f"Factory default cancelled for {cam_name}")
            return
        
        # Get password
        password = self.get_password("Factory Default", 
            f"Enter current password for {cam_name}:")
        if not password:
            return
        
        self.notebook.select(self.log_tab)
        self.log(f"\n⚠️ FACTORY DEFAULTING: {cam_name} ({cam_ip})")
        
        def run():
            success = False

            self.log(f"  Sending {self.protocol.BRAND_NAME} factory reset...")
            try:
                if self.protocol.factory_reset(cam_ip, password):
                    success = True
                    self.log(f"  ✓ Factory reset command accepted")
            except requests.exceptions.Timeout:
                success = True
                self.log("  ✓ Camera stopped responding (likely rebooting)")
            except Exception as e:
                self.log(f"  Factory reset failed: {e}")

            if success:
                self.log(f"\n  Camera is rebooting to factory defaults.")
                self.log(f"  This takes 2-5 minutes. Camera will come back at factory IP.")
                
                # Ask to remove from Camera List (on main thread)
                cam_mac = cam.get('mac', '').upper().replace(':', '').replace('-', '').strip()
                def ask_remove():
                    if messagebox.askyesno("Remove from Camera List?",
                        f"Camera '{cam_name}' has been factory defaulted.\n"
                        f"Its old programming is now invalid.\n\n"
                        f"Remove it from the Camera List?\n\n"
                        f"(It will reappear in Discovered once it reboots)"):
                        # Remove ALL entries with this MAC (cleans up duplicates)
                        if cam_mac:
                            before = len(self.camera_data.cameras)
                            self.camera_data.cameras = [
                                c for c in self.camera_data.cameras
                                if c.get('mac', '').upper().replace(':', '').replace('-', '').strip() != cam_mac
                            ]
                            removed = before - len(self.camera_data.cameras)
                        else:
                            # No MAC — just remove by index
                            try:
                                self.camera_data.cameras.pop(idx)
                                removed = 1
                            except IndexError:
                                removed = 0
                        self.camera_data.save()
                        self.refresh_camera_list()
                        self.log(f"  Removed {removed} entry/entries for {cam_name} from Camera List")
                self.root.after(0, ask_remove)
                # Camera takes 45-60s to reboot after factory default — stagger rescans
                self.log("  Waiting for camera to reboot... (rescanning at 30s, 60s, 90s)")
                self.root.after(30000, lambda: self.background_scan(force=True, quiet=True))
                self.root.after(60000, lambda: self.background_scan(force=True, quiet=True))
                self.root.after(90000, lambda: self.background_scan(force=True, quiet=True))
            else:
                self.log(f"\n  ✗ Factory default FAILED — could not reach camera or wrong password.")
        
        threading.Thread(target=run, daemon=True).start()
    
    def smart_import(self):
        """Smart import with column detection"""
        dialog = SmartImportDialog(self.root)
        if dialog.result:
            added = updated = 0
            for cam in dialog.result:
                result = self.camera_data.upsert(cam)
                if result == 'added':
                    added += 1
                else:
                    updated += 1
            self.camera_data.save()
            self.camera_data.dedup_camera_list()
            self.camera_data.save()
            self.refresh_camera_list()
            msg = []
            if added:
                msg.append(f"{added} added")
            if updated:
                msg.append(f"{updated} updated")
            summary = ', '.join(msg)
            self.log(f"Imported cameras: {summary}")
            messagebox.showinfo("Import Complete", f"Imported cameras: {summary}")
    
    def get_local_network_info(self):
        """Get local IP address, subnet mask, and suggest appropriate scan range.
        Works on any Windows PC — isolated networks, any locale, any adapter name.
        Uses factory_ip to pick the correct adapter on multi-adapter systems."""
        import subprocess

        local_ip = None
        subnet_mask = None
        suggested_range = "192.168.1.1-254"
        self._detected_iface_index = None  # numeric interface index for netsh
        self._detected_local_ip = None

        # Get factory IP for subnet matching
        try:
            factory_ip = self.settings.get('general', 'factory_ip') if hasattr(self, 'settings') else '192.168.0.90'
        except:
            factory_ip = '192.168.0.90'

        # --- Step 1: Get ALL IPv4 addresses via PowerShell (locale-independent) ---
        # Returns IP + PrefixLength + InterfaceIndex in one query.
        # We score candidates to pick the right adapter — not VPN/tunnel.
        try:
            cmd = ("Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | "
                   "Select-Object IPAddress,PrefixLength,InterfaceIndex | "
                   "ConvertTo-Json -Compress")
            result = subprocess.run(
                ['powershell', '-NoProfile', '-Command', cmd],
                capture_output=True, text=True, timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode == 0 and result.stdout.strip():
                entries = json.loads(result.stdout)
                if isinstance(entries, dict):
                    entries = [entries]

                candidates = []
                for e in entries:
                    ip = e.get('IPAddress', '')
                    prefix = int(e.get('PrefixLength', 24))
                    idx = e.get('InterfaceIndex')

                    # Skip unsuitable addresses
                    if not ip or ip.startswith('127.') or ip.startswith('169.254.') or ip == '0.0.0.0':
                        continue

                    # Compute subnet mask from prefix length
                    mask_int = (0xFFFFFFFF << (32 - prefix)) & 0xFFFFFFFF
                    mask = '.'.join(str((mask_int >> (8 * i)) & 0xFF) for i in [3, 2, 1, 0])

                    score = 0
                    ip_parts = [int(x) for x in ip.split('.')]

                    # Best: same subnet as factory_ip (this is the camera's adapter)
                    if factory_ip:
                        try:
                            fip_parts = [int(x) for x in factory_ip.split('.')]
                            mask_parts = [int(x) for x in mask.split('.')]
                            if all((a & m) == (b & m) for a, b, m in zip(ip_parts, fip_parts, mask_parts)):
                                score += 100
                        except:
                            pass

                    # Penalize CGNAT / Tailscale (100.64.0.0/10)
                    if ip_parts[0] == 100 and 64 <= ip_parts[1] <= 127:
                        score -= 50

                    # Prefer private IPs
                    if ip_parts[0] == 192 and ip_parts[1] == 168:
                        score += 10
                    elif ip_parts[0] == 10:
                        score += 10
                    elif ip_parts[0] == 172 and 16 <= ip_parts[1] <= 31:
                        score += 10

                    candidates.append({'ip': ip, 'mask': mask, 'index': idx, 'score': score})

                if candidates:
                    candidates.sort(key=lambda c: c['score'], reverse=True)
                    best = candidates[0]
                    local_ip = best['ip']
                    subnet_mask = best['mask']
                    self._detected_iface_index = best['index']
        except:
            pass

        # --- Fallback methods if PowerShell failed ---

        if not local_ip:
            # Method A: socket trick (fast, needs internet route)
            try:
                s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
                s.settimeout(1)
                s.connect(("8.8.8.8", 80))
                candidate = s.getsockname()[0]
                s.close()
                # Reject CGNAT/Tailscale
                parts = [int(x) for x in candidate.split('.')]
                if not (parts[0] == 100 and 64 <= parts[1] <= 127):
                    local_ip = candidate
            except:
                pass

        if not local_ip:
            # Method B: hostname resolution (works on isolated networks)
            try:
                for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
                    ip = info[4][0]
                    if ip.startswith('127.') or ip.startswith('169.254.'):
                        continue
                    parts = [int(x) for x in ip.split('.')]
                    if parts[0] == 100 and 64 <= parts[1] <= 127:
                        continue
                    local_ip = ip
                    break
            except:
                pass

        if not local_ip:
            # Method C: regex-parse ipconfig output (locale-independent)
            local_ip, subnet_mask = self._parse_ipconfig_for_ip()

        # --- Step 2: Get subnet mask + interface index if not from PowerShell ---

        if local_ip and not subnet_mask:
            ps = self._query_interface_powershell(local_ip)
            if ps:
                subnet_mask = ps.get('mask')
                if self._detected_iface_index is None:
                    self._detected_iface_index = ps.get('index')

        if local_ip and not subnet_mask:
            _, subnet_mask = self._parse_ipconfig_for_ip(target_ip=local_ip)

        if not subnet_mask:
            subnet_mask = '255.255.255.0'

        if local_ip and self._detected_iface_index is None:
            self._detected_iface_index = self._find_interface_index(local_ip)

        # Cache local IP for route/mDNS methods
        if local_ip:
            self._detected_local_ip = local_ip

        # Calculate scan range
        if local_ip:
            parts = local_ip.split('.')
            suggested_range = f"{parts[0]}.{parts[1]}.{parts[2]}.1-254"

        return local_ip, subnet_mask, suggested_range

    def _parse_ipconfig_for_ip(self, target_ip=None):
        """Parse ipconfig output using regex only — works on any Windows locale.
        If target_ip given, finds its subnet mask.
        If target_ip is None, returns the first usable (IP, mask) pair."""
        import subprocess
        try:
            result = subprocess.run(['ipconfig'], capture_output=True, text=True,
                                   creationflags=subprocess.CREATE_NO_WINDOW)
            ip_re = re.compile(r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})')
            prev_ip = None
            for line in result.stdout.split('\n'):
                m = ipre.search(line)
                if not m:
                    # blank line or header — reset pair tracking
                    if not line.strip() or (line and not line[0].isspace()):
                        prev_ip = None
                    continue
                val = m.group(1)
                if val.startswith('255.'):
                    # This is a subnet mask — prev_ip is the matching address
                    if target_ip and prev_ip == target_ip:
                        return prev_ip, val
                    if not target_ip and prev_ip and not prev_ip.startswith('127.') and not prev_ip.startswith('169.254.'):
                        return prev_ip, val
                    prev_ip = None
                else:
                    prev_ip = val
        except:
            pass
        return None, None

    def _query_interface_powershell(self, target_ip):
        """Use PowerShell Get-NetIPAddress to get subnet mask + interface index.
        Returns dict with 'mask' and 'index', or None. Locale-independent (JSON)."""
        import subprocess
        try:
            cmd = (f"Get-NetIPAddress -IPAddress '{target_ip}' -ErrorAction SilentlyContinue | "
                   f"Select-Object InterfaceIndex,PrefixLength | ConvertTo-Json -Compress")
            result = subprocess.run(
                ['powershell', '-NoProfile', '-Command', cmd],
                capture_output=True, text=True, timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode == 0 and result.stdout.strip():
                data = json.loads(result.stdout)
                if isinstance(data, list):
                    data = data[0]
                prefix = int(data.get('PrefixLength', 24))
                mask_int = (0xFFFFFFFF << (32 - prefix)) & 0xFFFFFFFF
                mask = '.'.join(str((mask_int >> (8 * i)) & 0xFF) for i in [3, 2, 1, 0])
                return {'mask': mask, 'index': data.get('InterfaceIndex')}
        except:
            pass
        return None

    def _find_interface_index(self, target_ip):
        """Find the Windows interface index for a given IP address.
        Parses 'route print -4' which is locale-independent (uses numbers/IPs)."""
        import subprocess
        try:
            result = subprocess.run(
                ['route', 'print', '-4'],
                capture_output=True, text=True, timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW)
            output = result.stdout

            # Build idx→name map from Interface List section
            # Lines like: " 12...b8 a4 4f xx xx xx ......Intel(R) Ethernet"
            iface_indices = []
            in_iface_list = False
            for line in output.split('\n'):
                if '===' in line:
                    if in_iface_list:
                        in_iface_list = False
                    continue
                if 'Interface List' in line or line.strip().startswith('Interface'):
                    in_iface_list = True
                    continue
                if in_iface_list:
                    m = re.match(r'\s*(\d+)\.\.\.', line)
                    if m:
                        idx = int(m.group(1))
                        if idx != 1:  # Skip loopback (always index 1)
                            iface_indices.append(idx)

            # Try to match target_ip in routing table to find its interface
            # Route table has columns: Destination, Netmask, Gateway, Interface, Metric
            # A route with our IP as "Interface" tells us which adapter it's on
            for line in output.split('\n'):
                if target_ip not in line:
                    continue
                parts = line.split()
                # Find our target_ip as the Interface column (usually 4th)
                for i, p in enumerate(parts):
                    if p == target_ip and i >= 3:
                        # This is likely the Interface column
                        # We know this route belongs to our interface
                        # Return the first non-loopback index (most common case)
                        if iface_indices:
                            return iface_indices[0]

            # No route match — just return first non-loopback interface
            if iface_indices:
                return iface_indices[0]
        except:
            pass
        return None
    
    def background_scan(self, force=False, quiet=False):
        """Scan network for cameras. Runs on startup, periodically, and after operations."""
        if self.startup_scan_complete and not force and not quiet:
            return
        
        # Prevent overlapping scans
        if getattr(self, '_scan_running', False):
            return
        self._scan_running = True
        
        self.startup_scan_complete = True

        if not quiet:
            self.log("Scanning network for cameras...")

        def _scan_thread():
            self._run_scan(force, quiet)
        threading.Thread(target=_scan_thread, daemon=True).start()

    def _run_scan(self, force=False, quiet=False):
        """Scan logic — runs in a background thread to avoid blocking GUI."""
        from concurrent.futures import ThreadPoolExecutor, as_completed

        # =====================================================================
        # PHASE 1: DHCP + mDNS Discovery (finds link-local cameras)
        # =====================================================================
        discovered_cameras = {}  # key=MAC_no_colons → camera dict

        # 1a. DHCP snooping — fastest method: cameras broadcast DHCP DISCOVER
        #     with hostname, model, MAC. Works on any subnet (Layer 2 broadcast).
        if not quiet:
            self.log("  Phase 1a: DHCP snooping (listening for camera broadcasts)...")
        try:
            dhcp_results = AxisDHCPDiscovery.discover(timeout=4)
            for cam in dhcp_results:
                key = cam.get('mac', '').upper().replace(':', '')
                if key:
                    discovered_cameras[key] = cam
                    if not quiet:
                        self.log(f"    [DHCP] {cam.get('model', '?')} | MAC: {cam.get('mac', '')} | {cam.get('hostname', '')}")
        except Exception as e:
            if not quiet:
                self.log(f"    DHCP snooping failed: {e}")

        # 1b. mDNS — finds cameras with IPs (including 169.254.x.x link-local)
        if not quiet:
            self.log("  Phase 1b: mDNS discovery...")
        try:
            mdns_results = AxisMDNSDiscovery.discover(timeout=3)
            for cam in mdns_results:
                key = cam.get('mac', '').upper().replace(':', '') or cam.get('ip', '')
                if key:
                    if key in discovered_cameras:
                        # Merge mDNS data into DHCP result (mDNS has IP)
                        existing = discovered_cameras[key]
                        for field in ['ip', 'ipv6', 'model', 'hostname']:
                            if cam.get(field) and not existing.get(field):
                                existing[field] = cam[field]
                    else:
                        discovered_cameras[key] = cam
                    if not quiet:
                        self.log(f"    [mDNS] {cam.get('model', '?')} @ {cam.get('ip', '?')} (SN: {cam.get('serial', '')})")
            if not quiet and not mdns_results:
                self.log("    No cameras found via mDNS")
        except Exception as e:
            if not quiet:
                self.log(f"    mDNS failed: {e}")

        if not quiet and not discovered_cameras:
            self.log("    No cameras found in Phase 1")

        # If we found link-local cameras, try targeted mDNS to resolve IPs
        # NOTE: do NOT add link-local route here — that modifies network adapters
        # and should only happen during programming when the user explicitly chooses
        has_linklocal = any(c.get('ip', '').startswith('169.254.') for c in discovered_cameras.values())
        has_dhcp_only = any(not c.get('ip') for c in discovered_cameras.values())
        if has_linklocal or has_dhcp_only:
            # Phase 1c: Targeted mDNS on correct interface for DHCP-only cameras
            # This resolves IPs that regular mDNS missed (multi-adapter issue)
            if has_dhcp_only and getattr(self, '_linklocal_route_active', False):
                if not quiet:
                    self.log("  Phase 1c: Resolving IPs via targeted mDNS...")
                try:
                    resolved_cams = self._resolve_linklocal_cameras(timeout=4)
                    for cam in resolved_cams:
                        key = cam.get('mac', '').upper().replace(':', '') or cam.get('ip', '')
                        if key and key in discovered_cameras:
                            # Fill in the IP from targeted mDNS
                            existing = discovered_cameras[key]
                            for field in ['ip', 'ipv6', 'model', 'hostname']:
                                if cam.get(field) and not existing.get(field):
                                    existing[field] = cam[field]
                            if not quiet:
                                self.log(f"    [targeted mDNS] {existing.get('model', '?')} @ {existing.get('ip', '?')}")
                        elif key:
                            discovered_cameras[key] = cam
                            if not quiet:
                                self.log(f"    [targeted mDNS] {cam.get('model', '?')} @ {cam.get('ip', '?')}")
                except Exception as e:
                    if not quiet:
                        self.log(f"    Targeted mDNS failed: {e}")

        # Get network info from THIS PC
        local_ip, subnet_mask, _ = self.get_local_network_info()
        if not local_ip:
            if not quiet:
                self.log("Could not detect network")
            # Even if no local network, mDNS may have found cameras
            if discovered_cameras:
                found = list(discovered_cameras.values())
                self.root.after(0, lambda: self._on_scan_complete(found, quiet))
            else:
                self._scan_running = False
                self.root.after(0, self._reset_rescan_btn)
            return

        parts = local_ip.split('.')

        if not quiet:
            self.log(f"  Phase 2: HTTP scan | Your IP: {local_ip} | Subnet: {subnet_mask}")
        
        # Build scan ranges based on subnet
        ips = []
        if subnet_mask == '255.255.0.0':
            # /16 network: scan local /24 + .0.x (factory defaults) + common factory subnets
            for i in range(1, 255):
                ips.append(f"{parts[0]}.{parts[1]}.{parts[2]}.{i}")  # local /24
                ips.append(f"{parts[0]}.{parts[1]}.0.{i}")           # x.x.0.x subnet
            # Also scan 192.168.0.x and 192.168.1.x (common factory defaults)
            if parts[0] != '192' or parts[1] != '168':
                for i in range(1, 255):
                    ips.append(f"192.168.0.{i}")
                    ips.append(f"192.168.1.{i}")
            ips = list(set(ips))
        else:
            ips = [f"{parts[0]}.{parts[1]}.{parts[2]}.{i}" for i in range(1, 255)]
        
        # Add any IPs discovered via mDNS (including link-local 169.254.x.x)
        # so we can try to get network config from them
        for key, cam in discovered_cameras.items():
            mdns_ip = cam.get('ip', '')
            if mdns_ip and mdns_ip not in ips:
                ips.append(mdns_ip)
        
        # Collect passwords to try for network config
        passwords_to_try = self.password_data.get_all()[:]
        user_count = len(passwords_to_try)
        for p in ['pass', 'admin', 'root', 'Admin123', 'password', 'service']:
            if p not in passwords_to_try:
                passwords_to_try.append(p)
        default_count = len(passwords_to_try) - user_count
        
        if not quiet:
            if user_count:
                self.log(f"Scanning {len(ips)} IPs ({user_count} saved + {default_count} default passwords)")
            else:
                self.log(f"Scanning {len(ips)} IPs ({default_count} default passwords)")
        
        found = []
        
        def parse_network_params(text, cam):
            """Parse param.cgi Network output into cam dict"""
            for line in text.split('\n'):
                line = line.strip()
                if '=' not in line:
                    continue
                key, val = line.split('=', 1)
                val = val.strip()
                if 'SubnetMask' in key and val:
                    cam['subnet'] = val
                elif 'DefaultRouter' in key and val:
                    cam['gateway'] = val
                elif 'BootProto' in key and val:
                    cam['dhcp'] = 'Yes' if val.lower() == 'dhcp' else 'No'
        
        def check_camera(ip):
            """Detect camera of the selected brand and get its network config"""
            # Step 1: Protocol-based detection (no auth)
            cam = self.protocol.get_discovery_info(ip, timeout=2)

            # Step 1b: If auth required, try known passwords for more details
            if cam and cam.get('model') == '(Auth Required)' and passwords_to_try:
                username = self.protocol.DEFAULT_USER
                max_attempts = 2 if self.protocol.BRAND_KEY == 'hanwha' else 8
                for pwd in passwords_to_try[:max_attempts]:
                    try:
                        if self.protocol.test_password(ip, username, pwd):
                            model = self.protocol.get_model_noauth(ip)
                            if model:
                                cam['model'] = model
                            serial = self.protocol.get_serial(ip, pwd)
                            if serial and serial != 'UNKNOWN':
                                cam['serial'] = serial
                            cam['_auth_pwd'] = pwd
                            break
                    except LockoutError:
                        break
                    except:
                        continue

            if not cam:
                return None

            # Step 2: Axis-specific network config via param.cgi
            if self.protocol.BRAND_KEY == 'axis':
                try:
                    r = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                        params={"action": "list", "group": "Network"},
                        timeout=1.5)
                    if r.status_code == 200:
                        if 'SubnetMask' in r.text:
                            parse_network_params(r.text, cam)
                        for line in r.text.split('\n'):
                            if 'root.Network.HostName=' in line:
                                hostname = line.split('=', 1)[1].strip()
                                if hostname:
                                    cam['hostname'] = hostname
                                    cam['original_hostname'] = hostname
                                break
                except:
                    pass

                if not cam.get('subnet') or not cam.get('hostname'):
                    pwds_to_check = []
                    if cam.get('_auth_pwd'):
                        pwds_to_check.append(cam['_auth_pwd'])
                    for p in passwords_to_try[:8]:
                        if p not in pwds_to_check:
                            pwds_to_check.append(p)

                    for pwd in pwds_to_check:
                        try:
                            r = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                                params={"action": "list", "group": "Network"},
                                auth=HTTPDigestAuth("root", pwd), timeout=1.5)
                            if r.status_code == 200:
                                if 'SubnetMask' in r.text and not cam.get('subnet'):
                                    parse_network_params(r.text, cam)
                                if not cam.get('hostname'):
                                    for line in r.text.split('\n'):
                                        if 'root.Network.HostName=' in line:
                                            hostname = line.split('=', 1)[1].strip()
                                            if hostname:
                                                cam['hostname'] = hostname
                                            break
                                break
                            elif r.status_code == 401:
                                continue
                        except:
                            break

            # Clean up internal tracking field
            if '_auth_pwd' in cam:
                del cam['_auth_pwd']

            # Set camera name from hostname
            cam['name'] = cam.get('hostname', cam['ip'])

            # MAC from serial
            if cam.get('serial') and len(cam['serial']) == 12 and not cam.get('mac'):
                cam['mac'] = ':'.join(cam['serial'][i:i+2] for i in range(0, 12, 2))

            return cam
        
        def do_scan():
            with ThreadPoolExecutor(max_workers=100) as executor:
                futures = {executor.submit(check_camera, ip): ip for ip in ips}
                for future in as_completed(futures):
                    try:
                        result = future.result()
                        if result:
                            found.append(result)
                    except:
                        pass
            
            # Merge mDNS results with HTTP scan results
            # mDNS provides model/serial/MAC, HTTP provides gateway/subnet/hostname
            merged = {}
            
            # First, add all mDNS cameras (these might be link-local)
            for key, cam in discovered_cameras.items():
                merged[key] = cam.copy()
            
            # Then merge/add HTTP scan results
            for cam in found:
                cam_mac = cam.get('mac', '').upper().replace(':', '').replace('-', '')
                cam_serial = cam.get('serial', '').upper()
                cam_ip = cam.get('ip', '')
                
                # Find matching mDNS entry
                match_key = None
                for key, mcam in discovered_cameras.items():
                    mcam_mac = mcam.get('mac', '').upper().replace(':', '').replace('-', '')
                    mcam_serial = mcam.get('serial', '').upper()
                    mcam_ip = mcam.get('ip', '')
                    
                    if (cam_mac and mcam_mac and cam_mac == mcam_mac) or \
                       (cam_serial and mcam_serial and cam_serial == mcam_serial) or \
                       (cam_ip and mcam_ip and cam_ip == mcam_ip):
                        match_key = key
                        break
                
                if match_key:
                    # Merge: HTTP provides network config, mDNS provides model/serial
                    existing = merged[match_key]
                    for field in ['ip', 'gateway', 'subnet', 'dhcp', 'hostname', 'original_hostname', 'brand']:
                        if cam.get(field) and not existing.get(field):
                            existing[field] = cam[field]
                    # HTTP might have better model info if mDNS didn't get it
                    if cam.get('model') and (not existing.get('model') or existing.get('model') == '(Auth Required)'):
                        existing['model'] = cam['model']
                else:
                    # New camera only found via HTTP
                    key = cam_mac or cam_serial or cam_ip
                    if key:
                        merged[key] = cam
            
            # Deduplicate by MAC — same physical camera should appear once
            seen_macs = set()
            all_cameras = []
            for cam in merged.values():
                mac = cam.get('mac', '').upper().replace(':', '').replace('-', '')
                if mac:
                    if mac in seen_macs:
                        continue
                    seen_macs.add(mac)
                all_cameras.append(cam)
            self.root.after(0, lambda: self._on_scan_complete(all_cameras, quiet))
        
        threading.Thread(target=do_scan, daemon=True).start()
    
    def _on_scan_complete(self, cameras, quiet=False):
        """Handle scan completion - update UI with found cameras"""
        self._scan_running = False
        self._reset_rescan_btn()
        if cameras:
            if not quiet:
                self.log(f"✓ Found {len(cameras)} camera(s) on network!")
            
            # Populate discovered tab
            self.discovered_cameras = cameras
            self.refresh_discovered_list()
            
            needs_auth = []
            updated = 0
            for cam in cameras:
                gateway = cam.get('gateway', '')
                subnet = cam.get('subnet', '')
                model = cam.get('model', '')
                
                missing = []
                if not gateway:
                    missing.append("gateway")
                if not subnet:
                    missing.append("subnet")
                if model == '(Auth Required)':
                    missing.append("model/serial")
                
                # Mark link-local cameras specially
                is_linklocal = cam.get('ip', '').startswith('169.254.')
                linklocal_marker = " [LINK-LOCAL]" if is_linklocal else ""
                
                if not quiet:
                    self.log(f"  {model or '?'} @ {cam['ip']}{linklocal_marker}"
                            + (f" | {cam.get('hostname', '')}" if cam.get('hostname') else "")
                            + (f" | SN: {cam.get('serial', '')}" if cam.get('serial') else "")
                            + (f" | GW: {gateway}" if gateway else "")
                            + (f" | (needs auth for {', '.join(missing)})" if missing and not is_linklocal else ""))
                
                # Also enrich any matching cameras already in the list
                cam_mac_norm = cam.get('mac', '').upper().replace(':', '').replace('-', '').strip()
                for i, existing in enumerate(self.camera_data.cameras):
                    ex_mac_norm = existing.get('mac', '').upper().replace(':', '').replace('-', '').strip()
                    if (existing.get('ip') == cam.get('ip') or 
                        (cam.get('serial') and existing.get('serial') == cam.get('serial')) or
                        (cam_mac_norm and ex_mac_norm and cam_mac_norm == ex_mac_norm)):
                        for field in ['model', 'serial', 'mac', 'gateway', 'subnet', 'dhcp', 'hostname', 'brand']:
                            new_val = cam.get(field, '')
                            old_val = existing.get(field, '')
                            if new_val and (not old_val or old_val == '(Auth Required)'):
                                existing[field] = new_val
                        updated += 1
                        break
                
                if missing and not is_linklocal:
                    needs_auth.append(cam)
            
            if updated:
                self.camera_data.save()
                self.refresh_camera_list()
                if not quiet:
                    self.log(f"Enriched {updated} camera(s) in Camera List with scan data")
            
            if not quiet:
                self.log(f"See Discovered tab for all {len(cameras)} camera(s)")
            
            # Show or hide password bar on Discovered tab
            if needs_auth and not quiet:
                self.notebook.select(self.discovered_tab)
                self.show_discovered_password_bar(needs_auth)
                self.log(f"⚠ {len(needs_auth)} camera(s) need password — see Discovered tab")
            else:
                # Password list resolved everything (or nothing needed auth) — clear bar
                for w in self.discovered_password_bar.winfo_children():
                    w.destroy()
        else:
            if not quiet:
                self.log("No cameras found on local network")

        # Schedule next periodic scan (90 seconds)
        self._schedule_periodic_scan()
    
    def _schedule_periodic_scan(self):
        """Schedule the next quiet background rescan"""
        if hasattr(self, '_periodic_scan_id') and self._periodic_scan_id:
            self.root.after_cancel(self._periodic_scan_id)
        self._periodic_scan_id = self.root.after(90000, lambda: self.background_scan(force=True, quiet=True))
        # Start countdown
        self._rescan_seconds_left = 90
        self._tick_rescan_countdown()
    
    def _tick_rescan_countdown(self):
        """Update the rescan countdown display every second"""
        if hasattr(self, '_countdown_tick_id') and self._countdown_tick_id:
            self.root.after_cancel(self._countdown_tick_id)
            self._countdown_tick_id = None
        if self._rescan_seconds_left > 0:
            self.rescan_countdown_var.set(f"Next rescan in {self._rescan_seconds_left}s")
            self._rescan_seconds_left -= 1
            self._countdown_tick_id = self.root.after(1000, self._tick_rescan_countdown)
        else:
            self.rescan_countdown_var.set("Rescanning...")
            self._countdown_tick_id = None
    
    def rescan_after_operation(self):
        """Trigger a quiet rescan shortly after an operation completes"""
        if hasattr(self, '_post_op_scan_id') and self._post_op_scan_id:
            self.root.after_cancel(self._post_op_scan_id)
        self._post_op_scan_id = self.root.after(10000, lambda: self.background_scan(force=True, quiet=True))
    
    def show_discovered_password_bar(self, cameras_needing_auth):
        """Show a bar on the Discovered tab to enter password for network info"""
        
        # Clear any existing bar content
        for w in self.discovered_password_bar.winfo_children():
            w.destroy()
        
        bar = tk.Frame(self.discovered_password_bar, bg='#FFF3CD', relief='solid', bd=1)
        bar.pack(fill=tk.X, pady=(0, 5))
        
        inner = tk.Frame(bar, bg='#FFF3CD')
        inner.pack(fill=tk.X, padx=10, pady=8)
        
        tk.Label(inner, text=f"⚠ {len(cameras_needing_auth)} camera(s) need authentication for full info", 
                bg='#FFF3CD', font=('Helvetica', 10, 'bold')).pack(side=tk.LEFT)
        
        tk.Label(inner, text="  Password:", bg='#FFF3CD').pack(side=tk.LEFT, padx=(15, 0))
        
        pwd_var = tk.StringVar()
        pwd_entry = tk.Entry(inner, textvariable=pwd_var, show='•', width=20, font=('Courier', 10))
        pwd_entry.pack(side=tk.LEFT, padx=(5, 5))
        
        def try_password():
            pwd = pwd_var.get().strip()
            if not pwd:
                return
            
            try_btn.config(state='disabled', text='Checking...')
            bar.update()
            
            def do_auth():
                updated = 0
                for cam in cameras_needing_auth:
                    ip = cam.get('ip', '')
                    if not ip:
                        continue
                    try:
                        auth = HTTPDigestAuth("root", pwd)
                        got_something = False
                        
                        # Get model if missing
                        if not cam.get('model') or cam['model'] == '(Auth Required)':
                            try:
                                r = requests.post(f"http://{ip}/axis-cgi/basicdeviceinfo.cgi",
                                    json={"apiVersion": "1.0", "method": "getAllProperties"},
                                    auth=auth, timeout=2)
                                if r.status_code == 200:
                                    data = r.json()
                                    if 'data' in data and 'propertyList' in data['data']:
                                        props = data['data']['propertyList']
                                        cam['model'] = props.get('ProdFullName', props.get('ProdShortName', ''))
                                        if not cam.get('serial'):
                                            cam['serial'] = props.get('SerialNumber', '')
                                        got_something = True
                            except:
                                pass
                        
                        # Get network info - use full Network group
                        r = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                            params={"action": "list", "group": "Network"},
                            auth=auth, timeout=3)
                        if r.status_code == 200:
                            self.root.after(0, lambda t=r.text[:500], i=ip: self.log(f"  [{i}] Network params: {t}"))
                            for line in r.text.split('\n'):
                                line = line.strip()
                                if '=' not in line:
                                    continue
                                key, val = line.split('=', 1)
                                val = val.strip()
                                if 'DefaultRouter' in key and val and not cam.get('gateway'):
                                    cam['gateway'] = val
                                    got_something = True
                                elif 'SubnetMask' in key and val and not cam.get('subnet'):
                                    cam['subnet'] = val
                                    got_something = True
                                elif key.endswith('.BootProto') and val:
                                    cam['dhcp'] = 'Yes' if val.lower() == 'dhcp' else 'No'
                                elif 'root.Network.HostName' == key and val and not cam.get('hostname'):
                                    cam['hostname'] = val
                                    got_something = True
                        else:
                            self.root.after(0, lambda s=r.status_code, i=ip: self.log(f"  [{i}] param.cgi returned {s}"))
                        
                        # Derive MAC from serial
                        if cam.get('serial') and len(cam['serial']) == 12 and not cam.get('mac'):
                            cam['mac'] = ':'.join(cam['serial'][j:j+2] for j in range(0, 12, 2))
                        
                        # Update name from hostname
                        if cam.get('hostname'):
                            cam['name'] = cam['hostname']
                        
                        if got_something:
                            updated += 1
                        
                        self.root.after(0, lambda c=dict(cam): self.log(f"  Result: gw={c.get('gateway','')}, subnet={c.get('subnet','')}, model={c.get('model','')}"))
                    except Exception as e:
                        self.root.after(0, lambda err=str(e), i=ip: self.log(f"  [{i}] Error: {err}"))
                
                self.root.after(0, lambda: on_auth_complete(updated, pwd))
            
            def on_auth_complete(count, password):
                if count > 0:
                    # Refresh discovered list with new data
                    self.refresh_discovered_list()
                    
                    # Also enrich any matching cameras in the Camera List
                    enriched = 0
                    for cam in cameras_needing_auth:
                        cam_mac_norm = cam.get('mac', '').upper().replace(':', '').replace('-', '').strip()
                        for existing in self.camera_data.cameras:
                            ex_mac_norm = existing.get('mac', '').upper().replace(':', '').replace('-', '').strip()
                            if (existing.get('ip') == cam.get('ip') or
                                (cam_mac_norm and ex_mac_norm and cam_mac_norm == ex_mac_norm)):
                                for field in ['model', 'serial', 'mac', 'gateway', 'subnet', 'dhcp', 'hostname', 'brand']:
                                    new_val = cam.get(field, '')
                                    old_val = existing.get(field, '')
                                    if new_val and (not old_val or old_val == '(Auth Required)'):
                                        existing[field] = new_val
                                enriched += 1
                                break
                        self.camera_data.save()
                        self.refresh_camera_list()
                    
                    self.log(f"✓ Updated {count}/{len(cameras_needing_auth)} discovered camera(s) with network info")
                    
                    # Save password to password list if not already there
                    existing_pwds = self.password_data.get_all()
                    if password not in existing_pwds:
                        self.password_data.passwords.append(password)
                        self.password_data.save()
                        self.log(f"  Saved password to password list")
                    
                    dismiss_bar()
                else:
                    self.log("✗ Password didn't work for any cameras")
                    try_btn.config(state='normal', text='Try Password')
                    pwd_entry.delete(0, tk.END)
                    pwd_entry.focus_set()
            
            threading.Thread(target=do_auth, daemon=True).start()
        
        def dismiss_bar():
            for w in self.discovered_password_bar.winfo_children():
                w.destroy()
        
        try_btn = tk.Button(inner, text="Try Password", command=try_password,
                           bg='#4CAF50', fg='white', font=('Helvetica', 9, 'bold'),
                           padx=10, cursor='hand2')
        try_btn.pack(side=tk.LEFT, padx=(5, 5))
        
        # Dismiss button
        tk.Button(inner, text="✕", command=dismiss_bar, bg='#FFF3CD', 
                 relief='flat', font=('Helvetica', 10), cursor='hand2').pack(side=tk.RIGHT)
        
        # Enter key triggers try
        pwd_entry.bind('<Return>', lambda e: try_password())
        pwd_entry.focus_set()
    
    def discover_cameras(self):
        """Discover Axis cameras on the network via IP scan"""
        self.scan_ip_range()
    
    def scan_ip_range(self):
        """Scan a range of IPs for Axis cameras with parallel scanning"""
        from concurrent.futures import ThreadPoolExecutor, as_completed
        
        # Get local network info
        local_ip, subnet_mask, suggested_range = self.get_local_network_info()
        
        # Show scan configuration dialog
        scan_dialog = tk.Toplevel(self.root)
        scan_dialog.title("Network Scan Configuration")
        scan_dialog.geometry("600x560")
        scan_dialog.transient(self.root)
        scan_dialog.grab_set()
        
        # Main frame with buttons pinned at bottom
        outer = ttk.Frame(scan_dialog)
        outer.pack(fill=tk.BOTH, expand=True)
        
        frame = ttk.Frame(outer, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="📡 Scan Network for Axis Cameras", 
                 font=('Helvetica', 14, 'bold')).pack(anchor=tk.W)
        
        # Network info display
        info_frame = ttk.LabelFrame(frame, text="Your Network", padding="8")
        info_frame.pack(fill=tk.X, pady=(8, 8))
        
        if local_ip:
            ttk.Label(info_frame, text=f"IP Address: {local_ip}", 
                     font=('Courier', 10)).grid(row=0, column=0, sticky='w')
        if subnet_mask:
            ttk.Label(info_frame, text=f"Subnet Mask: {subnet_mask}", 
                     font=('Courier', 10)).grid(row=0, column=1, sticky='w', padx=(20, 0))
            
            # Explain subnet size
            if subnet_mask == '255.255.255.0':
                size_text = "/24 (254 hosts)"
            elif subnet_mask == '255.255.0.0':
                size_text = "/16 (65,534 hosts)"
            elif subnet_mask == '255.0.0.0':
                size_text = "/8 (16 million hosts)"
            else:
                size_text = ""
            if size_text:
                ttk.Label(info_frame, text=size_text, foreground='blue').grid(row=0, column=2, padx=(10, 0))
        
        ttk.Label(frame, text="Enter IP range to scan:", font=('Helvetica', 10)).pack(anchor=tk.W, pady=(5, 0))
        
        # Range input
        range_var = tk.StringVar(value=suggested_range)
        range_entry = ttk.Entry(frame, textvariable=range_var, width=45, font=('Courier', 11))
        range_entry.pack(fill=tk.X, pady=(5, 8))
        
        # Build dynamic examples based on detected network
        examples_frame = ttk.LabelFrame(frame, text="Quick Presets (click to use)", padding="5")
        examples_frame.pack(fill=tk.X, pady=(0, 8))
        
        examples = []
        if local_ip:
            parts = local_ip.split('.')
            examples.append((f"Your /24 subnet:", f"{parts[0]}.{parts[1]}.{parts[2]}.1-254", "~254 IPs, ~10 sec"))
            
            if subnet_mask == '255.255.0.0':
                examples.append((f"Full /16 network:", f"{parts[0]}.{parts[1]}.0.1-{parts[0]}.{parts[1]}.255.254", "~65K IPs, ~20 min"))
                examples.append((f"First few /24s:", f"{parts[0]}.{parts[1]}.0.1-{parts[0]}.{parts[1]}.3.254", "~1K IPs, ~30 sec"))
                examples.append((f"x.x.0.x range:", f"{parts[0]}.{parts[1]}.0.1-254", "~254 IPs, ~10 sec"))
            elif subnet_mask == '255.0.0.0':
                examples.append((f"Your /16 block:", f"{parts[0]}.{parts[1]}.0.1-{parts[0]}.{parts[1]}.255.254", "~65K IPs"))
        
        examples.append(("Custom small range:", "10.0.0.1-50", "50 IPs"))
        
        for i, (label, example, note) in enumerate(examples):
            row = ttk.Frame(examples_frame)
            row.pack(fill=tk.X, pady=1)
            ttk.Label(row, text=label, width=18).pack(side=tk.LEFT)
            btn = ttk.Button(row, text=example, width=30,
                           command=lambda e=example: range_var.set(e))
            btn.pack(side=tk.LEFT, padx=(5, 10))
            ttk.Label(row, text=note, foreground='gray', font=('Helvetica', 8)).pack(side=tk.LEFT)
        
        # Thread count
        thread_frame = ttk.Frame(frame)
        thread_frame.pack(fill=tk.X, pady=(5, 0))
        ttk.Label(thread_frame, text="Parallel threads:").pack(side=tk.LEFT)
        thread_var = tk.IntVar(value=50)
        thread_spin = ttk.Spinbox(thread_frame, from_=10, to=200, width=5, textvariable=thread_var)
        thread_spin.pack(side=tk.LEFT, padx=(5, 10))
        ttk.Label(thread_frame, text="(higher = faster but more network load)", 
                 foreground='gray').pack(side=tk.LEFT)
        
        result = {'start': False}
        
        def start_scan():
            result['start'] = True
            result['range'] = range_var.get()
            result['threads'] = thread_var.get()
            scan_dialog.destroy()
        
        def cancel():
            scan_dialog.destroy()
        
        # Buttons pinned at bottom with separator
        ttk.Separator(outer, orient='horizontal').pack(fill=tk.X, padx=10)
        
        btn_frame = ttk.Frame(outer, padding="10")
        btn_frame.pack(fill=tk.X)
        
        start_btn = tk.Button(btn_frame, text="🔍 START SCAN", command=start_scan,
                             bg='#4CAF50', fg='white', font=('Helvetica', 12, 'bold'),
                             padx=20, pady=5, cursor='hand2')
        start_btn.pack(side=tk.RIGHT, padx=5)
        
        cancel_btn = tk.Button(btn_frame, text="Cancel", command=cancel,
                              font=('Helvetica', 10), padx=10)
        cancel_btn.pack(side=tk.RIGHT, padx=5)
        
        # Enter key starts scan
        scan_dialog.bind('<Return>', lambda e: start_scan())
        range_entry.focus_set()
        
        scan_dialog.wait_window()
        
        if not result.get('start'):
            return
        
        ip_range = result['range']
        num_threads = result['threads']
        
        # Parse range - support multiple formats
        ips = []
        try:
            if '-' in ip_range:
                # Check if it's full IP-to-IP format (10.0.0.1-10.0.255.254)
                if ip_range.count('.') > 3:
                    # Full range: 10.0.0.1-10.0.255.254
                    start_ip, end_ip = ip_range.split('-')
                    start_parts = list(map(int, start_ip.split('.')))
                    end_parts = list(map(int, end_ip.split('.')))
                    
                    # Generate all IPs in range
                    current = start_parts.copy()
                    while current <= end_parts:
                        ips.append('.'.join(map(str, current)))
                        # Increment
                        current[3] += 1
                        for i in range(3, 0, -1):
                            if current[i] > 255:
                                current[i] = 0
                                current[i-1] += 1
                        if current[0] > 255:
                            break
                else:
                    # Short format: 10.0.7.1-254
                    parts = ip_range.split('-')
                    base_ip = parts[0].rsplit('.', 1)[0]
                    start = int(parts[0].rsplit('.', 1)[1])
                    end = int(parts[1])
                    ips = [f"{base_ip}.{i}" for i in range(start, end + 1)]
            else:
                ips = [ip_range]
        except Exception as e:
            messagebox.showerror("Invalid Range", f"Could not parse IP range: {e}")
            return
        
        if len(ips) > 10000:
            if not messagebox.askyesno("Large Scan", 
                f"This will scan {len(ips)} IP addresses.\n"
                f"With {num_threads} threads, this may take several minutes.\n\n"
                "Continue?"):
                return
        
        # Show progress
        progress = tk.Toplevel(self.root)
        progress.title("Scanning Network...")
        progress.geometry("500x220")
        progress.transient(self.root)
        
        frame = ttk.Frame(progress, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text=f"Scanning {len(ips)} IP addresses ({num_threads} threads)...", 
                 font=('Helvetica', 12)).pack(pady=(0, 10))
        
        progress_bar = ttk.Progressbar(frame, mode='determinate', length=350, maximum=len(ips))
        progress_bar.pack(pady=10)
        
        status_var = tk.StringVar(value="Starting scan...")
        ttk.Label(frame, textvariable=status_var).pack()
        
        found_var = tk.StringVar(value="Found: 0 cameras")
        ttk.Label(frame, textvariable=found_var, font=('Helvetica', 10, 'bold'), 
                 foreground='green').pack(pady=(5, 0))
        
        cancel_flag = [False]
        
        def cancel_scan():
            cancel_flag[0] = True
            progress.destroy()
        
        ttk.Button(frame, text="Cancel", command=cancel_scan).pack(pady=(10, 0))
        
        found = []
        scanned_count = [0]
        
        def get_camera_info(ip):
            """Get as much info as possible from a camera without auth"""
            if cancel_flag[0]:
                return None
                
            cam = {'ip': ip, 'name': ip}
            
            # Try basicdeviceinfo via getAllUnrestrictedProperties — works
            # without auth on factory AND configured cameras
            try:
                r = requests.post(f"http://{ip}/axis-cgi/basicdeviceinfo.cgi",
                    json={"apiVersion": "1.0", "method": "getAllUnrestrictedProperties"},
                    timeout=1.5)
                if r.status_code == 200:
                    data = r.json()
                    if 'data' in data and 'propertyList' in data['data']:
                        props = data['data']['propertyList']
                        cam['model'] = props.get('ProdFullName', props.get('ProdShortName', ''))
                        cam['serial'] = props.get('SerialNumber', '')
                        if cam['serial'] and len(cam['serial']) == 12:
                            cam['mac'] = ':'.join(cam['serial'][i:i+2] for i in range(0, 12, 2))
                        return cam
            except:
                pass
            
            # Fallback: try param.cgi Brand group
            try:
                r = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                    params={"action": "list", "group": "Brand"},
                    timeout=1.5)
                if r.status_code == 200:
                    for line in r.text.split('\n'):
                        if 'Brand.ProdFullName=' in line:
                            cam['model'] = line.split('=')[1].strip()
                        elif 'Brand.ProdShortName=' in line and not cam.get('model'):
                            cam['model'] = line.split('=')[1].strip()
                    # Name stays as-is (set elsewhere from serial)
                    return cam
                elif r.status_code == 401:
                    cam['model'] = '(Auth Required)'
                    return cam
            except:
                pass
            
            # Last resort: check if it responds at all to Axis endpoints
            try:
                r = requests.get(f"http://{ip}/axis-cgi/jpg/image.cgi", timeout=1)
                if r.status_code == 401:
                    cam['model'] = '(Auth Required)'
                    return cam
            except:
                pass
            
            return None
        
        def do_scan():
            with ThreadPoolExecutor(max_workers=num_threads) as executor:
                futures = {executor.submit(get_camera_info, ip): ip for ip in ips}
                
                for future in as_completed(futures):
                    if cancel_flag[0]:
                        break
                    
                    scanned_count[0] += 1
                    ip = futures[future]
                    
                    try:
                        cam = future.result()
                        if cam:
                            found.append(cam)
                            if progress.winfo_exists():
                                progress.after(0, lambda: found_var.set(f"Found: {len(found)} cameras"))
                    except:
                        pass
                    
                    if progress.winfo_exists():
                        progress.after(0, lambda v=scanned_count[0]: progress_bar.configure(value=v))
                        if scanned_count[0] % 10 == 0:
                            progress.after(0, lambda s=f"Scanned {scanned_count[0]}/{len(ips)} IPs...": status_var.set(s))
            
            if progress.winfo_exists():
                progress.after(0, on_scan_complete)
        
        def on_scan_complete():
            progress.destroy()
            if found:
                dialog = DiscoveryResultsDialog(self.root, found, self.settings)
                if dialog.result:
                    # Add to Discovered tab
                    self.discovered_cameras = dialog.result
                    self.refresh_discovered_list()
                    
                    # Enrich matching cameras in Camera List
                    enriched = 0
                    for cam in dialog.result:
                        cam_mac_norm = cam.get('mac', '').upper().replace(':', '').replace('-', '').strip()
                        for existing in self.camera_data.cameras:
                            ex_mac_norm = existing.get('mac', '').upper().replace(':', '').replace('-', '').strip()
                            if (existing.get('ip') == cam.get('ip') or
                                (ex_mac_norm and cam_mac_norm and ex_mac_norm == cam_mac_norm) or
                                (existing.get('serial') and existing.get('serial') == cam.get('serial'))):
                                for field in ['model', 'serial', 'mac', 'gateway', 'subnet', 'dhcp', 'hostname', 'brand']:
                                    new_val = cam.get(field, '')
                                    old_val = existing.get(field, '')
                                    if new_val and (not old_val or old_val == '(Auth Required)'):
                                        existing[field] = new_val
                                enriched += 1
                                break
                    
                    if enriched:
                        self.camera_data.save()
                        self.refresh_camera_list()
                        self.log(f"Enriched {enriched} camera(s) in Camera List")
                    
                    self.log(f"Discovered {len(dialog.result)} camera(s) — see Discovered tab")
            else:
                messagebox.showinfo("Scan Complete", 
                    f"Scanned {len(ips)} addresses.\n"
                    "No cameras found in the specified range.")
        
        threading.Thread(target=do_scan, daemon=True).start()
    
    def export_cameras(self):
        filepath = filedialog.asksaveasfilename(
            title="Export Cameras",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if filepath:
            try:
                self.camera_data.export_to_csv(filepath)
                self.log(f"Exported cameras to {filepath}")
                messagebox.showinfo("Export Complete", f"Exported to {filepath}")
            except Exception as e:
                messagebox.showerror("Export Error", str(e))
    
    def _build_triplett_csv_lines(self):
        """Build Triplett-format CSV lines from camera list. Returns list of strings."""
        cameras = self.camera_data.get_all()
        if not cameras:
            return []
        lines = []
        for i, cam in enumerate(cameras):
            num = cam.get('number', '') or str(i + 1)
            ip = cam.get('new_ip', '') or cam.get('ip', '')
            gateway = cam.get('gateway', '')
            subnet = cam.get('subnet', '255.255.255.0')
            model = cam.get('model', '')
            if model == '(Auth Required)':
                model = ''
            hostname = cam.get('hostname', '') or ''
            lines.append(f'{num},{ip},{gateway},{subnet},{model},{hostname}')
        return lines
    
    def _get_triplett_address(self):
        """Parse and validate the Triplett address field. Returns (ip, port) or None."""
        raw = self.triplett_addr_var.get().strip()
        if not raw:
            messagebox.showwarning("Missing Address", 
                "Type the address shown on the Triplett screen.\n\nExample: ftp://192.168.1.50:2121")
            return None
        
        # Strip protocol prefix
        cleaned = re.sub(r'^(ftp|http|https)://', '', raw)
        
        if ':' not in cleaned:
            messagebox.showwarning("Missing Port", 
                f"Include the port number — type the full address shown on the Triplett.\n\n"
                f"Example: ftp://{cleaned}:2121")
            return None
        
        parts = cleaned.rsplit(':', 1)
        ip = parts[0].strip()
        try:
            port = int(parts[1].strip().rstrip('/'))
        except ValueError:
            messagebox.showerror("Invalid Port", 
                f"Port '{parts[1].strip()}' is not a number.\n\nType exactly what the Triplett shows.")
            return None
        
        octets = ip.split('.')
        if len(octets) != 4:
            messagebox.showerror("Invalid Address", 
                f"'{ip}' doesn't look like an IP address.\n\nType exactly what the Triplett shows.")
            return None
        
        # Remember working address
        self.settings.set('general', 'android_ip', raw)
        return (ip, port)
    
    def _ftp_send_data(self, ip, port, filename, data_bytes, btn, orig_text):
        """FTP push bytes to Triplett. Runs in thread."""
        def do_ftp():
            try:
                self.root.after(0, lambda: self.log(f"Connecting to {ip}:{port}..."))
                ftp = ftplib.FTP()
                ftp.connect(ip, port, timeout=10)
                ftp.login()  # anonymous
                # FTP root IS /storage/emulated/0/Download — no CWD needed
                from io import BytesIO
                ftp.storbinary(f'STOR {filename}', BytesIO(data_bytes))
                ftp.quit()
                
                self.root.after(0, lambda: self.log(f"✓ Sent {filename} to {ip}:{port}"))
                self.root.after(0, lambda: messagebox.showinfo("Sent", 
                    f"✓ Sent {filename} to Triplett\n\n"
                    f"Address: {ip}:{port}"))
            except ConnectionRefusedError:
                self.root.after(0, lambda: self.log(f"✗ Connection refused at {ip}:{port}"))
                self.root.after(0, lambda: messagebox.showerror("Connection Refused", 
                    f"Nothing is listening at {ip}:{port}\n\n"
                    f"• Is the FTP server running on the Triplett?\n"
                    f"• Does the address match what the Triplett shows?"))
            except socket.timeout:
                self.root.after(0, lambda: self.log(f"✗ Timeout connecting to {ip}:{port}"))
                self.root.after(0, lambda: messagebox.showerror("Timeout", 
                    f"No response from {ip}:{port}\n\n"
                    f"• Is the Triplett on the same network?\n"
                    f"• Is the IP address correct?"))
            except Exception as e:
                self.root.after(0, lambda: self.log(f"✗ FTP error: {e}"))
                self.root.after(0, lambda: messagebox.showerror("FTP Error", 
                    f"Failed to send to {ip}:{port}\n\n{e}"))
            finally:
                self.root.after(0, lambda: btn.config(state='normal', text=orig_text))
        
        btn.config(state='disabled', text='Sending...')
        threading.Thread(target=do_ftp, daemon=True).start()
    
    def triplett_send_cameras(self):
        """Send camera list CSV to Triplett via FTP."""
        addr = self._get_triplett_address()
        if not addr:
            return
        lines = self._build_triplett_csv_lines()
        if not lines:
            messagebox.showwarning("Send Cameras", "No cameras in the list.")
            return
        data = '\n'.join(lines) + '\n'
        self.log(f"Sending {len(lines)} cameras as cameras.csv...")
        self._ftp_send_data(addr[0], addr[1], 'cameras.csv', data.encode('utf-8'),
                           self.triplett_cam_btn, '📋 Send Camera List')
    
    def triplett_send_passwords(self):
        """Send passwords.txt to Triplett via FTP."""
        addr = self._get_triplett_address()
        if not addr:
            return
        passwords = self.password_data.get_all()
        if not passwords:
            messagebox.showwarning("Send Passwords", "No passwords in the list.")
            return
        data = '\n'.join(passwords) + '\n'
        self.log(f"Sending {len(passwords)} passwords as passwords.txt...")
        self._ftp_send_data(addr[0], addr[1], 'passwords.txt', data.encode('utf-8'),
                           self.triplett_pwd_btn, '🔑 Send Passwords')
    
    def triplett_send_file(self):
        """Send any file to Triplett via FTP."""
        addr = self._get_triplett_address()
        if not addr:
            return
        filepath = filedialog.askopenfilename(
            title="Select file to send to Triplett",
            filetypes=[("All Files", "*.*"), ("CSV Files", "*.csv"), ("Text Files", "*.txt")]
        )
        if not filepath:
            return
        try:
            with open(filepath, 'rb') as f:
                data = f.read()
            filename = os.path.basename(filepath)
            self.log(f"Sending {filename} ({len(data)} bytes)...")
            self._ftp_send_data(addr[0], addr[1], filename, data,
                               self.triplett_file_btn, '📁 Send File...')
        except Exception as e:
            messagebox.showerror("File Error", f"Could not read file:\n\n{e}")
    
    def triplett_retrieve(self):
        """Retrieve files from Triplett via FTP into the export folder's triplett/ subdir."""
        addr = self._get_triplett_address()
        if not addr:
            return
        ip, port = addr
        
        self.triplett_recv_btn.config(state='disabled', text='Retrieving...')
        
        def do_retrieve():
            try:
                self.root.after(0, lambda: self.log(f"Connecting to {ip}:{port} for retrieval..."))
                ftp = ftplib.FTP()
                ftp.connect(ip, port, timeout=10)
                ftp.login()  # anonymous
                
                # List all files (FTP root IS /storage/emulated/0/Download)
                filenames = ftp.nlst()
                
                # Filter for files we want
                targets = []
                for fn in filenames:
                    fn_lower = fn.lower()
                    if fn_lower.startswith('img_') and fn_lower.endswith('.zip'):
                        targets.append(fn)
                    elif fn_lower.startswith('log_') and fn_lower.endswith('.txt'):
                        targets.append(fn)
                    elif fn_lower == 'successful_passwords.txt':
                        targets.append(fn)
                    elif fn_lower == 'results.txt':
                        targets.append(fn)
                
                if not targets:
                    self.root.after(0, lambda: self.log("No image zips, passwords, or results found on Triplett."))
                    self.root.after(0, lambda: messagebox.showinfo("Nothing Found", 
                        "No files to retrieve.\n\n"
                        "Looking for: img_*.zip, log_*.txt, successful_passwords.txt, results.txt\n\n"
                        "Run operations on the Triplett first."))
                    return
                
                # Create output directory
                os.makedirs(TRIPLETT_DIR, exist_ok=True)
                
                downloaded = []
                for fn in targets:
                    local_path = os.path.join(TRIPLETT_DIR, fn)
                    with open(local_path, 'wb') as f:
                        ftp.retrbinary(f'RETR {fn}', f.write)
                    downloaded.append(fn)
                    self.root.after(0, lambda fn=fn: self.log(f"  ✓ Retrieved {fn}"))
                
                ftp.quit()
                
                summary = f"Retrieved {len(downloaded)} file(s) → {TRIPLETT_DIR}\n\n"
                summary += '\n'.join(f"  • {fn}" for fn in downloaded)

                self.root.after(0, lambda: self.log(
                    f"✓ Retrieved {len(downloaded)} file(s) from Triplett → {TRIPLETT_DIR}"))
                self.root.after(0, lambda: messagebox.showinfo("Retrieved", summary))
                
            except ConnectionRefusedError:
                self.root.after(0, lambda: self.log(f"✗ Connection refused at {ip}:{port}"))
                self.root.after(0, lambda: messagebox.showerror("Connection Refused", 
                    f"Nothing is listening at {ip}:{port}\n\n"
                    f"• Is the FTP server running on the Triplett?\n"
                    f"• Does the address match what the Triplett shows?"))
            except socket.timeout:
                self.root.after(0, lambda: self.log(f"✗ Timeout connecting to {ip}:{port}"))
                self.root.after(0, lambda: messagebox.showerror("Timeout", 
                    f"No response from {ip}:{port}\n\n"
                    f"• Is the Triplett on the same network?\n"
                    f"• Is the IP address correct?"))
            except Exception as e:
                self.root.after(0, lambda: self.log(f"✗ FTP retrieve error: {e}"))
                self.root.after(0, lambda: messagebox.showerror("FTP Error", 
                    f"Failed to retrieve from {ip}:{port}\n\n{e}"))
            finally:
                self.root.after(0, lambda: self.triplett_recv_btn.config(
                    state='normal', text='📥 Retrieve from Triplett'))
        
        threading.Thread(target=do_retrieve, daemon=True).start()
    
    # ========================================================================
    # PASSWORD LIST MANAGEMENT
    # ========================================================================
    def refresh_password_list(self):
        self.password_listbox.delete(0, tk.END)
        passwords = self.password_data.get_all()
        visible = self.passwords_visible.get()
        for pwd in passwords:
            if visible:
                self.password_listbox.insert(tk.END, pwd)
            else:
                self.password_listbox.insert(tk.END, '•' * len(pwd))
        self.password_status.set(f"{len(passwords)} passwords in list")
    
    def toggle_password_visibility(self):
        self.passwords_visible.set(not self.passwords_visible.get())
        if self.passwords_visible.get():
            self.show_hide_btn.config(text="👁 Hide")
        else:
            self.show_hide_btn.config(text="👁 Show")
        self.refresh_password_list()
    
    def add_password(self):
        pwd = self.new_password_var.get().strip()
        if pwd:
            self.password_data.add(pwd)
            self.new_password_var.set("")
            self.refresh_password_list()
            self.log(f"Added password to list")
    
    def mass_add_passwords(self):
        text = self.mass_password_text.get('1.0', tk.END).strip()
        if not text:
            return
        added = 0
        for line in text.split('\n'):
            pwd = line.strip()
            if pwd:
                existing = self.password_data.get_all()
                if pwd not in existing:
                    self.password_data.add(pwd)
                    added += 1
        self.mass_password_text.delete('1.0', tk.END)
        self.refresh_password_list()
        self.log(f"Bulk added {added} password(s)")
    
    def add_password_quick(self, pwd):
        self.password_data.add(pwd)
        self.refresh_password_list()
        self.log(f"Added common password to list")
    
    def delete_password(self):
        selected = self.password_listbox.curselection()
        if selected:
            self.password_data.delete(selected[0])
            self.refresh_password_list()
    
    def clear_all_passwords(self):
        if messagebox.askyesno("Confirm", "Delete ALL passwords from the list?"):
            self.password_data.clear()
            self.refresh_password_list()

    # ========================================================================
    # ADDITIONAL USERS MANAGEMENT
    # ========================================================================
    def refresh_additional_users_list(self):
        for item in self.users_tree.get_children():
            self.users_tree.delete(item)
        users = self.additional_users_data.get_all()
        for u in users:
            self.users_tree.insert('', tk.END, values=(u['username'], u['password'], u['role']))
        self.additional_users_status.set(f"{len(users)} additional user{'s' if len(users) != 1 else ''}")

    def add_additional_user(self):
        name = self.new_user_name_var.get().strip()
        pwd = self.new_user_pwd_var.get().strip()
        role = self.new_user_role_var.get()
        if not name:
            messagebox.showwarning("Required", "Username is required.")
            return
        if not pwd:
            messagebox.showwarning("Required", "Password is required.")
            return
        if self.additional_users_data.add(name, pwd, role):
            self.new_user_name_var.set("")
            self.new_user_pwd_var.set("")
            self.refresh_additional_users_list()
            self.log(f"Added additional user: {name} ({role})")
        else:
            messagebox.showwarning("Duplicate", f"User '{name}' already exists.")

    def delete_additional_user(self):
        selected = self.users_tree.selection()
        if selected:
            idx = self.users_tree.index(selected[0])
            self.additional_users_data.delete(idx)
            self.refresh_additional_users_list()

    def clear_additional_users(self):
        if messagebox.askyesno("Confirm", "Delete ALL additional users?"):
            self.additional_users_data.clear()
            self.refresh_additional_users_list()

    # ========================================================================
    # LOGGING
    # ========================================================================
    def log(self, message, _persist_to_file=True):
        self.log_queue.put(message)
        # Auto-persist to wizard-run log file (#11). status_log() passes
        # _persist_to_file=False because it already wrote — avoids double lines.
        if _persist_to_file:
            fh = getattr(self, '_wizard_run_log_fh', None)
            if fh:
                try:
                    fh.write(f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
                    fh.flush()
                except Exception:
                    pass
    
    def process_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                timestamp = datetime.now().strftime("%H:%M:%S")
                self.log_text.insert(tk.END, f"[{timestamp}] {msg}\n")
                self.log_text.see(tk.END)
        except queue.Empty:
            pass
        self.root.after(100, self.process_log_queue)
    
    def clear_log(self):
        self.log_text.delete(1.0, tk.END)
    
    def log_copy_selection(self):
        """Copy selected text from log"""
        try:
            text = self.log_text.get(tk.SEL_FIRST, tk.SEL_LAST)
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
        except tk.TclError:
            pass  # No selection
    
    def log_copy_all(self):
        """Copy entire log content"""
        text = self.log_text.get('1.0', 'end-1c')
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
    
    def on_close(self):
        """Close all child windows and exit the application"""
        self.cancel_flag = True
        # Destroy all toplevel windows
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Toplevel):
                widget.destroy()
        self.root.destroy()
        import sys
        sys.exit(0)
    
    def save_log(self):
        # The toolkit has TWO log widgets: self.log_text (legacy/general tab)
        # and self._status_log_text (new wizard view). Older versions only
        # saved self.log_text — which was empty when the wizard had been the
        # only thing run, producing an empty file. Save BOTH with section
        # headers so the user gets whichever has content (or both).
        legacy = self.log_text.get('1.0', 'end-1c').strip() if hasattr(self, 'log_text') else ''
        wizard = self._status_log_text.get('1.0', 'end-1c').strip() if hasattr(self, '_status_log_text') else ''
        if not legacy and not wizard:
            messagebox.showinfo("Save Log", "No log content to save (both general log and wizard log are empty).")
            return
        filepath = filedialog.asksaveasfilename(
            title="Save Log",
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        if not filepath:
            return
        sections = []
        if wizard:
            sections.append(f"========== WIZARD LOG ({len(wizard.splitlines())} lines) ==========\n{wizard}\n")
        if legacy:
            sections.append(f"========== GENERAL LOG ({len(legacy.splitlines())} lines) ==========\n{legacy}\n")
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write("\n".join(sections))
        self.log(f"Log saved to {filepath}")
    
    # ========================================================================
    # UI HELPERS
    # ========================================================================
    def update_display(self, camera="----", status=""):
        self.current_camera_label.config(text=camera)
        self.current_status_label.config(text=status)
    
    def update_preview(self, image_data=None):
        if not HAS_PIL:
            return
        if image_data is None:
            self.preview_label.config(image='', text="No image")
            self.preview_image = None
            return
        try:
            img = Image.open(BytesIO(image_data))
            img.thumbnail((320, 240), Image.Resampling.LANCZOS)
            self.preview_image = ImageTk.PhotoImage(img)
            self.preview_label.config(image=self.preview_image, text="")
        except:
            self.preview_label.config(image='', text="Preview error")
    
    def clear_preview(self):
        self.root.after(0, lambda: self.update_preview(None))
    
    def enable_cancel(self, enable=True):
        state = 'normal' if enable else 'disabled'
        self.cancel_btn.config(state=state)
        self.log_cancel_btn.config(state=state)
    
    def cancel_operation(self):
        self.cancel_flag = True
        self.log("Cancelling operation...")
    
    def get_password(self, title="Password", prompt="Enter password:"):
        return PasswordDialog(self.root, title, prompt).result
    
    # ========================================================================
    # DIALOGS
    # ========================================================================
    def show_welcome(self):
        msg = """Welcome to CCTV IP Toolkit!

This tool helps you configure IP cameras quickly and efficiently.

QUICK START:
1. The app auto-scans your network — check the 'Discovered' tab
2. Copy cameras to 'Camera List' and edit their settings
3. Go to 'Operations' tab and run what you need
4. Smart Import: paste CSV data or use File → Smart Import

Need help? Click Help → Quick Start Guide"""
        messagebox.showinfo("Welcome!", msg)
    
    def show_quick_start(self):
        w = tk.Toplevel(self.root)
        w.title("Quick Start Guide")
        w.geometry("700x650")
        t = scrolledtext.ScrolledText(w, font=('Helvetica', 11), wrap=tk.WORD)
        t.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        t.insert(tk.END, """
CCTV IP TOOLKIT - QUICK START GUIDE
====================================

DISCOVERED TAB (Auto-Scan)
--------------------------
The app automatically scans your network for IP cameras
(Axis, Bosch, and more) on startup and every 90 seconds.

• Cameras found on your network appear in the Discovered tab
• "ON LIST" column shows which ones are already in your Camera List
• "Copy All New" button adds only cameras not yet on your list
• "Copy Selected" lets you pick specific cameras to add
• Rescan runs automatically after operations complete

ADDING CAMERAS TO THE LIST
---------------------------
Option 1: Copy from the Discovered tab (easiest)
Option 2: File → Smart Import (paste or import CSV data)
Option 3: Click "Add Camera" on the Camera List tab

Smart Import auto-detects columns from CSV/spreadsheet data.
Just paste rows and it figures out which columns are IP,
name, gateway, etc.

CAMERA LIST
-----------
Your working list of cameras. Double-click or press Enter to edit.

• Camera Name — your local reference (not sent to camera)
• Hostname — the network hostname ON the camera
• IP Address — current IP (read-only when editing; use New IP to change)
• New IP — set this to change a camera's IP via Update Cameras
• DHCP — toggle in the editor to enable/disable

Keyboard: Delete = remove, Enter = edit, Ctrl+A = select all

OPERATIONS
----------
Go to the "Operations" tab and click what you need:

• PROGRAM NEW CAMERAS
  For factory-fresh cameras. Discovers cameras via DHCP/mDNS
  or connects at the factory default IP. Sets the static IP,
  creates a user (Axis), and optionally sets the hostname.

• UPDATE CAMERAS
  Push any changes you've made in the editor: IP changes,
  hostname changes, or DHCP on/off. The wizard detects what
  changed and only pushes what's needed.

• PING TEST — Quick reachability check

• CAPTURE IMAGES — Download a snapshot from each camera
  Images are timestamped and watermarked (marked as
  "NOT FROM CAMERA OVERLAY" so it's clear).

• CHANGE PASSWORDS — Set a new password on all cameras

• VALIDATE PASSWORD — Test if a password works

• BATCH PASSWORD TEST — Try multiple passwords to find
  the right one (add passwords in the "Passwords" tab first)

PASSWORDS TAB
-------------
Manage passwords for batch testing. Passwords are never
shown in the operations log — only masked values appear.

SETTINGS
--------
File → Settings to configure:
  • Default username (usually "root")
  • Factory default IP for new cameras
  • Warning dialog preferences

KEYBOARD SHORTCUTS
------------------
Camera List:
  Delete    — Delete selected camera(s)
  Enter     — Edit selected camera
  Ctrl+A    — Select all

Discovered List:
  Enter     — Copy selected to Camera List
  Ctrl+A    — Select all

Log Tab:
  Right-click — Copy selection, Copy All, Clear

NEED MORE HELP?
---------------
Email: axisprogrammer@thelostping.net
""")
        t.config(state=tk.DISABLED)
    
    # ------------------------------------------------------------------
    # Update checking
    # ------------------------------------------------------------------
    def _fetch_latest_release(self, timeout=6):
        """Return (tag, body, html_url) of the latest published release, or None."""
        try:
            req = urllib.request.Request(
                GITHUB_LATEST_API,
                headers={"User-Agent": f"CCTVIPToolkit/{APP_VERSION}",
                         "Accept": "application/vnd.github+json"},
            )
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            tag = (data.get("tag_name") or "").lstrip("vV").strip()
            body = (data.get("body") or "").strip()
            url = data.get("html_url") or GITHUB_RELEASES_PAGE
            return tag, body, url
        except Exception:
            return None

    @staticmethod
    def _version_tuple(v):
        """'4.2' -> (4, 2). Handles 4.2.1, v4.2-beta1, etc. gracefully."""
        parts = []
        for chunk in re.split(r"[.\-+]", (v or "").lstrip("vV")):
            m = re.match(r"^(\d+)", chunk)
            if m:
                parts.append(int(m.group(1)))
        return tuple(parts) or (0,)

    def check_for_update(self, silent=False):
        """Hit GitHub's latest-release API, compare tags. If newer available, show
        dialog with release notes + Download / Remind Later buttons. When silent=True,
        stay quiet if there's nothing new (used for the startup auto-check)."""
        result = self._fetch_latest_release()
        if not result:
            if not silent:
                messagebox.showwarning(
                    "Update Check Failed",
                    "Couldn't reach GitHub to check for updates.\n\n"
                    f"Running version: {APP_VERSION}\n\n"
                    f"You can check manually at:\n{GITHUB_RELEASES_PAGE}",
                )
            return
        latest_tag, body, url = result
        latest_tup = self._version_tuple(latest_tag)
        current_tup = self._version_tuple(APP_VERSION)

        if latest_tup <= current_tup:
            if not silent:
                messagebox.showinfo(
                    "You're Up to Date",
                    f"Running v{APP_VERSION} - latest published is v{latest_tag}.\n\nYou're current.",
                )
            return

        # Newer version exists. If user already dismissed this specific version, stay silent.
        if silent and self.settings.get('general', 'last_dismissed_version') == latest_tag:
            return

        self._show_update_dialog(latest_tag, body, url)

    def _show_update_dialog(self, latest_tag, body, url):
        """Toplevel showing version diff + release notes + action buttons."""
        w = tk.Toplevel(self.root)
        w.title("Update Available")
        w.geometry("640x520")
        w.transient(self.root)
        w.grab_set()
        w.update_idletasks()
        px = self.root.winfo_x() + (self.root.winfo_width() - 640) // 2
        py = self.root.winfo_y() + (self.root.winfo_height() - 520) // 2
        w.geometry(f"640x520+{max(px, 0)}+{max(py, 0)}")

        ttk.Label(w, text=f"New version: v{latest_tag}",
                  font=('Helvetica', 16, 'bold')).pack(pady=(18, 4))
        ttk.Label(w, text=f"You're on v{APP_VERSION}",
                  foreground='gray', font=('Helvetica', 10)).pack(pady=(0, 10))

        notes_frame = ttk.LabelFrame(w, text="Release Notes", padding=8)
        notes_frame.pack(fill=tk.BOTH, expand=True, padx=18, pady=6)
        txt = scrolledtext.ScrolledText(notes_frame, wrap=tk.WORD, height=14,
                                        font=('Helvetica', 10))
        txt.pack(fill=tk.BOTH, expand=True)
        txt.insert('1.0', body if body else "(no release notes published)")
        txt.config(state=tk.DISABLED)

        btns = ttk.Frame(w)
        btns.pack(fill=tk.X, padx=18, pady=(6, 14))

        def open_download():
            # Route through fieldtoolkit.com tracker so upgrades count in analytics
            # the same as fresh downloads. The tracker 302s to the GitHub asset.
            import webbrowser
            tracked_url = f"{FIELDTOOLKIT_DOWNLOAD_URL}&v=v{latest_tag}"
            webbrowser.open(tracked_url)

        def open_release_notes():
            import webbrowser
            webbrowser.open(url)

        def remind_later():
            self.settings.set('general', 'last_dismissed_version', latest_tag)
            w.destroy()

        ttk.Button(btns, text=f"Download v{latest_tag}", command=open_download).pack(side=tk.LEFT)
        ttk.Button(btns, text="Release Notes on GitHub", command=open_release_notes).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(btns, text="Remind Me Later", command=remind_later).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(btns, text="Close", command=w.destroy).pack(side=tk.RIGHT)

    # ------------------------------------------------------------------
    # What's New (first launch of a new version)
    # ------------------------------------------------------------------
    WHATS_NEW = {
        "4.2.8": (
            "What's new in v4.2.8",
            [
                "• Critical bug fix: after a successful programming, if the user hadn't yet unplugged the previous camera (still on factory IP), the toolkit would IMMEDIATELY try to program the next slot onto that same camera — auth would fail, the network step would appear to succeed but actually do nothing useful, and the script would bail with PROGRAMMING COMPLETE: 1 succeeded.",
                "• Root cause #1: ARP query (the seen_macs check from v4.2.7) returned None because arp_unpin had cleared the static entry and the OS dynamic ARP cache had aged out — bypassing the seen_macs guard entirely.",
                "• Root cause #2: even with ARP working, there was no positive 'wait for previous camera to leave' before next discovery.",
                "• Fix #1: get_mac_from_arp now actively pings the IP first to force ARP resolution, and retries up to 3 times. ARP-based detection is reliable again.",
                "• Fix #2: after every successful camera completion, the toolkit now waits for the factory IP to become unreachable before starting the next discovery. Logs a heartbeat every 30 seconds during the wait. Cancel-aware. Applied to BOTH the named-list flow and the unprogrammed-cameras flow (with continue-dialog confirmation).",
                "• Net effect: even if you click \"continue\" on the dialog before unplugging the camera, the toolkit will not advance until the previous camera is physically gone from factory IP.",
            ],
        ),
        "4.2.7": (
            "What's new in v4.2.7",
            [
                "• Bug fix: when a previously-programmed camera was still rebooting on the factory IP, the toolkit appeared to flash through cameras every 3 seconds (\"Waiting for camera 17: KBW-005\" incrementing rapidly even though only one camera had been done).",
                "• Root cause: the \"already programmed, waiting for reboot\" branch was bouncing back to the outer loop, which re-printed the banner and bumped the camera counter on every retry.",
                "• Fix: the wait now happens IN PLACE — the toolkit silently waits for the old camera to leave the factory IP (with a heartbeat log every 30s), then re-enters discovery for the SAME slot. The camera number stays accurate.",
                "• Same fix applied to both the auto-name-from-list flow and the unprogrammed-cameras flow.",
            ],
        ),
        "4.2.6": (
            "What's new in v4.2.6",
            [
                "• Update dialog: 'Download' button now pulls the new EXE directly from fieldtoolkit.com instead of routing you to the GitHub release page.",
                "• 'Release Notes on GitHub' added as a separate button so you can read the changelog without grabbing the EXE.",
                "• Internal: source file renamed to cctv_toolkit.py to match the multi-brand reality (Axis + Bosch + Hanwha). No effect on the app or your data.",
            ],
        ),
        "4.2.4": (
            "What's new in v4.2.4",
            [
                "• Heavy comment pass on the vendor protocol code (Axis, Bosch, Hanwha).",
                "• Field-tech context, vendor-quirk notes, and the why-it's-written-this-way is now in the source for anyone reading the public mirror.",
                "• Zero functional changes — same app, same APIs, same workflow.",
            ],
        ),
        "4.2.3": (
            "What's new in v4.2.3",
            [
                "• New canonical home: https://fieldtoolkit.com (was cctv.thelostping.net).",
                "• The old URL still works — it 301 redirects to fieldtoolkit.com.",
                "• No functional app changes from 4.2.2.",
            ],
        ),
        "4.2.2": (
            "What's new in v4.2.2",
            [
                "• New app icon — cleaner, more professional mark.",
                "• No functional changes from 4.2.1.",
            ],
        ),
        "4.2.1": (
            "What's new in v4.2.1",
            [
                "• Patch release: corrects a stray release-note bullet that mentioned an unrelated project.",
                "• Same functionality as v4.2 otherwise.",
            ],
        ),
        "4.2": (
            "What's new in v4.2",
            [
                "• Built-in update checker — Help menu + silent check on startup.",
                "• First-launch \"What's New\" popup (this dialog) so version bumps don't sneak past you.",
                "• Data split in v4.1 carried forward: config in %APPDATA%, exports in Documents.",
                "• Triplett Android integration stays dormant -- placeholder in the UI, code preserved for future re-enable.",
            ],
        ),
        "4.1": (
            "What's new in v4.1",
            [
                "• Config moved to %APPDATA%\\CCTVIPToolkit\\ so upgrades never wipe your password list or camera list.",
                "• Exports moved to Documents\\CCTV Toolkit (user-configurable in Settings).",
                "• File menu: Open Export Folder + Open Config Folder replace the old Open Data Folder.",
                "• One-time migration copies any legacy ./data/ next to the .exe into both new locations, leaves the original as a safety net.",
                "• Triplett Android integration UI hidden pending a future release (code kept intact).",
            ],
        ),
    }

    def show_whats_new(self, version=None):
        """Show the release notes for a specific version (default: current APP_VERSION).
        Called automatically on first launch of a new version; also available from Help menu."""
        v = version or APP_VERSION
        entry = self.WHATS_NEW.get(v)
        if not entry:
            messagebox.showinfo("What's New",
                                f"No local release notes for v{v}.\n\nFull changelog on GitHub:\n{GITHUB_RELEASES_PAGE}")
            return
        title, bullets = entry
        messagebox.showinfo(title, "\n".join(bullets) + f"\n\nFull changelog: {GITHUB_RELEASES_PAGE}")

    def show_about(self):
        messagebox.showinfo("About", f"""CCTV IP Toolkit v{APP_VERSION}

Created by Brian Preston

Features:
• Auto-discover cameras on network (Axis, Bosch)
• Program new cameras (IP, user, hostname, DHCP)
• Update cameras (IP, hostname, DHCP changes)
• Smart CSV import with auto-detection
• Batch ping, image capture, password testing
• Background rescanning every 90 seconds
• Timestamped image snapshots

https://buymeacoffee.com/thelostping""")
    
    def show_lldp_discovery(self):
        LldpDiscoveryDialog(self.root)

    def show_settings(self):
        w = tk.Toplevel(self.root)
        w.title("Settings")
        w.geometry("650x550")
        w.transient(self.root)
        w.grab_set()
        
        # Center on parent
        w.update_idletasks()
        px = self.root.winfo_x() + (self.root.winfo_width() - 650) // 2
        py = self.root.winfo_y() + (self.root.winfo_height() - 550) // 2
        w.geometry(f"650x550+{max(px,0)}+{max(py,0)}")
        
        # Scrollable interior
        canvas = tk.Canvas(w, highlightthickness=0)
        scrollbar = ttk.Scrollbar(w, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        frame = ttk.Frame(canvas, padding="25")
        canvas.create_window((0, 0), window=frame, anchor='nw')
        frame.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind_all('<MouseWheel>', lambda e: canvas.yview_scroll(-1 * (e.delta // 120), 'units'))
        
        ttk.Label(frame, text="Settings", font=('Helvetica', 18, 'bold')).pack(pady=(0, 20))
        
        # Defaults section
        defaults_frame = ttk.LabelFrame(frame, text="Default Values", padding="15")
        defaults_frame.pack(fill=tk.X, pady=(0, 15))
        
        settings_entries = {}
        defaults = [
            ('Username:', 'username'),
        ]
        
        for i, (label, key) in enumerate(defaults):
            ttk.Label(defaults_frame, text=label, font=('Helvetica', 10)).grid(row=i, column=0, sticky='w', pady=5)
            entry = ttk.Entry(defaults_frame, width=25, font=('Helvetica', 10))
            entry.insert(0, self.settings.get('general', key))
            entry.grid(row=i, column=1, sticky='w', padx=(15, 0), pady=5)
            settings_entries[key] = entry
        
        # Export folder section
        export_frame = ttk.LabelFrame(frame, text="Export Folder (CSVs, screenshots, FTP pulls)", padding="15")
        export_frame.pack(fill=tk.X, pady=(0, 15))

        current_export = self.settings.get('general', 'export_dir') or str(_default_export_dir())
        export_var = tk.StringVar(value=current_export)
        export_row = ttk.Frame(export_frame)
        export_row.pack(fill=tk.X)
        export_entry = ttk.Entry(export_row, textvariable=export_var, font=('Helvetica', 10))
        export_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        def pick_export_dir():
            chosen = filedialog.askdirectory(
                parent=w,
                title="Choose export folder",
                initialdir=export_var.get() or str(Path.home()),
            )
            if chosen:
                export_var.set(chosen)
        ttk.Button(export_row, text="Browse…", command=pick_export_dir).pack(side=tk.LEFT, padx=(8, 0))

        def reset_export_dir():
            export_var.set(str(_default_export_dir()))
        ttk.Button(export_row, text="Reset", command=reset_export_dir).pack(side=tk.LEFT, padx=(4, 0))

        ttk.Label(export_frame,
                  text="Config (passwords, camera list, settings) always stays in your user profile and survives upgrades.",
                  foreground='gray', font=('Helvetica', 9), wraplength=540, justify=tk.LEFT).pack(anchor=tk.W, pady=(8, 0))

        # Warnings section
        warnings_frame = ttk.LabelFrame(frame, text="Show Warnings", padding="15")
        warnings_frame.pack(fill=tk.X, pady=(0, 15))
        
        warning_vars = {}
        warnings = [
            ('show_incomplete_camera_warning', 'Incomplete camera data warning'),
            ('show_batch_test_explanation', 'Batch test explanation'),
            ('show_programming_intro', 'Programming introduction'),
        ]
        
        for key, label in warnings:
            var = tk.BooleanVar(value=self.settings.get_bool('warnings', key))
            ttk.Checkbutton(warnings_frame, text=label, variable=var).pack(anchor=tk.W, pady=3)
            warning_vars[key] = var
        
        # Save button
        def save_settings():
            for key, entry in settings_entries.items():
                self.settings.set('general', key, entry.get())
            for key, var in warning_vars.items():
                self.settings.set('warnings', key, str(var.get()).lower())
            # Export dir: blank = default
            new_export = (export_var.get() or '').strip()
            default_export = str(_default_export_dir())
            persisted = '' if (not new_export or new_export == default_export) else new_export
            self.settings.set('general', 'export_dir', persisted)
            self.settings.apply_export_dir()
            w.destroy()
            messagebox.showinfo("Saved", f"Settings saved.\n\nExport folder: {EXPORT_DIR}")
        
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="💾 Save Settings", command=save_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=w.destroy).pack(side=tk.LEFT, padx=5)
    
    def open_export_folder(self):
        """Open the folder where CSVs, screenshots, and FTP pulls land."""
        EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(EXPORT_DIR))
        except AttributeError:
            import subprocess
            subprocess.Popen(['xdg-open' if sys.platform.startswith('linux') else 'open', str(EXPORT_DIR)])

    def open_config_folder(self):
        """Open the private folder with saved passwords, camera list, and settings."""
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(CONFIG_DIR))
        except AttributeError:
            import subprocess
            subprocess.Popen(['xdg-open' if sys.platform.startswith('linux') else 'open', str(CONFIG_DIR)])
    
    # ========================================================================
    # OPERATION WIZARDS
    # ========================================================================
    def validate_cameras_for_programming(self):
        """Check if camera list is ready for programming"""
        cameras = self.camera_data.get_valid_for_programming()
        if not cameras:
            all_cams = self.camera_data.get_all()
            if not all_cams:
                messagebox.showwarning("No Cameras", 
                    "Your camera list is empty!\n\n"
                    "Go to the 'Camera List' tab and add cameras first.\n\n"
                    "Each camera needs:\n"
                    "• Camera Name\n"
                    "• IP Address\n"
                    "• Gateway\n"
                    "• Subnet Mask")
                self.notebook.select(self.cameras_tab)  # Switch to camera list tab
                return None
            else:
                # Have cameras but none valid
                incomplete = [c['name'] for c in all_cams if not c.get('gateway') or not c.get('subnet')]
                if incomplete:
                    messagebox.showwarning("Incomplete Data",
                        f"These cameras are missing Gateway or Subnet:\n\n"
                        f"{', '.join(incomplete[:5])}{'...' if len(incomplete) > 5 else ''}\n\n"
                        "Programming requires:\n"
                        "• IP Address\n"
                        "• Gateway\n"
                        "• Subnet Mask\n\n"
                        "Please edit these cameras in the Camera List tab.")
                    self.notebook.select(self.cameras_tab)
                    return None
                
                processed = [c['name'] for c in all_cams if c.get('processed')]
                if processed:
                    if messagebox.askyesno("All Processed",
                        "All cameras are marked as processed.\n\n"
                        "Would you like to clear the processed flags\n"
                        "and program them again?"):
                        self.reset_status()
                        return self.camera_data.get_valid_for_programming()
                return None
        return cameras
    
    def validate_cameras_for_basic_ops(self):
        """Check if camera list has basic data for ping/capture/etc"""
        cameras = self.camera_data.get_valid_for_basic_ops()
        if not cameras:
            messagebox.showwarning("No Cameras",
                "Your camera list is empty or has no valid IPs!\n\n"
                "Go to the 'Camera List' tab and add cameras.\n\n"
                "At minimum, each camera needs:\n"
                "• Camera Name\n"
                "• IP Address")
            self.notebook.select(self.cameras_tab)
            return None
        return cameras
    
    def start_program_wizard_classic(self):
        """Classic programming flow — original combined-options dialog + log-only UI."""
        cameras = self.validate_cameras_for_programming()
        if not cameras:
            return

        # Show intro if enabled
        if self.settings.get_bool('warnings', 'show_programming_intro'):
            factory_ip = self.settings.get('general', 'factory_ip')
            dialog = WarningDialog(self.root, "Program New Cameras (Classic)",
                f"Ready to program {len(cameras)} camera(s).\n\n"
                "This will:\n"
                f"1. Discover cameras via DHCP/mDNS (finds link-local 169.254.x.x)\n"
                f"   or connect to factory IP ({factory_ip})\n"
                "2. ARP-pin to lock onto that specific camera\n"
                "3. Verify model matches (if specified)\n"
                "4. Create system user with your password\n"
                "5. Set static IP and disable DHCP\n"
                "6. Capture serial/MAC after programming\n"
                "7. Optionally create additional users\n\n"
                "Supports factory default IPs and link-local cameras.\n"
                "Multiple cameras can be connected simultaneously\n"
                "- each programmed one at a time.",
                'show_programming_intro', self.settings)
            if not dialog.result:
                return
        
        # Get password
        password = self.get_password("Set Camera Password", 
            "Enter the password to set on all cameras:")
        if not password:
            return
        
        confirm = self.get_password("Confirm Password", "Confirm password:")
        if password != confirm:
            messagebox.showerror("Mismatch", "Passwords don't match!")
            return
        
        # Get factory IP and hostname option
        prog_opts = ProgramOptionsDialog(self.root,
            factory_ip=self.protocol.FACTORY_IP,
            additional_users_count=len(self.additional_users_data.get_all()))
        if not prog_opts.result:
            return
        factory_ip = prog_opts.result['factory_ip']
        discovery_mode = prog_opts.result.get('discovery_mode', 'both')
        set_hostname = prog_opts.result['set_hostname']
        add_additional_users = prog_opts.result.get('add_additional_users', False)
        selected_iface = prog_opts.result.get('interface')

        # Override detected interface if user selected one — persist in settings
        if selected_iface:
            self._detected_iface_index = selected_iface['index']
            self._detected_local_ip = selected_iface['ip']
            self.settings.set('general', 'interface_index', str(selected_iface['index']))
            self.log(f"Using interface: {selected_iface['label']}")

        # Start programming
        self.notebook.select(self.log_tab)  # Switch to log tab
        self.cancel_flag = False
        self.enable_cancel(True)

        def run():
            wizard_log_path = self._open_wizard_run_log()
            username = self.protocol.DEFAULT_USER
            total_ok = total_fail = 0

            # v4.3 #13 — Building Reports sticker counter
            _r0 = (prog_opts.result if (prog_opts and prog_opts.result) else {})
            br_enabled = bool(_r0.get('add_br_stickers', False))
            br_counter = None
            if br_enabled:
                try:
                    br_counter = int(str(_r0.get('br_first_label', '')).strip())
                except (ValueError, TypeError):
                    br_enabled = False
                    self.log("⚠ Building Reports stickers requested but seed value is not a number — skipping.")

            _ensure_output_csv_header()

            if wizard_log_path:
                self.log(f"Wizard run log: {wizard_log_path}")
            if br_enabled:
                self.log(f"Building Reports stickers: starting at #{br_counter}")
            self.log(f"Discovery mode: {discovery_mode}")
            if factory_ip:
                self.log(f"Factory IP: {factory_ip}")

            # Pre-add link-local route when discovery mode includes link-local cameras
            if discovery_mode in ('mdns', 'both'):
                self.log("Adding link-local route for camera discovery...")
                self.add_linklocal_route()

            # Track which cameras still need programming
            remaining = list(cameras)  # copy — we'll remove as we go
            programmed_count = 0
            seen_macs = set()  # Track MACs we've already programmed to avoid re-hitting same camera
            consecutive_skips = 0
            # Suppress banner re-print + counter bump when re-entering the outer
            # loop after a wait-for-reboot wait (2026-04-30 bug fix).
            last_announced_remaining_count = None

            while remaining and not self.cancel_flag:
                pinned_mac = None
                camera_ip = None  # Will be factory_ip OR link-local from discovery

                # Only bump counter + re-announce when remaining shrinks (real new slot).
                if len(remaining) != last_announced_remaining_count:
                    programmed_count += 1
                    self.root.after(0, lambda: self.update_display("Waiting...",
                        f"{len(remaining)} cameras remaining"))
                    self.log(f"\n{'='*50}")
                    self.log(f"Waiting for unprogrammed camera...")
                    if discovery_mode == 'mdns':
                        self.log(f"  Checking: DHCP + mDNS discovery (link-local)")
                    elif discovery_mode == 'factory':
                        self.log(f"  Checking: factory IP ({factory_ip}) only")
                    else:
                        self.log(f"  Checking: factory IP ({factory_ip}) + DHCP/mDNS (link-local)")
                    self.log(f"{len(remaining)} cameras remaining to program")
                    self.log(f"{'='*50}")
                    last_announced_remaining_count = len(remaining)

                # Wait for a camera based on discovery mode
                while not self.cancel_flag:
                    # Try factory default IP if enabled (fast 1s ping — local LAN)
                    if discovery_mode in ('factory', 'both'):
                        if factory_ip and self.ping_camera(factory_ip, timeout_ms=1000):
                            camera_ip = factory_ip
                            self.log(f"Camera found at {factory_ip}")
                            break

                    if discovery_mode in ('mdns', 'both'):
                        # Phase 1: DHCP snooping — listen for camera broadcasts
                        # Cameras send DHCP DISCOVER every few seconds with MAC/model/hostname.
                        # This is passive and works on any subnet (Layer 2 broadcast).
                        dhcp_found_mac = None
                        dhcp_found_brand = None
                        try:
                            dhcp_cams = AxisDHCPDiscovery.discover(timeout=3)
                            for dc in dhcp_cams:
                                dc_mac = dc.get('mac', '').upper().replace(':', '').replace('-', '')
                                if dc_mac and dc_mac not in seen_macs:
                                    dhcp_found_mac = dc.get('mac', '')
                                    dhcp_found_brand = dc.get('brand', self.protocol.BRAND_KEY)
                                    self.log(f"Camera detected via DHCP broadcast:")
                                    self.log(f"  Brand: {dhcp_found_brand} | Model: {dc.get('model', '?')} | MAC: {dc.get('mac', '?')}")
                                    self.add_linklocal_route()
                                    break
                        except Exception:
                            pass

                        # Brand shortcut — try factory IP for DHCP-discovered camera
                        if dhcp_found_mac and not camera_ip:
                            try:
                                if self.ping_camera(factory_ip, timeout_ms=1000):
                                    info = self.protocol.get_discovery_info(factory_ip, timeout=2)
                                    if info:
                                        camera_ip = factory_ip
                                        pinned_mac = dhcp_found_mac
                                        self.log(f"Camera found at {factory_ip}")
                                        self.log(f"  Model: {info.get('model', '?')}")
                                        break
                            except Exception:
                                pass

                        # Phase 2: targeted mDNS to resolve camera IP (Axis cameras)
                        # Binds to 169.254.100.1 so multicast goes out on correct adapter.
                        if not camera_ip:
                            try:
                                resolved_cams = self._resolve_linklocal_cameras(
                                    target_mac=dhcp_found_mac, timeout=4)
                                for mc in resolved_cams:
                                    mc_ip = mc.get('ip', '')
                                    mc_mac = mc.get('mac', '').upper().replace(':', '').replace('-', '')
                                    if mc_mac and mc_mac in seen_macs:
                                        continue
                                    if mc_ip and self.ping_camera(mc_ip, timeout_ms=1000):
                                        camera_ip = mc_ip
                                        pinned_mac = mc.get('mac', '')
                                        self.log(f"Camera found: {mc_ip}")
                                        self.log(f"  Model: {mc.get('model', '?')} | Serial: {mc.get('serial', '?')}")
                                        break
                                if camera_ip:
                                    break
                            except Exception:
                                pass

                        # Phase 3: fallback — regular mDNS (finds non-link-local cameras too)
                        if not camera_ip:
                            try:
                                mdns_cams = AxisMDNSDiscovery.discover(timeout=2)
                                for mc in mdns_cams:
                                    mc_ip = mc.get('ip', '')
                                    mc_mac = mc.get('mac', '').upper().replace(':', '').replace('-', '')
                                    if mc_mac and mc_mac in seen_macs:
                                        continue
                                    if mc_ip.startswith('169.254.'):
                                        self.add_linklocal_route()
                                    if mc_ip and self.ping_camera(mc_ip, timeout_ms=1000):
                                        camera_ip = mc_ip
                                        pinned_mac = mc.get('mac', '')
                                        self.log(f"Camera found via mDNS: {mc_ip}")
                                        self.log(f"  Model: {mc.get('model', '?')} | Serial: {mc.get('serial', '?')}")
                                        break
                                if camera_ip:
                                    break
                            except Exception:
                                pass

                    time.sleep(1)
                    
                if self.cancel_flag:
                    break

                # ---- MAC freshness gate (v4.3 #9) ----
                # Backend-only loop. Don't announce "programming new camera"
                # to the operator until probe_unrestricted has confirmed the
                # discovered MAC is NOT in seen_macs. Brian 2026-05-03 spec:
                # "I see a camera ip-(no matter what it is), scan- Can I see
                # the mac? yes, Mac is/is not pre-existing, if not pre-existing
                # 'Hey User! I found a new camera, programming now!'"
                # This eliminates the phantom-reprogram race where a just-
                # programmed camera transients through factory/link-local IPs
                # during its reboot and the wizard would have re-grabbed it.
                mac_settled = False
                last_heartbeat = time.time()
                while not self.cancel_flag and not mac_settled:
                    pre_probe = self.protocol.probe_unrestricted(camera_ip)
                    mac = pre_probe.get('mac')
                    if mac:
                        mac_norm = mac.upper().replace(':', '').replace('-', '')
                        if mac_norm not in seen_macs:
                            pinned_mac = mac
                            mac_settled = True
                            break
                    if not self.ping_camera(camera_ip, timeout_ms=1000):
                        self.log(f"Camera left {camera_ip} — re-running discovery...")
                        camera_ip = None
                        pinned_mac = None
                        break
                    now_t = time.time()
                    if now_t - last_heartbeat >= 30:
                        seen_str = (mac.upper() if mac else 'unreadable')
                        self.log(f"  Waiting for new camera... (current at {camera_ip}: {seen_str} — already programmed; unplug + plug next)")
                        last_heartbeat = now_t
                    time.sleep(2)

                if not mac_settled:
                    continue

                # ARP pin — lock onto this specific camera's MAC
                if not pinned_mac:
                    pinned_mac = self.get_mac_from_arp(camera_ip)
                if pinned_mac:
                    # Check if we already programmed this MAC (camera hasn't rebooted yet)
                    if pinned_mac.upper().replace(':', '').replace('-', '') in seen_macs:
                        # Wait IN PLACE for the old camera to leave factory IP. Don't bounce
                        # back to the outer while — that increments programmed_count and
                        # re-prints "Waiting for camera N" on every 3s tick, which looks
                        # like the toolkit is flashing through cameras (2026-04-30 bug).
                        self.log(f"Camera {pinned_mac} just programmed — waiting for it to leave factory IP {camera_ip}...")
                        self.arp_unpin(camera_ip)
                        wait_start = time.time()
                        last_heartbeat = wait_start
                        while not self.cancel_flag:
                            if not self.ping_camera(camera_ip, timeout_ms=800):
                                self.log(f"  Old camera left factory IP after {int(time.time() - wait_start)}s — ready for next one")
                                break
                            now_t = time.time()
                            if now_t - last_heartbeat >= 30:
                                self.log(f"  ...still rebooting ({int(now_t - wait_start)}s elapsed). Plug in next camera if ready.")
                                last_heartbeat = now_t
                            time.sleep(2)
                        camera_ip = None
                        pinned_mac = None
                        continue

                    if self.arp_pin(camera_ip, pinned_mac):
                        self.log(f"Camera detected! ARP pinned to {pinned_mac}")
                    else:
                        self.log(f"Camera detected! MAC: {pinned_mac} (ARP pin failed — proceeding)")
                else:
                    self.log("Camera detected! (could not read MAC from ARP table)")
                
                # Single no-auth probe: model + serial(=MAC for Axis) + firmware + hardware
                probe = self.protocol.probe_unrestricted(camera_ip)
                actual_model = probe.get('model')
                if actual_model:
                    self.log(f"Camera model: {actual_model}")
                else:
                    self.log("Camera model: could not query (will match any entry)")
                # If ARP didn't give us MAC, harvest from the probe (Axis serial == MAC)
                if probe.get('mac') and not pinned_mac:
                    pinned_mac = probe['mac']
                    self.log(f"MAC (from probe, ARP missed): {pinned_mac}")

                actual_firmware = probe.get('firmware') or 'UNKNOWN'
                if not probe.get('firmware'):
                    try:
                        actual_firmware = self.protocol.get_firmware(camera_ip, '') or 'UNKNOWN'
                    except Exception:
                        actual_firmware = 'UNKNOWN'
                if actual_firmware and actual_firmware != 'UNKNOWN':
                    self.log(f"Camera firmware: {actual_firmware}")

                # Find a matching camera entry from the remaining list
                cam = None
                cam_idx = None
                if actual_model:
                    norm_actual = actual_model.upper().replace('AXIS-', '').replace('AXIS ', '').strip()
                    for idx, c in enumerate(remaining):
                        c_model = c.get('model', '')
                        if not c_model:
                            # Entry has no model specified — matches anything
                            cam = c
                            cam_idx = idx
                            break
                        norm_expected = c_model.upper().replace('AXIS-', '').replace('AXIS ', '').strip()
                        if norm_expected in norm_actual or norm_actual in norm_expected:
                            cam = c
                            cam_idx = idx
                            break
                    
                    if not cam:
                        # Wrong model — tell user what we need
                        models_needed = sorted(set(c.get('model', '(any)') for c in remaining))
                        self.log(f"⚠ MODEL MISMATCH: got {actual_model}")
                        self.log(f"  Need: {', '.join(models_needed)}")
                        self.arp_unpin(camera_ip)
                        consecutive_skips += 1
                        
                        # Always show dialog on mismatch so user knows what to plug in
                        result = [None]
                        skip_count = consecutive_skips
                        def show_mismatch_warning():
                            msg = (f"Wrong camera model detected!\n\n"
                                   f"Connected camera: {actual_model}\n\n"
                                   f"Models still needed:\n")
                            # Show count per model
                            model_counts = {}
                            for c in remaining:
                                m = c.get('model', '(any)')
                                model_counts[m] = model_counts.get(m, 0) + 1
                            for m, count in sorted(model_counts.items()):
                                msg += f"  • {m}  ×{count}\n"
                            msg += (f"\nPlease connect a matching camera.\n\n"
                                    f"({skip_count} mismatch{'es' if skip_count != 1 else ''} so far)\n\n"
                                    "Try again?")
                            result[0] = messagebox.askyesno("Wrong Camera Model", msg)
                        self.root.after(0, show_mismatch_warning)
                        while result[0] is None and not self.cancel_flag:
                            time.sleep(0.1)
                        if not result[0]:
                            self.cancel_flag = True
                            break
                        
                        time.sleep(2)  # Brief delay before trying again
                        continue
                else:
                    # Probe returned no model — same v4.3 #8 confirm dialog as
                    # new wizard (askyesnocancel: Try Again / Proceed Anyway / Cancel)
                    next_name = remaining[0].get('name', '?')
                    next_model = remaining[0].get('model', '(any)') or '(any)'
                    result = [None]
                    def show_no_model_classic():
                        msg = (f"Could not identify the camera at {camera_ip}.\n\n"
                               f"Probe returned no model — camera may be rebooting, "
                               f"unreachable on HTTP, or have firmware that doesn't "
                               f"answer basicdeviceinfo.cgi.\n\n"
                               f"Yes    → Try again (re-probe — camera may respond now)\n"
                               f"No     → Proceed anyway, treat as: {next_name} ({next_model})\n"
                               f"Cancel → Bail the whole wizard run")
                        result[0] = messagebox.askyesnocancel(
                            "Cannot Confirm Camera Model", msg)
                    self.root.after(0, show_no_model_classic)
                    while result[0] is None and not self.cancel_flag:
                        time.sleep(0.1)
                    if self.cancel_flag or result[0] is None:
                        self.cancel_flag = True
                        self.log("Cancelled by user — bailing wizard.")
                        break
                    if result[0] is True:
                        self.log("Retrying probe for this camera slot...")
                        self.arp_unpin(camera_ip)
                        camera_ip = None
                        pinned_mac = None
                        time.sleep(2)
                        continue
                    self.log(f"Proceeding without model verification — assuming {next_name}.")
                    cam = remaining[0]
                    cam_idx = 0

                consecutive_skips = 0  # Reset skip counter on match
                
                cam_name = cam['name']
                static_ip = cam['ip']
                gateway = cam['gateway']
                subnet = cam['subnet']
                expected_model = cam.get('model', '')
                cidr = self.subnet_to_cidr(subnet)
                errors = []
                
                self.root.after(0, lambda n=cam_name, m=actual_model or expected_model: 
                    self.update_display(n, f"Programming... ({m})"))
                self.log(f"\nAssigned to: {cam_name} → {static_ip}")
                self.log(f"Programming from: {camera_ip}" + (" [link-local]" if camera_ip.startswith('169.254.') else ""))
                if expected_model:
                    self.log(f"Model match: expected {expected_model}, got {actual_model} ✓")
                
                # Program via protocol — brand-agnostic steps
                # Network change MUST be last — after that the camera may be unreachable
                cam['_program_ip'] = camera_ip  # Tell protocol which IP to program from
                steps = self.protocol.get_programming_steps(cam, password)

                # Split: auth steps first, network change last
                # Convention: "Setting gateway" / "Setting network" / "Setting password" are
                # identified by looking for network-change keywords
                network_keywords = ('gateway', 'network', 'ip', 'dhcp')
                auth_steps = []
                network_steps = []
                for desc, fn in steps:
                    desc_lower = desc.lower()
                    if any(kw in desc_lower for kw in network_keywords):
                        network_steps.append((desc, fn))
                    else:
                        auth_steps.append((desc, fn))

                # Phase 1: Auth steps (create user, set password)
                total_steps = len(auth_steps) + len(network_steps)
                extra_count = 0
                if add_additional_users and self.additional_users_data.get_all():
                    extra_count += len(self.additional_users_data.get_all())
                if set_hostname:
                    extra_count += 1
                total_steps += extra_count
                step_num = 0

                # Cancel checks at every step boundary (#12 hard-bail). Same as
                # new wizard: check before each step within phase + after each
                # phase to bail outer while.
                for desc, step_fn in auth_steps:
                    if self.cancel_flag: break
                    step_num += 1
                    self.log(f"[{step_num}/{total_steps}] {desc}")
                    if step_fn():
                        self.log("      ✓ Done.")
                    else:
                        self.log(f"      ✗ {desc} failed")
                        errors.append(desc.lower().split()[0])
                if self.cancel_flag:
                    self.log("Cancelled by user — bailing wizard.")
                    break

                # Phase 2: Additional users + hostname at CURRENT IP (before network change)
                if add_additional_users:
                    extra_users = self.additional_users_data.get_all()
                    if extra_users:
                        for eu in extra_users:
                            if self.cancel_flag: break
                            step_num += 1
                            eu_name = eu['username']
                            eu_pwd = eu['password']
                            eu_role = eu['role']
                            self.log(f"[{step_num}/{total_steps}] Creating user '{eu_name}' ({eu_role})")
                            result = self.protocol.add_user(camera_ip, password, eu_name, eu_pwd, eu_role)
                            if result:
                                self.log(f"      ✓ Done.")
                            else:
                                self.log(f"      ✗ Failed (may not be supported for {self.protocol.BRAND_NAME})")
                                errors.append(f"user:{eu_name}")
                if self.cancel_flag:
                    self.log("Cancelled by user — bailing wizard.")
                    break

                # Try to get serial while still at factory IP
                try:
                    pre_serial = self.protocol.get_serial(camera_ip, password)
                    if pre_serial and pre_serial != 'UNKNOWN':
                        cam['serial'] = pre_serial
                        if len(pre_serial) == 12:
                            cam['mac'] = ':'.join(pre_serial[j:j+2] for j in range(0, 12, 2))
                        self.log(f"Serial (pre-network): {pre_serial}")
                except:
                    pass

                if set_hostname and not self.cancel_flag:
                    step_num += 1
                    brand_prefix = self.protocol.BRAND_KEY
                    cam_number = cam.get('number', str(programmed_count))
                    s = cam.get('serial', 'unknown')
                    if s and s != 'UNKNOWN':
                        hostname = f"{cam_number}-{brand_prefix}-{s.lower()}"
                    else:
                        hostname = f"{cam_number}-{brand_prefix}-unknown"
                    self.log(f"[{step_num}/{total_steps}] Setting hostname: {hostname}")
                    result = self.protocol.set_hostname(camera_ip, password, hostname)
                    if result:
                        self.log("      ✓ Done.")
                        cam['hostname'] = hostname
                        cam['name'] = hostname
                    else:
                        self.log("      ✗ Hostname failed")
                        errors.append("hostname")
                if self.cancel_flag:
                    self.log("Cancelled by user — bailing wizard.")
                    break

                # Phase 3: Network change — LAST (camera may become unreachable).
                # Cancel mid-network-change can leave camera in partial state but
                # we still honor the operator's choice.
                for desc, step_fn in network_steps:
                    if self.cancel_flag:
                        self.log("  ⚠ Cancelled mid-network-change — camera may be in partial state.")
                        break
                    step_num += 1
                    self.log(f"[{step_num}/{total_steps}] {desc}")
                    if step_fn():
                        self.log("      ✓ Done.")
                    else:
                        self.log(f"      ✗ {desc} failed")
                        errors.append(desc.lower().split()[0])
                if self.cancel_flag:
                    self.log("Cancelled by user — bailing wizard.")
                    break

                # Track this MAC as programmed and release ARP pin
                if pinned_mac:
                    seen_macs.add(pinned_mac.upper().replace(':', '').replace('-', ''))
                    # Always save MAC from ARP immediately
                    if not cam.get('mac'):
                        cam['mac'] = pinned_mac
                self.arp_unpin(camera_ip)

                # Wait for camera to come back at new IP.
                # v4.3 #15 fix — Brian's empirical test 2026-05-03 confirmed
                # set_network DOES persist (HTTP responds at new IP within 15s
                # of the SOAP call), but pure-ICMP ping_camera was sometimes
                # failing the 45s window — likely some firmwares delay ICMP
                # echo response longer than HTTP comes up. Try ICMP first
                # (cheap), fall back to a no-auth HTTP probe via
                # protocol.probe_unrestricted (all brand classes implement it
                # — Axis returns rich data, others return model via base-class
                # fallback to get_model_noauth). Either succeeding proves the
                # camera is alive at the new IP.
                camera_reachable = False
                self.log(f"Waiting for camera at new IP ({static_ip})...")
                time.sleep(3)
                wait_count = 0
                while not self.cancel_flag and wait_count < 60:
                    if self.ping_camera(static_ip, timeout_ms=1500):
                        camera_reachable = True
                        break
                    # ICMP didn't answer — try HTTP probe (some firmwares
                    # serve HTTP before ICMP, or never serve ICMP)
                    try:
                        p = self.protocol.probe_unrestricted(static_ip)
                        if p and (p.get('mac') or p.get('model')):
                            camera_reachable = True
                            self.log(f"  Camera answered HTTP probe at {static_ip} (ICMP silent)")
                            break
                    except Exception:
                        pass
                    wait_count += 1
                    if wait_count % 10 == 0:
                        self.log(f"  Still waiting... ({wait_count}s)")
                    time.sleep(1)

                if self.cancel_flag:
                    break

                if not camera_reachable:
                    self.log(f"✗ Camera not responding at {static_ip} after {wait_count}s")
                    errors.append("unreachable")
                else:
                    self.log(f"Camera online at {static_ip} (after {wait_count + 3}s)")
                    time.sleep(1)

                # v4.3 #10 — ONVIF user teardown + optional rename:
                #   keep=False              → delete ONVIF root (default — clean)
                #   keep=True, no creds     → don't delete, root remains
                #   keep=True, custom creds → delete root, create custom (rename)
                # Classic's older ProgramOptionsDialog doesn't expose custom
                # creds yet, so the rename path never triggers from classic —
                # but the code is wired anyway so it Just Works if/when the
                # older dialog gets the fields.
                if camera_reachable and hasattr(self.protocol, 'delete_onvif_user'):
                    _r = (prog_opts.result if (prog_opts and prog_opts.result) else {})
                    keep_onvif = bool(_r.get('keep_onvif_user', False))
                    custom_u = (_r.get('onvif_username') or '').strip()
                    custom_p = (_r.get('onvif_password') or '').strip()
                    if not keep_onvif:
                        if self.protocol.delete_onvif_user(static_ip, password, 'root'):
                            self.log("ONVIF user deleted — camera ends with VAPIX root only.")
                        else:
                            self.log("⚠ ONVIF user delete failed — leftover ONVIF account on camera.")
                    elif custom_u and custom_p and hasattr(self.protocol, 'add_onvif_user'):
                        if not self.protocol.delete_onvif_user(static_ip, password, 'root'):
                            self.log("⚠ Could not delete ONVIF root before custom-user create — proceeding anyway.")
                        if self.protocol.add_onvif_user(static_ip, password, custom_u, custom_p):
                            self.log(f"ONVIF user replaced: 'root' → '{custom_u}'.")
                        else:
                            self.log(f"⚠ Could not create ONVIF user '{custom_u}' — camera may have NO ONVIF user.")
                    else:
                        self.log("ONVIF user kept as 'root' (operator-requested).")

                # Get serial/MAC if camera is reachable
                serial = 'UNKNOWN'
                if camera_reachable:
                    serial = self.protocol.get_serial(static_ip, password)
                    self.log(f"Serial: {serial}")

                    if serial and serial != 'UNKNOWN':
                        cam['serial'] = serial
                        if len(serial) == 12:
                            cam['mac'] = ':'.join(serial[j:j+2] for j in range(0, 12, 2))
                            self.log(f"MAC: {cam['mac']}")
                    elif pinned_mac:
                        cam['mac'] = pinned_mac
                        self.log(f"MAC (from ARP): {pinned_mac}")
                else:
                    # Not reachable (different subnet or timed out) — use ARP MAC
                    if pinned_mac:
                        cam['mac'] = pinned_mac
                        mac_clean = pinned_mac.upper().replace(':', '').replace('-', '')
                        cam['serial'] = mac_clean
                        self.log(f"MAC (from ARP): {pinned_mac}")
                        serial = mac_clean
                
                # Update model from actual if entry had none
                if actual_model and not expected_model:
                    cam['model'] = actual_model

                # Post-network: try to get image if reachable + retry firmware with auth
                if camera_reachable:
                    img = self.protocol.get_image(static_ip, username, password)
                    if img:
                        self.root.after(0, lambda d=img: self.update_preview(d))
                    if actual_firmware == 'UNKNOWN':
                        try:
                            fw = self.protocol.get_firmware(static_ip, password)
                            if fw and fw != 'UNKNOWN':
                                actual_firmware = fw
                                self.log(f"Camera firmware: {actual_firmware}")
                        except Exception:
                            pass

                cam['firmware'] = actual_firmware

                # Save updated camera data
                self.camera_data.save()
                self.root.after(0, self.refresh_camera_list)

                # v4.3 #13 — assign Building Reports sticker number if enabled.
                # Assigned BEFORE CSV write so the row carries the label. Only
                # increments on successful programming (errors → no sticker
                # assigned, no counter bump, operator can re-program later).
                br_label = ''
                if br_enabled and br_counter is not None and not errors:
                    br_label = str(br_counter)
                    self.log(f"Building Reports sticker: #{br_label}  ← peel and apply to {cam.get('name', cam_name)}")
                    br_counter += 1

                # Save to output CSV (always — success or partial fail)
                cam_mac = cam.get('mac', pinned_mac or '')
                with open(OUTPUT_CSV, 'a', newline='') as f:
                    csv.writer(f).writerow([cam.get('name', cam_name), static_ip, serial, cam_mac,
                                           actual_model or expected_model, actual_firmware,
                                           br_label,
                                           datetime.now().isoformat()])

                # Mark based on results
                cam_mac = cam.get('mac', pinned_mac or 'unknown')
                if 'unreachable' in errors:
                    self._mark_cam_failed(cam, ', '.join(errors))
                    total_fail += 1
                    self.log(f"\n*** CAMERA {cam_name} FAILED: {', '.join(errors)} ***")
                    self.log(f"    IP: {static_ip}, Serial: {serial}, MAC: {cam_mac}")
                elif errors:
                    self._mark_cam_failed(cam, ', '.join(errors))
                    total_fail += 1
                    self.log(f"\n*** CAMERA {cam_name} PARTIAL FAIL: {', '.join(errors)} ***")
                    self.log(f"    IP: {static_ip}, Serial: {serial}, MAC: {cam_mac}")
                else:
                    idx = self.camera_data.get_all().index(cam)
                    self.camera_data.mark_processed(idx)
                    total_ok += 1
                    self.log(f"\n*** CAMERA {cam_name} COMPLETE ***")
                    self.log(f"    IP: {static_ip}, DHCP: DISABLED, Serial: {serial}, MAC: {cam_mac}")
                
                # Remove from remaining list
                remaining.pop(cam_idx)
                
                # Show continue dialog if more cameras
                if remaining:
                    next_name = remaining[0]['name']
                    next_model = remaining[0].get('model', '')
                    status_msg = f"Camera {cam_name} {'complete' if not errors else 'PARTIAL FAIL'}!"
                    result = [None]
                    def show():
                        result[0] = ContinueDialog(self.root, status_msg,
                                                   next_name, next_model, img).result
                    self.root.after(0, show)
                    while result[0] is None and not self.cancel_flag:
                        time.sleep(0.1)
                    if not result[0]:
                        self.cancel_flag = True
                        break
                    self.clear_preview()

                    # 2026-04-30 hot fix: even after user confirms continue, ensure
                    # the previous camera has actually left factory IP before next
                    # discovery. Otherwise a too-quick "yes" click while the camera
                    # is still plugged in causes the next slot to program onto it.
                    if factory_ip:
                        wait_start = time.time()
                        last_heartbeat = wait_start
                        announced = False
                        while not self.cancel_flag:
                            if not self.ping_camera(factory_ip, timeout_ms=800):
                                if announced:
                                    self.log(f"  Previous camera left factory IP after {int(time.time() - wait_start)}s")
                                break
                            now_t = time.time()
                            if not announced and now_t - wait_start >= 5:
                                self.log(f"  Waiting for {cam_name} to leave factory IP {factory_ip} (unplug it / let it reboot)...")
                                announced = True
                                last_heartbeat = now_t
                            elif announced and now_t - last_heartbeat >= 30:
                                self.log(f"  ...still at factory IP ({int(now_t - wait_start)}s elapsed)")
                                last_heartbeat = now_t
                            time.sleep(2)

            # Clean up link-local route if we added one
            self.remove_linklocal_route()

            self.log(f"\n{'='*50}")
            self.log(f"PROGRAMMING COMPLETE: {total_ok} succeeded, {total_fail} failed")
            if remaining:
                self.log(f"  {len(remaining)} cameras not programmed")
            self.log(f"Results saved to {OUTPUT_CSV}")
            self.log(f"{'='*50}")
            self.root.after(0, lambda: self.update_display("DONE", f"{total_ok} OK, {total_fail} failed"))
            self.root.after(0, lambda: self.enable_cancel(False))
            self.root.after(0, self.refresh_camera_list)
            self.root.after(0, self.rescan_after_operation)
            self.clear_preview()
            self._close_wizard_run_log()

        threading.Thread(target=run, daemon=True).start()

    # ========================================================================
    # NEW PROGRAMMING FLOW (step-by-step wizard + live status view)
    # ========================================================================
    def start_program_wizard(self):
        """Step-by-step programming flow with live checklist UI."""
        cameras = self.validate_cameras_for_programming()
        if not cameras:
            return

        # Show wizard
        wiz = ProgramWizardDialog(self.root,
            brand_name=self.protocol.BRAND_NAME,
            factory_ip=self.protocol.FACTORY_IP,
            camera_count=len(cameras),
            additional_users_count=len(self.additional_users_data.get_all()))
        if not wiz.result:
            return

        opts = wiz.result
        password = opts['password']
        factory_ip = opts['factory_ip']
        discovery_mode = opts['discovery_mode']
        set_hostname = opts['set_hostname']
        add_additional_users = opts['add_additional_users']
        selected_iface = opts['interface']

        # Persist factory IP if user changed it
        if factory_ip and factory_ip != self.protocol.FACTORY_IP:
            self.protocol.FACTORY_IP = factory_ip
            try:
                key = 'bosch_factory_ip' if self.protocol.BRAND_KEY == 'bosch' else 'factory_ip'
                self.settings.set('general', key, factory_ip)
            except Exception:
                pass

        # Override interface if user picked one
        if selected_iface:
            self._detected_iface_index = selected_iface['index']
            self._detected_local_ip = selected_iface['ip']
            self.settings.set('general', 'interface_index', str(selected_iface['index']))

        # Determine which checklist steps will actually run for this config
        used_steps = ['discover', 'pin', 'verify_model', 'firmware', 'auth']
        if add_additional_users and self.additional_users_data.get_all():
            used_steps.append('extra_users')
        if set_hostname:
            used_steps.append('hostname')
        used_steps.extend(['network', 'verify_online', 'capture'])

        # Switch to status tab and prep the UI
        self.notebook.select(self.status_tab)
        self.cancel_flag = False
        self.enable_cancel(True)
        self.status_enable_cancel(True)
        self.root.after(0, lambda: self.status_reset_steps(used_steps))
        self.root.after(0, lambda: self.status_set_camera('—', f'0 of {len(cameras)} done'))
        self.root.after(0, lambda: self.status_set_banner(
            'STARTING…', f'Programming {len(cameras)} {self.protocol.BRAND_NAME} camera(s)', '#2196F3'))
        self.root.after(0, lambda: self.status_set_preview(None))

        def _ui(fn, *args, **kwargs):
            self.root.after(0, lambda: fn(*args, **kwargs))

        def run():
            wizard_log_path = self._open_wizard_run_log()
            username = self.protocol.DEFAULT_USER
            total_ok = total_fail = 0

            # v4.3 #13 — Building Reports sticker counter
            br_enabled = bool(opts.get('add_br_stickers', False))
            br_counter = None
            if br_enabled:
                try:
                    br_counter = int(str(opts.get('br_first_label', '')).strip())
                except (ValueError, TypeError):
                    br_enabled = False
                    self.status_log("⚠ Building Reports stickers requested but seed value is not a number — skipping.")

            _ensure_output_csv_header()

            if wizard_log_path:
                self.status_log(f"Wizard run log: {wizard_log_path}")
            if br_enabled:
                self.status_log(f"Building Reports stickers: starting at #{br_counter}")
            self.status_log(f"Discovery mode: {discovery_mode}")
            if factory_ip:
                self.status_log(f"Factory IP: {factory_ip}")

            if discovery_mode in ('mdns', 'both'):
                self.status_log("Adding link-local route for camera discovery...")
                self.add_linklocal_route()

            remaining = list(cameras)
            programmed_count = 0
            seen_macs = set()
            consecutive_skips = 0

            # Banner suppression: only print "Waiting for camera N" once per real new
            # camera. When we re-enter the outer loop after waiting for the previous
            # camera to leave factory IP (already-programmed branch below), the slot
            # hasn't changed and the user shouldn't see "Waiting for camera 17: KBW-005"
            # flashing past — that was the 2026-04-30 bug. last_announced_name tracks
            # what we last announced; only re-announce when remaining[0] actually changes.
            last_announced_name = None

            while remaining and not self.cancel_flag:
                pinned_mac = None
                camera_ip = None

                next_cam = remaining[0]
                next_name = next_cam['name']
                next_model = next_cam.get('model', '')

                if next_name != last_announced_name:
                    # Real new slot — bump counter, print banner + log header.
                    programmed_count += 1
                    _ui(self.status_set_banner, 'PLUG IN CAMERA',
                        f"Plug in: {next_name}" + (f"  ({next_model})" if next_model else ''),
                        '#FF9800')
                    _ui(self.status_set_camera, next_name,
                        f"{programmed_count - 1} of {len(cameras)} done · {len(remaining)} remaining")
                    _ui(self.status_reset_steps, used_steps)
                    _ui(self.status_set_step, 'discover', 'active')
                    self.status_log(f"\n{'=' * 50}")
                    self.status_log(f"Waiting for camera {programmed_count}: {next_name}")
                    self.status_log(f"{'=' * 50}")
                    last_announced_name = next_name
                else:
                    # Re-entry after wait-for-reboot. Stay quiet, just put discover step
                    # back to active state.
                    _ui(self.status_set_step, 'discover', 'active')

                # ---- Discovery phase ----
                while not self.cancel_flag:
                    if discovery_mode in ('factory', 'both'):
                        if factory_ip and self.ping_camera(factory_ip, timeout_ms=1000):
                            camera_ip = factory_ip
                            self.status_log(f"Camera found at factory IP {factory_ip}")
                            break

                    if discovery_mode in ('mdns', 'both'):
                        dhcp_found_mac = None
                        try:
                            dhcp_cams = AxisDHCPDiscovery.discover(timeout=3)
                            for dc in dhcp_cams:
                                dc_mac = dc.get('mac', '').upper().replace(':', '').replace('-', '')
                                if dc_mac and dc_mac not in seen_macs:
                                    dhcp_found_mac = dc.get('mac', '')
                                    self.status_log(f"DHCP broadcast: {dc.get('model', '?')} MAC {dhcp_found_mac}")
                                    self.add_linklocal_route()
                                    break
                        except Exception:
                            pass

                        if dhcp_found_mac and not camera_ip:
                            try:
                                if self.ping_camera(factory_ip, timeout_ms=1000):
                                    info = self.protocol.get_discovery_info(factory_ip, timeout=2)
                                    if info:
                                        camera_ip = factory_ip
                                        pinned_mac = dhcp_found_mac
                                        self.status_log(f"Camera at {factory_ip}  Model: {info.get('model', '?')}")
                                        break
                            except Exception:
                                pass

                        if not camera_ip:
                            try:
                                resolved_cams = self._resolve_linklocal_cameras(
                                    target_mac=dhcp_found_mac, timeout=4)
                                for mc in resolved_cams:
                                    mc_ip = mc.get('ip', '')
                                    mc_mac = mc.get('mac', '').upper().replace(':', '').replace('-', '')
                                    if mc_mac and mc_mac in seen_macs:
                                        continue
                                    if mc_ip and self.ping_camera(mc_ip, timeout_ms=1000):
                                        camera_ip = mc_ip
                                        pinned_mac = mc.get('mac', '')
                                        self.status_log(f"Camera at {mc_ip}  Model: {mc.get('model', '?')}")
                                        break
                                if camera_ip:
                                    break
                            except Exception:
                                pass

                        if not camera_ip:
                            try:
                                mdns_cams = AxisMDNSDiscovery.discover(timeout=2)
                                for mc in mdns_cams:
                                    mc_ip = mc.get('ip', '')
                                    mc_mac = mc.get('mac', '').upper().replace(':', '').replace('-', '')
                                    if mc_mac and mc_mac in seen_macs:
                                        continue
                                    if mc_ip.startswith('169.254.'):
                                        self.add_linklocal_route()
                                    if mc_ip and self.ping_camera(mc_ip, timeout_ms=1000):
                                        camera_ip = mc_ip
                                        pinned_mac = mc.get('mac', '')
                                        self.status_log(f"Camera (mDNS): {mc_ip}  Model: {mc.get('model', '?')}")
                                        break
                                if camera_ip:
                                    break
                            except Exception:
                                pass

                    time.sleep(1)

                if self.cancel_flag:
                    break

                # ---- MAC freshness gate (v4.3 #9) ----
                # Don't announce "programming new camera" until the backend has
                # confirmed the discovered camera's MAC is NOT in seen_macs.
                # Brian 2026-05-03: previous fix (pre_probe + seen_macs) leaked
                # because if probe FAILED to return a MAC at all, the seen_macs
                # check was silently bypassed. Inverted logic: REQUIRE a fresh
                # MAC before proceeding. No fresh MAC → keep polling silently.
                # This also covers the "camera doing a quick reset, transient
                # appearance during reboot" case Brian flagged: the just-
                # programmed cam keeps showing up, but its MAC stays in
                # seen_macs, so the wizard quietly waits until operator
                # actually unplugs A and plugs in B.
                mac_settled = False
                last_heartbeat = time.time()
                while not self.cancel_flag and not mac_settled:
                    pre_probe = self.protocol.probe_unrestricted(camera_ip)
                    mac = pre_probe.get('mac')
                    if mac:
                        mac_norm = mac.upper().replace(':', '').replace('-', '')
                        if mac_norm not in seen_macs:
                            # Fresh camera! Proceed.
                            pinned_mac = mac
                            mac_settled = True
                            break
                    # MAC unreachable OR stale → keep waiting silently
                    if not self.ping_camera(camera_ip, timeout_ms=1000):
                        # Camera left this IP — re-discover (might be a new IP now)
                        self.status_log(f"Camera left {camera_ip} — re-running discovery...")
                        camera_ip = None
                        pinned_mac = None
                        _ui(self.status_set_step, 'discover', 'active')
                        break
                    # Heartbeat every 30s so operator knows we're alive
                    now_t = time.time()
                    if now_t - last_heartbeat >= 30:
                        seen_str = (mac.upper() if mac else 'unreadable')
                        self.status_log(f"  Waiting for new camera... (current at {camera_ip}: {seen_str} — already programmed; unplug + plug next)")
                        last_heartbeat = now_t
                    time.sleep(2)

                if not mac_settled:
                    # Inner loop bailed via re-discover OR cancel — outer loop continues
                    continue

                _ui(self.status_set_step, 'discover', 'ok', camera_ip)
                _ui(self.status_set_banner, 'PROGRAMMING…',
                    f"{next_name}  →  working on it", '#2196F3')

                # ---- Factory-default-before-program (v4.3 reuse-camera workflow) ----
                # Operator opted to wipe the camera before applying the new
                # config. Wipes via the existing root password they provided.
                # Solves the case where Brian's reusing a camera from a prior
                # site (has old password + old config) and wants a clean program.
                # Also addresses the "couldn't hold the factory-reset button
                # before wizard wrote to it" race — wizard does the wipe FOR you.
                factory_first = bool(opts.get('factory_first', False))
                existing_pwd = opts.get('existing_root_pwd') or ''
                if factory_first and existing_pwd and hasattr(self.protocol, 'factory_reset'):
                    self.status_log(f"Factory-resetting {pinned_mac} via existing password...")
                    if self.protocol.factory_reset(camera_ip, existing_pwd):
                        self.status_log("  ✓ Reset issued. Waiting for camera to come back...")
                        # Camera reboots, loses everything. Try original IP first
                        # (might come back same via DHCP), then 192.168.0.90 fallback.
                        old_ip = camera_ip
                        target_mac_norm = pinned_mac.upper().replace(':', '').replace('-', '')
                        camera_ip = None
                        for attempt in range(60):  # up to 120s
                            if self.cancel_flag: break
                            for try_ip in (old_ip, '192.168.0.90'):
                                if self.ping_camera(try_ip, timeout_ms=1000):
                                    p = self.protocol.probe_unrestricted(try_ip)
                                    p_mac = (p.get('mac') or '').upper().replace(':', '').replace('-', '')
                                    if p_mac == target_mac_norm:
                                        camera_ip = try_ip
                                        break
                            if camera_ip:
                                break
                            time.sleep(2)
                        if not camera_ip:
                            self.status_log("  ✗ Camera didn't reappear within 120s — bailing this slot")
                            errors.append('factory_reset_no_return')
                            continue
                        self.status_log(f"  ✓ Camera back at {camera_ip} (factory state)")
                    else:
                        self.status_log("  ✗ Factory reset failed — wrong existing password? bailing this slot")
                        errors.append('factory_reset_failed')
                        continue

                # ---- ARP pin ----
                _ui(self.status_set_step, 'pin', 'active')
                if not pinned_mac:
                    pinned_mac = self.get_mac_from_arp(camera_ip)
                if pinned_mac:
                    if pinned_mac.upper().replace(':', '').replace('-', '') in seen_macs:
                        # Previously-programmed camera is still on the factory IP because it
                        # hasn't finished rebooting yet. Wait IN PLACE until it's gone — DO NOT
                        # bounce back to the outer loop, that re-banners and increments the
                        # camera counter every 3s ("Waiting for camera 17: KBW-005" flashing).
                        # Bug fix 2026-04-30 — Brian saw the toolkit appear to flash through
                        # 17 cameras in seconds after one real success.
                        self.status_log(f"MAC {pinned_mac} just programmed — waiting for it to leave factory IP {camera_ip}...")
                        self.arp_unpin(camera_ip)
                        _ui(self.status_set_step, 'pin', 'pending', 'reboot...')
                        wait_start = time.time()
                        last_heartbeat = wait_start
                        while not self.cancel_flag:
                            # Camera off factory IP? -> previous one rebooted, retry discovery
                            if not self.ping_camera(camera_ip, timeout_ms=800):
                                self.status_log(f"  Old camera left factory IP after {int(time.time() - wait_start)}s — ready for next one")
                                break
                            # Same MAC still there? Heartbeat every 30s, then keep waiting.
                            now_t = time.time()
                            if now_t - last_heartbeat >= 30:
                                self.status_log(f"  ...still rebooting ({int(now_t - wait_start)}s elapsed). Plug in next camera if ready.")
                                last_heartbeat = now_t
                            time.sleep(2)
                        # Reset state and re-enter discovery for THIS slot (no counter bump).
                        camera_ip = None
                        pinned_mac = None
                        _ui(self.status_set_step, 'discover', 'active')
                        continue
                    if self.arp_pin(camera_ip, pinned_mac):
                        self.status_log(f"ARP pinned to {pinned_mac}")
                        _ui(self.status_set_step, 'pin', 'ok', pinned_mac)
                    else:
                        self.status_log(f"ARP pin failed for {pinned_mac} (continuing)")
                        _ui(self.status_set_step, 'pin', 'ok', pinned_mac)
                else:
                    self.status_log("No MAC from ARP")
                    _ui(self.status_set_step, 'pin', 'fail', 'no MAC')

                # ---- Verify model ----
                # Single no-auth probe via basicdeviceinfo.cgi getAllUnrestrictedProperties
                # — works on factory AND password-locked cameras. Returns model + serial
                # (which IS the MAC for Axis) + firmware + hardware in one round trip.
                # This eliminates the "expected model, got ?" path when ARP pin missed:
                # the probe always gives us MAC if the camera is reachable on HTTP.
                _ui(self.status_set_step, 'verify_model', 'active')
                probe = self.protocol.probe_unrestricted(camera_ip)
                actual_model = probe.get('model')
                if actual_model:
                    self.status_log(f"Camera model: {actual_model}")
                else:
                    self.status_log("Could not query model — will match any entry")
                # If ARP didn't pin a MAC, harvest from the probe and pin retroactively
                if probe.get('mac') and not pinned_mac:
                    pinned_mac = probe['mac']
                    self.status_log(f"MAC (from probe, ARP missed): {pinned_mac}")
                    if self.arp_pin(camera_ip, pinned_mac):
                        _ui(self.status_set_step, 'pin', 'ok', pinned_mac)

                # ---- Read firmware (use probe value if we got it, else fall back) ----
                _ui(self.status_set_step, 'firmware', 'active')
                actual_firmware = probe.get('firmware') or 'UNKNOWN'
                if not probe.get('firmware'):
                    try:
                        actual_firmware = self.protocol.get_firmware(camera_ip, '') or 'UNKNOWN'
                    except Exception:
                        actual_firmware = 'UNKNOWN'
                if actual_firmware and actual_firmware != 'UNKNOWN':
                    self.status_log(f"Firmware: {actual_firmware}")
                    _ui(self.status_set_step, 'firmware', 'ok', actual_firmware)
                else:
                    _ui(self.status_set_step, 'firmware', 'pending', '(retry after auth)')

                # ---- Match camera entry ----
                cam = None
                cam_idx = None
                if actual_model:
                    norm_actual = actual_model.upper().replace('AXIS-', '').replace('AXIS ', '').strip()
                    for idx, c in enumerate(remaining):
                        c_model = c.get('model', '')
                        if not c_model:
                            cam = c
                            cam_idx = idx
                            break
                        norm_expected = c_model.upper().replace('AXIS-', '').replace('AXIS ', '').strip()
                        if norm_expected in norm_actual or norm_actual in norm_expected:
                            cam = c
                            cam_idx = idx
                            break

                    if not cam:
                        # Defensive: actual_model could theoretically be empty here even
                        # though we entered this branch on truthy actual_model — guard so
                        # the UI never shows "got None" or bare "got " (Brian's "got ?")
                        shown_model = actual_model if actual_model else "(probe failed)"
                        self.status_log(f"⚠ MODEL MISMATCH: got {shown_model}")
                        _ui(self.status_set_step, 'verify_model', 'fail', f'got {shown_model}')
                        _ui(self.status_set_banner, 'WRONG MODEL',
                            f'Got {shown_model} — plug in a different camera', '#F44336')
                        self.arp_unpin(camera_ip)
                        consecutive_skips += 1

                        result = [None]
                        skip_count = consecutive_skips
                        models_needed = sorted(set(c.get('model', '(any)') for c in remaining))
                        def show_mismatch():
                            msg = (f"Wrong camera model detected!\n\n"
                                   f"Connected: {actual_model}\n\n"
                                   f"Models still needed:\n")
                            model_counts = {}
                            for c in remaining:
                                m = c.get('model', '(any)')
                                model_counts[m] = model_counts.get(m, 0) + 1
                            for m, count in sorted(model_counts.items()):
                                msg += f"  • {m}  ×{count}\n"
                            msg += (f"\n({skip_count} mismatch{'es' if skip_count != 1 else ''} so far)\n\n"
                                    "Try again?")
                            result[0] = messagebox.askyesno("Wrong Camera Model", msg)
                        self.root.after(0, show_mismatch)
                        while result[0] is None and not self.cancel_flag:
                            time.sleep(0.1)
                        if not result[0]:
                            self.cancel_flag = True
                            break
                        time.sleep(2)
                        continue
                else:
                    # Probe returned no model at all — camera mid-reboot, weird
                    # firmware, network glitch, etc. v4.3 #8: replace the silent
                    # "take first remaining" fallback with an explicit confirm
                    # so the operator chooses what to do. askyesnocancel:
                    #   Yes    = Try again (re-probe the same slot — camera
                    #            may respond on the next pass)
                    #   No     = Proceed anyway, assume next-in-list
                    #   Cancel = Bail the whole run
                    next_name = remaining[0].get('name', '?')
                    next_model = remaining[0].get('model', '(any)') or '(any)'
                    result = [None]
                    def show_no_model():
                        msg = (f"Could not identify the camera at {camera_ip}.\n\n"
                               f"Probe returned no model — camera may be rebooting, "
                               f"unreachable on HTTP, or have firmware that doesn't "
                               f"answer basicdeviceinfo.cgi.\n\n"
                               f"Yes    → Try again (re-probe — camera may respond now)\n"
                               f"No     → Proceed anyway, treat as: {next_name} ({next_model})\n"
                               f"Cancel → Bail the whole wizard run")
                        result[0] = messagebox.askyesnocancel(
                            "Cannot Confirm Camera Model", msg)
                    self.root.after(0, show_no_model)
                    while result[0] is None and not self.cancel_flag:
                        time.sleep(0.1)
                    if self.cancel_flag or result[0] is None:
                        # Cancel pressed — bail
                        self.cancel_flag = True
                        self.status_log("Cancelled by user — bailing wizard.")
                        break
                    if result[0] is True:
                        # Try again — re-loop discovery for this slot
                        self.status_log("Retrying probe for this camera slot...")
                        self.arp_unpin(camera_ip)
                        camera_ip = None
                        pinned_mac = None
                        _ui(self.status_set_step, 'discover', 'active')
                        time.sleep(2)
                        continue
                    # result[0] is False → Proceed anyway with first remaining
                    self.status_log(f"Proceeding without model verification — assuming {next_name}.")
                    cam = remaining[0]
                    cam_idx = 0

                consecutive_skips = 0
                _ui(self.status_set_step, 'verify_model', 'ok', actual_model or '(unverified)')

                cam_name = cam['name']
                static_ip = cam['ip']
                gateway = cam['gateway']
                subnet = cam['subnet']
                expected_model = cam.get('model', '')
                cidr = self.subnet_to_cidr(subnet)
                errors = []

                _ui(self.status_set_camera, cam_name,
                    f"Programming → {static_ip}  ({programmed_count} of {len(cameras)})")
                self.status_log(f"\nAssigned to: {cam_name} → {static_ip}")
                self.status_log(f"Programming from: {camera_ip}" + (" [link-local]" if camera_ip.startswith('169.254.') else ""))

                # ---- Build steps and split ----
                cam['_program_ip'] = camera_ip
                steps = self.protocol.get_programming_steps(cam, password)
                network_keywords = ('gateway', 'network', 'ip', 'dhcp')
                auth_steps = []
                network_steps = []
                for desc, fn in steps:
                    desc_lower = desc.lower()
                    if any(kw in desc_lower for kw in network_keywords):
                        network_steps.append((desc, fn))
                    else:
                        auth_steps.append((desc, fn))

                # ---- Phase 1: Auth ----
                # Cancel checks at every step boundary (#12 hard-bail). Brian's
                # 2026-04-30 complaint: cancel was advancing to the next step
                # ("Cancel" mid-"Attempting to write user" → moved to "Setting
                # IP" anyway). Now: check before each step, and check after the
                # phase to bail the outer while loop immediately.
                _ui(self.status_set_step, 'auth', 'active')
                auth_ok = True
                for desc, step_fn in auth_steps:
                    if self.cancel_flag: break
                    self.status_log(f"  {desc}")
                    if step_fn():
                        self.status_log("    ✓ Done.")
                    else:
                        self.status_log(f"    ✗ {desc} failed")
                        errors.append(desc.lower().split()[0])
                        auth_ok = False
                _ui(self.status_set_step, 'auth', 'ok' if auth_ok else 'fail')
                if self.cancel_flag:
                    self.status_log("Cancelled by user — bailing wizard.")
                    break

                # ---- Phase 2a: Additional users ----
                if add_additional_users:
                    extra_users = self.additional_users_data.get_all()
                    if extra_users:
                        _ui(self.status_set_step, 'extra_users', 'active')
                        eu_ok = True
                        for eu in extra_users:
                            if self.cancel_flag: break
                            self.status_log(f"  Creating user '{eu['username']}' ({eu['role']})")
                            if self.protocol.add_user(camera_ip, password,
                                                       eu['username'], eu['password'], eu['role']):
                                self.status_log("    ✓ Done.")
                            else:
                                self.status_log(f"    ✗ Failed")
                                errors.append(f"user:{eu['username']}")
                                eu_ok = False
                        _ui(self.status_set_step, 'extra_users', 'ok' if eu_ok else 'fail')
                        if self.cancel_flag:
                            self.status_log("Cancelled by user — bailing wizard.")
                            break

                # Try to get serial while still at factory IP
                try:
                    pre_serial = self.protocol.get_serial(camera_ip, password)
                    if pre_serial and pre_serial != 'UNKNOWN':
                        cam['serial'] = pre_serial
                        if len(pre_serial) == 12:
                            cam['mac'] = ':'.join(pre_serial[j:j+2] for j in range(0, 12, 2))
                        self.status_log(f"Pre-network serial: {pre_serial}")
                except Exception:
                    pass

                # ---- Phase 2b: Hostname ----
                if set_hostname and not self.cancel_flag:
                    _ui(self.status_set_step, 'hostname', 'active')
                    brand_prefix = self.protocol.BRAND_KEY
                    cam_number = cam.get('number', str(programmed_count))
                    s = cam.get('serial', 'unknown')
                    if s and s != 'UNKNOWN':
                        hostname = f"{cam_number}-{brand_prefix}-{s.lower()}"
                    else:
                        hostname = f"{cam_number}-{brand_prefix}-unknown"
                    self.status_log(f"  Setting hostname: {hostname}")
                    if self.protocol.set_hostname(camera_ip, password, hostname):
                        self.status_log("    ✓ Done.")
                        cam['hostname'] = hostname
                        cam['name'] = hostname
                        _ui(self.status_set_step, 'hostname', 'ok', hostname)
                    else:
                        self.status_log("    ✗ Hostname failed")
                        errors.append("hostname")
                        _ui(self.status_set_step, 'hostname', 'fail')
                if self.cancel_flag:
                    self.status_log("Cancelled by user — bailing wizard.")
                    break

                # ---- Phase 3: Network change ----
                # Network change is the riskiest cancel point — partial set_network
                # leaves the camera in a half-configured state. Still honor the
                # cancel; operator chose to bail. Camera may need factory reset.
                _ui(self.status_set_step, 'network', 'active')
                net_ok = True
                for desc, step_fn in network_steps:
                    if self.cancel_flag:
                        self.status_log("  ⚠ Cancelled mid-network-change — camera may be in partial state.")
                        break
                    self.status_log(f"  {desc}")
                    if step_fn():
                        self.status_log("    ✓ Done.")
                    else:
                        self.status_log(f"    ✗ {desc} failed")
                        errors.append(desc.lower().split()[0])
                        net_ok = False
                _ui(self.status_set_step, 'network', 'ok' if net_ok else 'fail')
                if self.cancel_flag:
                    self.status_log("Cancelled by user — bailing wizard.")
                    break

                # Track this MAC and release pin
                if pinned_mac:
                    seen_macs.add(pinned_mac.upper().replace(':', '').replace('-', ''))
                    if not cam.get('mac'):
                        cam['mac'] = pinned_mac
                self.arp_unpin(camera_ip)

                # ---- Wait for camera at new IP ----
                # v4.3 #15 fix — see classic wizard for full rationale. ICMP
                # first (cheap), HTTP probe via protocol.probe_unrestricted
                # as the reliable backup (works on all brands).
                _ui(self.status_set_step, 'verify_online', 'active')
                camera_reachable = False
                self.status_log(f"Waiting for camera at {static_ip}...")
                time.sleep(3)
                wait_count = 0
                while not self.cancel_flag and wait_count < 60:
                    if self.ping_camera(static_ip, timeout_ms=1500):
                        camera_reachable = True
                        break
                    try:
                        p = self.protocol.probe_unrestricted(static_ip)
                        if p and (p.get('mac') or p.get('model')):
                            camera_reachable = True
                            self.status_log(f"  Camera answered HTTP probe at {static_ip} (ICMP silent)")
                            break
                    except Exception:
                        pass
                    wait_count += 1
                    if wait_count % 10 == 0:
                        self.status_log(f"  Still waiting... ({wait_count}s)")
                    time.sleep(1)

                if self.cancel_flag:
                    break

                if not camera_reachable:
                    self.status_log(f"✗ Camera not responding at {static_ip} after {wait_count}s")
                    errors.append("unreachable")
                    _ui(self.status_set_step, 'verify_online', 'fail', f'{wait_count}s timeout')
                else:
                    self.status_log(f"Camera online at {static_ip} (after {wait_count + 3}s)")
                    _ui(self.status_set_step, 'verify_online', 'ok', f'{wait_count + 3}s')
                    time.sleep(1)

                # v4.3 #10: ONVIF user teardown — see classic wizard for full
                # rationale. New wizard's ProgramWizardDialog exposes the
                # "Keep ONVIF user" checkbox + optional custom user/password.
                # Behavior matrix:
                #   keep=False           → delete ONVIF root (default — clean state)
                #   keep=True, no creds  → don't delete, root remains
                #   keep=True, custom    → delete root, create custom (rename)
                if camera_reachable and hasattr(self.protocol, 'delete_onvif_user'):
                    keep_onvif = bool(opts.get('keep_onvif_user', False))
                    custom_u = (opts.get('onvif_username') or '').strip()
                    custom_p = (opts.get('onvif_password') or '').strip()
                    if not keep_onvif:
                        if self.protocol.delete_onvif_user(static_ip, password, 'root'):
                            self.status_log("ONVIF user deleted — camera ends with VAPIX root only.")
                        else:
                            self.status_log("⚠ ONVIF user delete failed — leftover ONVIF account on camera.")
                    elif custom_u and custom_p and hasattr(self.protocol, 'add_onvif_user'):
                        # Rename pattern: delete transient root, create operator-named user
                        if not self.protocol.delete_onvif_user(static_ip, password, 'root'):
                            self.status_log("⚠ Could not delete ONVIF root before custom-user create — proceeding anyway.")
                        if self.protocol.add_onvif_user(static_ip, password, custom_u, custom_p):
                            self.status_log(f"ONVIF user replaced: 'root' → '{custom_u}'.")
                        else:
                            self.status_log(f"⚠ Could not create ONVIF user '{custom_u}' — camera may have NO ONVIF user.")
                    else:
                        self.status_log("ONVIF user kept as 'root' (operator-requested).")

                # ---- Capture serial / MAC / image ----
                _ui(self.status_set_step, 'capture', 'active')
                serial = 'UNKNOWN'
                if camera_reachable:
                    serial = self.protocol.get_serial(static_ip, password)
                    self.status_log(f"Serial: {serial}")
                    if serial and serial != 'UNKNOWN':
                        cam['serial'] = serial
                        if len(serial) == 12:
                            cam['mac'] = ':'.join(serial[j:j+2] for j in range(0, 12, 2))
                            self.status_log(f"MAC: {cam['mac']}")
                    elif pinned_mac:
                        cam['mac'] = pinned_mac
                        self.status_log(f"MAC (from ARP): {pinned_mac}")
                else:
                    if pinned_mac:
                        cam['mac'] = pinned_mac
                        mac_clean = pinned_mac.upper().replace(':', '').replace('-', '')
                        cam['serial'] = mac_clean
                        self.status_log(f"MAC (from ARP): {pinned_mac}")
                        serial = mac_clean

                if actual_model and not expected_model:
                    cam['model'] = actual_model

                if camera_reachable:
                    img = self.protocol.get_image(static_ip, username, password)
                    if img:
                        _ui(self.status_set_preview, img)
                    # Retry firmware with auth if we didn't get it before
                    if actual_firmware == 'UNKNOWN':
                        try:
                            fw = self.protocol.get_firmware(static_ip, password)
                            if fw and fw != 'UNKNOWN':
                                actual_firmware = fw
                                self.status_log(f"Firmware: {actual_firmware}")
                                _ui(self.status_set_step, 'firmware', 'ok', actual_firmware)
                        except Exception:
                            pass

                if actual_firmware == 'UNKNOWN':
                    _ui(self.status_set_step, 'firmware', 'fail', 'UNKNOWN')

                cam['firmware'] = actual_firmware
                _ui(self.status_set_step, 'capture', 'ok',
                    f"{serial} / {cam.get('mac', '?')}")

                self.camera_data.save()
                _ui(self.refresh_camera_list)

                # v4.3 #13 — assign Building Reports sticker number if enabled
                br_label = ''
                if br_enabled and br_counter is not None and not errors:
                    br_label = str(br_counter)
                    self.status_log(f"Building Reports sticker: #{br_label}  ← peel and apply to {cam.get('name', cam_name)}")
                    br_counter += 1

                # ---- Write CSV row ----
                cam_mac = cam.get('mac', pinned_mac or '')
                with open(OUTPUT_CSV, 'a', newline='') as f:
                    csv.writer(f).writerow([cam.get('name', cam_name), static_ip, serial, cam_mac,
                                           actual_model or expected_model, actual_firmware,
                                           br_label,
                                           datetime.now().isoformat()])

                # ---- Mark success/fail ----
                if 'unreachable' in errors:
                    self._mark_cam_failed(cam, ', '.join(errors))
                    total_fail += 1
                    self.status_log(f"\n*** {cam_name} FAILED: {', '.join(errors)} ***")
                    _ui(self.status_set_banner, f'FAILED  ({total_ok} OK / {total_fail} fail)',
                        f"{cam_name} failed — plug in next camera", '#F44336')
                elif errors:
                    self._mark_cam_failed(cam, ', '.join(errors))
                    total_fail += 1
                    self.status_log(f"\n*** {cam_name} PARTIAL FAIL: {', '.join(errors)} ***")
                    _ui(self.status_set_banner, f'PARTIAL  ({total_ok} OK / {total_fail} fail)',
                        f"{cam_name} had errors — plug in next camera", '#FF9800')
                else:
                    idx = self.camera_data.get_all().index(cam)
                    self.camera_data.mark_processed(idx)
                    total_ok += 1
                    self.status_log(f"\n*** {cam_name} COMPLETE ***")
                    _ui(self.status_set_banner, f'DONE  ({total_ok} OK / {total_fail} fail)',
                        f"{cam_name} complete — plug in next camera", '#4CAF50')

                remaining.pop(cam_idx)

                if remaining and not self.cancel_flag:
                    # No continue dialog — banner tells the user what to do
                    _ui(self.status_set_camera, '—',
                        f'{programmed_count} of {len(cameras)} done · {len(remaining)} remaining')
                    time.sleep(2)  # brief pause so user can see "DONE" before "PLUG IN"

                    # 2026-04-30 hot fix: wait for the just-programmed camera to leave
                    # factory IP before entering next discovery. Without this guard, if
                    # the user hasn't unplugged yet, the next iteration finds the SAME
                    # camera at factory IP, ARP query (post-unpin) returns None so the
                    # seen_macs check is bypassed, and we try to program the next slot
                    # onto the previous still-plugged-in camera. Auth fails, network
                    # call appears to "succeed" but actually does nothing useful, then
                    # the script bails. Belt-and-suspenders: ping is the source of truth
                    # for "is the previous camera still here", since ARP cache is stale
                    # after arp_unpin.
                    if factory_ip:
                        wait_start = time.time()
                        last_heartbeat = wait_start
                        announced = False
                        while not self.cancel_flag:
                            if not self.ping_camera(factory_ip, timeout_ms=800):
                                if announced:
                                    self.status_log(f"  Previous camera left factory IP after {int(time.time() - wait_start)}s")
                                break
                            now_t = time.time()
                            if not announced and now_t - wait_start >= 5:
                                self.status_log(f"  Waiting for {cam_name} to leave factory IP {factory_ip} (unplug it / let it reboot)...")
                                announced = True
                                last_heartbeat = now_t
                            elif announced and now_t - last_heartbeat >= 30:
                                self.status_log(f"  ...still at factory IP ({int(now_t - wait_start)}s elapsed)")
                                last_heartbeat = now_t
                            time.sleep(2)

            self.remove_linklocal_route()

            self.status_log(f"\n{'=' * 50}")
            self.status_log(f"PROGRAMMING COMPLETE: {total_ok} succeeded, {total_fail} failed")
            if remaining:
                self.status_log(f"  {len(remaining)} cameras not programmed")
            self.status_log(f"Results saved to {OUTPUT_CSV}")
            self.status_log(f"{'=' * 50}")

            if self.cancel_flag:
                _ui(self.status_set_banner, 'CANCELLED',
                    f'{total_ok} OK / {total_fail} failed / {len(remaining)} skipped', '#9E9E9E')
            elif total_fail == 0 and not remaining:
                _ui(self.status_set_banner, 'ALL DONE!',
                    f'All {total_ok} cameras programmed successfully', '#4CAF50')
            else:
                _ui(self.status_set_banner,
                    f'FINISHED  ({total_ok} OK / {total_fail} fail)',
                    f'{len(remaining)} cameras not programmed' if remaining else '',
                    '#FF9800' if total_fail else '#4CAF50')
            _ui(self.status_set_camera, '—', f'{total_ok + total_fail} of {len(cameras)} processed')
            _ui(self.enable_cancel, False)
            _ui(self.status_enable_cancel, False)
            _ui(self.refresh_camera_list)
            _ui(self.rescan_after_operation)
            self._close_wizard_run_log()

        threading.Thread(target=run, daemon=True).start()

    def start_factory_default_wizard(self):
        """v4.3 — Standalone factory default. Asks for IP + existing root
        password, fires factory_reset, polls for camera to come back. Useful
        for one-off cleanups outside the programming flow (camera came in
        from a prior site and needs to be wiped before it goes on the shelf
        as inventory)."""
        if not hasattr(self.protocol, 'factory_reset'):
            messagebox.showinfo("Factory Default",
                                f"{self.protocol.BRAND_NAME} protocol does not support factory_reset yet.")
            return
        ip = simpledialog.askstring("Factory Default",
            "Camera IP to wipe:", parent=self.root)
        if not ip:
            return
        ip = ip.strip()
        password = simpledialog.askstring("Factory Default",
            f"Existing root password for {ip}:", show='*', parent=self.root)
        if password is None:
            return
        if not messagebox.askyesno("Factory Default — confirm",
            f"This will WIPE the camera at {ip} back to factory state. ALL config, "
            f"users, network settings, certs will be lost. Continue?"):
            return

        self.cancel_flag = False
        self.enable_cancel(True)

        def run():
            self.log(f"\n{'='*60}")
            self.log(f"FACTORY DEFAULT — {ip}")
            self.log(f"{'='*60}")
            # Capture identity before reset so we can recognize the camera coming back
            probe = self.protocol.probe_unrestricted(ip)
            target_mac = (probe.get('mac') or '').upper().replace(':', '').replace('-', '')
            if target_mac:
                self.log(f"Camera identified: {probe.get('model','?')}  serial={probe.get('serial','?')}")
            else:
                self.log("⚠ Could not identify camera before reset (continuing anyway)")
            self.log(f"Firing factory_reset on {ip}...")
            ok = self.protocol.factory_reset(ip, password)
            if not ok:
                self.log("✗ Factory reset call failed (wrong password? unsupported endpoint?)")
                self.root.after(0, lambda: self.enable_cancel(False))
                return
            self.log("  ✓ Reset issued. Waiting for camera to come back (up to 120s)...")
            found_at = None
            for attempt in range(60):
                if self.cancel_flag:
                    self.log("Cancelled by user.")
                    break
                for try_ip in (ip, '192.168.0.90'):
                    if self.ping_camera(try_ip, timeout_ms=1000):
                        p = self.protocol.probe_unrestricted(try_ip)
                        p_mac = (p.get('mac') or '').upper().replace(':', '').replace('-', '')
                        if not target_mac or p_mac == target_mac:
                            found_at = try_ip
                            break
                if found_at:
                    break
                time.sleep(2)
            if found_at:
                self.log(f"\n✓ Camera back at {found_at} (factory state — no users yet)")
                self.log("  Use 'Program New Cameras' to apply a new config.")
                self.root.after(0, lambda: self.update_display("DONE", f"Camera factory-reset, now at {found_at}"))
            else:
                self.log("\n⚠ Camera didn't reappear within 120s")
                self.log("  Reset may have succeeded but camera is on a different IP.")
                self.log("  Check the network or wait + scan manually.")
                self.root.after(0, lambda: self.update_display("DONE", "Reset issued, camera not seen"))
            self.root.after(0, lambda: self.enable_cancel(False))

        threading.Thread(target=run, daemon=True).start()

    def start_confirm_wizard(self):
        """v4.3 #12 — Confirm Programming. Audit each camera in the list against
        what was expected: IP matches? Auth (root + master password) returns
        200? DHCP is off? MAC matches the expected serial? Read-only — never
        mutates the camera. Reports per-camera pass/fail to the main log
        widget AND writes EXPORT_DIR/verification_<timestamp>.csv."""
        cameras = self.camera_data.get_all()
        if not cameras:
            messagebox.showinfo("Confirm Programming",
                                "No cameras in the list to verify.")
            return
        if not hasattr(self.protocol, 'verify_camera_state'):
            messagebox.showinfo("Confirm Programming",
                                f"{self.protocol.BRAND_NAME} protocol does not support verification yet.")
            return
        password = simpledialog.askstring("Confirm Programming",
            f"Master password (root) to verify {len(cameras)} camera(s):",
            show='*', parent=self.root)
        if not password:
            return

        self.cancel_flag = False
        self.enable_cancel(True)

        def run():
            from datetime import datetime as _dt
            ts = _dt.now().strftime('%Y%m%d_%H%M%S')
            report_path = EXPORT_DIR / f'verification_{ts}.csv'
            EXPORT_DIR.mkdir(parents=True, exist_ok=True)
            header = ['CameraName', 'ExpectedIP', 'ActualIP', 'IPMatch', 'AuthOK',
                      'DHCPOff', 'ActualMAC', 'ExpectedMAC', 'MACMatch',
                      'Model', 'Firmware', 'Status', 'Timestamp']
            ok_count = warn_count = fail_count = 0
            self.log(f"\n{'='*60}")
            self.log(f"CONFIRM PROGRAMMING — {len(cameras)} camera(s)")
            self.log(f"Report: {report_path}")
            self.log(f"{'='*60}")
            with open(report_path, 'w', newline='', encoding='utf-8') as f:
                w = csv.writer(f)
                w.writerow(header)
                for cam in cameras:
                    if self.cancel_flag:
                        self.log("Cancelled by user.")
                        break
                    name = cam.get('name', '?')
                    expected_ip = cam.get('ip', '')
                    expected_mac = (cam.get('mac') or '').upper().replace('-', ':')
                    self.log(f"\n--- {name} (expected {expected_ip}) ---")
                    if not expected_ip:
                        self.log("  ⚠ No expected IP in list — skipping.")
                        w.writerow([name, '', '', '', '', '', '', expected_mac, '', '', '', 'NO_EXPECTED_IP', _dt.now().isoformat()])
                        warn_count += 1
                        continue
                    state = self.protocol.verify_camera_state(expected_ip, password)
                    actual_mac = (state.get('mac') or '').upper()
                    actual_ip = state.get('actual_ip') or ''
                    ip_match = (actual_ip == expected_ip) if actual_ip else state.get('reachable', False)
                    mac_match = (actual_mac and expected_mac and actual_mac == expected_mac)
                    # Status decision
                    issues = []
                    if not state['reachable']:
                        issues.append('UNREACHABLE')
                    if state['reachable'] and not state['auth_ok']:
                        issues.append('AUTH_FAIL')
                    if state['reachable'] and state['auth_ok'] and state['dhcp_off'] is False:
                        issues.append('DHCP_ON')
                    if expected_mac and actual_mac and not mac_match:
                        issues.append('MAC_MISMATCH')
                    if state['reachable'] and actual_ip and actual_ip != expected_ip:
                        issues.append('IP_MISMATCH')
                    status = 'OK' if not issues else ('FAIL' if 'UNREACHABLE' in issues or 'AUTH_FAIL' in issues or 'MAC_MISMATCH' in issues else 'WARN')
                    # Log
                    self.log(f"  reachable={state['reachable']}  auth_ok={state['auth_ok']}  "
                             f"dhcp_off={state['dhcp_off']}  actual_ip={actual_ip or '?'}  "
                             f"actual_mac={actual_mac or '?'}")
                    if status == 'OK':
                        self.log(f"  ✓ OK")
                        ok_count += 1
                    elif status == 'WARN':
                        self.log(f"  ⚠ WARN: {', '.join(issues)}")
                        warn_count += 1
                    else:
                        self.log(f"  ✗ FAIL: {', '.join(issues)}")
                        fail_count += 1
                    w.writerow([name, expected_ip, actual_ip,
                                'Y' if ip_match else 'N',
                                'Y' if state['auth_ok'] else 'N',
                                'Y' if state['dhcp_off'] else ('N' if state['dhcp_off'] is False else '?'),
                                actual_mac, expected_mac,
                                'Y' if mac_match else ('N' if expected_mac and actual_mac else '?'),
                                state.get('model') or '', state.get('firmware') or '',
                                status, _dt.now().isoformat()])
            self.log(f"\n{'='*60}")
            self.log(f"VERIFICATION DONE  —  ✓{ok_count}  ⚠{warn_count}  ✗{fail_count}  of {len(cameras)}")
            self.log(f"Report saved: {report_path}")
            self.log(f"{'='*60}")
            self.root.after(0, lambda: self.update_display("DONE", f"✓{ok_count} ⚠{warn_count} ✗{fail_count}"))
            self.root.after(0, lambda: self.enable_cancel(False))

        threading.Thread(target=run, daemon=True).start()

    def start_update_wizard(self):
        """Smart update — pushes any changes (IP, hostname, DHCP) to cameras"""
        cameras = self.camera_data.get_all()
        if not cameras:
            messagebox.showwarning("No Cameras", "No cameras in the Camera List.")
            return
        
        # Figure out what each camera needs
        updates = []
        for cam in cameras:
            if not cam.get('ip') or cam.get('processed'):
                continue
            # Use pending list from editor if available, otherwise detect changes
            pending = cam.get('pending', [])
            tasks = []
            if 'ip' in pending or cam.get('new_ip'):
                tasks.append('ip')
            if 'hostname' in pending:
                tasks.append('hostname')
            if 'dhcp' in pending:
                if cam.get('dhcp', '').lower() == 'yes':
                    tasks.append('dhcp_on')
                else:
                    tasks.append('dhcp_off')
            if tasks:
                updates.append((cam, tasks))
        
        if not updates:
            messagebox.showinfo("Nothing to Update",
                "No cameras have changes to push.\n\n"
                "Updatable fields:\n"
                "• New IP — set in the 'New IP' column\n"
                "• Hostname — set in the editor\n"
                "• DHCP — toggle in the editor")
            return
        
        # Build summary
        summary_lines = []
        for cam, tasks in updates[:8]:
            parts = []
            if 'ip' in tasks:
                parts.append(f"IP → {cam['new_ip']}")
            if 'hostname' in tasks:
                parts.append(f"hostname → {cam['hostname']}")
            if 'dhcp_on' in tasks:
                parts.append("DHCP on")
            if 'dhcp_off' in tasks:
                parts.append("DHCP off")
            summary_lines.append(f"  {cam['name']}: {', '.join(parts)}")
        
        summary = "\n".join(summary_lines)
        if len(updates) > 8:
            summary += f"\n  ... and {len(updates) - 8} more"
        
        if not messagebox.askyesno("Confirm Update",
            f"Push changes to {len(updates)} camera(s)?\n\n{summary}"):
            return
        
        password = self.get_password("Camera Password", "Enter current camera password:")
        if not password:
            return
        
        self.notebook.select(self.log_tab)
        self.cancel_flag = False
        self.enable_cancel(True)
        
        def run():
            try:
                ok = fail = 0
                for cam, tasks in updates:
                    if self.cancel_flag:
                        break
                    
                    current_ip = cam['ip']
                    cam_name = cam.get('name', current_ip)
                    errors = []
                    self.root.after(0, lambda n=cam_name: self.update_display(n, "Updating..."))
                    self.log(f"\nUpdating {cam_name} ({current_ip})")

                    # 1. Network (IP + subnet + gateway + disable DHCP)
                    if 'ip' in tasks:
                        gateway = cam.get('gateway', '')
                        subnet = cam.get('subnet', '255.255.255.0')
                        new_ip = cam['new_ip']

                        self.root.after(0, lambda n=cam_name: self.update_display(n, "Setting network..."))
                        self.log(f"  Setting IP: {current_ip} → {new_ip}")
                        if not self.protocol.set_network(current_ip, password, new_ip, subnet, gateway):
                            self.log(f"  ✗ Network config failed")
                            errors.append("ip")
                            self._mark_cam_failed(cam, "IP change failed")
                            fail += 1
                            continue

                        # Wait for camera at new IP for subsequent operations
                        if 'hostname' in tasks or 'dhcp_on' in tasks or 'dhcp_off' in tasks:
                            self.root.after(0, lambda n=cam_name: self.update_display(n, "Waiting for camera at new IP..."))
                            came_online = False
                            for attempt in range(15):
                                if self.cancel_flag:
                                    break
                                if self.ping_camera(new_ip):
                                    came_online = True
                                    break
                                time.sleep(2)
                            if came_online:
                                time.sleep(1)
                                current_ip = new_ip
                            elif not self.cancel_flag:
                                self.log(f"  ⚠ Camera not responding at {new_ip} — skipping remaining updates")
                                errors.append("unreachable at new IP")
                                self._mark_cam_failed(cam, "unreachable at new IP")
                                fail += 1
                                continue

                    # 2. Hostname
                    if 'hostname' in tasks and not self.cancel_flag:
                        target_ip = cam['new_ip'] if 'ip' in tasks else current_ip
                        hostname = cam['hostname']
                        self.root.after(0, lambda n=cam_name: self.update_display(n, "Setting hostname..."))
                        self.log(f"  Setting hostname: {hostname}")

                        if 'ip' not in tasks:
                            if not self.ping_camera(target_ip):
                                self.log(f"  ✗ Camera not responding at {target_ip}")
                                errors.append("unreachable")
                                self._mark_cam_failed(cam, "unreachable")
                                fail += 1
                                continue

                        result = self.protocol.set_hostname(target_ip, password, hostname)
                        if result:
                            self.log("  ✓ Hostname set")
                        else:
                            self.log("  Hostname failed, retrying...")
                            time.sleep(3)
                            result = self.protocol.set_hostname(target_ip, password, hostname)
                            if result:
                                self.log("  ✓ Hostname set on retry")
                            else:
                                self.log("  ✗ Hostname set failed")
                                errors.append("hostname")

                    # 3. DHCP
                    if ('dhcp_on' in tasks or 'dhcp_off' in tasks) and not self.cancel_flag:
                        target_ip = cam['new_ip'] if 'ip' in tasks else current_ip
                        enable = 'dhcp_on' in tasks
                        action = "Enabling" if enable else "Disabling"
                        self.root.after(0, lambda n=cam_name, a=action: self.update_display(n, f"{a} DHCP..."))
                        self.log(f"  {action} DHCP")
                        result = self.protocol.set_dhcp(target_ip, password, enable=enable)
                        if result:
                            self.log(f"  ✓ DHCP {'enabled' if enable else 'disabled'}")
                        else:
                            self.log(f"  ✗ DHCP {'enable' if enable else 'disable'} failed")
                            errors.append("dhcp")
                    
                    # Final status
                    if errors:
                        self.log(f"  ⚠ Partial failure: {', '.join(errors)}")
                        self._mark_cam_failed(cam, ', '.join(errors))
                        fail += 1
                    else:
                        self.log("  ✓ All changes applied")
                        self._mark_cam_processed(cam)
                        ok += 1
                    
            except Exception as e:
                self.log(f"Error during update: {e}")
            finally:
                self.log(f"\nUpdate complete: {ok} succeeded, {fail} failed")
                self.root.after(0, lambda: self.update_display("DONE", f"{ok} OK, {fail} failed"))
                self.root.after(0, lambda: self.enable_cancel(False))
                self.root.after(0, self.rescan_after_operation)
        
        threading.Thread(target=run, daemon=True).start()
    
    def _mark_cam_processed(self, cam):
        """Mark a camera as processed, promote new_ip to ip, clear pending, and refresh"""
        try:
            cam['pending'] = []
            # Promote new_ip to ip and clear new_ip
            if cam.get('new_ip'):
                cam['ip'] = cam['new_ip']
                cam['new_ip'] = ''
            idx = self.camera_data.get_all().index(cam)
            self.camera_data.mark_processed(idx)
            self.camera_data.save()
            self.root.after(0, self.refresh_camera_list)
        except ValueError:
            pass
    
    def _mark_cam_failed(self, cam, reason=''):
        """Mark a camera as failed and refresh"""
        try:
            idx = self.camera_data.get_all().index(cam)
            self.camera_data.mark_failed(idx, reason)
            self.root.after(0, self.refresh_camera_list)
        except ValueError:
            pass
    
    def start_capture_wizard(self):
        """Wizard for capturing images"""
        cameras = self.validate_cameras_for_basic_ops()
        if not cameras:
            return
        
        password = self.get_password("Capture Images", "Enter camera password:")
        if not password:
            return
        
        os.makedirs(IMAGES_DIR, exist_ok=True)
        self.notebook.select(self.log_tab)
        self.cancel_flag = False
        self.enable_cancel(True)
        
        def run():
            username = self.settings.get('general', 'username')
            captured = failed = 0
            
            for cam in cameras:
                if self.cancel_flag:
                    break
                self.root.after(0, lambda n=cam['name']: self.update_display(n, "Capturing..."))
                
                img = self.protocol.get_image(cam['ip'], self.protocol.DEFAULT_USER, password)
                fn = os.path.join(IMAGES_DIR, f"{self.sanitize_filename(cam['name'])}.jpg")
                
                if img:
                    img = self.watermark_image(img, cam['name'])
                    with open(fn, 'wb') as f:
                        f.write(img)
                    self.log(f"[OK] {cam['name']} → {fn}")
                    captured += 1
                    self.root.after(0, lambda d=img: self.update_preview(d))
                else:
                    self.log(f"[FAIL] {cam['name']}")
                    failed += 1
            
            self.log(f"\nCapture complete: {captured} OK, {failed} failed")
            self.log(f"Images saved to {IMAGES_DIR}")
            self.root.after(0, lambda: self.update_display("DONE", f"{captured} OK, {failed} failed"))
            self.root.after(0, lambda: self.enable_cancel(False))
            self.clear_preview()
        
        threading.Thread(target=run, daemon=True).start()
    
    def start_ping_wizard(self):
        """Wizard for ping test"""
        cameras = self.validate_cameras_for_basic_ops()
        if not cameras:
            return
        
        self.notebook.select(self.log_tab)
        self.cancel_flag = False
        self.enable_cancel(True)
        
        def run():
            results = []
            ok = fail = 0
            
            for cam in cameras:
                if self.cancel_flag:
                    break
                self.root.after(0, lambda n=cam['name'], ip=cam['ip']: self.update_display(n, f"Pinging {ip}..."))
                
                if self.ping_camera(cam['ip']):
                    self.log(f"[OK] {cam['name']} ({cam['ip']})")
                    results.append((cam['name'], cam['ip'], 'Success'))
                    ok += 1
                else:
                    self.log(f"[FAIL] {cam['name']} ({cam['ip']})")
                    results.append((cam['name'], cam['ip'], 'Failed'))
                    fail += 1
            
            with open(PING_RESULTS, 'w', newline='') as f:
                w = csv.writer(f)
                w.writerow(['CameraName', 'IPAddress', 'Status'])
                w.writerows(results)
            
            self.log(f"\nPing complete: {ok} OK, {fail} failed")
            self.log(f"Results saved to {PING_RESULTS}")
            self.root.after(0, lambda: self.update_display("DONE", f"{ok} OK, {fail} failed"))
            self.root.after(0, lambda: self.enable_cancel(False))
        
        threading.Thread(target=run, daemon=True).start()
    
    def start_validate_wizard(self):
        """Wizard for validating a password"""
        cameras = self.validate_cameras_for_basic_ops()
        if not cameras:
            return
        
        password = self.get_password("Validate Password", "Enter password to test:")
        if not password:
            return
        
        self.notebook.select(self.log_tab)
        self.cancel_flag = False
        self.enable_cancel(True)
        
        def run():
            username = self.settings.get('general', 'username')
            ok = fail = 0
            
            for cam in cameras:
                if self.cancel_flag:
                    break
                self.root.after(0, lambda n=cam['name']: self.update_display(n, "Testing password..."))
                
                if self.protocol.test_password(cam['ip'], self.protocol.DEFAULT_USER, password):
                    self.log(f"[OK] {cam['name']}")
                    ok += 1
                else:
                    self.log(f"[FAIL] {cam['name']}")
                    fail += 1

            self.log(f"\n{ok} OK, {fail} failed")
            self.root.after(0, lambda: self.update_display("DONE", f"{ok} OK, {fail} failed"))
            self.root.after(0, lambda: self.enable_cancel(False))
        
        threading.Thread(target=run, daemon=True).start()
    
    def start_change_password_wizard(self):
        """Wizard for changing passwords"""
        cameras = self.validate_cameras_for_basic_ops()
        if not cameras:
            return
        
        old_pwd = self.get_password("Change Passwords", "CURRENT password:")
        if not old_pwd:
            return
        
        new_pwd = self.get_password("Change Passwords", "NEW password:")
        if not new_pwd:
            return
        
        confirm = self.get_password("Change Passwords", "CONFIRM new password:")
        if new_pwd != confirm:
            messagebox.showerror("Mismatch", "Passwords don't match!")
            return
        
        if not messagebox.askyesno("Confirm", f"Change password on {len(cameras)} camera(s)?"):
            return
        
        self.notebook.select(self.log_tab)
        self.cancel_flag = False
        self.enable_cancel(True)
        
        def run():
            username = self.settings.get('general', 'username')
            ok = fail = 0
            
            for cam in cameras:
                if self.cancel_flag:
                    break
                self.root.after(0, lambda n=cam['name']: self.update_display(n, "Changing password..."))
                
                if self.protocol.change_password(cam['ip'], self.protocol.DEFAULT_USER, old_pwd, new_pwd):
                    self.log(f"[OK] {cam['name']}")
                    ok += 1
                else:
                    self.log(f"[FAIL] {cam['name']}")
                    fail += 1
            
            self.log(f"\n{ok} changed, {fail} failed")
            self.root.after(0, lambda: self.update_display("DONE", f"{ok} changed, {fail} failed"))
            self.root.after(0, lambda: self.enable_cancel(False))
        
        threading.Thread(target=run, daemon=True).start()
    
    def start_batch_test_wizard(self):
        """Wizard for batch password testing"""
        cameras = self.validate_cameras_for_basic_ops()
        if not cameras:
            return
        
        passwords = self.password_data.get_all()
        if not passwords:
            if self.settings.get_bool('warnings', 'show_batch_test_explanation'):
                dialog = WarningDialog(self.root, "Batch Password Test",
                    "This feature tests multiple passwords against your cameras "
                    "to find unknown credentials.\n\n"
                    "You need to add passwords to test first!\n\n"
                    "Go to the 'Passwords' tab and add potential passwords.",
                    'show_batch_test_explanation', self.settings)
            else:
                messagebox.showinfo("No Passwords",
                    "Add passwords to the 'Passwords' tab first.")
            self.notebook.select(self.passwords_tab)
            return
        
        if not messagebox.askyesno("Confirm",
            f"Test {len(passwords)} password(s) against {len(cameras)} camera(s)?\n\n"
            "This may take a while."):
            return
        
        self.notebook.select(self.log_tab)
        self.cancel_flag = False
        self.enable_cancel(True)
        
        def run():
            username = self.settings.get('general', 'username')
            found = []
            total_cams = len(cameras)
            total_pwds = len(passwords)
            
            for ci, cam in enumerate(cameras, 1):
                if self.cancel_flag:
                    break
                cam_label = f"{cam['name']} ({cam['ip']})"
                self.log(f"\nTesting {cam_label} (camera {ci}/{total_cams})...")
                self.root.after(0, lambda n=cam['name'], c=ci, t=total_cams: 
                    self.update_display(n, f"Camera {c}/{t} — testing passwords..."))
                
                matched = False
                for pi, pwd in enumerate(passwords, 1):
                    if self.cancel_flag:
                        break
                    masked = pwd[0] + '*' * (len(pwd) - 1) if len(pwd) > 1 else '*'
                    if self.protocol.test_password(cam['ip'], self.protocol.DEFAULT_USER, pwd):
                        self.log(f"  Trying {masked} ({pi}/{total_pwds})... ✓ FOUND")
                        found.append((cam['name'], cam['ip'], pwd))
                        matched = True
                        break
                    else:
                        self.log(f"  Trying {masked} ({pi}/{total_pwds})... ✗")
                
                if not matched and not self.cancel_flag:
                    self.log(f"  No password matched for {cam_label}")
            
            self.log("")
            if found:
                with open(SUCCESSFUL_PASSWORDS, 'w', newline='') as f:
                    w = csv.writer(f)
                    w.writerow(['CameraName', 'IPAddress', 'Password'])
                    w.writerows(found)
                self.log(f"Found {len(found)} password(s) → {SUCCESSFUL_PASSWORDS}")
            else:
                self.log("No passwords found")
            self.log(f"Summary: {len(found)} found, {total_cams - len(found)} no match")
            
            self.root.after(0, lambda: self.update_display("DONE", f"Found {len(found)}/{total_cams}"))
            self.root.after(0, lambda: self.enable_cancel(False))
        
        threading.Thread(target=run, daemon=True).start()
    
    # ========================================================================
    # CAMERA API FUNCTIONS (unchanged from original)
    # ========================================================================
    def sanitize_filename(self, name):
        return re.sub(r'[\\/:*?"<>|]', '_', str(name))
    
    def subnet_to_cidr(self, subnet):
        if not subnet:
            return 24
        return sum(bin(int(x)).count('1') for x in subnet.split('.'))
    
    def get_mac_from_arp(self, ip):
        """Get MAC address from ARP table after actively resolving the IP.

        Hardened 2026-04-30: previously returned None when arp_unpin had cleared
        the static entry and the dynamic ARP cache had aged out. Now does an
        active ping FIRST to force the OS to ARP-resolve, then reads the table.
        Retries up to 3 times because Windows occasionally needs more than one
        ping to populate the table.
        """
        import subprocess
        for attempt in range(3):
            try:
                # Force ARP resolution by pinging. Windows populates the ARP cache
                # only when traffic is exchanged with the IP.
                self.ping_camera(ip, timeout_ms=600)
            except Exception:
                pass
            try:
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
                result = subprocess.run(['arp', '-a', ip],
                    capture_output=True, text=True, timeout=5, startupinfo=startupinfo,
                    creationflags=subprocess.CREATE_NO_WINDOW)
                if result.returncode == 0:
                    for line in result.stdout.split('\n'):
                        line = line.strip()
                        if ip in line:
                            # Windows ARP output: "192.168.0.90  aa-bb-cc-dd-ee-ff  dynamic"
                            parts = line.split()
                            for part in parts:
                                if len(part.replace('-', '').replace(':', '')) == 12 and part != ip:
                                    return part.upper().replace('-', ':')
            except Exception:
                pass
            # Brief backoff before retry
            if attempt < 2:
                import time as _t
                _t.sleep(0.4)
        return None
    
    def arp_pin(self, ip, mac):
        """Pin ARP entry so all traffic to IP goes to specific MAC.
        Note: requires Run as Administrator on Windows for arp -s to work.
        If not admin, pin will fail silently and programming proceeds without it."""
        import subprocess
        try:
            # Windows needs MAC with dashes
            mac_dashes = mac.replace(':', '-').lower()
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            # First delete any existing entry
            subprocess.run(['arp', '-d', ip], capture_output=True, timeout=5,
                startupinfo=startupinfo, creationflags=subprocess.CREATE_NO_WINDOW)
            # Add static entry
            result = subprocess.run(['arp', '-s', ip, mac_dashes],
                capture_output=True, text=True, timeout=5, startupinfo=startupinfo,
                creationflags=subprocess.CREATE_NO_WINDOW)
            return result.returncode == 0
        except:
            return False
    
    def arp_unpin(self, ip):
        """Remove static ARP entry"""
        import subprocess
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            subprocess.run(['arp', '-d', ip], capture_output=True, timeout=5,
                startupinfo=startupinfo, creationflags=subprocess.CREATE_NO_WINDOW)
        except:
            pass

    # =========================================================================
    # LINK-LOCAL ROUTE — reach 169.254.x.x cameras from a static-IP PC
    # Adds a temporary secondary 169.254.100.1 address to the detected
    # interface. This is needed because cameras at 169.254.x.x won't respond
    # to ARP from IPs outside their subnet. The secondary address gives us
    # a source IP the camera can route back to. Your existing IP stays.
    # Uses interface index (a number) — works on any Windows PC.
    # Removed automatically at end of session. Non-persistent on reboot.
    # =========================================================================

    LINKLOCAL_IP = '169.254.100.1'
    LINKLOCAL_MASK = '255.255.0.0'

    def _has_linklocal_route(self):
        """Check if 169.254.100.1 is already on the detected interface.
        Must check the specific IP on the specific interface — other adapters
        often have random 169.254.x.x addresses that don't help us."""
        import subprocess
        iface_idx = getattr(self, '_detected_iface_index', None)
        if iface_idx is None:
            return False
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            result = subprocess.run(
                ['netsh', 'interface', 'ipv4', 'show', 'addresses', str(iface_idx)],
                capture_output=True, text=True, timeout=5,
                startupinfo=startupinfo, creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode == 0 and self.LINKLOCAL_IP in result.stdout:
                return True
        except:
            pass
        return False

    def _get_interface_index(self):
        """Get the interface index for the active network adapter.
        Returns an integer or None. Checks: cached value, saved setting, then auto-detect."""
        idx = getattr(self, '_detected_iface_index', None)
        if idx is not None:
            return idx
        # Check saved preference from settings
        try:
            saved = self.settings.get('general', 'interface_index')
            if saved:
                self._detected_iface_index = int(saved)
                return self._detected_iface_index
        except:
            pass
        local_ip, _, _ = self.get_local_network_info()
        return getattr(self, '_detected_iface_index', None)

    def add_linklocal_route(self):
        """Add a temporary secondary 169.254.100.1 to the detected interface.
        Uses PowerShell New-NetIPAddress which adds a secondary address
        WITHOUT converting DHCP to static (netsh add address does that).
        Your existing IP and DHCP lease stay untouched.
        Removed at end of session. Non-persistent: gone on reboot.
        Returns True on success."""
        import subprocess

        if self._has_linklocal_route():
            self._linklocal_route_active = True
            self._linklocal_iface_idx = self._get_interface_index()
            return True  # Already on the correct interface

        iface_idx = self._get_interface_index()
        if iface_idx is None:
            self.log("  Link-local: FAILED — could not detect interface")
            return False

        self._linklocal_route_active = False

        # PowerShell New-NetIPAddress adds a secondary address cleanly.
        # Unlike 'netsh add address', it does NOT convert DHCP to static.
        try:
            cmd = (f"New-NetIPAddress -InterfaceIndex {iface_idx} "
                   f"-IPAddress '{self.LINKLOCAL_IP}' -PrefixLength 16 "
                   f"-SkipAsSource $true -PolicyStore ActiveStore "
                   f"-ErrorAction Stop")
            result = subprocess.run(
                ['powershell', '-NoProfile', '-Command', cmd],
                capture_output=True, text=True, timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW)
            if result.returncode == 0:
                self._linklocal_route_active = True
                self._linklocal_iface_idx = iface_idx
                self.log(f"  Link-local: {self.LINKLOCAL_IP} on interface {iface_idx}")
                time.sleep(1)
                return True
            # Address may already exist on this interface
            if self._has_linklocal_route():
                self._linklocal_route_active = True
                self._linklocal_iface_idx = iface_idx
                self.log(f"  Link-local: {self.LINKLOCAL_IP} already on interface {iface_idx}")
                return True
        except:
            pass

        self.log("  Link-local: FAILED — Run as Administrator")
        return False

    def remove_linklocal_route(self):
        """Remove the temporary 169.254.100.1 secondary address."""
        import subprocess

        if not getattr(self, '_linklocal_route_active', False):
            return

        iface_idx = getattr(self, '_linklocal_iface_idx', None)
        if iface_idx is None:
            return

        try:
            cmd = (f"Remove-NetIPAddress -IPAddress '{self.LINKLOCAL_IP}' "
                   f"-InterfaceIndex {iface_idx} -Confirm:$false "
                   f"-ErrorAction SilentlyContinue")
            subprocess.run(
                ['powershell', '-NoProfile', '-Command', cmd],
                capture_output=True, timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW)
            self._linklocal_route_active = False
            self.log(f"  Link-local: removed {self.LINKLOCAL_IP}")
        except:
            pass

    def _resolve_linklocal_cameras(self, target_mac=None, timeout=4):
        """Probe for link-local cameras via mDNS on the correct interface.
        Uses IP_MULTICAST_IF bound to 169.254.100.1 to force multicast out
        on the link-local interface — critical on multi-adapter systems.
        If target_mac is given, only return cameras matching that MAC.
        Returns list of camera dicts with ip, mac, model, serial."""
        cameras = []
        seen = set()

        # Use the link-local IP for multicast if available, else fall back to detected IP
        multicast_ip = self.LINKLOCAL_IP if getattr(self, '_linklocal_route_active', False) else None
        if not multicast_ip:
            multicast_ip = getattr(self, '_detected_local_ip', None)
        if not multicast_ip:
            multicast_ip, _, _ = self.get_local_network_info()
        if not multicast_ip:
            return cameras

        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM, socket.IPPROTO_UDP)
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            sock.setsockopt(socket.IPPROTO_IP, socket.IP_MULTICAST_TTL, 255)

            # Force multicast out on link-local interface
            sock.setsockopt(socket.IPPROTO_IP, socket.IP_MULTICAST_IF,
                            socket.inet_aton(multicast_ip))

            sock.bind(('', AxisMDNSDiscovery.MDNS_PORT))

            # Join mDNS multicast on our interface
            mreq = socket.inet_aton(AxisMDNSDiscovery.MDNS_ADDR) + socket.inet_aton(multicast_ip)
            sock.setsockopt(socket.IPPROTO_IP, socket.IP_ADD_MEMBERSHIP, mreq)

            sock.settimeout(0.5)

            # Send mDNS queries
            for service in AxisMDNSDiscovery.SERVICE_TYPES:
                try:
                    query = AxisMDNSDiscovery.build_mdns_query(service)
                    sock.sendto(query, (AxisMDNSDiscovery.MDNS_ADDR, AxisMDNSDiscovery.MDNS_PORT))
                except Exception:
                    pass

            # Listen for responses
            end_time = time.time() + timeout
            queries_sent = 1
            while time.time() < end_time:
                try:
                    data, addr = sock.recvfrom(4096)
                    camera = AxisMDNSDiscovery.parse_mdns_response(data, addr[0])
                    if camera and camera.get('ip'):
                        cam_mac = camera.get('mac', '').upper().replace(':', '')
                        key = cam_mac or camera['ip']
                        if key not in seen:
                            seen.add(key)
                            if target_mac:
                                target_clean = target_mac.upper().replace(':', '').replace('-', '')
                                if cam_mac != target_clean:
                                    continue
                            cameras.append(camera)
                except socket.timeout:
                    if queries_sent < 3:
                        for service in AxisMDNSDiscovery.SERVICE_TYPES:
                            try:
                                query = AxisMDNSDiscovery.build_mdns_query(service)
                                sock.sendto(query, (AxisMDNSDiscovery.MDNS_ADDR, AxisMDNSDiscovery.MDNS_PORT))
                            except Exception:
                                pass
                        queries_sent += 1
                    continue
                except Exception:
                    continue

            sock.close()
        except Exception:
            pass
        return cameras

    def ping_camera(self, ip, timeout_ms=None):
        import subprocess
        if timeout_ms is None:
            timeout_ms = TIMEOUT * 1000
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            cmd = ['ping', '-n', '1', '-w', str(timeout_ms), ip]
            result = subprocess.run(cmd,
                capture_output=True, timeout=(timeout_ms / 1000) + 2, startupinfo=startupinfo,
                creationflags=subprocess.CREATE_NO_WINDOW)
            # Check stdout for "Reply from" — returncode alone can be unreliable
            # when there are multiple adapters or "Destination host unreachable"
            output = result.stdout.decode('utf-8', errors='ignore') if result.stdout else ''
            if 'Reply from' in output and 'unreachable' not in output.lower():
                return True
            return result.returncode == 0 and 'Reply from' in output
        except:
            return False
    
    def watermark_image(self, img_bytes, cam_name=''):
        """Add timestamp watermark to bottom-right corner of image"""
        if not HAS_PIL:
            return img_bytes
        try:
            img = Image.open(BytesIO(img_bytes))
            draw = ImageDraw.Draw(img)
            
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            text = f"{cam_name}  {timestamp}" if cam_name else timestamp
            disclaimer = "NOT FROM CAMERA OVERLAY — added by CCTV IP Toolkit"
            
            # Pick font size based on image width
            img_w, img_h = img.size
            font_size = max(14, img_w // 50)
            disclaimer_font_size = max(10, font_size // 2)
            try:
                font = ImageFont.truetype("arial.ttf", font_size)
                small_font = ImageFont.truetype("arial.ttf", disclaimer_font_size)
            except:
                try:
                    font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", font_size)
                    small_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", disclaimer_font_size)
                except:
                    font = ImageFont.load_default()
                    small_font = font
            
            # Measure main text
            bbox = draw.textbbox((0, 0), text, font=font)
            text_w = bbox[2] - bbox[0]
            text_h = bbox[3] - bbox[1]
            
            # Measure disclaimer
            dbbox = draw.textbbox((0, 0), disclaimer, font=small_font)
            disc_w = dbbox[2] - dbbox[0]
            disc_h = dbbox[3] - dbbox[1]
            
            # Position: bottom-right with padding
            pad = 10
            total_w = max(text_w, disc_w)
            total_h = text_h + disc_h + 4
            x = img_w - total_w - pad
            y = img_h - total_h - pad
            
            # Draw dark background rectangle for readability
            draw.rectangle([x - 6, y - 4, img_w - pad + 6, img_h - pad + 4], fill=(0, 0, 0, 180))
            draw.text((x, y), text, fill=(255, 255, 255), font=font)
            draw.text((x, y + text_h + 2), disclaimer, fill=(200, 200, 200), font=small_font)
            
            buf = BytesIO()
            img.save(buf, format='JPEG', quality=95)
            return buf.getvalue()
        except:
            return img_bytes
    

# ============================================================================
# MAIN
# ============================================================================
def _ensure_admin():
    """Re-launch as Administrator if not already elevated.
    Needed for link-local route and ARP pinning."""
    import ctypes, sys
    try:
        if ctypes.windll.shell32.IsUserAnAdmin():
            return True
    except:
        return True  # Not Windows — skip elevation

    # Re-launch ourselves elevated
    # Works for both .py scripts and PyInstaller .exe
    if getattr(sys, 'frozen', False):
        # Running as compiled .exe
        exe = sys.executable
        args = ' '.join(sys.argv[1:])
    else:
        # Running as .py script
        exe = sys.executable  # python.exe
        args = '"' + '" "'.join(sys.argv) + '"'

    try:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", exe, args, None, 1)
    except:
        pass
    sys.exit(0)


if __name__ == "__main__":
    _ensure_admin()
    root = tk.Tk()
    app = CCTVToolkitApp(root)
    root.mainloop()
