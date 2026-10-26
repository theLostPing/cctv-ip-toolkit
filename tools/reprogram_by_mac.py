#!/usr/bin/env python3
"""
reprogram_by_mac.py - put factory-defaulted Axis cameras back the way they were,
one camera at a time, matched by MAC against the toolkit's programmed_cameras.csv.

    python reprogram_by_mac.py check --csv MACs.xlsx --here 3    # who is still fine, who is gone (plugged into switch 3)
    python reprogram_by_mac.py run --password XXXX   # sit and wait, program each camera as it shows up
    python reprogram_by_mac.py ports [--switch IP]    # which switch port each MAC is on -> switch_ports.csv
        add --plan-base 10.20.48.0 to check ports against the numbering plan and to name
        unlisted cameras from their port (-> proposed_cameras.csv, review, then run --csv it)

What it does for every camera it finds that is in the list and not already right:
    1. sets the root password (factory camera has none)
    2. sets static IP / mask / gateway, DHCP off      (the camera reboots itself)
    3. waits for it at the new IP and checks the MAC is still the same camera
    4. sets hostname + Bonjour name to the camera name (KBO-006)
    5. adds the extra user (admin / admin by default)
    6. writes a row to reprogrammed_cameras.csv next to the input CSV

A MAC that is not in the list is reported and NEVER touched.
A camera already sitting at its own IP with a working password is left alone,
except that the extra user is added if it is missing.

Needs: Python 3.8+ and `requests` (pip install requests). Run as Administrator
so it can add temporary IP aliases (192.168.0.x and the camera subnet) to your
NIC; without admin it still works if you set those addresses by hand.
Axis only. Part of Brian Preston's CCTV IP Toolkit (fieldtoolkit.com), 2026-10-26.
"""
import argparse
import csv
import ctypes
import ipaddress
import json
import os
import re
import socket
import subprocess
import sys
import time
from datetime import datetime

try:
    import requests
    from requests.auth import HTTPDigestAuth
except ImportError:
    print("Missing 'requests'. Run:  python -m pip install requests")
    sys.exit(2)

try:
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
except Exception:
    pass

TIMEOUT = 5
FACTORY_IP = '192.168.0.90'
LINKLOCAL_IP = '169.254.100.1'
AXIS_OUIS = ('00:40:8c', 'ac:cc:8e', 'b8:a4:4f', 'e8:27:25')
APPDATA_CFG = os.path.join(os.environ.get('APPDATA', ''), 'CCTVIPToolkit')
def _default_csv_candidates():
    import glob
    home = os.environ.get('USERPROFILE', '')
    cands = [os.path.join(home, 'Documents', 'CCTV Toolkit', 'programmed_cameras.csv')]
    cands += glob.glob(os.path.join(home, 'OneDrive*', 'Documents', 'CCTV Toolkit', 'programmed_cameras.csv'))
    return cands

LOG_FH = None


def log(msg):
    line = f"{datetime.now().strftime('%H:%M:%S')}  {msg}"
    print(line, flush=True)
    if LOG_FH:
        try:
            LOG_FH.write(line + '\n')
            LOG_FH.flush()
        except Exception:
            pass


def norm_mac(s):
    s = re.sub(r'[^0-9A-Fa-f]', '', s or '')
    if len(s) != 12:
        return ''
    s = s.upper()
    return ':'.join(s[i:i + 2] for i in range(0, 12, 2))


def is_admin():
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def ps(cmd, timeout=20):
    r = subprocess.run(['powershell', '-NoProfile', '-Command', cmd],
                       capture_output=True, text=True, timeout=timeout,
                       creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    return r.returncode, (r.stdout or '') + (r.stderr or '')


# ---------------------------------------------------------------- camera list
def _norm_hdr(h):
    return re.sub(r'\s+', ' ', str(h or '')).strip().upper()


def _read_table(path):
    """Return a list of rows (lists of strings) from .csv or .xlsx."""
    if path.lower().endswith(('.xlsx', '.xlsm')):
        try:
            import openpyxl
        except ImportError:
            raise SystemExit("Reading .xlsx needs openpyxl (pip install openpyxl) - or save the sheet as CSV.")
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.worksheets[0]
        return [[('' if c is None else str(c)) for c in row] for row in ws.iter_rows(values_only=True)]
    with open(path, newline='', encoding='utf-8-sig') as f:
        return [list(r) for r in csv.reader(f)]


def load_list(path, since, prefix, only_net=None):
    """Load the camera list from any of:
       - the toolkit's run log programmed_cameras.csv (latest row per MAC on/after --since)
       - the toolkit's Camera List export / proposed_cameras.csv
       - the Lenel programming sheet (.xlsx or .csv): SWITCH, SWITCH PORT, CAMERA IP ADDRESS,
         CAMERA GATEWAY, CAMERA SUBNET, DEVICE MODEL, CAD ID, CAMERA / DEVICE MAC ADDRESS
    Each camera: name, ip, mac, model, gateway, subnet, switch, port, ts."""
    table = _read_table(path)
    if not table:
        raise SystemExit(f"No rows in {path}")
    # find the header row: the first row that has a MAC column and a name column
    hdr_i = None
    fallback = None
    for i, row in enumerate(table[:60]):
        H = [_norm_hdr(c) for c in row]
        if any('MAC' in h for h in H) and any(h in ('CAD ID', 'CAMERANAME', 'CAMERA NAME') for h in H):
            if 'SWITCH PORT' in H or 'CAMERA IP ADDRESS' in H:
                hdr_i = i          # the camera table of the Lenel sheet
                break
            if fallback is None:
                fallback = i
    if hdr_i is None:
        hdr_i = fallback
    if hdr_i is None:
        raise SystemExit(f"Could not find a header row with a MAC column in {path}")
    H = [_norm_hdr(c) for c in table[hdr_i]]

    def find(*cands):
        for c in cands:
            for j, h in enumerate(H):
                if h == c:
                    return j
        for c in cands:
            for j, h in enumerate(H):
                if c in h:
                    return j
        return None
    j_name = find('CAD ID', 'CAMERANAME', 'CAMERA NAME')
    j_mac = find('CAMERA / DEVICE MAC ADDRESS', 'MACADDRESS', 'MAC ADDRESS', 'MAC')
    j_ip = find('CAMERA IP ADDRESS', 'NEW IP', 'IPADDRESS', 'IP ADDRESS')
    j_gw = find('CAMERA GATEWAY', 'GATEWAY')
    j_sn = find('CAMERA SUBNET', 'SUBNET')
    j_model = find('DEVICE MODEL', 'MODEL')
    j_ts = find('TIMESTAMP')
    j_sw = find('SWITCH')
    j_port = find('SWITCH PORT')
    j_actual = find('L2POE PORT ACTUAL')
    j_desc = find('CAMERA NAME / LOCATION', 'DEVICE NAME / LOCATION')
    is_lenel = j_sw is not None and j_port is not None and j_sw != j_port
    kept = {}
    skipped_net = set()
    skipped_rows = 0
    for row in table[hdr_i + 1:]:
        def g(j):
            return row[j].strip() if (j is not None and j < len(row) and row[j] is not None) else ''
        name, mac, ip = g(j_name), norm_mac(g(j_mac)), g(j_ip)
        ts = g(j_ts)
        if since and ts and ts[:10] < since:
            continue
        if prefix and not name.upper().startswith(prefix.upper()):
            continue
        if not (mac and ip and name):
            if name or mac:
                skipped_rows += 1
            continue
        try:
            ipaddress.ip_address(ip)
        except ValueError:
            skipped_rows += 1
            continue
        if only_net is not None and ipaddress.ip_address(ip) not in only_net:
            skipped_net.add(name)
            continue
        kept[mac] = {'name': name, 'ip': ip, 'mac': mac, 'model': g(j_model), 'ts': ts,
                     'gateway': g(j_gw), 'subnet': g(j_sn), 'switch': g(j_sw) if is_lenel else '',
                     'port': g(j_port) if is_lenel else '', 'port_actual': g(j_actual),
                     'desc': g(j_desc)}
    byname = {}
    for c in kept.values():
        byname.setdefault(c['name'], []).append(c)
    for n, v in byname.items():
        if len(v) > 1:
            v.sort(key=lambda c: c['ts'])
            for old in v[:-1]:
                del kept[old['mac']]
            log(f"WARNING name {n} had {len(v)} MACs, keeping the latest ({v[-1]['mac']})")
    if skipped_net:
        log(f"ignored {len(skipped_net)} name(s) outside {only_net}: {', '.join(sorted(skipped_net))}")
    if skipped_rows:
        log(f"ignored {skipped_rows} row(s) without name+MAC+IP")
    cams = sorted(kept.values(), key=lambda c: ipaddress.ip_address(c['ip']))
    kind = 'Lenel sheet' if is_lenel else ('toolkit log' if j_ts is not None else 'camera list')
    log(f"loaded {len(cams)} cameras from the {kind} {os.path.basename(path)}")
    return cams


def load_log(path):
    """MAC -> latest {name, ip, ts} from a toolkit run log (for cross-reference)."""
    out = {}
    if not path or not os.path.exists(path):
        return out
    try:
        with open(path, newline='', encoding='utf-8-sig') as f:
            for r in csv.DictReader(f):
                mac = norm_mac(r.get('MACAddress', ''))
                if mac:
                    out[mac] = {'name': (r.get('CameraName') or '').strip(), 'ip': (r.get('IPAddress') or '').strip(),
                                'ts': (r.get('Timestamp') or '')[:16]}
    except Exception as e:
        log(f"  could not read log {path}: {e}")
    return out


def switch_no(name):
    m = re.search(r'(\d+)\s*$', name or '')
    return int(m.group(1)) if m else None


def cam_gateway(a, cam):
    if cam.get('gateway'):
        return cam['gateway']
    if a.gateway and a.gateway != 'auto':
        return a.gateway
    net = ipaddress.ip_network(f"{cam['ip']}/{cam_subnet(a, cam)}", strict=False)
    return str(net.network_address + 1)


def cam_subnet(a, cam):
    return cam.get('subnet') or a.subnet or '255.255.255.0'


def cam_net(a, cam):
    return ipaddress.ip_network(f"{cam['ip']}/{cam_subnet(a, cam)}", strict=False)


def load_password_candidates(explicit):
    cands = []
    if explicit:
        cands.append(explicit)
    try:
        p = os.path.join(APPDATA_CFG, 'passwords.json')
        if os.path.exists(p):
            with open(p) as f:
                for pw in json.load(f):
                    if pw and pw not in cands:
                        cands.append(pw)
    except Exception:
        pass
    for pw in ('admin', 'pass', 'root', 'Admin123', 'password'):
        if pw not in cands:
            cands.append(pw)
    return cands


# ---------------------------------------------------------------- Axis VAPIX
def probe(ip, timeout=TIMEOUT):
    """No-auth identity probe. Works factory and locked. Returns dict or None."""
    try:
        r = requests.post(f"http://{ip}/axis-cgi/basicdeviceinfo.cgi",
                          json={"apiVersion": "1.0", "method": "getAllUnrestrictedProperties"},
                          timeout=timeout)
        if r.status_code != 200:
            return None
        props = (r.json().get('data') or {}).get('propertyList') or {}
        serial = (props.get('SerialNumber') or '').strip()
        return {'ip': ip, 'mac': norm_mac(serial), 'serial': serial,
                'model': (props.get('ProdNbr') or '').strip(),
                'firmware': (props.get('Version') or '').strip()}
    except Exception:
        return None


def is_factory(ip):
    """True if root has no password yet (param.cgi answers without auth)."""
    try:
        r = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                         params={"action": "list", "group": "Brand"}, timeout=TIMEOUT)
        return r.status_code == 200 and 'Brand' in r.text
    except Exception:
        return False


def test_password(ip, user, pw):
    try:
        r = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                         params={"action": "list", "group": "Brand"},
                         auth=HTTPDigestAuth(user, pw), timeout=TIMEOUT)
        return r.status_code == 200 and 'Brand' in r.text
    except Exception:
        return False


def find_password(ip, cands):
    for pw in cands:
        if test_password(ip, 'root', pw):
            return pw
    return None


def create_root(ip, password):
    """Factory camera: FW 11+ takes pwdgrp.cgi open; FW 10 wants digest with the new pwd."""
    url = f"http://{ip}/axis-cgi/pwdgrp.cgi"
    params = {"action": "add", "user": "root", "pwd": password,
              "grp": "root", "sgrp": "admin:operator:viewer:ptz"}
    try:
        r = requests.get(url, params=params, timeout=TIMEOUT)
        if r.status_code == 401:
            r = requests.get(url, params=params, auth=HTTPDigestAuth("root", password), timeout=TIMEOUT)
        if r.status_code != 200:
            log(f"    pwdgrp.cgi -> HTTP {r.status_code} {r.text[:120]!r}")
            return False
        low = (r.text or '').strip().lower()
        if low.startswith('error') or 'error:' in low[:60]:
            log(f"    pwdgrp.cgi -> {r.text[:120]!r}")
            return False
    except Exception as e:
        log(f"    pwdgrp.cgi raised {type(e).__name__}: {e}")
        return False
    time.sleep(2)
    # ONVIF copy of root so a VMS can log in too (best effort)
    soap = ('<?xml version="1.0"?><Envelope xmlns="http://www.w3.org/2003/05/soap-envelope"><Header/><Body>'
            '<CreateUsers xmlns="http://www.onvif.org/ver10/device/wsdl" xmlns:tt="http://www.onvif.org/ver10/schema">'
            f'<User><tt:Username>root</tt:Username><tt:Password>{password}</tt:Password>'
            '<tt:UserLevel>Administrator</tt:UserLevel></User></CreateUsers></Body></Envelope>')
    for _ in range(3):
        try:
            r = requests.post(f"http://{ip}/vapix/services", data=soap,
                              headers={"Content-Type": "application/soap+xml"},
                              auth=HTTPDigestAuth("root", password), timeout=15)
            if r.status_code == 401:
                time.sleep(3)
                continue
            break
        except Exception:
            time.sleep(3)
    return True


def user_exists(ip, password, username):
    try:
        r = requests.get(f"http://{ip}/axis-cgi/pwdgrp.cgi", params={"action": "get"},
                         auth=HTTPDigestAuth("root", password), timeout=TIMEOUT)
        if r.status_code != 200:
            return None
        names = set()
        for line in r.text.splitlines():
            if '=' in line:
                names.update(x.strip().strip('"') for x in line.split('=', 1)[1].split(','))
        return username in names
    except Exception:
        return None


def add_user(ip, password, username, user_pw, role='Administrator'):
    sgrp = {'Administrator': 'admin:operator:viewer:ptz',
            'Operator': 'operator:viewer:ptz', 'Viewer': 'viewer'}.get(role, 'admin:operator:viewer:ptz')
    try:
        r = requests.get(f"http://{ip}/axis-cgi/pwdgrp.cgi",
                         params={"action": "add", "user": username, "pwd": user_pw,
                                 "grp": "users", "sgrp": sgrp},
                         auth=HTTPDigestAuth("root", password), timeout=TIMEOUT)
        if r.status_code != 200:
            return False
        low = (r.text or '').strip().lower()
        if low.startswith('error') or 'error:' in low[:60]:
            if 'exist' in low:
                return True
            log(f"    add user -> {r.text[:120]!r}")
            return False
    except Exception as e:
        log(f"    add user raised {type(e).__name__}: {e}")
        return False
    level = 'Administrator' if role == 'Administrator' else ('Operator' if role == 'Operator' else 'User')
    soap = ('<?xml version="1.0"?><Envelope xmlns="http://www.w3.org/2003/05/soap-envelope"><Header/><Body>'
            '<CreateUsers xmlns="http://www.onvif.org/ver10/device/wsdl" xmlns:tt="http://www.onvif.org/ver10/schema">'
            f'<User><tt:Username>{username}</tt:Username><tt:Password>{user_pw}</tt:Password>'
            f'<tt:UserLevel>{level}</tt:UserLevel></User></CreateUsers></Body></Envelope>')
    try:
        requests.post(f"http://{ip}/vapix/services", data=soap,
                      headers={"Content-Type": "application/soap+xml"},
                      auth=HTTPDigestAuth("root", password), timeout=15)
    except Exception:
        pass
    return True


def set_hostname(ip, password, name):
    clean = re.sub(r'[^a-zA-Z0-9-]', '-', name).strip('-')[:63]
    ok = False
    for key in ('root.Network.HostName', 'root.Network.Bonjour.FriendlyName'):
        try:
            r = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                             params={"action": "update", key: clean},
                             auth=HTTPDigestAuth("root", password), timeout=TIMEOUT)
            if key.endswith('HostName'):
                ok = r.status_code == 200 and not r.text.strip().startswith('# Error')
        except Exception:
            pass
    return ok


def set_network(ip, password, new_ip, subnet, gateway, firmware=''):
    """Same two paths the toolkit uses: ONVIF SOAP first on OS 12+, VAPIX param.cgi
    first on anything older (ONVIF gateway write takes 60-90 s on 10.x/11.x)."""
    auth = HTTPDigestAuth("root", password)
    cidr = ipaddress.ip_network(f"0.0.0.0/{subnet}").prefixlen
    m = re.match(r'^(\d+)', firmware or '')
    prefer_vapix = bool(m) and int(m.group(1)) < 12

    def vapix():
        try:
            r = requests.get(f"http://{ip}/axis-cgi/param.cgi",
                             params={"action": "update", "root.Network.BootProto": "none",
                                     "root.Network.IPAddress": new_ip,
                                     "root.Network.SubnetMask": subnet,
                                     "root.Network.DefaultRouter": gateway},
                             auth=auth, timeout=10)
            ok = r.status_code == 200 and not (r.text or '').strip().startswith('# Error')
            log(f"    VAPIX network update -> HTTP {r.status_code} {r.text[:80]!r}")
            return ok
        except (requests.exceptions.ConnectionError, requests.exceptions.ReadTimeout):
            log("    VAPIX network update -> connection dropped mid-reply (camera moved) - treating as OK")
            return True
        except Exception as e:
            log(f"    VAPIX network update raised {type(e).__name__}: {e}")
            return False

    def onvif():
        gw = ('<?xml version="1.0"?><Envelope xmlns="http://www.w3.org/2003/05/soap-envelope"><Header/><Body>'
              '<SetNetworkDefaultGateway xmlns="http://www.onvif.org/ver10/device/wsdl">'
              f'<IPv4Address>{gateway}</IPv4Address></SetNetworkDefaultGateway></Body></Envelope>')
        ifc = ('<?xml version="1.0"?><Envelope xmlns="http://www.w3.org/2003/05/soap-envelope" '
               'xmlns:tds="http://www.onvif.org/ver10/device/wsdl" xmlns:tt="http://www.onvif.org/ver10/schema"><Header/><Body>'
               '<tds:SetNetworkInterfaces><tds:InterfaceToken>eth0</tds:InterfaceToken>'
               '<tds:NetworkInterface><tt:Enabled>true</tt:Enabled><tt:IPv4><tt:Enabled>true</tt:Enabled>'
               f'<tt:Manual><tt:Address>{new_ip}</tt:Address><tt:PrefixLength>{cidr}</tt:PrefixLength></tt:Manual>'
               '<tt:DHCP>false</tt:DHCP></tt:IPv4></tds:NetworkInterface></tds:SetNetworkInterfaces></Body></Envelope>')
        try:
            requests.post(f"http://{ip}/vapix/services", data=gw,
                          headers={"Content-Type": "application/soap+xml"}, auth=auth, timeout=15)
            r = requests.post(f"http://{ip}/vapix/services", data=ifc,
                              headers={"Content-Type": "application/soap+xml"}, auth=auth, timeout=15)
            ok = r.status_code == 200 and 'Fault' not in (r.text or '')
            log(f"    ONVIF network update -> HTTP {r.status_code} {'ok' if ok else r.text[:120]!r}")
            return ok
        except Exception as e:
            log(f"    ONVIF network update raised {type(e).__name__}: {e}")
            return False

    order = (vapix, onvif) if prefer_vapix else (onvif, vapix)
    for fn in order:
        if fn():
            return True
    return False


# ---------------------------------------------------------------- discovery
def local_ipv4s():
    out = []
    try:
        rc, txt = ps("Get-NetIPAddress -AddressFamily IPv4 | Where-Object {$_.IPAddress -notlike '127.*'} | "
                     "ForEach-Object { \"$($_.InterfaceIndex) $($_.IPAddress) $($_.PrefixLength) $($_.AddressState)\" }")
        for line in txt.splitlines():
            p = line.split()
            if len(p) >= 3:
                out.append({'idx': int(p[0]), 'ip': p[1], 'prefix': int(p[2]), 'state': p[3] if len(p) > 3 else ''})
    except Exception:
        pass
    return out


def ssdp_discover(iface_ips, timeout=3.0):
    """M-SEARCH on every local address; return the set of Axis IPs that answered."""
    found = set()
    sock = None
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM, socket.IPPROTO_UDP)
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            sock.bind(('', 1900))
        except OSError:
            sock.bind(('', 0))
        for jip in (iface_ips or ['0.0.0.0']):
            try:
                sock.setsockopt(socket.IPPROTO_IP, socket.IP_ADD_MEMBERSHIP,
                                socket.inet_aton('239.255.255.250') + socket.inet_aton(jip))
            except OSError:
                pass
        sock.setsockopt(socket.IPPROTO_IP, socket.IP_MULTICAST_TTL, 4)
        sock.settimeout(0.5)
        msearch = ("M-SEARCH * HTTP/1.1\r\nHOST: 239.255.255.250:1900\r\nMAN: \"ssdp:discover\"\r\n"
                   "MX: 2\r\nST: ssdp:all\r\n\r\n").encode()
        for jip in (iface_ips or ['0.0.0.0']):
            try:
                sock.setsockopt(socket.IPPROTO_IP, socket.IP_MULTICAST_IF, socket.inet_aton(jip))
                sock.sendto(msearch, ('239.255.255.250', 1900))
            except OSError:
                pass
        end = time.time() + timeout
        while time.time() < end:
            try:
                data, addr = sock.recvfrom(8192)
            except socket.timeout:
                continue
            except Exception:
                continue
            if b'AXIS' in data.upper():
                found.add(addr[0])
    finally:
        if sock:
            try:
                sock.close()
            except Exception:
                pass
    return found


def arp_axis_ips():
    found = set()
    try:
        rc, txt = ps("Get-NetNeighbor -AddressFamily IPv4 | Where-Object {$_.State -ne 'Unreachable'} | "
                     "ForEach-Object { \"$($_.IPAddress) $($_.LinkLayerAddress)\" }")
        for line in txt.splitlines():
            p = line.split()
            if len(p) == 2 and norm_mac(p[1]).lower()[:8] in AXIS_OUIS:
                found.add(p[0])
    except Exception:
        pass
    return found


def tcp_open(ip, port=80, timeout=0.7):
    try:
        with socket.create_connection((ip, port), timeout=timeout):
            return True
    except Exception:
        return False


def discover(iface_ips, extra_ips=()):
    ips = set()
    ips |= ssdp_discover(iface_ips)
    ips |= arp_axis_ips()
    for ip in (FACTORY_IP,) + tuple(extra_ips):
        if tcp_open(ip):
            ips.add(ip)
    return ips


# ---------------------------------------------------------------- NIC aliases
class Aliases:
    def __init__(self, iface_idx, keep):
        self.idx = iface_idx
        self.keep = keep
        self.added = []
        self.routes = []

    def ensure(self, ip, prefix, skip_as_source=False):
        for a in local_ipv4s():
            if a['ip'] == ip:
                return True
        if not is_admin():
            log(f"  not running as Administrator - cannot add {ip}/{prefix}; add it by hand on your NIC")
            return False
        sas = ' -SkipAsSource $true' if skip_as_source else ''
        rc, txt = ps(f"New-NetIPAddress -InterfaceIndex {self.idx} -IPAddress '{ip}' -PrefixLength {prefix} "
                     f"-PolicyStore ActiveStore{sas} -ErrorAction Stop | Out-Null")
        if rc == 0 or 'already exists' in txt.lower() or 'DuplicateObject' in txt:
            self.added.append(ip)
            log(f"  added temporary address {ip}/{prefix} on interface {self.idx}")
            time.sleep(2)
            return True
        log(f"  could not add {ip}/{prefix}: {txt.strip()[:160]}")
        return False

    def cleanup(self):
        if self.keep:
            return
        for dest in self.routes:
            ps(f"Remove-NetRoute -DestinationPrefix '{dest}' -InterfaceIndex {self.idx} -Confirm:$false -ErrorAction SilentlyContinue")
        for ip in self.added:
            ps(f"Remove-NetIPAddress -IPAddress '{ip}' -Confirm:$false -ErrorAction SilentlyContinue")
            log(f"  removed temporary address {ip}")


def pick_iface(arg):
    addrs = local_ipv4s()
    if arg:
        for a in addrs:
            if str(a['idx']) == str(arg) or a['ip'] == arg:
                return a['idx']
        rc, txt = ps(f"(Get-NetAdapter -Name '{arg}' -ErrorAction SilentlyContinue).ifIndex")
        if txt.strip().isdigit():
            return int(txt.strip())
        raise SystemExit(f"Interface {arg!r} not found. Use --list-ifaces.")
    # default: the first UP physical adapter with a cable
    rc, txt = ps("Get-NetAdapter -Physical | Where-Object {$_.Status -eq 'Up' -and $_.MediaType -eq '802.3'} | "
                 "Sort-Object ifIndex | Select-Object -First 1 -ExpandProperty ifIndex")
    if txt.strip().isdigit():
        return int(txt.strip())
    raise SystemExit("No wired adapter is Up. Plug into the camera switch, or pass --iface.")


def free_host_ip(net, avoid):
    """Pick an unused address for this laptop inside `net`: walk down from just
    below the top of the range (the top host is the switch's own address on this
    site, the bottom host the gateway), skipping listed cameras and anything
    that answers."""
    avoid = set(avoid) | {str(net.network_address + 1), str(net.broadcast_address - 1)}
    top = int(net.broadcast_address) - 2
    bottom = int(net.network_address) + 2
    tried = 0
    for h in range(top, bottom - 1, -1):
        ip = str(ipaddress.ip_address(h))
        if ip in avoid:
            continue
        tried += 1
        if tried > 12:
            break
        if not tcp_open(ip, 80, 0.3) and subprocess.run(['ping', '-n', '1', '-w', '400', ip], capture_output=True,
                                                          creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0)).returncode != 0:
            return ip
    return None


def here_net(a, cams):
    """The /27 (or whatever) of the switch the laptop is plugged into: --here as CIDR,
    or --here N = switch number N of the list."""
    if not a.here:
        return None
    if '/' in a.here:
        return ipaddress.ip_network(a.here, strict=False)
    n = int(a.here)
    nets = {}
    for c in cams:
        sw = switch_no(c.get('switch', ''))
        if sw == n:
            nets[cam_net(a, c)] = nets.get(cam_net(a, c), 0) + 1
    if not nets:
        raise SystemExit(f"--here {n}: no camera in the list belongs to switch {n}; give --here as a CIDR instead")
    return max(nets, key=nets.get)


def setup_access(a, cams, al):
    """Give the laptop an address on the local switch's subnet, the Axis factory
    subnet and link-local, plus a route to the other camera subnets through the
    local gateway. Returns the local net (or None)."""
    if a.no_alias:
        return None
    al.ensure('192.168.0.89', 24)
    al.ensure(LINKLOCAL_IP, 16, skip_as_source=True)
    hn = here_net(a, cams)
    if hn is None:
        log("  no --here given: not adding an address on the camera subnet (pass --here <switch no or CIDR>)")
        return None
    have = [x['ip'] for x in local_ipv4s() if ipaddress.ip_address(x['ip']) in hn]
    if not have:
        mine = free_host_ip(hn, {c['ip'] for c in cams})
        if mine:
            al.ensure(mine, hn.prefixlen)
        else:
            log(f"  could not find a free address in {hn} for this laptop")
            return hn
    gw = str(hn.network_address + 1)
    others = {cam_net(a, c) for c in cams} - {hn}
    if others and not a.no_route:
        if subprocess.run(['ping', '-n', '1', '-w', '700', gw], capture_output=True,
                          creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0)).returncode == 0:
            for on in sorted(others, key=lambda n: int(n.network_address)):
                rc, txt = ps(f"New-NetRoute -DestinationPrefix '{on}' -InterfaceIndex {al.idx} -NextHop '{gw}' "
                             f"-PolicyStore ActiveStore -ErrorAction SilentlyContinue | Out-Null")
                al.routes.append(str(on))
            log(f"  routes to {len(others)} other camera subnet(s) via {gw} added (temporary)")
        else:
            log(f"  gateway {gw} does not answer ping: cameras on OTHER switches will not be reachable from here")
    return hn


# ---------------------------------------------------------------- SNMP (pure python, v2c)


def _len(n):
    if n < 0x80:
        return bytes([n])
    b = n.to_bytes((n.bit_length() + 7) // 8, 'big')
    return bytes([0x80 | len(b)]) + b


def _tlv(tag, body):
    return bytes([tag]) + _len(len(body)) + body


def _int(v, tag=0x02):
    if v == 0:
        return _tlv(tag, b'\x00')
    b = v.to_bytes((v.bit_length() + 8) // 8, 'big', signed=True)
    return _tlv(tag, b)


def _oid(s):
    parts = [int(x) for x in s.strip('.').split('.')]
    out = bytes([40 * parts[0] + parts[1]])
    for p in parts[2:]:
        if p < 0x80:
            out += bytes([p])
        else:
            chunks = []
            while p:
                chunks.insert(0, p & 0x7f)
                p >>= 7
            out += bytes([c | 0x80 for c in chunks[:-1]] + [chunks[-1]])
    return _tlv(0x06, out)


def _rd(buf, pos):
    tag = buf[pos]
    ln = buf[pos + 1]
    pos += 2
    if ln & 0x80:
        n = ln & 0x7f
        ln = int.from_bytes(buf[pos:pos + n], 'big')
        pos += n
    return tag, buf[pos:pos + ln], pos + ln


def _dec_oid(b):
    first = b[0]
    parts = [first // 40, first % 40]
    v = 0
    for c in b[1:]:
        v = (v << 7) | (c & 0x7f)
        if not c & 0x80:
            parts.append(v)
            v = 0
    return '.'.join(map(str, parts))


def _dec_val(tag, b):
    if tag in (0x02, 0x41, 0x42, 0x43, 0x46):
        return int.from_bytes(b, 'big', signed=(tag == 0x02)) if b else 0
    if tag == 0x04:
        return b
    if tag == 0x06:
        return _dec_oid(b)
    if tag == 0x40:
        return '.'.join(map(str, b))
    if tag in (0x80, 0x81, 0x82):
        return None  # noSuchObject / noSuchInstance / endOfMibView
    return b


def _request(ip, community, oid, pdu_tag, timeout=2.0, reqid=None):
    reqid = reqid or int(time.time() * 1000) & 0x7fffffff
    vb = _tlv(0x30, _tlv(0x30, _oid(oid) + b'\x05\x00'))
    pdu = _tlv(pdu_tag, _int(reqid) + _int(0) + _int(0) + vb)
    msg = _tlv(0x30, _int(1) + _tlv(0x04, community.encode()) + pdu)
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    s.settimeout(timeout)
    try:
        s.sendto(msg, (ip, 161))
        data, _ = s.recvfrom(65535)
    finally:
        s.close()
    tag, body, _ = _rd(data, 0)
    _, _, p = _rd(body, 0)          # version
    _, _, p = _rd(body, p)          # community
    ptag, pdu, _ = _rd(body, p)
    _, _, p = _rd(pdu, 0)           # request id
    _, err, p = _rd(pdu, p)         # error status
    _, _, p = _rd(pdu, p)           # error index
    if err and int.from_bytes(err, 'big'):
        return None, None
    _, vbl, _ = _rd(pdu, p)
    _, vb0, _ = _rd(vbl, 0)
    otag, ob, q = _rd(vb0, 0)
    vtag, vb, _ = _rd(vb0, q)
    return _dec_oid(ob), (vtag, _dec_val(vtag, vb))


def snmp_get(ip, community, oid, timeout=2.0):
    try:
        o, v = _request(ip, community, oid, 0xa0, timeout)
        return None if v is None else v[1]
    except Exception:
        return None


def snmp_walk(ip, community, base, timeout=2.0, limit=5000):
    """Yield (oid, value) for every object under `base`."""
    out = []
    cur = base
    for _ in range(limit):
        try:
            o, v = _request(ip, community, cur, 0xa1, timeout)
        except Exception:
            break
        if not o or not (o == base or o.startswith(base + '.')) or v is None or v[0] == 0x82:
            break
        out.append((o, v[1]))
        cur = o
    return out


# bridge MIB helpers --------------------------------------------------------
OID_SYSDESCR = '1.3.6.1.2.1.1.1.0'
OID_SYSNAME = '1.3.6.1.2.1.1.5.0'
OID_IFNAME = '1.3.6.1.2.1.31.1.1.1.1'
OID_IFDESCR = '1.3.6.1.2.1.2.2.1.2'
OID_BASEPORT_IFINDEX = '1.3.6.1.2.1.17.1.4.1.2'
OID_DOT1Q_FDB_PORT = '1.3.6.1.2.1.17.7.1.2.2.1.2'
OID_DOT1D_FDB_PORT = '1.3.6.1.2.1.17.4.3.1.2'


def mac_table(ip, community='public', timeout=2.0):
    """Return {MAC: port_name} from the switch's forwarding table, or None if no SNMP."""
    if snmp_get(ip, community, OID_SYSDESCR, timeout) is None:
        return None
    ifnames = {}
    for o, v in snmp_walk(ip, community, OID_IFNAME, timeout) or snmp_walk(ip, community, OID_IFDESCR, timeout):
        ifnames[int(o.rsplit('.', 1)[1])] = v.decode(errors='ignore') if isinstance(v, bytes) else str(v)
    port2if = {}
    for o, v in snmp_walk(ip, community, OID_BASEPORT_IFINDEX, timeout):
        port2if[int(o.rsplit('.', 1)[1])] = int(v)
    rows = snmp_walk(ip, community, OID_DOT1Q_FDB_PORT, timeout)
    nmac = 6
    if not rows:
        rows = snmp_walk(ip, community, OID_DOT1D_FDB_PORT, timeout)
    table = {}
    for o, port in rows:
        idx = o.split('.')
        mac = ':'.join('%02X' % int(x) for x in idx[-nmac:])
        port = int(port)
        if port == 0:
            continue
        ifi = port2if.get(port, port)
        table[mac] = ifnames.get(ifi, f'port {port}')
    return table



# ---------------------------------------------------------------- LLDP (pktmon)
def lldp_listen(seconds=35):
    """Capture LLDP frames with pktmon (built into Windows 10/11, needs admin).
    Returns a list of dicts: switch_name, port_id, port_desc, mgmt_ip, system_desc."""
    import struct
    import tempfile
    sp = subprocess
    kw = dict(capture_output=True, timeout=15, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    etl = tempfile.mktemp(suffix='.etl')
    pcap = tempfile.mktemp(suffix='.pcapng')
    results = []
    try:
        sp.run(['pktmon', 'stop'], **kw)
        sp.run(['pktmon', 'filter', 'remove'], **kw)
        r = sp.run(['pktmon', 'filter', 'add', 'LLDP', '-d', '0x88CC'], **kw)
        if r.returncode != 0:
            r = sp.run(['pktmon', 'filter', 'add', 'LLDP', '-t', '0x88CC'], **kw)
        if r.returncode != 0:
            log(f"  pktmon filter failed: {r.stderr.decode(errors='replace')[:120]}")
            return results
        r = sp.run(['pktmon', 'start', '-c', '--pkt-size', '512', '-f', etl], **kw)
        if r.returncode != 0:
            log(f"  pktmon start failed: {r.stderr.decode(errors='replace')[:120]}")
            return results
        log(f"  listening for LLDP for {seconds}s (switches announce every 30s) ...")
        time.sleep(seconds)
        sp.run(['pktmon', 'stop'], **kw)
        time.sleep(0.3)
        sp.run(['pktmon', 'etl2pcap', etl, '-o', pcap], **kw)
        if not os.path.exists(pcap) or os.path.getsize(pcap) < 100:
            return results
        with open(pcap, 'rb') as f:
            data = f.read()
        packets = []
        off = 0
        while off + 12 <= len(data):
            btype = struct.unpack('<I', data[off:off + 4])[0]
            blen = struct.unpack('<I', data[off + 4:off + 8])[0]
            if blen < 12 or off + blen > len(data):
                break
            if btype == 6 and off + 28 <= len(data):
                cap_len = struct.unpack('<I', data[off + 20:off + 24])[0]
                if off + 28 + cap_len <= len(data):
                    packets.append(data[off + 28:off + 28 + cap_len])
            off += blen
        seen = set()
        for frame in packets:
            info = _parse_lldp(frame)
            if info and (info['switch_name'], info['port_id']) not in seen:
                seen.add((info['switch_name'], info['port_id']))
                results.append(info)
    except Exception as e:
        log(f"  LLDP capture error: {e}")
    finally:
        try:
            sp.run(['pktmon', 'filter', 'remove'], **kw)
        except Exception:
            pass
        for f in (etl, pcap):
            try:
                os.remove(f)
            except Exception:
                pass
    return results


def _parse_lldp(frame):
    import struct
    if len(frame) < 16:
        return None
    ethertype = struct.unpack('!H', frame[12:14])[0]
    payload_off = 14
    if ethertype == 0x8100 and len(frame) >= 18:
        ethertype = struct.unpack('!H', frame[16:18])[0]
        payload_off = 18
    if ethertype != 0x88CC:
        for i in range(12, min(64, len(frame) - 1)):
            if frame[i] == 0x88 and frame[i + 1] == 0xCC:
                payload_off = i + 2
                ethertype = 0x88CC
                break
    if ethertype != 0x88CC:
        return None
    src = frame[6:12]
    r = {'switch_mac': ':'.join(f'{b:02X}' for b in src), 'switch_name': '', 'port_id': '',
         'port_desc': '', 'system_desc': '', 'mgmt_ip': '', 'chassis_id': ''}
    off = payload_off
    while off + 2 <= len(frame):
        hdr = struct.unpack('!H', frame[off:off + 2])[0]
        ttype, tlen = (hdr >> 9) & 0x7F, hdr & 0x1FF
        off += 2
        if ttype == 0 or off + tlen > len(frame):
            break
        val = frame[off:off + tlen]
        off += tlen
        try:
            if ttype == 1 and len(val) > 1:
                r['chassis_id'] = ':'.join(f'{b:02X}' for b in val[1:7]) if val[0] == 4 else val[1:].decode('utf-8', 'replace').rstrip('\x00')
            elif ttype == 2 and len(val) > 1:
                r['port_id'] = ':'.join(f'{b:02X}' for b in val[1:7]) if val[0] == 3 else val[1:].decode('utf-8', 'replace').rstrip('\x00')
            elif ttype == 4:
                r['port_desc'] = val.decode('utf-8', 'replace').rstrip('\x00')
            elif ttype == 5:
                r['switch_name'] = val.decode('utf-8', 'replace').rstrip('\x00')
            elif ttype == 6:
                r['system_desc'] = val.decode('utf-8', 'replace').rstrip('\x00')
            elif ttype == 8 and len(val) >= 6 and val[0] >= 5 and val[1] == 1:
                r['mgmt_ip'] = socket.inet_ntoa(val[2:6])
        except Exception:
            pass
    return r


# ---------------------------------------------------------------- port map
OUI_NAMES = {
    'b8:a4:4f': 'Axis', '00:40:8c': 'Axis', 'ac:cc:8e': 'Axis', 'e8:27:25': 'Axis',
    '00:07:5f': 'Bosch', '00:09:18': 'Hanwha', 'e4:30:22': 'Hanwha', '00:06:8e': 'HID',
    '00:17:fc': 'Suprema', '00:17:61': 'ZKTeco', '00:1b:8f': 'Edgecore', '70:72:cf': 'Edgecore',
    '00:12:cf': 'Edgecore', 'd0:c7:c0': 'Edgecore', '00:e0:4b': 'Accton/Edgecore',
    '00:0c:c8': 'Honeywell', '00:24:b2': 'Honeywell', '00:50:c2': 'IEEE-registered (small vendor)',
    '00:1e:c0': 'Microchip', '00:04:a3': 'Microchip', 'a4:cf:12': 'Espressif',
    '00:50:f2': 'Microsoft', '00:06:9b': 'Gentner/ClearOne',
}


def vendor_of(mac):
    return OUI_NAMES.get(mac.lower()[:8], '')


def arp_table():
    out = {}
    try:
        rc, txt = ps("Get-NetNeighbor -AddressFamily IPv4 | Where-Object {$_.State -ne 'Unreachable'} | "
                     "ForEach-Object { \"$($_.IPAddress) $($_.LinkLayerAddress)\" }")
        for line in txt.splitlines():
            p = line.split()
            if len(p) == 2 and norm_mac(p[1]):
                out.setdefault(norm_mac(p[1]), p[0])
    except Exception:
        pass
    return out


def _port_sort_key(name):
    nums = re.findall(r'\d+', name or '')
    return (0, int(nums[-1])) if nums else (1, name or '')



# ---------------------------------------------------------------- numbering plan
def cam_number(name):
    m = re.search(r'(\d+)\s*$', name or '')
    return int(m.group(1)) if m else None


def plan_ip(a, n):
    """Camera number -> IP under the site plan: cameras are numbered in blocks of
    --plan-per-switch (one switch each); each block owns --plan-block addresses
    starting at --plan-base + --plan-first. Default plan = 24 cameras per switch,
    32 addresses per block (8 spare), first camera at base+67."""
    base = int(ipaddress.ip_address(a.plan_base))
    k = (n - 1) // a.plan_per_switch
    return str(ipaddress.ip_address(base + a.plan_first + (n - 1 - k * a.plan_per_switch) + k * a.plan_block))


def plan_port(a, n):
    return (n - 1) % a.plan_per_switch + 1


def plan_switch(a, n):
    return (n - 1) // a.plan_per_switch + 1


def plan_ok(a, cams):
    """True if every listed camera's IP matches the plan - otherwise we do not propose anything."""
    bad = []
    for c in cams:
        n = cam_number(c['name'])
        if n is None or plan_ip(a, n) != c['ip']:
            bad.append(c['name'])
    if bad:
        log(f"  plan check: {len(bad)} listed camera(s) do not follow the numbering plan: {', '.join(bad[:8])}")
    return not bad


def parse_mac_table_text(path):
    """Parse a pasted `show mac-address-table` (Edgecore/Cisco-style) into {MAC: 'Port N'}.
    Tolerant: any line with a MAC and a port token like 'Eth 1/ 5', 'Eth1/5', '1/5', 'Gi1/0/5', 'Port 5'."""
    table = {}
    mac_re = re.compile(r'([0-9A-Fa-f]{2}[:\-]){5}[0-9A-Fa-f]{2}|([0-9A-Fa-f]{4}\.){2}[0-9A-Fa-f]{4}')
    port_res = [re.compile(r'(?:Eth(?:ernet)?\s*)?\b\d+/(?:\d+/)?\s*(\d+)\b', re.I),
                re.compile(r'\bPort\s*(\d+)\b', re.I),
                re.compile(r'\b(?:Gi|Fa|Te)\S*?(\d+)\b', re.I)]
    with open(path, encoding='utf-8', errors='ignore') as f:
        for line in f:
            m = mac_re.search(line)
            if not m:
                continue
            mac = norm_mac(m.group(0))
            rest = line[:m.start()] + ' ' + line[m.end():]
            port = None
            for pr in port_res:
                pm = pr.search(rest)
                if pm:
                    port = int(pm.group(1))
                    break
            if mac and port is not None:
                table[mac] = f'Port {port}'
    return table


# ---------------------------------------------------------------- telnet CLI fetch
def telnet_mac_table(ip, user, password, timeout=25, raw_out=None):
    """Log in to the switch CLI over telnet (port 23), run 'show mac-address-table',
    return the raw text (pager handled). Edgecore/Accton and Cisco-style prompts."""
    IAC, DONT, DO, WONT, WILL, SB, SE = 255, 254, 253, 252, 251, 250, 240

    def negotiate(sock, data):
        """Strip telnet IAC sequences, refusing every option."""
        out = bytearray()
        i = 0
        while i < len(data):
            b = data[i]
            if b != IAC:
                out.append(b)
                i += 1
                continue
            if i + 1 >= len(data):
                break
            cmd = data[i + 1]
            if cmd in (DO, DONT, WILL, WONT) and i + 2 < len(data):
                opt = data[i + 2]
                reply = WONT if cmd == DO else (DONT if cmd == WILL else None)
                if reply is not None:
                    try:
                        sock.sendall(bytes([IAC, reply, opt]))
                    except Exception:
                        pass
                i += 3
            elif cmd == SB:
                j = data.find(bytes([IAC, SE]), i)
                i = (j + 2) if j >= 0 else len(data)
            else:
                i += 2
        return bytes(out)

    def read_until(sock, patterns, limit):
        buf = b''
        end = time.time() + limit
        while time.time() < end:
            try:
                chunk = sock.recv(4096)
            except socket.timeout:
                continue
            if not chunk:
                break
            buf += negotiate(sock, chunk)
            low = buf.lower()
            for pat in patterns:
                if pat in low:
                    return buf, pat
        return buf, None

    sock = socket.create_connection((ip, 23), timeout=5)
    sock.settimeout(1.0)
    text = b''
    try:
        buf, hit = read_until(sock, (b'username:', b'login:', b'user name:', b'password:'), 8)
        text += buf
        if hit and hit != b'password:':
            sock.sendall(user.encode() + b'\r\n')
            buf, hit = read_until(sock, (b'password:',), 8)
            text += buf
        if hit == b'password:':
            sock.sendall(password.encode() + b'\r\n')
        buf, hit = read_until(sock, (b'#', b'>', b'invalid', b'failed', b'denied', b'username:', b'login:'), 10)
        text += buf
        if hit in (b'invalid', b'failed', b'denied', b'username:', b'login:'):
            raise RuntimeError('login refused')
        if hit == b'>':
            sock.sendall(b'enable\r\n')
            buf, hit2 = read_until(sock, (b'password:', b'#'), 6)
            text += buf
            if hit2 == b'password:':
                sock.sendall(password.encode() + b'\r\n')
                buf, _ = read_until(sock, (b'#',), 6)
                text += buf
        sock.sendall(b'terminal length 0\r\n')
        read_until(sock, (b'#',), 4)
        sock.sendall(b'show mac-address-table\r\n')
        out = b''
        end = time.time() + timeout
        quiet = 0
        mac_re = re.compile(rb'([0-9A-Fa-f]{2}[:\-]){5}[0-9A-Fa-f]{2}|([0-9A-Fa-f]{4}\.){2}[0-9A-Fa-f]{4}')
        while time.time() < end:
            try:
                chunk = sock.recv(8192)
            except socket.timeout:
                quiet += 1
                if out and quiet >= 3:
                    break
                continue
            except OSError:
                break          # switch closed the session: whatever we have is the answer
            if not chunk:
                break
            quiet = 0
            chunk = negotiate(sock, chunk)
            out += chunk
            if b'--more--' in out[-200:].lower() or b'-- more --' in out[-200:].lower():
                try:
                    sock.sendall(b' ')
                except OSError:
                    break
            stripped = out.rstrip()
            if stripped.endswith((b'#', b'>')) and mac_re.search(out) and b'more' not in out[-40:].lower():
                # prompt came back after the listing; give a late page half a second
                if time.time() + 1 < end:
                    time.sleep(0.5)
                    try:
                        more = sock.recv(8192)
                        if more:
                            out += negotiate(sock, more)
                            continue
                    except (socket.timeout, OSError):
                        pass
                break
        try:
            sock.sendall(b'exit\r\n')
        except Exception:
            pass
    finally:
        try:
            sock.close()
        except Exception:
            pass
    raw = out.decode('utf-8', 'ignore').replace('\x08', '')
    raw = re.sub(r'\x1b\[[0-9;]*[A-Za-z]', '', raw)
    if raw_out:
        try:
            with open(raw_out, 'w', encoding='utf-8') as f:
                f.write(raw)
        except Exception:
            pass
    return raw

def cmd_ports(cams, a):
    """Who is on which switch port. Needs the switch's SNMP read community (default 'public')."""
    bymac = {c['mac']: c for c in cams}
    al = Aliases(pick_iface(a.iface), a.keep_alias) if not a.no_alias else None
    hn = setup_access(a, cams, al) if al else None
    switch_ip = a.switch
    lldp = []
    if (not switch_ip and not a.mac_table and not (a.telnet and a.here)) or a.lldp:
        if is_admin():
            lldp = lldp_listen(a.lldp_seconds)
            for i in lldp:
                log(f"  LLDP: switch {i['switch_name'] or i['chassis_id']} says this laptop is on port "
                    f"{i['port_id']} {('(' + i['port_desc'] + ')') if i['port_desc'] else ''} mgmt IP {i['mgmt_ip'] or '?'}  {i['system_desc'][:60]}")
            if not switch_ip:
                for i in lldp:
                    if i['mgmt_ip']:
                        switch_ip = i['mgmt_ip']
                        break
        else:
            log("  not Administrator, skipping LLDP")
    tried = []
    table = None
    if a.mac_table:
        table = parse_mac_table_text(a.mac_table)
        log(f"  MAC table read from {a.mac_table}: {len(table)} entries")
        switch_ip = switch_ip or 'pasted'
    elif a.telnet:
        if ':' not in a.telnet:
            raise SystemExit('--telnet must be user:password')
        tuser, tpass = a.telnet.split(':', 1)
        tip = switch_ip or (str(hn.broadcast_address - 1) if hn is not None else None)
        if not tip:
            raise SystemExit('--telnet needs --switch <ip> or --here so I know which switch to log in to')
        raw_path = os.path.join(os.path.dirname(a.results), f"mac_table_{tip.replace('.', '_')}.txt")
        try:
            log(f"  logging in to {tip} over telnet as {tuser} ...")
            raw = telnet_mac_table(tip, tuser, tpass, raw_out=raw_path)
            table = parse_mac_table_text(raw_path)
            log(f"  'show mac-address-table' from {tip}: {len(table)} entries (raw saved to {raw_path})")
            if not table:
                log("  nothing parsed - open the raw file and send it to me")
            switch_ip = tip
        except Exception as e:
            log(f"  telnet to {tip} failed: {e}")
            table = None
    cands = [] if table is not None else ([switch_ip] if switch_ip else [])
    if hn is not None:
        cands.append(str(hn.broadcast_address - 1))   # this site: the switch is the top host of its /27
    cands += [a.gateway if a.gateway and a.gateway != 'auto' else '', '192.168.2.10', '192.168.1.1', '192.168.0.1']
    for cand in cands:
        if not cand or cand in tried:
            continue
        tried.append(cand)
        for comm in [a.community] + [c for c in ('Public', 'public', 'private') if c != a.community]:
            t = mac_table(cand, comm, timeout=1.5)
            if t is not None:
                descr = snmp_get(cand, comm, OID_SYSDESCR)
                name = snmp_get(cand, comm, OID_SYSNAME)
                log(f"  switch {cand} (community {comm!r}): {name!r} {descr!r}, {len(t)} MAC entries")
                table, switch_ip = t, cand
                break
        if table is not None:
            break
    arp = arp_table()
    # wake the ARP table and the switch's FDB so idle devices show up
    live = {}
    for ip in discover([x['ip'] for x in local_ipv4s()]):
        info = probe(ip, timeout=2)
        if info and info['mac']:
            live[info['mac']] = info
    for c in cams:
        if c['mac'] not in live and tcp_open(c['ip'], 80, 0.4):
            info = probe(c['ip'], timeout=2)
            if info and info['mac']:
                live[info['mac']] = info
    arp.update({m: i['ip'] for m, i in live.items()})
    if table is None:
        log("  No SNMP answer from any switch candidate. Give me the switch IP with --switch and its read "
            "community with --community, or enable SNMP on it (Edgecore: admin/admin web UI -> SNMP).")
        log("  Fallback: 'ports --snapshot before.json', unplug ONE port, then "
            "'ports --diff before.json --diff-port N' to record which MAC vanished.")
        table = {}
    my_macs = set()
    try:
        rc, txt = ps("Get-NetAdapter | ForEach-Object { $_.MacAddress }")
        my_macs = {norm_mac(x) for x in txt.split() if norm_mac(x)}
    except Exception:
        pass
    rows = []
    macs = set(table) | set(live) | set(bymac)
    for mac in macs:
        c = bymac.get(mac)
        info = live.get(mac)
        port = table.get(mac, '')
        if mac in my_macs:
            who = 'THIS LAPTOP'
        elif c:
            who = c['name']
        elif info:
            who = f"Axis {info['model']} (NOT in list)"
        else:
            who = vendor_of(mac) or 'unknown device'
        status = ''
        if c:
            if not info:
                status = 'not answering'
            elif info['ip'] != c['ip']:
                status = f"wrong IP (at {info['ip']})"
            elif is_factory(c['ip']):
                status = 'factory (no password)'
            else:
                status = 'OK'
        rows.append({'Port': port, 'MAC': mac, 'Vendor': vendor_of(mac), 'IP': (info['ip'] if info else arp.get(mac, '')),
                     'Device': who, 'TargetIP': c['ip'] if c else '', 'Model': (info['model'] if info else (c['model'] if c else '')),
                     'Status': status, 'Switch': switch_ip or ''})
    rows.sort(key=lambda r: (_port_sort_key(r['Port']), r['Device']))

    # --- expected ports: from the Lenel sheet when the list has them, else from the numbering plan
    def port_no(pname):
        nums = re.findall(r'\d+', pname or '')
        return int(nums[-1]) if nums else None
    has_sheet = any(c.get('port') for c in cams)
    use_plan = (not has_sheet) and a.plan_base and plan_ok(a, cams)
    sw_no = a.switch_number
    if not sw_no and (has_sheet or use_plan):
        votes = {}
        for r in rows:
            if r['TargetIP'] and r['Port']:
                c = bymac.get(r['MAC'])
                n = cam_number(c['name']) if c else None
                k = switch_no(c.get('switch', '')) if (c and has_sheet) else (plan_switch(a, n) if n else None)
                if k:
                    votes[k] = votes.get(k, 0) + 1
        if votes:
            sw_no = max(votes, key=votes.get)
            log(f"  this looks like switch #{sw_no} ({votes[sw_no]} of {sum(votes.values())} listed cameras on it agree)"
                + ("" if len(votes) == 1 else f"; others vote {votes}"))
    if has_sheet and sw_no:
        at_port = {}
        for c in cams:
            if switch_no(c.get('switch', '')) == sw_no and c.get('port'):
                at_port[port_no(c['port'])] = c
    else:
        at_port = {}
    proposals = []
    for r in rows:
        r['ExpectedPort'] = ''
        r['Proposed'] = ''
        r['LogName'] = ''
        lg = a.log.get(r['MAC'])
        if lg:
            r['LogName'] = f"{lg['name']} {lg['ip']}"
        pn = port_no(r['Port'])
        c = bymac.get(r['MAC'])
        if c:
            if has_sheet:
                exp, exp_sw = port_no(c.get('port')), switch_no(c.get('switch', ''))
            elif use_plan:
                n = cam_number(c['name'])
                exp, exp_sw = (plan_port(a, n), plan_switch(a, n)) if n else (None, None)
            else:
                exp, exp_sw = None, None
            if exp:
                r['ExpectedPort'] = f"{exp_sw}/{exp}" if exp_sw else str(exp)
            if pn and exp and sw_no and exp_sw and (exp_sw, exp) != (sw_no, pn):
                r['Status'] = (r['Status'] + '; ' if r['Status'] else '') + f"sheet puts it on switch {exp_sw} port {exp}"
            if lg and lg['name'] != c['name']:
                r['Status'] = (r['Status'] + '; ' if r['Status'] else '') + f"5/5 log named this MAC {lg['name']}"
        elif r['Device'].startswith('Axis') and pn and sw_no:
            if has_sheet:
                sc = at_port.get(pn)
                if sc:
                    r['Proposed'] = f"{sc['name']} {sc['ip']} (sheet expected MAC {sc['mac']} here)"
                    proposals.append({'Camera Name': sc['name'], 'IP Address': sc['ip'], 'Gateway': cam_gateway(a, sc),
                                      'Subnet': cam_subnet(a, sc), 'Model': r['Model'], 'New IP': sc['ip'],
                                      'MAC Address': r['MAC'], 'Processed': 'No', 'Port': r['Port'], 'Switch': sw_no,
                                      'SheetMAC': sc['mac'], 'LogName': r['LogName']})
            elif use_plan and pn <= a.plan_per_switch:
                n = (sw_no - 1) * a.plan_per_switch + pn
                name = a.plan_name.format(n=n)
                gw_ip = plan_ip(a, n)
                fake = {'ip': gw_ip, 'gateway': '', 'subnet': ''}
                r['Proposed'] = f"{name} {gw_ip}"
                proposals.append({'Camera Name': name, 'IP Address': gw_ip, 'Gateway': cam_gateway(a, fake),
                                  'Subnet': cam_subnet(a, fake), 'Model': r['Model'], 'New IP': gw_ip,
                                  'MAC Address': r['MAC'], 'Processed': 'No', 'Port': r['Port'], 'Switch': sw_no,
                                  'SheetMAC': '', 'LogName': r['LogName']})
    if proposals:
        pout = os.path.join(os.path.dirname(a.results), 'proposed_cameras.csv')
        new_file = not os.path.exists(pout)
        have = set()
        if not new_file:
            with open(pout, newline='') as f:
                have = {norm_mac(x.get('MAC Address', '')) for x in csv.DictReader(f)}
        with open(pout, 'a', newline='') as f:
            w = csv.DictWriter(f, fieldnames=list(proposals[0].keys()))
            if new_file:
                w.writeheader()
            for pr in proposals:
                if pr['MAC Address'] not in have:
                    w.writerow(pr)
        log(f"  {len(proposals)} unlisted Axis unit(s) given a name + IP from their port -> {pout}  (REVIEW it, then: run --csv {pout})")

    out = os.path.join(os.path.dirname(a.results), 'switch_ports.csv')
    with open(out, 'w', newline='') as f:
        w = csv.DictWriter(f, fieldnames=['Port', 'Device', 'IP', 'TargetIP', 'MAC', 'Vendor', 'Model', 'Status', 'ExpectedPort', 'LogName', 'Proposed', 'Switch'])
        w.writeheader()
        w.writerows(rows)
    print(f"\n{'Port':10} {'Device':28} {'IP':15} {'MAC':18} {'Status / proposal'}")
    for r in rows:
        if r['Port'] or (r['IP'] and not r['TargetIP']):
            tail = r['Status'] + ((' -> ' + r['Proposed']) if r['Proposed'] else '')
            print(f"{r['Port']:10} {r['Device'][:28]:28} {r['IP']:15} {r['MAC']:18} {tail}")
    off = [r['Device'] for r in rows if r['TargetIP'] and not r['Port']]
    if off:
        print(f"({len(off)} listed cameras are not on this switch; full table in the CSV)")
    if al:
        al.cleanup()
    on_switch = [r for r in rows if r['Port']]
    log(f"{len(on_switch)} MACs on the switch, {sum(1 for r in on_switch if r['TargetIP'])} are list cameras, "
        f"{sum(1 for r in on_switch if r['Vendor'] and r['Vendor'] != 'Axis')} other known vendors. Saved {out}")
    if a.snapshot:
        with open(a.snapshot, 'w') as f:
            json.dump(sorted(macs - my_macs), f)
        log(f"snapshot of {len(macs - my_macs)} MACs saved to {a.snapshot}")
    if a.diff:
        with open(a.diff) as f:
            before = set(json.load(f))
        gone = before - macs
        log(f"MACs that disappeared since {a.diff}: " + (', '.join(
            f"{m} ({bymac[m]['name'] if m in bymac else vendor_of(m) or '?'})" for m in sorted(gone)) or 'none'))
        if a.diff_port and gone:
            with open(os.path.join(os.path.dirname(a.results), 'switch_ports_manual.csv'), 'a', newline='') as f:
                w = csv.writer(f)
                for m in sorted(gone):
                    w.writerow([a.diff_port, bymac[m]['name'] if m in bymac else '', m, datetime.now().isoformat()])
            log(f"  recorded as port {a.diff_port} in switch_ports_manual.csv")


# ---------------------------------------------------------------- the work
def wait_for(ip, mac, seconds):
    end = time.time() + seconds
    while time.time() < end:
        info = probe(ip, timeout=2)
        if info and info['mac'] == mac:
            return info
        time.sleep(2)
    return None


def program_camera(cam, found, password, a, extra):
    """found = probe dict at the IP where the camera lives right now."""
    ip = found['ip']
    log(f"PROGRAMMING {cam['name']} ({cam['mac']}) found at {ip}, target {cam['ip']}  fw {found.get('firmware') or '?'}")
    if cam['model'] and found['model'] and cam['model'] != found['model']:
        log(f"  NOTE model is {found['model']}, list says {cam['model']} - continuing (same MAC)")
    factory = is_factory(ip)
    cur_pw = None
    if factory:
        log("  step 1: factory camera, setting root password")
        if not create_root(ip, password):
            log("  FAILED to set root password"); return False
        cur_pw = password
    else:
        cur_pw = find_password(ip, a.pw_candidates)
        if not cur_pw:
            log("  FAILED: camera has a root password and none of the candidates work. Factory-reset it (button) and try again.")
            return False
        if cur_pw != password:
            log(f"  step 1: root password is {cur_pw!r}; changing it to the one you gave")
            try:
                r = requests.get(f"http://{ip}/axis-cgi/pwdgrp.cgi",
                                 params={"action": "update", "user": "root", "pwd": password},
                                 auth=HTTPDigestAuth("root", cur_pw), timeout=TIMEOUT)
                if r.status_code == 200 and not r.text.strip().lower().startswith('error'):
                    cur_pw = password
                else:
                    log(f"  could not change root password ({r.status_code}); keeping {cur_pw!r}")
            except Exception as e:
                log(f"  could not change root password: {e}")
    if not test_password(ip, 'root', cur_pw):
        time.sleep(3)
        if not test_password(ip, 'root', cur_pw):
            log("  FAILED: root password did not take"); return False
    log("  step 1 ok")

    gw, sn = cam_gateway(a, cam), cam_subnet(a, cam)
    tnet = cam_net(a, cam)
    if ip != cam['ip']:
        if a.al is not None and not a.no_alias and not any(ipaddress.ip_address(x['ip']) in tnet for x in local_ipv4s()) \
                and (a.here_net is None or tnet == a.here_net):
            mine = free_host_ip(tnet, {c['ip'] for c in a.cams})
            if mine:
                a.al.ensure(mine, tnet.prefixlen)
        log(f"  step 2: static {cam['ip']} / {sn} gw {gw}")
        if not set_network(ip, cur_pw, cam['ip'], sn, gw, found.get('firmware', '')):
            log("  FAILED to set network"); return False
        log(f"  waiting up to {a.wait}s for the camera at {cam['ip']} ...")
        back = wait_for(cam['ip'], cam['mac'], a.wait)
        if not back:
            log(f"  FAILED: nothing with MAC {cam['mac']} answered at {cam['ip']} (is {tnet} reachable from this laptop?)")
            return False
        log("  step 2 ok, camera is at its IP")
    else:
        log("  step 2: already at its IP")
    tip = cam['ip']
    # password may need a moment after reboot
    for _ in range(10):
        if test_password(tip, 'root', cur_pw):
            break
        time.sleep(2)

    if not a.no_hostname:
        log(f"  step 3: hostname {cam['name']}")
        if not set_hostname(tip, cur_pw, cam['name']):
            log("  hostname did not take (not fatal)")
    if extra:
        u, p, role = extra
        log(f"  step 4: user {u} ({role})")
        if user_exists(tip, cur_pw, u):
            log(f"  user {u} already there")
        elif not add_user(tip, cur_pw, u, p, role):
            log(f"  FAILED to add user {u}")
            return False
    write_result(a, cam, tip, found)
    log(f"DONE {cam['name']} -> {tip}")
    return True


def write_result(a, cam, ip, found):
    path = a.results
    new = not os.path.exists(path)
    with open(path, 'a', newline='') as f:
        w = csv.writer(f)
        if new:
            w.writerow(['CameraName', 'IPAddress', 'SerialNumber', 'MACAddress', 'Model', 'Firmware',
                        'BuildingReportsLabel', 'Timestamp'])
        w.writerow([cam['name'], ip, found.get('serial', ''), cam['mac'], found.get('model', ''),
                    found.get('firmware', ''), '', datetime.now().isoformat()])


def cmd_check(cams, a, extra):
    al = Aliases(pick_iface(a.iface), a.keep_alias) if not a.no_alias else None
    if al:
        setup_access(a, cams, al)
    ok, missing, wrong = [], [], []
    print(f"\nChecking {len(cams)} cameras at their IPs ...\n")
    print(f"{'name':10} {'ip':14} {'status'}")
    for cam in cams:
        info = probe(cam['ip'], timeout=2) if tcp_open(cam['ip'], 80, 0.6) else None
        if not info:
            missing.append(cam); print(f"{cam['name']:10} {cam['ip']:14} NO ANSWER"); continue
        if info['mac'] != cam['mac']:
            lg = a.log.get(info['mac'])
            wrong.append(cam); print(f"{cam['name']:10} {cam['ip']:14} DIFFERENT CAMERA there ({info['mac']}, {info['model']}{', log: ' + lg['name'] if lg else ''})"); continue
        if is_factory(cam['ip']):
            wrong.append(cam); print(f"{cam['name']:10} {cam['ip']:14} right camera but NO PASSWORD (factory) - run will fix"); continue
        pw = find_password(cam['ip'], a.pw_candidates)
        if not pw:
            wrong.append(cam); print(f"{cam['name']:10} {cam['ip']:14} password unknown"); continue
        extra_s = ''
        if extra:
            e = user_exists(cam['ip'], pw, extra[0])
            extra_s = f", user {extra[0]} {'present' if e else 'MISSING' if e is False else '?'}"
        lg = a.log.get(cam['mac'])
        lg_s = f"  (log: {lg['name']})" if lg and lg['name'] != cam['name'] else ''
        ok.append(cam); print(f"{cam['name']:10} {cam['ip']:14} OK  fw {info['firmware']}  root pw {pw!r}{extra_s}{lg_s}")
    if al:
        al.cleanup()
    print(f"\nOK {len(ok)}   no answer {len(missing)}   needs work {len(wrong)}")
    if missing:
        print("No answer (defaulted, unpowered, or not on this network): " + ', '.join(c['name'] for c in missing))
    if wrong:
        print("Needs work: " + ', '.join(c['name'] for c in wrong))
    print("\nTip: 'run' programs whatever shows up at 192.168.0.90 / link-local / DHCP and matches the list.")


def cmd_run(cams, a, extra):
    if not a.password:
        raise SystemExit("run needs --password (the root password to set). The one used in May is in "
                         f"{os.path.join(APPDATA_CFG, 'passwords.json')} (toolkit: Tools -> Users & Passwords).")
    bymac = {c['mac']: c for c in cams}
    done = set()
    iface = pick_iface(a.iface)
    al = Aliases(iface, a.keep_alias)
    a.al, a.cams = al, cams
    a.here_net = setup_access(a, cams, al)
    iface_ips = [x['ip'] for x in local_ipv4s() if x['idx'] == iface]
    # cameras already fine are done before we start
    log(f"{len(cams)} cameras in the list. Looking for the ones already OK ...")
    for c in cams:
        if tcp_open(c['ip'], 80, 0.5):
            info = probe(c['ip'], timeout=2)
            if info and info['mac'] == c['mac'] and not is_factory(c['ip']) and find_password(c['ip'], a.pw_candidates):
                done.add(c['mac'])
    log(f"{len(done)} already programmed, {len(cams) - len(done)} to do.")
    log(f"Listening on interface {iface} ({', '.join(iface_ips)}). Plug cameras in; Ctrl-C to stop.")

    seen_unknown = set()
    failed = {}
    try:
        while len(done) < len(cams):
            ips = discover(iface_ips)
            for ip in sorted(ips):
                info = probe(ip, timeout=2)
                if not info or not info['mac']:
                    continue
                cam = bymac.get(info['mac'])
                if not cam:
                    if info['mac'] not in seen_unknown:
                        seen_unknown.add(info['mac'])
                        log(f"ignoring {info['model']} {info['mac']} at {ip} - not in the list")
                    continue
                if cam['mac'] in done:
                    continue
                if failed.get(cam['mac'], 0) >= a.max_tries:
                    continue
                if a.one_at_a_time:
                    input(f"\n  Found {cam['name']} at {ip}. Press Enter to program it (Ctrl-C to quit) ")
                if program_camera(cam, info, a.password, a, extra):
                    done.add(cam['mac'])
                else:
                    failed[cam['mac']] = failed.get(cam['mac'], 0) + 1
                    log(f"  will retry {cam['name']} next pass ({failed[cam['mac']]}/{a.max_tries})")
            left = [c['name'] for c in cams if c['mac'] not in done]
            log(f"-- {len(done)}/{len(cams)} done. Waiting for: {', '.join(left[:12])}{' ...' if len(left) > 12 else ''}")
            time.sleep(a.interval)
        log("ALL cameras in the list are programmed.")
    except KeyboardInterrupt:
        log("stopped by user")
    finally:
        al.cleanup()
        left = [c['name'] for c in cams if c['mac'] not in done]
        if left:
            log(f"still not done ({len(left)}): {', '.join(left)}")


def main():
    global LOG_FH
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('mode', nargs='?', default='run', choices=['check', 'run', 'ports', 'list', 'list-ifaces'])
    p.add_argument('--csv', help='camera list: the Lenel programming sheet (.xlsx/.csv), the toolkit run log programmed_cameras.csv, or a camera-list export')
    p.add_argument('--log', help='toolkit run log to cross-reference MACs against (default: programmed_cameras.csv if found)')
    p.add_argument('--here', help='the switch this laptop is plugged into: its number in the list (e.g. 3) or its camera subnet as CIDR (10.20.48.128/27)')
    p.add_argument('--only-net', help='ignore list rows outside this network, e.g. 10.20.48.0/23')
    p.add_argument('--no-route', action='store_true', help="don't add temporary routes to the other camera subnets")
    p.add_argument('--since', default='2026-05-05', help='ignore rows before this date (YYYY-MM-DD); "" for all')
    p.add_argument('--prefix', default='', help='only names starting with this, e.g. KB')
    p.add_argument('--password', help='root password to set on each camera')
    p.add_argument('--subnet', help='camera subnet mask when the list has none, e.g. 255.255.255.224')
    p.add_argument('--gateway', help='camera gateway when the list has none; "auto" = first host of each camera\'s subnet')
    p.add_argument('--extra-user', default='admin:admin', help='user:password to add on every camera ("" to skip)')
    p.add_argument('--extra-role', default='Administrator', choices=['Administrator', 'Operator', 'Viewer'])
    p.add_argument('--no-hostname', action='store_true', help="don't set hostname/Bonjour name to the camera name")
    p.add_argument('--iface', help='interface index, IP or adapter name to use (default: first wired adapter that is Up)')
    p.add_argument('--no-alias', action='store_true', help="don't add temporary IPs to the NIC")
    p.add_argument('--keep-alias', action='store_true', help='leave the temporary IPs on the NIC at exit')
    p.add_argument('--one-at-a-time', action='store_true', help='ask before programming each camera')
    p.add_argument('--wait', type=int, default=120, help='seconds to wait for a camera at its new IP')
    p.add_argument('--interval', type=float, default=4, help='seconds between discovery passes')
    p.add_argument('--max-tries', type=int, default=3)
    p.add_argument('--switch', help='ports: switch management IP (found via LLDP if not given)')
    p.add_argument('--switch-number', type=int, help='ports: which switch of the plan this is (1 = cameras 1-24); guessed from the listed cameras on it if omitted')
    p.add_argument('--plan-base', help='numbering plan: network base, e.g. 10.20.48.0 ("" = no plan)')
    p.add_argument('--plan-first', type=int, default=67, help='numbering plan: host offset of camera 1 (default 67)')
    p.add_argument('--plan-per-switch', type=int, default=24, help='numbering plan: cameras per switch (default 24)')
    p.add_argument('--plan-block', type=int, default=32, help='numbering plan: addresses per switch block (default 32)')
    p.add_argument('--plan-name', default='KBO-{n:03d}', help='numbering plan: name pattern for proposals (default KBO-{n:03d})')
    p.add_argument('--community', default='public', help='ports: SNMP read community')
    p.add_argument('--mac-table', help="ports: a text file with the switch's 'show mac-address-table' output, used instead of SNMP")
    p.add_argument('--telnet', help="ports: user:password to log in to the switch CLI over telnet and run 'show mac-address-table' (used instead of SNMP)")
    p.add_argument('--lldp', action='store_true', help='ports: always listen for LLDP even if --switch is given')
    p.add_argument('--lldp-seconds', type=int, default=35)
    p.add_argument('--snapshot', help='ports: save the set of MACs seen to this file')
    p.add_argument('--diff', help='ports: compare against a saved snapshot and report MACs that vanished')
    p.add_argument('--diff-port', help='ports: with --diff, record the vanished MACs as this port number')
    a = p.parse_args()

    if a.mode == 'list-ifaces':
        rc, txt = ps("Get-NetAdapter | Sort-Object ifIndex | Format-Table ifIndex,Name,Status,MacAddress,LinkSpeed -AutoSize | Out-String -Width 200")
        print(txt)
        for x in local_ipv4s():
            print(f"  if {x['idx']:3}  {x['ip']}/{x['prefix']}  {x['state']}")
        return

    csv_path = a.csv
    if not csv_path:
        for c in _default_csv_candidates():
            if os.path.exists(c):
                csv_path = c
                break
    if not csv_path or not os.path.exists(csv_path):
        raise SystemExit("Camera list not found. Pass --csv <path to programmed_cameras.csv>")
    a.results = os.path.join(os.path.dirname(os.path.abspath(csv_path)), 'reprogrammed_cameras.csv')
    LOG_FH = open(os.path.join(os.path.dirname(os.path.abspath(csv_path)), 'reprogram_log.txt'), 'a')
    log(f"=== {a.mode} ===  list {csv_path}")

    only = ipaddress.ip_network(a.only_net, strict=False) if a.only_net else None
    cams = load_list(csv_path, a.since, a.prefix, only)
    if not cams:
        raise SystemExit("List is empty after filtering (check --since / --prefix / --only-net).")
    missing_net = [c['name'] for c in cams if not (c.get('subnet') or a.subnet)]
    if missing_net and a.mode != 'list':
        raise SystemExit(f"{len(missing_net)} camera(s) have no subnet in the list - pass --subnet (and --gateway or --gateway auto)")
    log_path = a.log
    if not log_path:
        for c in [os.path.join(os.path.dirname(os.path.abspath(csv_path)), 'programmed_cameras.csv')] + _default_csv_candidates():
            if os.path.exists(c) and os.path.abspath(c) != os.path.abspath(csv_path):
                log_path = c
                break
    a.log = load_log(log_path)
    if a.log:
        log(f"cross-referencing {len(a.log)} MACs from the run log {log_path}")
    extra = None
    if a.extra_user:
        if ':' not in a.extra_user:
            raise SystemExit('--extra-user must be user:password')
        u, pw = a.extra_user.split(':', 1)
        extra = (u, pw, a.extra_role)
    a.pw_candidates = load_password_candidates(a.password)

    if a.mode == 'list':
        print(f"{len(cams)} cameras:")
        for c in cams:
            lg = a.log.get(c['mac'])
            lg_s = f"   log: {lg['name']} {lg['ip']}" if lg and (lg['name'] != c['name'] or lg['ip'] != c['ip']) else ''
            sw = f" sw{switch_no(c['switch'])}/p{c['port']}" if c.get('switch') else ''
            print(f"  {c['name']:10} {c['ip']:14} {c['mac']}  {c['model']:10}{sw}{lg_s}")
        return
    nets = sorted({str(cam_net(a, c)) for c in cams})
    log(f"{len(cams)} cameras in list, {cams[0]['ip']} .. {cams[-1]['ip']}, subnet(s) {', '.join(nets[:8])}{' ...' if len(nets) > 8 else ''}")
    if a.mode == 'check':
        cmd_check(cams, a, extra)
    elif a.mode == 'ports':
        cmd_ports(cams, a)
    else:
        cmd_run(cams, a, extra)


if __name__ == '__main__':
    main()
