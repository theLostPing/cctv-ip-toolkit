#!/usr/bin/env python3
"""
switchlib.py - talk to the PoE switch the cameras hang off.

Everything the CCTV IP Toolkit needs from a switch, vendor-agnostic where the
standards allow it and with a small per-vendor table where they do not:

    read_switch(ip, ...)      -> who is on which port (telnet CLI first, SNMP second)
    find_switch(...)          -> where is the switch (LLDP, the usual addresses, then a sweep)
    assign_by_port(...)       -> "the port decides the camera": sheet row for (switch, port)
                                 gives name/IP/gateway/mask, the switch gives the MAC
    poe_cycle(...)            -> power a port off and on (SNMP POWER-ETHERNET-MIB, or CLI)
    lldp_listen(...)          -> which switch/port is THIS PC on (pktmon, needs admin)

Pure python, no third-party modules. SNMP is v2c GET/GETNEXT/SET only.
Lessons from a 7-switch, 151-camera site (2026-10) are baked in:
  - the integrator sheet's MAC column was wrong for every camera; the physical
    port was right. assign_by_port trusts the port.
  - SNMP read community on the config said one thing, the switch answered to
    another. read_switch tries public / Public / private.
  - telnet timed out on one switch where SNMP still worked. Both are kept.
  - the management IP was "the top host of the subnet" on six switches and
    top-2 on the seventh. find_switch sweeps.
  - the uplink ports put a laptop in a different VLAN where it sees nothing;
    assign_by_port says so when this PC is not on a camera port.
Part of Brian Preston's CCTV IP Toolkit (fieldtoolkit.com).
"""
import ipaddress
import os
import re
import socket
import struct
import subprocess
import tempfile
import threading
import time
from concurrent.futures import ThreadPoolExecutor

_NO_WINDOW = getattr(subprocess, 'CREATE_NO_WINDOW', 0)

# ---------------------------------------------------------------- small helpers
OUI_NAMES = {
    'b8:a4:4f': 'Axis', '00:40:8c': 'Axis', 'ac:cc:8e': 'Axis', 'e8:27:25': 'Axis',
    '00:07:5f': 'Bosch', '00:09:18': 'Hanwha', 'e4:30:22': 'Hanwha', '00:16:6c': 'Hanwha',
    '00:06:8e': 'HID', '00:17:fc': 'Suprema', '00:17:61': 'ZKTeco',
    '00:1b:8f': 'Edgecore', '70:72:cf': 'Edgecore', '00:12:cf': 'Edgecore', 'd0:c7:c0': 'Edgecore',
    '00:e0:4b': 'Accton/Edgecore', '00:0c:c8': 'Honeywell', '00:24:b2': 'Honeywell',
    '00:1e:c0': 'Microchip', '00:04:a3': 'Microchip', 'a4:cf:12': 'Espressif',
    '00:50:f2': 'Microsoft', '00:06:9b': 'Gentner/ClearOne',
}
CAMERA_VENDORS = ('Axis', 'Bosch', 'Hanwha')


def norm_mac(s):
    """'b8-a4-4f-01-02-03' / 'B8A44F010203' / 'b8a4.4f01.0203' -> 'B8:A4:4F:01:02:03' ('' if not a MAC)."""
    s = re.sub(r'[^0-9A-Fa-f]', '', s or '')
    if len(s) != 12:
        return ''
    s = s.upper()
    return ':'.join(s[i:i + 2] for i in range(0, 12, 2))


def vendor_of(mac):
    return OUI_NAMES.get((mac or '').lower()[:8], '')


def is_camera_mac(mac, extra_ouis=()):
    v = vendor_of(mac)
    if v in CAMERA_VENDORS:
        return True
    m = (mac or '').lower().replace('-', ':')
    return any(m.startswith(o.lower().replace('-', ':')) for o in extra_ouis)


def port_num(pname):
    """'Eth 1/ 5' / 'Gi1/0/5' / 'Port 5' / '5' -> 5. None if there is no number."""
    nums = re.findall(r'\d+', str(pname or ''))
    return int(nums[-1]) if nums else None


def switch_no(name):
    """Trailing number of a switch name: 'IDF-2-L2POE03' -> 3, '3' -> 3."""
    m = re.search(r'(\d+)\s*$', str(name or ''))
    return int(m.group(1)) if m else None


def same_switch(a, b):
    """Does a camera row's 'switch' value refer to switch b? Exact text, or the
    same trailing number (the sheet said R180-L2POE01, the switch said R140-L2POE01)."""
    a, b = str(a or '').strip(), str(b or '').strip()
    if not a or not b:
        return False
    if a.lower() == b.lower():
        return True
    na, nb = switch_no(a), switch_no(b)
    return na is not None and na == nb


def ps(cmd, timeout=20):
    r = subprocess.run(['powershell', '-NoProfile', '-Command', cmd],
                       capture_output=True, text=True, timeout=timeout, creationflags=_NO_WINDOW)
    return r.returncode, (r.stdout or '') + (r.stderr or '')


def my_macs():
    """MACs of every adapter on this PC (so its own port is recognised in the table)."""
    out = set()
    try:
        rc, txt = ps("Get-NetAdapter | ForEach-Object { $_.MacAddress }")
        out = {norm_mac(x) for x in txt.split() if norm_mac(x)}
    except Exception:
        pass
    return out


def arp_table():
    """{MAC: IP} from the neighbour table (first IP seen per MAC)."""
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


def ping(ip, timeout_ms=400):
    try:
        r = subprocess.run(['ping', '-n', '1', '-w', str(timeout_ms), ip], capture_output=True,
                           text=True, timeout=timeout_ms / 1000 + 2, creationflags=_NO_WINDOW)
        return r.returncode == 0 and 'TTL=' in r.stdout
    except Exception:
        return False


def ping_sweep(network, timeout_ms=400, workers=64, cancel=None):
    """Ping every host of an ipaddress.ip_network; returns the IPs that answered."""
    hosts = [str(h) for h in network.hosts()]
    alive = []
    with ThreadPoolExecutor(max_workers=workers) as ex:
        for ip, ok in zip(hosts, ex.map(lambda h: (not (cancel and cancel.is_set())) and ping(h, timeout_ms), hosts)):
            if ok:
                alive.append(ip)
    return alive


def tcp_open(ip, port, timeout=0.7):
    try:
        with socket.create_connection((ip, port), timeout=timeout):
            return True
    except Exception:
        return False


# ---------------------------------------------------------------- SNMP v2c (tiny BER codec)
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


def _request(ip, community, oid, pdu_tag, timeout=2.0, value=None):
    """One varbind GET (0xa0) / GETNEXT (0xa1) / SET (0xa3). value for SET = (tag, bytes/int).
    Returns (oid, (tag, value)) or (None, None) on an SNMP error."""
    reqid = int(time.time() * 1000) & 0x7fffffff
    if value is None:
        vbval = b'\x05\x00'
    else:
        vtag, v = value
        vbval = _int(v, vtag) if isinstance(v, int) else _tlv(vtag, v)
    vb = _tlv(0x30, _tlv(0x30, _oid(oid) + vbval))
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


def snmp_set_int(ip, community, oid, value, timeout=2.0):
    """SET an INTEGER. True when the switch echoed the value back."""
    try:
        o, v = _request(ip, community, oid, 0xa3, timeout, value=(0x02, int(value)))
        return v is not None and v[1] == int(value)
    except Exception:
        return False


def snmp_walk(ip, community, base, timeout=2.0, limit=5000):
    """[(oid, value)] for every object under base."""
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


def _s(v):
    return v.decode(errors='ignore') if isinstance(v, bytes) else ('' if v is None else str(v))


OID_SYSDESCR = '1.3.6.1.2.1.1.1.0'
OID_SYSNAME = '1.3.6.1.2.1.1.5.0'
OID_IFNAME = '1.3.6.1.2.1.31.1.1.1.1'
OID_IFDESCR = '1.3.6.1.2.1.2.2.1.2'
OID_BASEPORT_IFINDEX = '1.3.6.1.2.1.17.1.4.1.2'
OID_DOT1Q_FDB_PORT = '1.3.6.1.2.1.17.7.1.2.2.1.2'
OID_DOT1D_FDB_PORT = '1.3.6.1.2.1.17.4.3.1.2'
OID_PETH_ADMIN = '1.3.6.1.2.1.105.1.1.1.3'       # pethPsePortAdminEnable.<group>.<port>  1=on 2=off
OID_PETH_STATUS = '1.3.6.1.2.1.105.1.1.1.6'      # pethPsePortDetectionStatus
SNMP_COMMUNITIES = ('public', 'Public', 'private')


def snmp_alive(ip, community, timeout=1.5):
    return snmp_get(ip, community, OID_SYSDESCR, timeout) is not None


def snmp_mac_table(ip, community='public', timeout=2.0):
    """{MAC: port_name} from the forwarding table, or None when SNMP does not answer."""
    if not snmp_alive(ip, community, timeout):
        return None
    ifnames = {}
    for o, v in snmp_walk(ip, community, OID_IFNAME, timeout) or snmp_walk(ip, community, OID_IFDESCR, timeout):
        ifnames[int(o.rsplit('.', 1)[1])] = _s(v)
    port2if = {}
    for o, v in snmp_walk(ip, community, OID_BASEPORT_IFINDEX, timeout):
        port2if[int(o.rsplit('.', 1)[1])] = int(v)
    rows = snmp_walk(ip, community, OID_DOT1Q_FDB_PORT, timeout)
    if not rows:
        rows = snmp_walk(ip, community, OID_DOT1D_FDB_PORT, timeout)
    table = {}
    for o, port in rows:
        idx = o.split('.')
        mac = ':'.join('%02X' % int(x) for x in idx[-6:])
        port = int(port)
        if port == 0:
            continue
        ifi = port2if.get(port, port)
        table[mac] = ifnames.get(ifi, f'port {port}')
    return table


def snmp_poe_ports(ip, community, timeout=2.0):
    """{port: admin_enabled(bool)} from POWER-ETHERNET-MIB, {} when the switch has no PoE MIB."""
    out = {}
    for o, v in snmp_walk(ip, community, OID_PETH_ADMIN, timeout):
        try:
            grp, port = o.rsplit('.', 2)[-2:]
            out[int(port)] = (int(v) == 1)
            out.setdefault('_group', int(grp))
        except Exception:
            pass
    return out


def snmp_poe_set(ip, community, port, enable, timeout=2.0, group=None):
    if group is None:
        group = snmp_poe_ports(ip, community, timeout).get('_group', 1)
    return snmp_set_int(ip, community, f'{OID_PETH_ADMIN}.{group}.{int(port)}', 1 if enable else 2, timeout)


# ---------------------------------------------------------------- CLI over telnet
CLI_PROFILES = {
    # name: how to shut the pager up, how to list the MAC table, how to flip PoE on a port.
    # {port} = the port number, {ifname} = the full interface name from the MAC table when known.
    'edgecore': {
        'label': 'Edgecore / Accton (ECS2100 etc.)',
        'pager_off': ['terminal length 0'],
        'mac_table': ['show mac-address-table'],
        'poe_off': ['configure', 'interface ethernet 1/{port}', 'no power inline', 'end'],
        'poe_on': ['configure', 'interface ethernet 1/{port}', 'power inline', 'end'],
    },
    'cisco': {
        'label': 'Cisco IOS / Small Business',
        'pager_off': ['terminal length 0'],
        'mac_table': ['show mac address-table', 'show mac-address-table'],
        'poe_off': ['configure terminal', 'interface {ifname}', 'power inline never', 'end'],
        'poe_on': ['configure terminal', 'interface {ifname}', 'power inline auto', 'end'],
    },
    'aruba': {
        'label': 'Aruba / HPE ProCurve',
        'pager_off': ['no page'],
        'mac_table': ['show mac-address'],
        'poe_off': ['configure', 'no interface {port} power-over-ethernet', 'exit'],
        'poe_on': ['configure', 'interface {port} power-over-ethernet', 'exit'],
    },
    'generic': {
        'label': 'Other (MAC table only)',
        'pager_off': ['terminal length 0'],
        'mac_table': ['show mac-address-table', 'show mac address-table', 'show mac-address'],
        'poe_off': [],
        'poe_on': [],
    },
}

_IAC, _DONT, _DO, _WONT, _WILL, _SB, _SE = 255, 254, 253, 252, 251, 250, 240
_MAC_RE = re.compile(rb'([0-9A-Fa-f]{2}[:\-]){5}[0-9A-Fa-f]{2}|([0-9A-Fa-f]{4}\.){2}[0-9A-Fa-f]{4}')


class TelnetCLI:
    """Minimal telnet client: IAC negotiation (refuses every option), login, enable,
    pager. Good enough for 'show mac-address-table' and a PoE toggle."""

    def __init__(self, ip, user, password, port=23, timeout=5, log=None):
        self.ip, self.user, self.password = ip, user, password
        self.log = log or (lambda m: None)
        self.sock = socket.create_connection((ip, port), timeout=timeout)
        self.sock.settimeout(1.0)
        self.transcript = b''

    def _negotiate(self, data):
        out = bytearray()
        i = 0
        while i < len(data):
            b = data[i]
            if b != _IAC:
                out.append(b)
                i += 1
                continue
            if i + 1 >= len(data):
                break
            cmd = data[i + 1]
            if cmd in (_DO, _DONT, _WILL, _WONT) and i + 2 < len(data):
                opt = data[i + 2]
                reply = _WONT if cmd == _DO else (_DONT if cmd == _WILL else None)
                if reply is not None:
                    try:
                        self.sock.sendall(bytes([_IAC, reply, opt]))
                    except Exception:
                        pass
                i += 3
            elif cmd == _SB:
                j = data.find(bytes([_IAC, _SE]), i)
                i = (j + 2) if j >= 0 else len(data)
            else:
                i += 2
        return bytes(out)

    def read_until(self, patterns, limit):
        buf = b''
        end = time.time() + limit
        while time.time() < end:
            try:
                chunk = self.sock.recv(4096)
            except socket.timeout:
                continue
            if not chunk:
                break
            buf += self._negotiate(chunk)
            self.transcript += buf[-len(chunk):]
            low = buf.lower()
            for pat in patterns:
                if pat in low:
                    return buf, pat
        return buf, None

    def send(self, line):
        self.sock.sendall(line.encode() + b'\r\n')

    def login(self):
        buf, hit = self.read_until((b'username:', b'login:', b'user name:', b'password:'), 8)
        if hit and hit != b'password:':
            self.send(self.user)
            buf, hit = self.read_until((b'password:',), 8)
        if hit == b'password:':
            self.send(self.password)
        buf, hit = self.read_until((b'#', b'>', b'invalid', b'failed', b'denied', b'username:', b'login:'), 10)
        if hit in (b'invalid', b'failed', b'denied', b'username:', b'login:'):
            raise RuntimeError('login refused')
        if hit == b'>':
            self.send('enable')
            buf, hit2 = self.read_until((b'password:', b'#'), 6)
            if hit2 == b'password:':
                self.send(self.password)
                self.read_until((b'#',), 6)
        return True

    def run(self, cmd, timeout=25, want_macs=False):
        """Send one command, return its output text with the pager handled."""
        self.send(cmd)
        out = b''
        end = time.time() + timeout
        quiet = 0
        while time.time() < end:
            try:
                chunk = self.sock.recv(8192)
            except socket.timeout:
                quiet += 1
                if out and quiet >= 3:
                    break
                continue
            except OSError:
                break
            if not chunk:
                break
            quiet = 0
            chunk = self._negotiate(chunk)
            out += chunk
            self.transcript += chunk
            tail = out[-200:].lower()
            if b'--more--' in tail or b'-- more --' in tail or b'<space>' in tail:
                try:
                    self.sock.sendall(b' ')
                except OSError:
                    break
                continue
            stripped = out.rstrip()
            if stripped.endswith((b'#', b'>')) and (not want_macs or _MAC_RE.search(out)) and b'more' not in tail[-40:]:
                if time.time() + 1 < end:
                    time.sleep(0.4)
                    try:
                        more = self.sock.recv(8192)
                        if more:
                            out += self._negotiate(more)
                            continue
                    except (socket.timeout, OSError):
                        pass
                break
        raw = out.decode('utf-8', 'ignore').replace('\x08', '')
        return re.sub(r'\x1b\[[0-9;]*[A-Za-z]', '', raw)

    def close(self):
        try:
            self.send('exit')
        except Exception:
            pass
        try:
            self.sock.close()
        except Exception:
            pass


def parse_mac_table_text(text):
    """'show mac-address-table' output (Edgecore/Cisco/Aruba style) -> {MAC: port_token}.
    Any line with a MAC and a port like 'Eth 1/ 5', 'Eth1/5', '1/5', 'Gi1/0/5', 'Port 5', 'A5'."""
    table = {}
    # xx:xx:xx:xx:xx:xx / xx-xx-… (Edgecore, Netgear), xxxx.xxxx.xxxx (Cisco), xxxxxx-xxxxxx (Aruba/HP)
    mac_re = re.compile(r'([0-9A-Fa-f]{2}[:\-]){5}[0-9A-Fa-f]{2}|([0-9A-Fa-f]{4}\.){2}[0-9A-Fa-f]{4}'
                        r'|\b[0-9A-Fa-f]{6}-[0-9A-Fa-f]{6}\b')
    port_res = [re.compile(r'((?:Eth(?:ernet)?\s*)?\b\d+/(?:\d+/)?\s*\d+)\b', re.I),
                re.compile(r'(Port\s*\d+)\b', re.I),
                re.compile(r'\b((?:Gi|Fa|Te|Tw|Hu)\S*?\d+)\b', re.I),
                re.compile(r'\b([A-Z]?\d{1,2})\b')]          # Aruba 'A5' / bare '9': first token after the MAC
    for line in text.splitlines():
        m = mac_re.search(line)
        if not m:
            continue
        mac = norm_mac(m.group(0))
        before, after = line[:m.start()], line[m.end():]
        port = None
        for i, pr in enumerate(port_res):
            pm = pr.search(before + ' ' + after) if i < 3 else pr.search(after)
            if pm:
                port = pm.group(1).strip()
                break
        if mac and port is not None and port_num(port) is not None:
            table[mac] = port
    return table


def cli_mac_table(ip, user, password, profile='generic', log=None):
    """Log in over telnet and fetch the MAC table. Returns (table, raw_text, profile_used)."""
    prof = CLI_PROFILES.get(profile) or CLI_PROFILES['generic']
    cli = TelnetCLI(ip, user, password, log=log)
    try:
        cli.login()
        for c in prof['pager_off']:
            cli.run(c, timeout=4)
        raw = ''
        for c in prof['mac_table']:
            raw = cli.run(c, timeout=25, want_macs=True)
            table = parse_mac_table_text(raw)
            if table:
                return table, raw, profile
        return {}, raw, profile
    finally:
        cli.close()


def cli_poe_set(ip, user, password, port, enable, profile='edgecore', ifname='', log=None):
    prof = CLI_PROFILES.get(profile) or {}
    cmds = prof.get('poe_on' if enable else 'poe_off') or []
    if not cmds:
        return False, f"no PoE command known for the '{profile}' switch type"
    cli = TelnetCLI(ip, user, password, log=log)
    try:
        cli.login()
        out = ''
        for c in cmds:
            out += cli.run(c.format(port=port, ifname=ifname or f'1/{port}'), timeout=6)
        bad = re.search(r'(invalid|incomplete|unrecognized|error|% )', out, re.I)
        return (not bad), (out[-300:] if bad else 'ok')
    finally:
        cli.close()


# ---------------------------------------------------------------- reading a switch
def read_switch(ip, community='', telnet_user='', telnet_password='', profile='generic',
                communities=SNMP_COMMUNITIES, log=None):
    """Who is on which port. Telnet CLI first when credentials are given (it gives
    the port names exactly as the switch prints them), SNMP as the fallback, every
    usual community tried. Returns a dict:
        {'ok': bool, 'via': 'telnet'|'snmp'|'', 'table': {MAC: port}, 'sysname': str,
         'sysdescr': str, 'community': str, 'raw': str, 'errors': [str]}"""
    log = log or (lambda m: None)
    res = {'ok': False, 'via': '', 'table': {}, 'sysname': '', 'sysdescr': '', 'community': '',
           'raw': '', 'errors': [], 'poe': {}}
    if telnet_user:
        try:
            log(f"  telnet {ip} as {telnet_user} ...")
            table, raw, used = cli_mac_table(ip, telnet_user, telnet_password, profile, log=log)
            res['raw'] = raw
            if table:
                res.update(ok=True, via='telnet', table=table)
                log(f"  'show mac-address-table': {len(table)} entries")
            else:
                res['errors'].append('telnet logged in but no MAC table was parsed')
        except Exception as e:
            res['errors'].append(f'telnet: {e}')
            log(f"  telnet failed: {e}")
    tried = [c for c in ([community] if community else []) + list(communities) if c]
    seen = set()
    for comm in tried:
        if comm in seen:
            continue
        seen.add(comm)
        if not snmp_alive(ip, comm, 1.5):
            continue
        res['community'] = comm
        res['sysname'] = _s(snmp_get(ip, comm, OID_SYSNAME))
        res['sysdescr'] = _s(snmp_get(ip, comm, OID_SYSDESCR))
        try:
            res['poe'] = snmp_poe_ports(ip, comm)
        except Exception:
            pass
        if not res['ok']:
            t = snmp_mac_table(ip, comm)
            if t:
                res.update(ok=True, via='snmp', table=t)
                log(f"  SNMP (community {comm!r}): {len(t)} MAC entries")
            else:
                res['errors'].append(f'SNMP {comm!r} answers but has no forwarding table')
        break
    if not res['community']:
        res['errors'].append('no SNMP answer with ' + '/'.join(tried))
    return res


def find_switch(subnets, hints=(), community='', telnet_user='', telnet_password='',
                lldp_seconds=0, cancel=None, log=None):
    """Candidates for the switch's management IP, best first. Each: {'ip', 'via', 'sysname', 'mac'}.
    subnets: iterable of ipaddress networks the cameras live in (the switch is usually in one).
    Order: LLDP management address (if lldp_seconds), explicit hints, the top hosts of each
    subnet (top, top-1, top-2) and its first host, then a ping sweep of each subnet keeping
    hosts whose MAC is NOT a camera vendor and that answer SNMP or telnet."""
    log = log or (lambda m: None)
    found = []
    seen = set()

    def consider(ip, via):
        if not ip or ip in seen or (cancel and cancel.is_set()):
            return
        seen.add(ip)
        name = ''
        ok = False
        for comm in [c for c in ([community] if community else []) + list(SNMP_COMMUNITIES) if c]:
            if snmp_alive(ip, comm, 1.0):
                name = _s(snmp_get(ip, comm, OID_SYSNAME))
                ok = True
                break
        if not ok and tcp_open(ip, 23, 0.6):
            ok = True
            via += ' (telnet port open)'
        if ok:
            found.append({'ip': ip, 'via': via, 'sysname': name})
            log(f"  switch candidate {ip} ({via}) {name!r}")

    if lldp_seconds:
        for i in lldp_listen(lldp_seconds, log=log):
            if i.get('mgmt_ip'):
                consider(i['mgmt_ip'], f"LLDP from {i.get('switch_name') or i.get('chassis_id')}")
    for h in hints:
        consider(h, 'given')
    nets = []
    for sn in subnets:
        try:
            nets.append(sn if isinstance(sn, (ipaddress.IPv4Network, ipaddress.IPv6Network)) else ipaddress.ip_network(sn, strict=False))
        except Exception:
            continue
    for net in nets:
        top = int(net.broadcast_address)
        for off in (1, 2, 3):
            consider(str(ipaddress.ip_address(top - off)), f'top host of {net}')
        consider(str(net.network_address + 1), f'first host of {net}')
    if found:
        return found
    camera_macs = set()
    for net in nets:
        if cancel and cancel.is_set():
            break
        log(f"  sweeping {net} for the switch ...")
        alive = ping_sweep(net, cancel=cancel)
        arp = {ip: mac for mac, ip in arp_table().items()}
        for ip in alive:
            mac = arp.get(ip, '')
            if mac and is_camera_mac(mac):
                camera_macs.add(mac)
                continue
            consider(ip, f'sweep of {net}' + (f', {vendor_of(mac)}' if vendor_of(mac) else ''))
    return found


# ---------------------------------------------------------------- PoE cycle
def poe_cycle(ip, port, community='private', telnet_user='', telnet_password='', profile='edgecore',
              off_seconds=6, ifname='', log=None):
    """Power a PoE port off, wait, power it on. SNMP SET first (standard MIB, any vendor
    that has it), then the vendor CLI. Returns (ok, how)."""
    log = log or (lambda m: None)
    for comm in [c for c in (community, 'private') if c]:
        try:
            ports = snmp_poe_ports(ip, comm)
        except Exception:
            ports = {}
        if int(port) in ports:
            grp = ports.get('_group', 1)
            if snmp_poe_set(ip, comm, port, False, group=grp):
                log(f"  PoE port {port} OFF (SNMP)")
                time.sleep(off_seconds)
                on = snmp_poe_set(ip, comm, port, True, group=grp)
                log(f"  PoE port {port} {'ON' if on else 'did not come back ON - check the switch!'} (SNMP)")
                return on, 'snmp'
            log(f"  SNMP SET refused with community {comm!r}")
    if telnet_user:
        ok, msg = cli_poe_set(ip, telnet_user, telnet_password, port, False, profile, ifname, log)
        if not ok:
            return False, f'cli: {msg}'
        log(f"  PoE port {port} OFF (CLI)")
        time.sleep(off_seconds)
        ok, msg = cli_poe_set(ip, telnet_user, telnet_password, port, True, profile, ifname, log)
        log(f"  PoE port {port} {'ON' if ok else 'did not come back ON - check the switch!'} (CLI)")
        return ok, 'cli' if ok else f'cli: {msg}'
    return False, 'no SNMP write community worked and no CLI login was given'


# ---------------------------------------------------------------- LLDP (pktmon)
def lldp_listen(seconds=35, log=None):
    """Capture LLDP frames with pktmon (Windows 10/11, needs admin).
    [{switch_name, port_id, port_desc, mgmt_ip, system_desc, chassis_id, switch_mac}]"""
    log = log or (lambda m: None)
    kw = dict(capture_output=True, timeout=15, creationflags=_NO_WINDOW)
    etl = tempfile.mktemp(suffix='.etl')
    pcap = tempfile.mktemp(suffix='.pcapng')
    results = []
    try:
        subprocess.run(['pktmon', 'stop'], **kw)
        subprocess.run(['pktmon', 'filter', 'remove'], **kw)
        r = subprocess.run(['pktmon', 'filter', 'add', 'LLDP', '-d', '0x88CC'], **kw)
        if r.returncode != 0:
            r = subprocess.run(['pktmon', 'filter', 'add', 'LLDP', '-t', '0x88CC'], **kw)
        if r.returncode != 0:
            log(f"  pktmon filter failed: {r.stderr.decode(errors='replace')[:120]}")
            return results
        r = subprocess.run(['pktmon', 'start', '-c', '--pkt-size', '512', '-f', etl], **kw)
        if r.returncode != 0:
            log(f"  pktmon start failed: {r.stderr.decode(errors='replace')[:120]}")
            return results
        log(f"  listening for LLDP for {seconds}s (switches announce every 30 s) ...")
        time.sleep(seconds)
        subprocess.run(['pktmon', 'stop'], **kw)
        time.sleep(0.3)
        subprocess.run(['pktmon', 'etl2pcap', etl, '-o', pcap], **kw)
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
            info = parse_lldp(frame)
            if info and (info['switch_name'], info['port_id']) not in seen:
                seen.add((info['switch_name'], info['port_id']))
                results.append(info)
    except Exception as e:
        log(f"  LLDP capture error: {e}")
    finally:
        try:
            subprocess.run(['pktmon', 'filter', 'remove'], **kw)
        except Exception:
            pass
        for f in (etl, pcap):
            try:
                os.remove(f)
            except Exception:
                pass
    return results


def parse_lldp(frame):
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


# ---------------------------------------------------------------- the port decides the camera
def rows_for_switch(cameras, switch_key):
    """Camera rows whose 'switch' field refers to switch_key, keyed by port number.
    switch_key '' = every row that has a port (single-switch job)."""
    out = {}
    for c in cameras:
        pn = port_num(c.get('port'))
        if pn is None:
            continue
        if switch_key and not same_switch(c.get('switch'), switch_key):
            continue
        out.setdefault(pn, []).append(c)
    return out


def assign_by_port(cameras, table, switch_key='', laptop_macs=(), live=None, camera_port_max=None,
                   extra_ouis=()):
    """Build the port map and the assignments.

    cameras : camera-list rows (dicts: name, ip, mac, model, switch, port, processed ...)
    table   : {MAC: port_token} from read_switch
    live    : optional {MAC: {'ip':..., 'model':...}} from a probe of the wire
    Returns {'ports': [row per port, sorted], 'assign': [(cam, mac, flags)], 'notes': [str],
             'laptop_ports': set, 'laptop_on_camera_port': bool|None, 'unlisted': [...]}
    Port row: {port, mac, vendor, ip, model, cam (row or None), status, flags, laptop}"""
    live = live or {}
    laptop_macs = {norm_mac(m) for m in laptop_macs if norm_mac(m)}
    by_port = {}
    for mac, pt in table.items():
        pn = port_num(pt)
        if pn is None:
            continue
        by_port.setdefault(pn, []).append((norm_mac(mac) or mac, pt))
    sheet = rows_for_switch(cameras, switch_key)
    cam_by_mac = {}
    for c in cameras:
        m = norm_mac(c.get('mac', ''))
        if m:
            cam_by_mac.setdefault(m, c)
    laptop_ports = {pn for pn, macs in by_port.items() if any(m in laptop_macs for m, _ in macs)}
    if camera_port_max is None:
        camera_port_max = max([pn for pn in sheet] + [0]) or max([pn for pn in by_port] + [0])
    ports, assign, notes, unlisted = [], [], [], []
    for pn in sorted(set(by_port) | set(sheet)):
        here = [(m, pt) for m, pt in by_port.get(pn, []) if m not in laptop_macs]
        rows = sheet.get(pn, [])
        sc = rows[0] if rows else None
        if len(rows) > 1:
            notes.append(f"port {pn}: the list has {len(rows)} cameras on this port ({', '.join(r.get('name', '?') for r in rows)}) - fix the list")
        cams_here = [(m, pt) for m, pt in here if is_camera_mac(m, extra_ouis)]
        others = [(m, pt) for m, pt in here if (m, pt) not in cams_here]
        row = {'port': pn, 'mac': '', 'vendor': '', 'ip': '', 'model': '', 'cam': sc, 'status': '',
               'flags': [], 'laptop': pn in laptop_ports, 'port_token': (here or by_port.get(pn) or [('', '')])[0][1]}
        if pn in laptop_ports and not here:
            row['status'] = 'THIS PC'
            if sc:
                row['flags'].append(f"the list expects {sc.get('name', '?')} here - move this PC and read again")
            ports.append(row)
            continue
        if not here:
            if sc:
                row['status'] = 'DARK' if not sc.get('processed') else 'DARK (was done)'
                row['flags'].append('no device on this port')
            else:
                continue  # nothing on the port and nothing expected: not interesting
            ports.append(row)
            continue
        if len(cams_here) > 1:
            row['mac'] = ', '.join(m for m, _ in cams_here)
            row['status'] = 'CABLING?'
            row['flags'].append(f"{len(cams_here)} cameras seen on one port - a switch behind it, or stale table")
            ports.append(row)
            continue
        if not cams_here:
            m, pt = others[0]
            row.update(mac=m, vendor=vendor_of(m) or 'unknown', ip=arp_ip(live, m))
            if sc:
                row['status'] = 'NOT A CAMERA'
                row['flags'].append(f"the list expects {sc.get('name', '?')} here but this is {vendor_of(m) or 'an unknown device'}")
            elif pn <= camera_port_max:
                row['status'] = 'OTHER DEVICE'
            else:
                row['status'] = 'UPLINK / OTHER'
            ports.append(row)
            continue
        m, pt = cams_here[0]
        info = live.get(m, {})
        row.update(mac=m, vendor=vendor_of(m), ip=info.get('ip', ''), model=info.get('model', ''))
        if sc:
            sheet_mac = norm_mac(sc.get('mac', ''))
            if sheet_mac == m:
                row['status'] = 'MATCH'
            elif sheet_mac:
                row['status'] = 'PORT WINS'
                row['flags'].append(f"list says {sheet_mac}; the port says {m}")
                other = cam_by_mac.get(m)
                if other is not None and other is not sc:
                    row['flags'].append(f"list calls this MAC {other.get('name', '?')} (port {other.get('port', '?')})")
            else:
                row['status'] = 'NEW MAC'
            if sc.get('processed'):
                row['status'] += ' (done)'
            if info.get('model') and sc.get('model') and not _model_match(info['model'], sc['model']):
                row['flags'].append(f"MODEL {info['model']} vs list {sc['model']}")
                row['status'] = 'MODEL?'
            assign.append((sc, m, list(row['flags'])))
        else:
            row['status'] = 'NOT ON LIST'
            other = cam_by_mac.get(m)
            if other is not None:
                row['cam'] = other
                row['flags'].append(f"the list has this MAC as {other.get('name', '?')} on switch {other.get('switch', '?')} port {other.get('port', '?')}")
            unlisted.append(row)
        ports.append(row)
    if laptop_macs:
        if laptop_ports:
            laptop_on_camera_port = all(p <= camera_port_max for p in laptop_ports) if camera_port_max else True
        else:
            laptop_on_camera_port = None
            notes.append("this PC's MAC is not in the switch table: you may be on an uplink/trunk port where it "
                         "sees nothing. Plug into a camera port (1-24 style) and read again.")
    else:
        laptop_on_camera_port = None
    return {'ports': ports, 'assign': assign, 'notes': notes, 'laptop_ports': laptop_ports,
            'laptop_on_camera_port': laptop_on_camera_port, 'unlisted': unlisted,
            'camera_port_max': camera_port_max}


def arp_ip(live, mac):
    return (live.get(mac) or {}).get('ip', '')


def _model_match(a, b):
    na = str(a).upper().replace('AXIS-', '').replace('AXIS ', '').strip()
    nb = str(b).upper().replace('AXIS-', '').replace('AXIS ', '').strip()
    return na in nb or nb in na


def next_pass_advice(result, cameras, switch_key=''):
    """Two passes per switch: this PC sits on one camera port, so that camera is done
    on the second pass from a port whose camera IS done. Never send the PC to a port
    whose camera is not finished. Returns a plain-English line or ''."""
    lp = result.get('laptop_ports') or set()
    if not lp:
        return ''
    sheet = rows_for_switch(cameras, switch_key)
    missing = []
    for p in sorted(lp):
        rows = sheet.get(p, [])
        if rows and not rows[0].get('processed'):
            missing.append((p, rows[0].get('name', '?')))
    if not missing:
        return f"This PC is on port {', '.join(map(str, sorted(lp)))}; nothing is expected there, so this pass can cover the whole switch."
    done_ports = sorted(p for p, rows in sheet.items() if rows and rows[0].get('processed') and p not in lp)
    p, name = missing[0]
    if done_ports:
        return (f"This PC is on port {p}, where the list expects {name}. When everything else is done: "
                f"move this PC to port {done_ports[0]} (its camera is finished), plug {name} into port {p}, "
                f"read the switch again and run Switch Loading again.")
    return (f"This PC is on port {p}, where the list expects {name}. Finish the other cameras first, then "
            f"move this PC to a port whose camera is DONE, plug {name} into port {p}, and run again.")


# ---------------------------------------------------------------- per-model reset guidance
RESET_GUIDE = [
    # (model prefix match, text). First match wins; keep the generic one last.
    ('AXIS M10', "Axis M10xx: unplug, hold the control button, plug back in, keep holding ~15-30 s until the status LED flashes amber, release. Back at 192.168.0.90 / DHCP in ~1 min."),
    ('AXIS M30', "Axis M30xx: control button is under the dome ring. Hold it while powering up until the LED flashes amber (15-30 s), release."),
    ('AXIS M32', "Axis M32xx: control button is under the dome ring. Hold it while powering up until the LED flashes amber (15-30 s), release."),
    ('AXIS P32', "Axis P32xx: loosen the dome, control button is on the main board next to the SD slot. Hold while powering up until the LED flashes amber (15-30 s)."),
    ('AXIS P33', "Axis P33xx: control button is next to the SD slot. Hold while powering up until the LED flashes amber (15-30 s)."),
    ('AXIS Q', "Axis Q-series: control button on the body. Hold while powering up until the LED flashes amber (15-30 s), release."),
    ('AXIS', "Axis: hold the control button while powering the camera up, keep holding until the status LED flashes amber (15-30 s), release. A red flash = held too short or too long; try again."),
    ('BOSCH', "Bosch: Tools -> Factory Default with its password, or hold the reset button 10 s with power on; it comes back on DHCP / 192.168.0.1 in ~30 s."),
    ('HANWHA', "Hanwha/Wisenet: hold the reset button ~5 s with power on until the LED blinks, release; back on 192.168.1.100 / DHCP."),
    ('', "Hold the camera's reset/control button while powering it up until the status LED blinks, then release."),
]


def reset_guide(model='', brand=''):
    key = f"{brand} {model}".upper()
    for pref, text in RESET_GUIDE:
        if not pref or pref in key:
            return text
    return RESET_GUIDE[-1][1]
