"""Strict anonymous sessions and explicit, fail-closed deployment boundaries.

The cookie is a bearer credential. Only its hash is stored; workspace IDs are
public browser-storage namespaces, never credentials or authorization inputs.
"""
from __future__ import annotations

import hashlib
import ipaddress
import os
import re
import secrets
import time
import math
import threading
from collections import OrderedDict, deque
from pathlib import Path
from urllib.parse import urlsplit

from .storage import Store

COOKIE = 'jakal_session'
SECURE_COOKIE = '__Host-jakal_session'


class SessionRateLimited(Exception):
    def __init__(self, retry_after: int):
        self.retry_after = retry_after


class SessionIssuance:
    def __init__(self, *, per_peer=60, total=600, max_peers=2048, clock=time.monotonic):
        self.per_peer, self.total, self.max_peers = per_peer, total, max_peers
        self.clock = clock
        self.peers = OrderedDict()
        self.recent = deque()
        self.lock = threading.Lock()

    def consume(self, peer: str):
        now = self.clock()
        key = hashlib.sha256(peer.encode()).digest()
        with self.lock:
            while self.recent and self.recent[0] <= now - 60:
                self.recent.popleft()
            entries = self.peers.get(key, deque())
            while entries and entries[0] <= now - 60:
                entries.popleft()
            if len(entries) >= self.per_peer or len(self.recent) >= self.total:
                oldest = entries[0] if len(entries) >= self.per_peer else self.recent[0]
                raise SessionRateLimited(max(1, math.ceil(oldest + 60 - now)))
            self.recent.append(now); entries.append(now)
            self.peers[key] = entries
            self.peers.move_to_end(key)
            while len(self.peers) > self.max_peers:
                self.peers.popitem(last=False)


class DeploymentPolicy:
    def __init__(self, repo: Path, data_dir: Path):
        self.production = os.environ.get('JAKAL_ENV', 'development') == 'production'
        public = os.environ.get('JAKAL_PUBLIC_ORIGIN', '').rstrip('/')
        self.origin = None
        if public:
            parsed = urlsplit(public)
            parsed.port  # Validate the configured port as well as the host.
            if parsed.scheme != 'https' or not parsed.hostname or parsed.username or parsed.password or parsed.path or parsed.query or parsed.fragment:
                raise ValueError('JAKAL_PUBLIC_ORIGIN must be a single HTTPS origin without a path.')
            self.origin = f'https://{parsed.netloc.lower()}'
        if self.production and not self.origin:
            raise ValueError('Production requires JAKAL_PUBLIC_ORIGIN=https://your-dedicated-host.')
        if self.production and '*' in os.environ.get('JAKAL_TRUSTED_PROXIES', ''):
            raise ValueError('Production requires explicit trusted proxy IPs, never a wildcard.')
        data = data_dir.resolve()
        # Never put private files inside a directory served as public assets.
        for public_root in (repo / 'web' / 'public', repo / 'web' / 'dist', repo / 'website'):
            public_root = public_root.resolve()
            if data.is_relative_to(public_root) or public_root.is_relative_to(data):
                raise ValueError('Private document storage must be separate from public web directories.')

    def allowed(self, request):
        try:
            authority = request.headers.get('host', '')
            if any(char.isspace() or ord(char) < 32 or ord(char) == 127 or char in '\\,' for char in authority):
                return False
            parsed = urlsplit(f"{request.url.scheme}://{authority}")
            parsed.port
            if parsed.username or parsed.password or parsed.path or parsed.query or parsed.fragment or not parsed.hostname:
                return False
            origin = f'{parsed.scheme}://{parsed.netloc.lower()}'
            if self.origin:
                return origin == self.origin
            peer = request.client.host if request.client else ''
            if peer != 'testclient' and not ipaddress.ip_address(peer).is_loopback:
                return False
            host = parsed.hostname.lower()
            return host in {'localhost', 'testserver'} or ipaddress.ip_address(host).is_loopback
        except ValueError:
            return False


class Sessions:
    def __init__(self, store: Store, *, allow_legacy: bool):
        self.store = store
        self.issuance = SessionIssuance()
        # One-time local upgrade for already uploaded documents. A supplied
        # random 64-character value does not establish an accepted session.
        with store.connect() as db:
            db.execute('''CREATE TABLE IF NOT EXISTS sessions (
                owner TEXT PRIMARY KEY, workspace TEXT NOT NULL, created REAL NOT NULL,
                expires REAL NOT NULL, touched REAL NOT NULL, revoked INTEGER NOT NULL DEFAULT 0)''')
            self.legacy = {row['owner'] for row in db.execute(
                'SELECT DISTINCT owner FROM documents WHERE created>?', (time.time() - store.ttl,))} if allow_legacy else set()

    def resolve(self, token: str, fence: str | None = None, peer: str = ''):
        now = time.time()
        owner = hashlib.sha256(token.encode()).hexdigest() if re.fullmatch('[a-f0-9]{64}', token) else ''
        with self.store.connect() as db:
            row = db.execute('SELECT * FROM sessions WHERE owner=?', (owner,)).fetchone()
            if row and not row['revoked'] and row['expires'] > now and row['created'] + 7 * 86400 > now:
                if fence and fence != row['workspace']:
                    return None
                if row['touched'] < now - 60:
                    db.execute('UPDATE sessions SET touched=?,expires=? WHERE owner=? AND revoked=0',
                        (now, min(now + self.store.ttl, row['created'] + 7 * 86400), owner))
                return token, owner, row['workspace']
            if fence:
                return None
            counted = row is not None or owner not in self.legacy
            if counted:
                self.issuance.consume(peer)
            # Serialize only issuance/migration, not normal session reads. A
            # parallel first request must reuse a just-migrated local session.
            db.execute('BEGIN IMMEDIATE')
            row = db.execute('SELECT * FROM sessions WHERE owner=?', (owner,)).fetchone()
            if row and not row['revoked'] and row['expires'] > now and row['created'] + 7 * 86400 > now:
                return token, owner, row['workspace']
            legacy_owner = owner
            if row or owner not in self.legacy:
                if not counted:
                    self.issuance.consume(peer)
                token = secrets.token_hex(32)
                owner = hashlib.sha256(token.encode()).hexdigest()
            workspace = secrets.token_hex(16)
            db.execute('INSERT OR IGNORE INTO sessions(owner,workspace,created,expires,touched) VALUES(?,?,?,?,?)',
                (owner, workspace, now, now + self.store.ttl, now))
            issued = db.execute('SELECT workspace FROM sessions WHERE owner=?', (owner,)).fetchone()
            self.legacy.discard(legacy_owner)
            return token, owner, issued['workspace']

    def revoke(self, owner: str):
        now = time.time()
        with self.store.connect() as db:
            db.execute('BEGIN IMMEDIATE')
            db.execute('UPDATE sessions SET revoked=1,expires=? WHERE owner=?', (now, owner))
            # Deny access immediately, including copied cookies and in-flight
            # jobs. The worker safely purges files once running work releases them.
            db.execute('UPDATE documents SET created=0 WHERE owner=?', (owner,))


def protect_storage(root: Path, *, production: bool = False):
    """POSIX process-private defaults; Windows requires a service-account ACL."""
    if os.name != 'nt':
        os.umask(0o077)
        root.chmod(0o700)
        for path in root.glob('app.sqlite3*'):
            path.chmod(0o600)
    elif production:
        import subprocess
        script = Path(__file__).resolve().parents[1] / 'scripts' / 'secure_private_storage.ps1'
        # A parent PowerShell 7 can export its module path; Windows PowerShell
        # 5 must resolve its own Security module rather than loading that binary.
        env = {key: value for key, value in os.environ.items() if key.upper() != 'PSMODULEPATH'}
        subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive',
            '-File', str(script), '-PrivateRoot', str(root)], check=True, timeout=120,
            env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW)
