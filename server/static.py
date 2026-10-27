"""Serve precompressed public build assets without runtime compression work."""
import re
import hashlib
from pathlib import Path
from pathlib import PurePosixPath

from starlette.exceptions import HTTPException
from starlette.staticfiles import StaticFiles

# These generated introductory demos are the only deliberately public documents.
PUBLIC_DOCUMENTS = {
    'sample.hwpx': '219f823825b5d896fc38ef11d9339a8fb14a3a0142c6d11a7acb5345f71039cd',
    'sample.hwp': 'e483cad81bbbf528e1afab093d467e0840fe4234695763eeeae23ebbaa70349e',
}
RESERVED_NAMES = {'con', 'prn', 'aux', 'nul', 'conin$', 'conout$', 'clock$',
    *(f'{prefix}{number}' for prefix in ('com', 'lpt') for number in '123456789¹²³')}


def accepts_gzip(header: str) -> bool:
    encodings = {}
    for entry in header.lower().split(','):
        parts = entry.strip().split(';')
        quality = 1.0
        for parameter in parts[1:]:
            if parameter.strip().startswith('q='):
                try:
                    quality = float(parameter.strip()[2:])
                except ValueError:
                    quality = 0.0
        encodings[parts[0].strip()] = quality if 0 <= quality <= 1 else 0.0
    return encodings.get('gzip', encodings.get('*', 0)) > 0


class WebStaticFiles(StaticFiles):
    async def get_response(self, path, scope):
        raw_path = scope.get('path', path)
        if any(ord(char) < 32 or ord(char) == 127 or char in '\\:*?"<>|' for char in raw_path):
            raise HTTPException(404)
        path = path.replace('\\', '/')
        parts = PurePosixPath(path.lower()).parts
        if any(part.startswith(' ') or part.endswith((' ', '.')) or re.search(r'~\d+(?:\.|$)', part) or part.split('.')[0] in RESERVED_NAMES for part in parts):
            raise HTTPException(404)
        filename = PurePosixPath(path.lower())
        while filename.suffix in {'.gz', '.br'}:
            filename = PurePosixPath(filename.stem)
        suffix = filename.suffix
        private_record = re.fullmatch(r'(?:snapshot-\d+|(?:job|result)-[a-f0-9]{32})\.json', filename.name)
        if any(part.startswith('.') or part in {'uploads', 'private-uploads', 'session-backups'} for part in parts) or suffix in {'.sqlite', '.sqlite3', '.db', '.key', '.pem'} or any(marker in path.lower() for marker in ('.sqlite3-', '.db-wal', '.db-shm')) or private_record:
            raise HTTPException(404)
        if suffix in {'.hwp', '.hwpx', '.pdf'}:
            expected = PUBLIC_DOCUMENTS.get(path)
            if not expected:
                raise HTTPException(404)
            full, stat = self.lookup_path(path)
            if not stat or hashlib.sha256(Path(full).read_bytes()).hexdigest() != expected:
                raise HTTPException(404)
        headers = dict(scope.get('headers', []))
        content_type = {'.js': 'text/javascript', '.css': 'text/css', '.wasm': 'application/wasm', '.svg': 'image/svg+xml'}.get(PurePosixPath(path).suffix)
        response = None
        if content_type and b'range' not in headers and accepts_gzip(headers.get(b'accept-encoding', b'').decode('latin-1')):
            try:
                candidate = await super().get_response(path + '.gz', scope)
                if candidate.status_code in {200, 304}:
                    response = candidate
                    response.headers['Content-Encoding'] = 'gzip'
                    response.headers['Content-Type'] = content_type
            except HTTPException as exc:
                if exc.status_code != 404:
                    raise
        if response is None:
            response = await super().get_response(path, scope)
        if content_type:
            response.headers['Vary'] = 'Accept-Encoding'
        if response.status_code >= 400:
            response.headers['Cache-Control'] = 'no-cache'
        elif re.fullmatch(r'(?:studio/)?assets/[^/]+-[A-Za-z0-9_-]{8,}\.(?:js|css|wasm)', path):
            response.headers['Cache-Control'] = 'public, max-age=31536000, immutable'
        elif path.startswith(('fonts/', 'studio/fonts/')):
            response.headers['Cache-Control'] = 'public, max-age=86400'
        else:
            response.headers['Cache-Control'] = 'no-cache'
        return response
