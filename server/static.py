"""Serve precompressed public build assets without runtime compression work."""
import re
from pathlib import PurePosixPath

from starlette.exceptions import HTTPException
from starlette.staticfiles import StaticFiles


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
        path = path.replace('\\', '/')
        headers = dict(scope.get('headers', []))
        content_type = {'.js': 'text/javascript', '.css': 'text/css', '.wasm': 'application/wasm', '.svg': 'image/svg+xml'}.get(PurePosixPath(path).suffix)
        response = None
        if content_type and b'range' not in headers and accepts_gzip(headers.get(b'accept-encoding', b'').decode('latin-1')):
            try:
                response = await super().get_response(path + '.gz', scope)
                response.headers['Content-Encoding'] = 'gzip'
                response.headers['Content-Type'] = content_type
            except HTTPException as exc:
                if exc.status_code != 404:
                    raise
        if response is None:
            response = await super().get_response(path, scope)
        if content_type:
            response.headers['Vary'] = 'Accept-Encoding'
        if re.fullmatch(r'(?:studio/)?assets/[^/]+-[A-Za-z0-9_-]{8,}\.(?:js|css|wasm)', path):
            response.headers['Cache-Control'] = 'public, max-age=31536000, immutable'
        elif path.startswith(('fonts/', 'studio/fonts/')):
            response.headers['Cache-Control'] = 'public, max-age=86400'
        else:
            response.headers['Cache-Control'] = 'no-cache'
        return response
