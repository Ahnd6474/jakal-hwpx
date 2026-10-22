"""Loopback health probe using the same public authority as the HTTPS proxy."""
from __future__ import annotations

import http.client
import json
import os
from urllib.parse import urlsplit

from .runtime import server_options


def probe() -> bool:
    # Never fetch a configured public URL or follow redirects from a probe.
    headers = {}
    origin = os.environ.get('JAKAL_PUBLIC_ORIGIN')
    if origin:
        public = urlsplit(origin)
        headers = {'Host': public.netloc, 'X-Forwarded-Proto': public.scheme}
    connection = http.client.HTTPConnection('127.0.0.1', server_options()['port'], timeout=4)
    try:
        connection.request('GET', '/api/health', headers=headers)
        response = connection.getresponse()
        body = response.read(4097)
        if response.status != 200 or len(body) > 4096:
            return False
        payload = json.loads(body)
        return isinstance(payload, dict) and payload.get('status') == 'ok'
    except (OSError, ValueError, http.client.HTTPException):
        return False
    finally:
        connection.close()


if __name__ == '__main__':
    raise SystemExit(0 if probe() else 1)
