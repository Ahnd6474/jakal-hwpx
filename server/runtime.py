"""Shared process configuration for local runs and container deployments."""
from __future__ import annotations

import os
import shutil
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]


def server_options() -> dict:
    # PaaS platforms commonly inject PORT. The explicit application setting wins.
    raw_port = os.environ.get('JAKAL_PORT', os.environ.get('PORT', '8000'))
    try:
        port = int(raw_port)
        if not 1 <= port <= 65535:
            raise ValueError
    except ValueError:
        raise ValueError('JAKAL_PORT / PORT must be an integer between 1 and 65535.') from None
    return {
        'host': os.environ.get('JAKAL_HOST', '127.0.0.1'),
        'port': port,
        # The durable SQLite queue supports exactly one API process per volume.
        'workers': 1,
        'forwarded_allow_ips': os.environ.get('JAKAL_TRUSTED_PROXIES', '127.0.0.1'),
        'access_log': False,
        'server_header': False,
    }


def check_runtime(repo: Path = REPO) -> None:
    environment = os.environ.get('JAKAL_ENV', 'development')
    if environment not in {'development', 'production'}:
        raise ValueError('JAKAL_ENV must be development or production.')
    if environment != 'production':
        return
    required = [
        'web/dist/index.html', 'web/dist/studio/index.html',
        'web/node_modules/@rhwp/core/rhwp.js',
        'web/node_modules/@rhwp/core/rhwp_bg.wasm',
    ]
    if any(not (repo / name).is_file() for name in required):
        raise RuntimeError('Production assets are incomplete. Build the web app and install the HWP converter first.')
    if not shutil.which(os.environ.get('JAKAL_NODE') or 'node'):
        raise RuntimeError('Production HWP conversion requires Node.js on PATH or JAKAL_NODE.')
