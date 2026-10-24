import asyncio
import socket
import threading
import time

import pytest
import uvicorn

from scripts.check_mcp import check_site
from server.app import create_app
from server.mcp_bridge import connection_url


def test_official_sdk_http_and_stdio_roundtrip(tmp_path):
    listener = socket.socket()
    listener.bind(('127.0.0.1', 0))
    port = listener.getsockname()[1]
    server = uvicorn.Server(uvicorn.Config(create_app(tmp_path / 'private'), log_level='critical', access_log=False))
    thread = threading.Thread(target=lambda: server.run(sockets=[listener]), daemon=True)
    thread.start()
    try:
        for _ in range(200):
            if server.started:
                break
            time.sleep(.05)
        assert server.started
        asyncio.run(check_site(f'http://127.0.0.1:{port}'))
    finally:
        server.should_exit = True
        thread.join(timeout=15)
        listener.close()
        assert not thread.is_alive()


@pytest.mark.parametrize('url', ['http://public.example', 'https://u:p@public.example', 'https://public.example/path', 'https://public.example?token=secret', 'file:///tmp/a'])
def test_stdio_bridge_rejects_unsafe_site_origins(url):
    with pytest.raises(ValueError):
        connection_url(url)
