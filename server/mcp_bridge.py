"""Run with python -m server.mcp_bridge; stdout is reserved for MCP messages."""
from __future__ import annotations

import asyncio
import ipaddress
import os
import sys
from urllib.parse import urlsplit

import httpx2 as httpx
from mcp.server.stdio import stdio_server

from .errors import WorkspaceError
from .mcp_adapter import create_mcp_server


def connection_url(value):
    parsed = urlsplit(value)
    try:
        local = parsed.hostname == 'localhost' or ipaddress.ip_address(parsed.hostname or '').is_loopback
    except ValueError:
        local = False
    if (parsed.scheme not in {'https', 'http'} or not parsed.hostname or parsed.username or parsed.password
        or parsed.path not in {'', '/'} or parsed.query or parsed.fragment or (parsed.scheme == 'http' and not local)):
        raise ValueError('Use an HTTPS site origin, or HTTP on loopback only.')
    parsed.port
    return value.rstrip('/')


async def serve():
    origin = connection_url(os.environ.get('JAKAL_WORKSPACE_URL', ''))
    token = os.environ.get('JAKAL_CONNECTION_TOKEN', '')
    if len(token) != 68 or not token.startswith('jkl_'):
        raise ValueError('Create a connection in the website and provide its token.')
    async with httpx.AsyncClient(base_url=origin, headers={'Authorization': 'Bearer ' + token},
        timeout=30, follow_redirects=False, trust_env=False) as client:
        async def permission(ctx):
            response = await client.get('/api/connected/info')
            if response.status_code != 200:
                raise WorkspaceError('connection_unavailable', '웹에서 연결 만료·해제 여부를 확인해 주세요.', 401)
            return response.json()['permission']

        async def execute(ctx, name, arguments):
            try:
                response = await client.post('/api/connected/tools/' + name, json=arguments)
            except httpx.HTTPError:
                raise WorkspaceError('connection_failed', '편집기에 연결할 수 없습니다. 실행 상태와 사이트 주소를 확인해 주세요.', 503) from None
            if response.status_code != 200:
                try:
                    error = response.json()
                except ValueError:
                    error = {}
                raise WorkspaceError(error.get('code', 'request_failed'), error.get('detail', '문서 작업을 처리하지 못했습니다.'), response.status_code)
            return response.json()

        server = create_mcp_server(execute, permission)
        async with stdio_server() as (read, write):
            await server.run(read, write, server.create_initialization_options())


def main():
    try:
        asyncio.run(serve())
    except (ValueError, OSError):
        print('Jakal 연결을 시작하지 못했습니다. 웹에서 발급한 연결 설정과 편집기 실행 상태를 확인해 주세요.', file=sys.stderr)
        raise SystemExit(1) from None


if __name__ == '__main__':
    main()
