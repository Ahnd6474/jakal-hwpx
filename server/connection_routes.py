from __future__ import annotations

import ipaddress
import os
from pathlib import Path
import sys

from fastapi import APIRouter, Request

from .connections import Approval, CreateConnection, Presence


def routes(connections):
    router = APIRouter()

    @router.post('/api/connections', status_code=201)
    def create(body: CreateConnection, request: Request):
        result = connections.create(request.state.owner, body)
        origin = str(request.base_url).rstrip('/')
        token = result['token']
        http_config = {'mcpServers': {'jakal': {'url': origin + '/mcp', 'headers': {'Authorization': 'Bearer ' + token}}}}
        try:
            local = request.url.hostname == 'localhost' or ipaddress.ip_address(request.url.hostname or '').is_loopback
        except ValueError:
            local = False
        stdio_config = {'mcpServers': {'jakal': {'command': sys.executable,
            'args': ['-I', '-B', '-X', 'utf8', str(Path(__file__).with_name('mcp_entry.py'))],
            'env': {'JAKAL_WORKSPACE_URL': origin, 'JAKAL_CONNECTION_TOKEN': token}}}} if local and os.environ.get('JAKAL_STDIO_CONFIG', '1') == '1' else None
        return result | {'http_config': http_config, 'stdio_config': stdio_config}

    @router.get('/api/documents/{identifier}/connections')
    def listing(identifier: str, request: Request):
        return {'connections': connections.list(request.state.owner, identifier),
            'state': connections.documents.state(identifier, request.state.owner),
            'proposals': connections.proposals(request.state.owner, identifier),
            'events': connections.events(request.state.owner, identifier),
            'editing': connections.editing(request.state.owner, identifier)}

    @router.delete('/api/connections/{identifier}')
    def revoke(identifier: str, request: Request):
        connections.revoke(request.state.owner, identifier)
        return {'revoked': True}

    @router.post('/api/documents/{identifier}/presence')
    def presence(identifier: str, body: Presence, request: Request):
        return connections.presence(request.state.owner, identifier, body)

    @router.delete('/api/documents/{identifier}/presence/{tab}')
    def release_presence(identifier: str, tab: str, request: Request):
        return connections.release_presence(request.state.owner, identifier, tab)

    @router.get('/api/proposals/{identifier}')
    def proposal(identifier: str, request: Request):
        return connections.review(request.state.owner, identifier)

    @router.post('/api/proposals/{identifier}/approve', status_code=202)
    def approve(identifier: str, body: Approval, request: Request):
        return connections.approve(request.state.owner, identifier, body)

    @router.post('/api/proposals/{identifier}/reject')
    def reject(identifier: str, request: Request):
        return connections.reject(request.state.owner, identifier)

    @router.get('/api/connected/info')
    def connected_info(request: Request):
        grant = request.state.connection
        return {key: grant[key] for key in ('id', 'document', 'label', 'permission', 'expires')}

    @router.post('/api/connected/tools/{name}')
    def tool(name: str, arguments: dict, request: Request):
        return connections.execute(request.state.connection, name, arguments)

    return router
