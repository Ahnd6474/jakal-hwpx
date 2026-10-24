"""Exercise real SDK clients against an isolated document in a running site."""
from __future__ import annotations

import argparse
import asyncio
import base64
from contextlib import asynccontextmanager
import secrets

import httpx2
from mcp import Client, StdioServerParameters
from mcp.client.streamable_http import streamable_http_client

from jakal_hwpx import HwpxDocument


async def check_site(origin, *, stdio=True, verify=True):
    async with httpx2.AsyncClient(base_url=origin, timeout=30, trust_env=False, follow_redirects=False, verify=verify) as browser:
        async def web(method, path, **kwargs):
            result = await browser.request(method, path, **kwargs)
            assert result.is_success, (method, path, result.status_code, result.text[:200])
            return result.json()

        async def wait(job):
            for _ in range(600):
                result = await web('GET', '/api/jobs/' + job)
                if result['status'] in {'completed', 'failed'}:
                    assert result['status'] == 'completed', result
                    return result
                await asyncio.sleep(.05)
            raise AssertionError('MCP document job timed out')

        source = HwpxDocument.blank()
        source.set_paragraph_text(0, 0, 'MCP 연결 검증 문서')
        source.append_table(rows=1, columns=1, cell_texts=[['표 보존']])
        uploaded = await web('POST', '/api/documents', files={'file': ('mcp-check.hwpx', source.compile())})
        document = uploaded['document_id']
        try:
            await wait(uploaded['job_id'])
            transports = ['http', 'stdio'] if stdio else ['http']
            for index, transport in enumerate(transports):
                grant = await web('POST', '/api/connections', json={'document_id': document, 'label': 'SDK ' + transport, 'permission': 'propose'})
                if transport == 'stdio':
                    config = grant['stdio_config']['mcpServers']['jakal']
                    endpoint = StdioServerParameters(**config)
                else:
                    @asynccontextmanager
                    async def authorized_http():
                        async with httpx2.AsyncClient(headers={'Authorization': 'Bearer ' + grant['token']}, timeout=30, trust_env=False, verify=verify) as http:
                            async with streamable_http_client(origin + '/mcp', http_client=http) as streams:
                                yield streams
                    endpoint = authorized_http()
                async with Client(endpoint, read_timeout_seconds=30) as mcp:
                    tools = await mcp.list_tools()
                    assert 'propose_changes' in {tool.name for tool in tools.tools}
                    assert 'apply_changes' not in {tool.name for tool in tools.tools}
                    outline = (await mcp.call_tool('get_document_outline', {})).structured_content
                    assert outline['id'] == document and outline['revision'] == index
                    text = f'MCP {transport} 편집 완료'
                    proposed = await mcp.call_tool('propose_changes', {'revision': index, 'settings_version': 0,
                        'request_key': secrets.token_hex(16), 'changes': [{'id': 's0-p0', 'text': text}]})
                    assert not proposed.is_error, proposed
                    proposal = proposed.structured_content['proposal_id']
                    assert (await web('GET', '/api/documents/' + document))['revision'] == index
                    accepted = await web('POST', f'/api/proposals/{proposal}/approve', json={'revision': index, 'settings_version': 0, 'targets': ['s0-p0']})
                    await wait(accepted['job_id'])
                    validated = await mcp.call_tool('validate_document', {'revision': index + 1, 'settings_version': 0, 'request_key': secrets.token_hex(16)})
                    assert not validated.is_error, validated
                    await wait(validated.structured_content['job_id'])
                    job = await mcp.call_tool('get_job', {'job_id': validated.structured_content['job_id']})
                    assert job.structured_content['status'] == 'completed'
                    exported = await mcp.call_tool('export_document', {'revision': index + 1})
                    assert not exported.is_error, exported
                    resource = next(content.resource for content in exported.content if content.type == 'resource')
                    document_copy = HwpxDocument.from_bytes(base64.b64decode(resource.blob))
                    assert document_copy.sections[0].text_fragments()[0] == text
                    assert document_copy.tables()[0].cell(0, 0).text == '표 보존'
                    await web('DELETE', '/api/connections/' + grant['id'])
                    if transport == 'stdio':
                        denied = await mcp.call_tool('get_document_outline', {})
                        assert denied.is_error
                print(f'MCP {transport}: official SDK discovery, scoped tools, web approval, edit, validation, export and revocation passed.', flush=True)
        finally:
            await browser.delete('/api/session')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('origin')
    parser.add_argument('--http-only', action='store_true')
    args = parser.parse_args()
    asyncio.run(check_site(args.origin.rstrip('/'), stdio=not args.http_only))


if __name__ == '__main__':
    main()
