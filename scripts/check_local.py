"""Check the shipped local controls, loopback publication and document flows."""
from __future__ import annotations

import argparse
import asyncio
import http.client
import ipaddress
import json
import os
from pathlib import Path
import secrets
import socket
import subprocess
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from check_deployment import DeploymentClient, check_deployment
from check_mcp import check_site


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('directory', type=Path)
    args = parser.parse_args()
    root = args.directory.resolve()
    assert (root / 'guide.html').is_file()
    project = 'jakal-preview-check-' + secrets.token_hex(5)
    with socket.socket() as listener:
        listener.bind(('127.0.0.1', 0))
        port = listener.getsockname()[1]
    networks = subprocess.check_output(['docker', 'network', 'ls', '-q'], text=True).split()
    configs = json.loads(subprocess.check_output(['docker', 'network', 'inspect', *networks], text=True)) if networks else []
    occupied = [ipaddress.ip_network(config['Subnet']) for network in configs for config in (network['IPAM']['Config'] or []) if config.get('Subnet')]
    subnet = next(ipaddress.ip_network(f'172.30.{part}.0/29') for part in range(100, 240)
        if not any(ipaddress.ip_network(f'172.30.{part}.0/29').overlaps(other) for other in occupied))
    env = os.environ | {'JAKAL_LOCAL_PROJECT': project, 'JAKAL_LOCAL_PORT': str(port),
        'JAKAL_LOCAL_SUBNET': str(subnet), 'JAKAL_LOCAL_PROXY_ADDRESS': str(subnet.network_address + 2),
        'JAKAL_LOCAL_APP_ADDRESS': str(subnet.network_address + 3)}
    compose = ['docker', 'compose', '--project-name', project, '--file', str(root / 'compose.local.yaml')]
    control = ['powershell.exe', '-NoProfile', '-NonInteractive', '-File', str(root / 'local.ps1'), '-Action'] if sys.platform == 'win32' else ['sh', str(root / 'local.sh')]

    def run(*command):
        subprocess.run(command, cwd=root, env=env, check=True)

    def containers(service):
        return subprocess.check_output([*compose, 'ps', '-q', service], env=env, text=True).strip()

    try:
        run(*control, 'start')
        app, gateway = containers('app'), containers('gateway')
        inspection = json.loads(subprocess.check_output(['docker', 'inspect', gateway], text=True))[0]
        assert inspection['HostConfig']['PortBindings'] == {'8080/tcp': [{'HostIp': '127.0.0.1', 'HostPort': str(port)}]}
        assert not json.loads(subprocess.check_output(['docker', 'inspect', app], text=True))[0]['HostConfig']['PortBindings']
        origin = f'http://127.0.0.1:{port}'
        os.environ['JAKAL_PUBLIC_ORIGIN'] = origin
        DeploymentClient.secure_cookie = False
        DeploymentClient.connection_factory = staticmethod(lambda: http.client.HTTPConnection('127.0.0.1', port, timeout=15))
        check_deployment(desktop=True)
        asyncio.run(check_site(origin, stdio=False))
        client = DeploymentClient()
        client.begin()
        from jakal_hwpx import HwpxDocument
        base = client.upload('restart.hwpx', HwpxDocument.blank().compile())
        result = client.json('POST', '/api/connections', {'document_id': base.rsplit('/', 1)[1], 'label': 'Docker local'}, status=201)
        assert result['stdio_config'] is None, 'A container must not advertise its Python path as a host command.'
        assert result['http_config']['mcpServers']['jakal']['url'] == origin + '/mcp'
        run(*control, 'stop')
        subprocess.run([*compose, 'up', '-d', '--wait', '--wait-timeout', '240'], cwd=root, env=env, check=True)
        assert client.json('GET', base)['status'] == 'ready'
        client.json('DELETE', '/api/session')
        print('Local kit passed: shipped controls, host-only ports, HWP/HWPX editing, MCP HTTP, session isolation and stop/restart persistence.', flush=True)
    except Exception:
        subprocess.run([*compose, 'logs', '--tail', '80'], cwd=root, env=env, check=False)
        raise
    finally:
        # This harness alone owns this random project and its generated data.
        subprocess.run([*compose, 'down', '--volumes'], cwd=root, env=env, check=True)


if __name__ == '__main__':
    main()
