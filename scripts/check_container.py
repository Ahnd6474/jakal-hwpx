"""Verify the production image with isolated volumes, restart and real TLS."""
from __future__ import annotations

import argparse
from pathlib import Path
import secrets
import subprocess
import time

REPO = Path(__file__).resolve().parents[1]


def docker(*args, source=None, check=True):
    return subprocess.run(['docker', *args], input=source, text=True, encoding='utf-8',
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=check)


def wait_healthy(container):
    deadline = time.monotonic() + 90
    while time.monotonic() < deadline:
        result = docker('exec', container, 'python', '-m', 'server.healthcheck', check=False)
        if result.returncode == 0:
            return
        if docker('inspect', '--format', '{{.State.Running}}', container).stdout.strip() != 'true':
            raise RuntimeError('Application exited before it became healthy.')
        time.sleep(.5)
    raise RuntimeError('Application did not become healthy.')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--image', default='jakal-workspace:validation')
    args = parser.parse_args()
    label = 'jakal-check-' + secrets.token_hex(6)
    app, gateway, volume, network = (label + suffix for suffix in ('-app', '-gateway', '-data', '-net'))
    config = REPO / '.artifacts' / 'validation' / (label + '.Caddyfile')
    config.parent.mkdir(parents=True, exist_ok=True)
    original = (REPO / 'deploy/Caddyfile').read_text(encoding='utf-8')
    # Exercise the production proxy configuration with a local test CA.
    config.write_text(original.replace('admin off', 'admin off\n    local_certs\n    skip_install_trust'), encoding='utf-8')
    module = (REPO / 'scripts/check_deployment.py').read_text(encoding='utf-8')
    shared = f"scope = {{'__name__': 'deployment_probe'}}\nexec(compile({module!r}, 'deployment_probe.py', 'exec'), scope)\n"
    shared += "Client = scope['DeploymentClient']\n"

    def execute(source):
        result = docker('exec', '-i', app, 'python', '-', source=source)
        if result.stdout:
            print(result.stdout.strip(), flush=True)

    try:
        docker('volume', 'create', volume)
        docker('network', 'create', '--internal', network)
        docker('run', '-d', '--name', gateway, '--network', network,
            '--network-alias', 'deployment.invalid', '--tmpfs', '/data', '--tmpfs', '/config',
            '--mount', f'type=bind,source={config},target=/etc/caddy/Caddyfile,readonly',
            '-e', 'JAKAL_DOMAIN=deployment.invalid', 'caddy:2-alpine')
        address = docker('inspect', '--format', '{{range .NetworkSettings.Networks}}{{.IPAddress}}{{end}}', gateway).stdout.strip()
        docker('run', '-d', '--name', app, '--network', network, '--network-alias', 'app',
            '--read-only', '--cap-drop', 'ALL', '--security-opt', 'no-new-privileges:true',
            '--tmpfs', '/tmp:rw,noexec,nosuid,size=268435456,uid=10001,gid=10001,mode=1770',
            '--mount', f'type=volume,source={volume},target=/var/lib/jakal',
            '-e', 'JAKAL_PUBLIC_ORIGIN=https://deployment.invalid',
            '-e', f'JAKAL_TRUSTED_PROXIES=127.0.0.1,{address}', args.image)
        wait_healthy(app)
        execute(shared + "scope['check_deployment']()\n")
        execute(shared + '''
import json, os, stat
from pathlib import Path
import pypdfium2 as pdfium
from PIL import Image
import jakal_hwpx as package
assert {'HeaderFooterBlock', 'ShapeObject', 'TableCell'} <= set(package.__all__)
assert all(getattr(package, name) is not None for name in package.__all__)
assert os.getuid() == 10001
assert stat.S_IMODE(Path('/var/lib/jakal').stat().st_mode) == 0o700
pdfium.PdfDocument.new().close()
Image.new('RGB', (2, 2)).save('/tmp/qa-image.jpg')
try:
    Path('/app/qa-write').write_bytes(b'test')
except OSError:
    pass
else:
    raise AssertionError('Application root is writable')
client = Client()
client.begin()
document = scope['HwpxDocument'].blank()
document.set_paragraph_text(0, 0, 'Container persistence check')
base = client.upload('persistence.hwpx', document.compile())
os.umask(0o077)
Path('/var/lib/jakal/qa-checkpoint.json').write_text(json.dumps({'headers':client.headers,'base':base}))
print('Non-root access, read-only filesystem, private volume, PDF and image runtimes passed.')
''')
        docker('restart', '--time', '15', app)
        wait_healthy(app)
        execute(shared + '''
import json
from pathlib import Path
checkpoint = Path('/var/lib/jakal/qa-checkpoint.json')
saved = json.loads(checkpoint.read_text())
client = Client()
client.headers = saved['headers']
assert client.json('GET', '/api/session')['workspace'] == saved['headers']['X-Jakal-Workspace']
assert client.json('GET', saved['base'])['status'] == 'ready'
assert client.request('GET', saved['base'] + '/original')[0] == 200
client.json('DELETE', saved['base'])
client.json('DELETE', '/api/session')
checkpoint.unlink()
print('Session and document access survived container restart on the persistent volume.')
''')
        authority = docker('exec', gateway, 'cat', '/data/caddy/pki/authorities/local/root.crt').stdout
        execute(shared + f'''
import http.client, ssl
context = ssl.create_default_context(cadata={authority!r})
Client.connection_factory = staticmethod(lambda: http.client.HTTPSConnection('deployment.invalid', 443, context=context, timeout=10))
redirect = http.client.HTTPConnection('deployment.invalid', 80, timeout=5)
redirect.request('GET', '/api/health')
response = redirect.getresponse()
assert response.status in (301, 302, 307, 308)
assert response.getheader('Location').startswith('https://deployment.invalid/')
response.read()
redirect.close()
scope['check_deployment'](via_gateway=True)
print('Verified TLS, HTTP redirect and spoofed forwarding-header handling through the production gateway.')
''')
        mcp_probe = (REPO / 'scripts/check_mcp.py').read_text(encoding='utf-8')
        execute(f'''
import asyncio, ssl
scope = {{'__name__': 'mcp_probe'}}
exec(compile({mcp_probe!r}, 'mcp_probe.py', 'exec'), scope)
context = ssl.create_default_context(cadata={authority!r})
asyncio.run(scope['check_site']('https://deployment.invalid', stdio=False, verify=context))
''')
    except subprocess.CalledProcessError as error:
        raise RuntimeError(error.stderr.strip() or error.stdout.strip() or 'Docker verification failed.') from None
    finally:
        for container in (app, gateway):
            docker('rm', '-f', container, check=False)
        docker('network', 'rm', network, check=False)
        docker('volume', 'rm', volume, check=False)
        config.unlink(missing_ok=True)


if __name__ == '__main__':
    main()
