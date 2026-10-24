"""Package the prebuilt website and portable server deployment controls."""
from __future__ import annotations

import argparse
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import tomllib

from build_desktop import REPO, archive_release, copy_application


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--skip-web-build', action='store_true')
    args = parser.parse_args()
    if not args.skip_web_build:
        subprocess.run(['npm.cmd' if sys.platform == 'win32' else 'npm', 'run', 'build'], cwd=REPO / 'web', check=True)
    if not (REPO / 'web/dist/studio/index.html').is_file():
        raise RuntimeError('Build the website before packaging.')
    version = tomllib.loads((REPO / 'pyproject.toml').read_text(encoding='utf-8'))['project']['version']
    work = REPO / '.artifacts/server-builds'
    work.mkdir(parents=True, exist_ok=True)
    stage = Path(tempfile.mkdtemp(prefix='build-', dir=work))
    root = stage / 'Jakal-Server'
    copy_application(root / 'app')
    subprocess.run([sys.executable, '-m', 'build', '--wheel', '--outdir', str(root / 'packages'), '--no-isolation'], cwd=REPO, check=True)
    for source, target in [('LICENSE', 'LICENSE'), ('THIRD_PARTY_NOTICES.md', 'THIRD_PARTY_NOTICES.md'),
            ('compose.yaml', 'compose.yaml'), ('.env.production.example', '.env.production.example'),
            ('deploy/Dockerfile.release', 'Dockerfile'), ('deploy/manage.sh', 'manage.sh'),
            ('deploy/manage.ps1', 'manage.ps1'), ('deploy/StartServer.cmd', 'StartServer.cmd')]:
        shutil.copy2(REPO / source, root / target)
    for name in ['local.sh', 'local.ps1', 'StartLocal.cmd', 'StopLocal.cmd', 'guide.html']:
        shutil.copy2(REPO / 'deploy' / name, root / name)
    shutil.copy2(REPO / 'compose.local.yaml', root / 'compose.local.yaml')
    (root / 'deploy').mkdir()
    shutil.copy2(REPO / 'deploy/Caddyfile', root / 'deploy/Caddyfile')
    shutil.copy2(REPO / 'deploy/Caddyfile.local', root / 'deploy/Caddyfile.local')
    (root / 'manage.sh').write_text((REPO / 'deploy/manage.sh').read_text(encoding='utf-8'), encoding='utf-8', newline='\n')
    (root / '.dockerignore').write_text('**\n!app/\n!app/**\n!packages/\n!packages/*.whl\n!LICENSE\n!THIRD_PARTY_NOTICES.md\n**/__pycache__/\n**/.env*\n**/*.sqlite*\n**/.web-data/\n**/uploads/\n', encoding='utf-8')
    (root / 'manage.sh').chmod(0o755)
    (root / 'local.sh').write_text((REPO / 'deploy/local.sh').read_text(encoding='utf-8'), encoding='utf-8', newline='\n')
    (root / 'local.sh').chmod(0o755)
    output = REPO / 'dist'
    output.mkdir(exist_ok=True)
    archive_release(root, output, f'Jakal-Server-{version}.zip', {'version': version, 'platform': 'docker'})


if __name__ == '__main__':
    main()
