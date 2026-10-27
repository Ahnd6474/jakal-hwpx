"""Check the actual Docker context without sending local runtime documents."""
from __future__ import annotations

import hashlib
import subprocess
from pathlib import Path

from check_private_artifacts import PUBLIC_DEMOS

REPO = Path(__file__).resolve().parents[1]


def main():
    destination = REPO / '.artifacts' / 'validation' / 'container-context'
    destination.mkdir(parents=True, exist_ok=True)
    subprocess.run(['docker', 'build', '--file', '-', '--output', f'type=local,dest={destination}', '.'],
        cwd=REPO, input=b'FROM scratch\nCOPY . /\n', check=True)
    names = {path.relative_to(destination).as_posix() for path in destination.rglob('*') if path.is_file()}
    expected = {'web/package-lock.json', 'web/src/main.tsx', 'web/studio/host-bridge.ts',
        'server/app.py', 'src/jakal_hwpx/__init__.py', *PUBLIC_DEMOS}
    assert expected <= names, f'Missing container build inputs: {sorted(expected - names)}'
    forbidden = {name for name in names if name.startswith(('web/node_modules/', 'web/dist/',
        'web/public/studio/', '.git/', '.web-data/', '.venv/', 'tests/', 'web_tests/'))
        or '/__pycache__/' in name or any(part.startswith('.env') for part in Path(name).parts)}
    public_files = {*PUBLIC_DEMOS, 'web/public/favicon.svg', 'web/public/404.html'}
    forbidden |= {name for name in names if name.startswith('web/public/')
        and name not in public_files and not name.startswith(('web/public/fonts/', 'web/public/licenses/'))}
    assert not forbidden, 'Local dependencies, build outputs or private runtime files entered the Docker context.'
    for name, digest in PUBLIC_DEMOS.items():
        assert hashlib.sha256((destination / name).read_bytes()).hexdigest() == digest
    print(f'Container context: {len(names)} application inputs; local dependencies and private runtime data excluded.')


if __name__ == '__main__':
    main()
