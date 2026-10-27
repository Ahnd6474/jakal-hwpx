"""Fail CI when known private runtime artifacts would be published in Git."""
from __future__ import annotations

import re
import hashlib
import subprocess
from pathlib import Path, PurePosixPath

# Deliberately public, generated introductory demos. Any change requires a
# fresh review; an uploaded document cannot take one of these filenames.
PUBLIC_DEMOS = {
    'web/public/sample.hwpx': '219f823825b5d896fc38ef11d9339a8fb14a3a0142c6d11a7acb5345f71039cd',
    'web/public/sample.hwp': 'e483cad81bbbf528e1afab093d467e0840fe4234695763eeeae23ebbaa70349e',
}


def private_artifact(name: str) -> bool:
    path = PurePosixPath(name.lower())
    if any(part in {'.web-data', '.web-test-data', 'private-uploads', 'uploads', 'session-backups'} for part in path.parts):
        return True
    base = path.name
    if (base == '.env' or base.startswith('.env.')) and not base.endswith('.example'):
        return True
    if path.suffix in {'.sqlite', '.sqlite3', '.pem', '.key'} or '.sqlite3-' in base:
        return True
    if re.fullmatch(r'(?:snapshot-\d+|job-[a-f0-9]{32}|result-[a-f0-9]{32})\.json', base):
        return True
    return name.startswith(('web/public/', 'web/dist/', 'website/')) and path.suffix in {'.hwp', '.hwpx', '.pdf'}


if __name__ == '__main__':
    tracked = subprocess.check_output(['git', 'ls-files', '-z']).decode('utf-8').split('\0')
    failures = [name for name in tracked if name and private_artifact(name)
        and not (name in PUBLIC_DEMOS and hashlib.sha256(Path(name).read_bytes()).hexdigest() == PUBLIC_DEMOS[name])]
    if failures:
        print('Private runtime/document artifacts must not be published:')
        print('\n'.join(failures))
        raise SystemExit(1)
    print('Private artifact publication check passed. Review sample documents separately for personal data.')
