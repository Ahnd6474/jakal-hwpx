"""Build an offline Windows package from explicit, verified application inputs."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import shutil
import subprocess
import sys
import tempfile
import tomllib
import urllib.request
import zipfile

from check_private_artifacts import PUBLIC_DEMOS, private_artifact

REPO = Path(__file__).resolve().parents[1]


def copy_application(app):
    (app / 'server').mkdir(parents=True)
    for file in (REPO / 'server').iterdir():
        if file.is_file() and file.suffix in {'.py', '.mjs'}:
            shutil.copy2(file, app / 'server' / file.name)
    shutil.copytree(REPO / 'web/dist', app / 'web/dist')
    core = app / 'web/node_modules/@rhwp/core'
    core.mkdir(parents=True)
    for name in ['package.json', 'rhwp.js', 'rhwp_bg.wasm', 'LICENSE']:
        shutil.copy2(REPO / 'web/node_modules/@rhwp/core' / name, core / name)


def archive_release(root, output_dir, filename, metadata):
    files = [p for p in validate_bundle(root) if p.name != 'package-manifest.json']
    metadata = metadata | {'files': {p.relative_to(root).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest() for p in files}}
    (root / 'package-manifest.json').write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding='utf-8')
    archive = output_dir / filename
    with zipfile.ZipFile(archive, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=6) as target:
        for file in sorted(root.rglob('*')):
            if file.is_file():
                name = root.name + '/' + file.relative_to(root).as_posix()
                if file.suffix == '.sh':
                    entry = zipfile.ZipInfo.from_file(file, name)
                    entry.create_system = 3
                    entry.external_attr = 0o100755 << 16
                    target.writestr(entry, file.read_bytes(), compress_type=zipfile.ZIP_DEFLATED, compresslevel=6)
                else:
                    target.write(file, name)
    digest = hashlib.sha256(archive.read_bytes()).hexdigest()
    archive.with_suffix('.zip.sha256').write_text(digest + '  ' + archive.name + '\n', encoding='ascii')
    print(f'Package: {archive}\nSize: {archive.stat().st_size:,} bytes\nUnpacked application: {root}', flush=True)
    return archive


def download_verified(item, cache):
    target = cache / PurePosixPath(item['url']).name
    if target.is_file() and hashlib.file_digest(target.open('rb'), 'sha256').hexdigest() == item['sha256']:
        return target
    pending = target.with_suffix('.pending')
    request = urllib.request.Request(item['url'], headers={'User-Agent': 'Jakal-Package-Builder'})
    with urllib.request.urlopen(request, timeout=90) as source, pending.open('wb') as output:
        shutil.copyfileobj(source, output)
    with pending.open('rb') as source:
        actual = hashlib.file_digest(source, 'sha256').hexdigest()
    if actual != item['sha256']:
        pending.unlink()
        raise RuntimeError('Runtime download checksum does not match the pinned release.')
    pending.replace(target)
    return target


def extract_python(archive, target):
    target.mkdir(parents=True)
    with zipfile.ZipFile(archive) as source:
        for member in source.infolist():
            path = PurePosixPath(member.filename)
            if path.is_absolute() or '..' in path.parts or '\\' in member.filename or ':' in member.filename:
                raise RuntimeError('Invalid runtime archive path.')
            source.extract(member, target)
    (target / 'python313._pth').write_text('python313.zip\n.\n../packages\n../../app\nimport site\n', encoding='utf-8')
    (target / 'sitecustomize.py').write_text('import site, sys\nfrom pathlib import Path\nsys.dont_write_bytecode = True\nsite.addsitedir(str(Path(__file__).resolve().parents[1] / "packages"))\n', encoding='utf-8')
    subprocess.run([str(target / 'python.exe'), '-I', '-B', '-X', 'utf8', '-c',
        "import py_compile; py_compile.compile('sitecustomize.py', cfile='sitecustomize.pyc', doraise=True)"], cwd=target, check=True)
    (target / 'sitecustomize.py').unlink()


def validate_bundle(root):
    files = sorted(path for path in root.rglob('*') if path.is_file())
    for file in files:
        name = file.relative_to(root).as_posix()
        if file.is_symlink() or private_artifact(name) or '/__pycache__/' in name:
            raise RuntimeError(f'Private or unexpected runtime artifact in package: {name}')
        if file.suffix.lower() in {'.hwp', '.hwpx', '.pdf'}:
            if name.startswith('runtime/packages/jakal_hwpx/bundled_hwp_profile/') and file.suffix == '.hwp':
                continue
            public_name = name.removeprefix('app/')
            public_name = public_name.replace('web/dist/', 'web/public/', 1)
            expected = PUBLIC_DEMOS.get(public_name)
            if expected is None or hashlib.sha256(file.read_bytes()).hexdigest() != expected:
                raise RuntimeError(f'Only pinned sample documents may be distributed: {name}')
    return files


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--skip-web-build', action='store_true')
    parser.add_argument('--output', type=Path, default=REPO / 'dist')
    args = parser.parse_args()
    if sys.platform != 'win32':
        raise SystemExit('The Windows launcher must be built on Windows.')
    version = tomllib.loads((REPO / 'pyproject.toml').read_text(encoding='utf-8'))['project']['version']
    if not all(part.isdigit() for part in version.split('.')):
        raise RuntimeError('Expected a numeric release version.')
    runtimes = json.loads((REPO / 'desktop/runtimes.json').read_text(encoding='utf-8'))
    cache = REPO / '.artifacts' / 'runtime-downloads'
    cache.mkdir(parents=True, exist_ok=True)
    # Download independent pinned runtimes before the build needs them.
    python_zip = download_verified(runtimes['python'], cache)
    node_zip = download_verified(runtimes['node'], cache)
    if not args.skip_web_build:
        subprocess.run(['npm.cmd', 'run', 'build'], cwd=REPO / 'web', check=True)
    if not (REPO / 'web/dist/studio/index.html').is_file():
        raise RuntimeError('Build the web app and native studio before packaging.')
    args.output.mkdir(parents=True, exist_ok=True)
    work = REPO / '.artifacts' / 'desktop-builds'
    work.mkdir(parents=True, exist_ok=True)
    stage = Path(tempfile.mkdtemp(prefix='build-', dir=work))
    root = stage / 'Jakal-Workspace'
    app = root / 'app'
    app.mkdir(parents=True)
    subprocess.run([sys.executable, '-m', 'build', '--wheel', '--outdir', str(stage / 'wheels'), '--no-isolation'], cwd=REPO, check=True)
    wheel = next((stage / 'wheels').glob('*.whl'))
    extract_python(python_zip, root / 'runtime/python')
    node_target = root / 'runtime/node'
    node_target.mkdir()
    with zipfile.ZipFile(node_zip) as source:
        prefix = 'node-v' + runtimes['node']['version'] + '-win-x64/'
        for name in ['node.exe', 'LICENSE']:
            (node_target / name).write_bytes(source.read(prefix + name))
    subprocess.run([sys.executable, '-m', 'pip', 'install', '--no-compile', '--only-binary=:all:',
        '--platform', 'win_amd64', '--python-version', '3.13', '--implementation', 'cp', '--abi', 'cp313',
        '--target', str(root / 'runtime/packages'), str(wheel) + '[web]'], cwd=REPO, check=True)
    copy_application(app)
    shutil.copy2(REPO / 'desktop/bootstrap.py', app / 'desktop_bootstrap.py')
    shutil.copy2(REPO / 'desktop/guide.html', root / 'guide.html')
    shutil.copy2(REPO / 'desktop/Start.cmd', root / '시작.cmd')
    for name in ['LICENSE', 'THIRD_PARTY_NOTICES.md']:
        shutil.copy2(REPO / name, root / name)
    # The installed runtime is fully local; user Python paths and site packages
    # are excluded by python313._pth and the launcher's isolated-mode flag.
    for file in (root / 'runtime/packages').glob('*.dist-info/RECORD'):
        if not file.is_file():
            raise RuntimeError('Incomplete dependency metadata.')
    compiler = Path(os.environ['WINDIR']) / 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
    assembly = stage / 'AssemblyInfo.cs'
    numeric = '.'.join((version.split('.') + ['0', '0', '0', '0'])[:4])
    assembly.write_text('using System.Reflection;\n[assembly: AssemblyVersion("' + numeric + '")]\n', encoding='utf-8')
    subprocess.run([str(compiler), '/nologo', '/target:winexe', '/platform:x64', '/codepage:65001',
        '/reference:System.Windows.Forms.dll', '/reference:System.Drawing.dll',
        '/win32manifest:' + str(REPO / 'desktop/launcher.manifest'),
        '/out:' + str(root / 'Jakal.exe'), str(REPO / 'desktop/Launcher.cs'), str(assembly)], check=True)
    archive_release(root, args.output, f'Jakal-Workspace-{version}-windows-x64.zip',
        {'version': version, 'platform': 'windows-x64', 'runtimes': runtimes})


if __name__ == '__main__':
    main()
