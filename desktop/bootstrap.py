"""Application-local Python entry point, supervised by the Windows launcher."""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import socket
import sys
import threading
import time


def acquire_lock(path: Path):
    path.touch(mode=0o600, exist_ok=True)
    stream = path.open('r+b')
    if not path.stat().st_size:
        stream.write(b'0')
        stream.flush()
    stream.seek(0)
    try:
        if os.name == 'nt':
            import msvcrt
            msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(stream, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        stream.close()
        raise RuntimeError('This document storage is already being used by another editor.') from None
    return stream


def bind_local(preferred: int):
    candidates = [preferred] if preferred == 0 else [preferred, *range(18987, 19008)]
    for port in dict.fromkeys(candidates):
        listener = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        if os.name == 'nt':
            listener.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        try:
            listener.bind(('127.0.0.1', port))
            listener.listen(128)
            return listener
        except OSError:
            listener.close()
    raise RuntimeError('No local editor port is available. Close another editor and try again.')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--data-root', type=Path, required=True)
    parser.add_argument('--port', type=int, default=18987)
    args = parser.parse_args()
    if not 0 <= args.port <= 65535:
        parser.error('Port must be between 0 and 65535.')
    root = args.data_root.resolve()
    root.mkdir(parents=True, exist_ok=True, mode=0o700)
    state = root / 'state'
    state.mkdir(exist_ok=True, mode=0o700)
    lock = acquire_lock(state / 'server.lock')
    ready = state / 'ready.json'
    preferences = state / 'preferences.json'
    app = Path(__file__).resolve().parent
    if not (app / 'server').is_dir():
        app = app.parent
    sys.path.insert(0, str(app))
    os.environ.update(JAKAL_ENV='development', JAKAL_HOST='127.0.0.1',
        JAKAL_DATA_DIR=str(root / 'documents'), JAKAL_TRUSTED_PROXIES='')
    os.environ.pop('JAKAL_PUBLIC_ORIGIN', None)
    os.environ.pop('JAKAL_COOKIE_SECURE', None)
    node = app.parent / 'runtime' / 'node' / 'node.exe'
    if node.is_file():
        os.environ['JAKAL_NODE'] = str(node)
    preferred = args.port
    if preferred == 18987:
        try:
            saved = json.loads(preferences.read_text(encoding='utf-8'))['port']
            if isinstance(saved, int) and 18987 <= saved <= 19007:
                preferred = saved
        except (OSError, ValueError, KeyError, TypeError):
            pass
    listener = bind_local(preferred)
    port = listener.getsockname()[1]
    os.environ['JAKAL_PORT'] = str(port)
    from server.files import atomic_json
    from jakal_hwpx import __version__
    import uvicorn
    server = uvicorn.Server(uvicorn.Config('server.app:create_app', factory=True,
        host='127.0.0.1', port=port, workers=1, access_log=False, server_header=False,
        forwarded_allow_ips='', log_config=None, timeout_graceful_shutdown=8))

    def watch_parent():
        try:
            while True:
                line = sys.stdin.readline()
                if not line or line.strip() == 'stop':
                    break
        finally:
            server.should_exit = True

    def publish_ready():
        while not server.started and not server.should_exit:
            time.sleep(.05)
        if server.started and not server.should_exit:
            atomic_json(preferences, {'port': port})
            atomic_json(ready, {'pid': os.getpid(), 'port': port, 'version': __version__})
            print(f'READY {port}', flush=True)

    threading.Thread(target=watch_parent, daemon=True).start()
    threading.Thread(target=publish_ready, daemon=True).start()
    try:
        server.run(sockets=[listener])
    finally:
        listener.close()
        try:
            if json.loads(ready.read_text(encoding='utf-8'))['pid'] == os.getpid():
                ready.unlink()
        except (OSError, ValueError, KeyError):
            pass
        lock.close()


if __name__ == '__main__':
    main()
