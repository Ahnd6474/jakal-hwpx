from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
import time
from pathlib import Path

from .storage import Store

REPO = Path(__file__).resolve().parents[1]


class Worker:
    def __init__(self, store: Store, timeout: int = 120):
        self.store = store
        self.timeout = timeout
        self.stop = threading.Event()
        self.thread = threading.Thread(target=self.loop, daemon=True, name="hwpx-jobs")
        self.process: subprocess.Popen | None = None

    def start(self):
        self.store.recover()
        self.thread.start()

    def close(self):
        self.stop.set()
        if self.process and self.process.poll() is None:
            self.process.kill()
        self.thread.join(timeout=5)

    def loop(self):
        next_cleanup = 0.0
        while not self.stop.is_set():
            try:
                if time.monotonic() >= next_cleanup:
                    self.store.cleanup()
                    next_cleanup = time.monotonic() + 60
                job = self.store.claim()
                if job:
                    self.run(job)
                    continue
            except Exception:
                import logging
                logging.getLogger(__name__).exception("Document worker failure")
            self.stop.wait(0.15)

    def run(self, job: dict):
        document = self.store.document(job["document"], job["owner"])
        if not document:
            self.store.finish(job, None, "문서 보관 기간이 만료되었습니다.")
            return
        directory = self.store.directory(job["document"])
        request_path = directory / f"job-{job['id']}.json"
        result_path = directory / f"result-{job['id']}.json"
        request_path.write_text(json.dumps({"directory": str(directory), "kind": job["kind"], "payload": json.loads(job["payload"]), "name": document["name"]}, ensure_ascii=False), encoding="utf-8")
        result = None
        error = None
        try:
            env = os.environ.copy()
            env["PYTHONPATH"] = str(REPO / "src") + os.pathsep + str(REPO)
            env["PYTHONIOENCODING"] = "utf-8"
            self.process = subprocess.Popen([sys.executable, "-m", "server.process_job", str(request_path), str(result_path)], cwd=REPO, env=env, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0)
            import psutil
            started = time.monotonic()
            while self.process.poll() is None:
                if self.stop.wait(0.05) or time.monotonic() - started > self.timeout:
                    error = "작업이 중단되었거나 제한 시간(120초)을 초과했습니다."
                    self.process.kill()
                    break
                try:
                    if psutil.Process(self.process.pid).memory_info().rss > 768 * 1024 * 1024:
                        error = "문서 처리 메모리 제한(768MB)을 초과했습니다."
                        self.process.kill()
                        break
                except psutil.NoSuchProcess:
                    pass
            self.process.wait(timeout=5)
            if not error:
                outcome = json.loads(result_path.read_text(encoding="utf-8"))
                result, error = outcome.get("result"), outcome.get("error")
        except Exception:
            error = "문서 처리 프로세스가 종료되었습니다. 다시 시도해 주세요."
        finally:
            if self.process and self.process.poll() is None:
                self.process.kill()
                self.process.wait(timeout=5)
            self.process = None
            self.store.finish(job, result, error)
            request_path.unlink(missing_ok=True)
            result_path.unlink(missing_ok=True)
