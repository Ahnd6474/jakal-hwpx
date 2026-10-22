from __future__ import annotations

import json
import logging
import os
import re
import subprocess
import sys
import threading
import time
from pathlib import Path

from .storage import Store

REPO = Path(__file__).resolve().parents[1]
logger = logging.getLogger(__name__)


class Worker:
    def __init__(self, store: Store, timeout: int = 120):
        self.store = store
        self.timeout = timeout
        self.stop = threading.Event()
        self.thread = threading.Thread(target=self.loop, daemon=True, name="hwpx-jobs")
        self.process: subprocess.Popen | None = None

    def start(self):
        for job in self.store.recover():
            self.cleanup_files(job)
        self.thread.start()

    def close(self):
        self.stop.set()
        if self.process and self.process.poll() is None:
            self.kill_process_tree()
        self.thread.join(timeout=5)

    def kill_process_tree(self):
        import psutil
        if not self.process:
            return
        try:
            process = psutil.Process(self.process.pid)
            children = process.children(recursive=True)
            for child in reversed(children):
                try:
                    child.kill()
                except psutil.NoSuchProcess:
                    pass
            process.kill()
        except psutil.NoSuchProcess:
            pass

    def loop(self):
        next_cleanup = 0.0
        unsettled = None
        while not self.stop.is_set():
            job = None
            try:
                if unsettled:
                    # Keep retrying the DB transition after temporary storage
                    # outages, without executing the document operation twice.
                    self.store.finish(unsettled, None, "작업 결과를 저장하지 못했습니다. 다시 시도해 주세요.")
                    unsettled = None
                if time.monotonic() >= next_cleanup:
                    self.store.cleanup()
                    next_cleanup = time.monotonic() + 60
                job = self.store.claim()
                if job:
                    self.run(job)
                    continue
            except Exception:
                logger.exception("Document worker failure")
                if job:
                    unsettled = job
            self.stop.wait(0.15)

    def run(self, job: dict):
        result = None
        error = None
        try:
            document = self.store.document(job["document"], job["owner"])
            if not document:
                error = "문서 보관 기간이 만료되었습니다."
                return
            directory = self.store.directory(job["document"])
            request_path = directory / f"job-{job['id']}.json"
            result_path = directory / f"result-{job['id']}.json"
            request_path.write_text(json.dumps({"directory": str(directory), "kind": job["kind"], "payload": json.loads(job["payload"]), "name": document["name"]}, ensure_ascii=False), encoding="utf-8")
            from .process_environment import parser_environment
            env = parser_environment(REPO)
            self.process = subprocess.Popen([sys.executable, "-m", "server.process_job", str(request_path), str(result_path)], cwd=REPO, env=env, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0)
            import psutil
            started = time.monotonic()
            while self.process.poll() is None:
                if self.stop.wait(0.05) or time.monotonic() - started > self.timeout:
                    error = f"작업이 중단되었거나 제한 시간({self.timeout}초)을 초과했습니다."
                    self.kill_process_tree()
                    break
                try:
                    parent = psutil.Process(self.process.pid)
                    processes = [parent, *parent.children(recursive=True)]
                    if sum(p.memory_info().rss for p in processes) > 768 * 1024 * 1024:
                        error = "문서 처리 메모리 제한(768MB)을 초과했습니다."
                        self.kill_process_tree()
                        break
                except psutil.NoSuchProcess:
                    pass
            exit_code = self.process.wait(timeout=5)
            if not error:
                if exit_code != 0:
                    raise ValueError("Document process exited unsuccessfully")
                outcome = json.loads(result_path.read_text(encoding="utf-8"))
                if isinstance(outcome, dict) and set(outcome) == {"result"} and isinstance(outcome["result"], dict):
                    result = outcome["result"]
                elif isinstance(outcome, dict) and set(outcome) == {"error"} and isinstance(outcome["error"], str) and outcome["error"].strip():
                    error = outcome["error"][:1500]
                else:
                    raise ValueError("Incomplete document process output")
        except Exception:
            logger.exception("Document job failed: %s", job["id"])
            error = "문서 처리 프로세스가 종료되었습니다. 다시 시도해 주세요."
        finally:
            try:
                if self.process and self.process.poll() is None:
                    self.kill_process_tree()
                    self.process.wait(timeout=5)
            except Exception:
                logger.exception("Could not stop document process")
                error = "문서 처리 프로세스를 중단했습니다. 다시 시도해 주세요."
            finally:
                self.process = None
                try:
                    self.store.finish(job, result, error)
                finally:
                    self.cleanup_files(job)

    def cleanup_files(self, job: dict):
        directory = self.store.directory(job["document"])
        paths = [directory / f"job-{job['id']}.json", directory / f"result-{job['id']}.json", directory / f"result-{job['id']}.json.tmp"]
        try:
            payload = json.loads(job["payload"])
            candidate = payload.get("candidate", "")
            if job["kind"] == "studio" and isinstance(candidate, str) and re.fullmatch(r"studio-[a-f0-9]{32}\.hwpx", candidate):
                paths.append(directory / candidate)
            if job['kind'] == 'layout':
                paths.extend(path for path in directory.glob('layout-*.hwpx*') if re.fullmatch(r'layout-[a-f0-9]{32}\.hwpx(?:\.tmp)?', path.name))
            revision = payload.get("revision", 0)
            if type(revision) is int and revision >= 0:
                revision += job["kind"] in {"edit", "studio", "restore", "layout"}
                paths.extend(directory / name for name in (f"snapshot-{revision}.json.tmp", f"snapshot-{revision}.tmp", f"revision-{revision}.hwpx.tmp", f"revision-{revision}.hwp.tmp"))
        except (ValueError, AttributeError):
            pass
        # These names are subprocess-only scratch files, never original files.
        paths.extend(directory / name for name in ("pending.hwp", "conversion-pending.hwpx", "conversion-pending.json"))
        for path in paths:
            try:
                path.unlink(missing_ok=True)
            except OSError:
                logger.warning("Could not remove temporary job file: %s", path.name)
