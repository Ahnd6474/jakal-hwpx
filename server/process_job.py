"""One bounded subprocess per job. Never executes uploaded or generated code."""
import json
import sys
from pathlib import Path
from .files import atomic_json


def main():
    request_path, result_path = map(Path, sys.argv[1:3])
    try:
        if sys.platform != "win32":
            import resource
            # The worker monitors aggregate RSS, including the Node converter.
            # V8 reserves a large virtual address range before allocating RAM;
            # an inherited RLIMIT_AS prevents it from starting on Linux.
            resource.setrlimit(resource.RLIMIT_CPU, (100, 100))
        from .engine import process
        request = json.loads(request_path.read_text(encoding="utf-8"))
        result = process(Path(request["directory"]), request["kind"], request["payload"], request["name"])
        output = {"result": result}
    except Exception as exc:
        # No server paths, tracebacks or environment details cross the API.
        if isinstance(exc, (ValueError, IndexError, KeyError)):
            message = str(exc) if isinstance(exc, ValueError) else "수정 대상을 찾을 수 없습니다. 문서를 다시 열어 주세요."
        else:
            message = "문서를 처리하지 못했습니다. 올바른 HWP·HWPX 파일인지 확인해 주세요."
        output = {"error": message[:1500]}
    atomic_json(result_path, output)


if __name__ == "__main__":
    main()
