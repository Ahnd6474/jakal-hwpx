from pathlib import Path


def document_name(value: str, extension: str) -> str:
    """A display/download filename, never a path inside the storage directory."""
    title = value.strip()
    if title.lower().endswith(('.hwp', '.hwpx')):
        title = Path(title).stem
    if not title or any(not c.isprintable() or c in '/\\:*?"<>|' for c in title):
        raise ValueError('파일 이름에는 경로 구분자나 특수문자를 사용할 수 없습니다.')
    title = title.strip('. ')
    if not title:
        raise ValueError('파일 이름을 입력해 주세요.')
    return title + '.' + extension
