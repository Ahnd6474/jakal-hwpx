"""Build the small, repository-owned document used by the web demo."""
from pathlib import Path

from jakal_hwpx import HwpxDocument, HwpDocument


def build():
    doc = HwpxDocument.blank()
    doc.set_metadata(title="문서 작업을 시작해 보세요", creator="Jakal Workspace")
    doc.set_paragraph_text(0, 0, "반가워요. 여기는 나만의 HWPX 워크스페이스입니다.")
    doc.append_paragraph("이 문단을 클릭해 내용을 바꿔보세요. 원본 파일은 별도로 보관됩니다.")
    doc.append_paragraph("문단과 표를 수정한 다음 ‘수정 저장’을 누르고, ‘검증하기’로 저장 결과를 확인하세요.")
    doc.append_table(rows=4, columns=3, cell_texts=[["작업", "내용", "상태"], ["문서 열기", "HWPX 파일 업로드", "완료"], ["내용 수정", "문단과 표의 텍스트 편집", "진행 중"], ["검증 · 다운로드", "결과를 검사하고 수정본 받기", "대기"]])
    doc.append_paragraph("표 탭에서는 각 셀의 텍스트를 바꿀 수 있습니다. 병합된 셀은 기존 구조대로 표시됩니다.")
    path = Path(__file__).resolve().parents[1] / "web" / "public" / "sample.hwpx"
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)
    print(path)
    hwp = HwpDocument.blank()
    hwp.append_paragraph("HWP 문서의 본문을 직접 수정할 수 있습니다.", section_index=0)
    hwp.append_paragraph("고급 설정에서 작업 메모를 저장하거나, 찾기·바꾸기로 여러 내용을 수정해 보세요.", section_index=0)
    hwp.append_table(rows=3, cols=2, cell_texts=[["작업", "상태"], ["문단·표 수정", "진행 중"], ["검증 후 다운로드", "대기"]], section_index=0)
    hwp.save(path.with_suffix(".hwp"))


if __name__ == "__main__":
    build()
