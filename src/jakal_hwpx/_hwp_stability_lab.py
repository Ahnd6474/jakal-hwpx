from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Callable

from ._hancom import convert_document
from .exceptions import HancomInteropError
from .hwp_binary import TAG_CTRL_HEADER, HwpBinaryDocument
from .hwp_document import HwpDocument
from .hwp_pure_profile import HwpPureProfile


def _control_ids(document: HwpDocument, section_index: int) -> list[str]:
    result: list[str] = []
    for record in document.binary_document().section_records(section_index):
        if record.tag_id == TAG_CTRL_HEADER and len(record.payload) >= 4:
            result.append(record.payload[:4][::-1].decode("latin1", errors="replace"))
    return result


@dataclass(frozen=True)
class HwpStabilityCase:
    name: str
    mutate: Callable[[HwpDocument], None]
    expected_texts: tuple[str, ...] = ()
    expected_control_ids: tuple[str, ...] = ()
    min_bindata_delta: int = 0
    min_section_record_delta: int = 0


@dataclass(frozen=True)
class HwpStabilityCaseResult:
    name: str
    ok: bool
    binary_ok: bool
    binary_errors: list[str]
    hancom_status: str
    hancom_ok: bool | None
    hancom_errors: list[str]
    hancom_artifacts: list[str]
    text_snapshot: str
    control_ids: list[str]
    bindata_count: int
    section_record_count: int
    reopened_text_snapshot: str
    reopened_control_ids: list[str]
    reopened_bindata_count: int
    reopened_section_record_count: int


def _validate_with_hancom(case_dir: Path, edited_path: Path) -> tuple[str, bool | None, list[str], list[str]]:
    artifacts: list[str] = []
    errors: list[str] = []
    try:
        hwpx_path = case_dir / f"{edited_path.stem}_hancom.hwpx"
        convert_document(edited_path, hwpx_path, "HWPX")
        artifacts.append(str(hwpx_path))

        roundtrip_hwp_path = case_dir / f"{edited_path.stem}_hancom_roundtrip.hwp"
        convert_document(hwpx_path, roundtrip_hwp_path, "HWP")
        artifacts.append(str(roundtrip_hwp_path))
    except HancomInteropError as error:
        message = str(error)
        lowered = message.lower()
        if "security module" in lowered or "comobject" in lowered or "com conversion failed" in lowered or "activeobject" in lowered:
            return ("skipped", None, [message], artifacts)
        return ("failed", False, [message], artifacts)
    return ("passed", True, errors, artifacts)


def _run_case(case: HwpStabilityCase, output_dir: Path, *, validate_with_hancom: bool) -> HwpStabilityCaseResult:
    document = HwpDocument.blank()
    profile = HwpPureProfile.load_bundled()
    section_index = profile.target_section_index
    binary = document.binary_document()

    before_bindata_count = binary.docinfo_model().id_mappings_record().bin_data_count
    before_section_record_count = len(binary.section_records(section_index))

    baseline_path = output_dir / f"{case.name}_baseline.hwp"
    document.save(baseline_path)
    baseline_reopened = HwpDocument.open(baseline_path)
    binary_errors: list[str] = []
    if not baseline_reopened.list_stream_paths():
        binary_errors.append("baseline reopen did not expose any streams")

    case.mutate(document)

    mutated_text = document.get_document_text()
    mutated_control_ids = _control_ids(document, section_index)
    mutated_bindata_count = document.binary_document().docinfo_model().id_mappings_record().bin_data_count
    mutated_section_record_count = len(document.binary_document().section_records(section_index))

    for expected_text in case.expected_texts:
        if expected_text not in mutated_text:
            binary_errors.append(f"missing mutated text: {expected_text}")
    for expected_control_id in case.expected_control_ids:
        if expected_control_id not in mutated_control_ids:
            binary_errors.append(f"missing mutated control: {expected_control_id}")
    if mutated_bindata_count - before_bindata_count < case.min_bindata_delta:
        binary_errors.append(
            f"bindata delta {mutated_bindata_count - before_bindata_count} was smaller than expected {case.min_bindata_delta}"
        )
    if mutated_section_record_count - before_section_record_count < case.min_section_record_delta:
        binary_errors.append(
            "section record growth "
            f"{mutated_section_record_count - before_section_record_count} was smaller than expected {case.min_section_record_delta}"
        )

    edited_path = output_dir / f"{case.name}_edited.hwp"
    document.save(edited_path)
    reopened = HwpDocument.open(edited_path)
    reopened_text = reopened.get_document_text()
    reopened_control_ids = _control_ids(reopened, section_index)
    reopened_bindata_count = reopened.binary_document().docinfo_model().id_mappings_record().bin_data_count
    reopened_section_record_count = len(reopened.binary_document().section_records(section_index))

    for expected_text in case.expected_texts:
        if expected_text not in reopened_text:
            binary_errors.append(f"missing reopened text: {expected_text}")
    for expected_control_id in case.expected_control_ids:
        if expected_control_id not in reopened_control_ids:
            binary_errors.append(f"missing reopened control: {expected_control_id}")
    if reopened_bindata_count != mutated_bindata_count:
        binary_errors.append(f"reopened bindata count {reopened_bindata_count} != mutated {mutated_bindata_count}")
    if reopened_section_record_count != mutated_section_record_count:
        binary_errors.append(f"reopened section record count {reopened_section_record_count} != mutated {mutated_section_record_count}")

    if validate_with_hancom:
        hancom_status, hancom_ok, hancom_errors, hancom_artifacts = _validate_with_hancom(output_dir, edited_path)
    else:
        hancom_status, hancom_ok, hancom_errors, hancom_artifacts = ("skipped", None, [], [])

    binary_ok = not binary_errors
    ok = binary_ok and hancom_status != "failed"

    return HwpStabilityCaseResult(
        name=case.name,
        ok=ok,
        binary_ok=binary_ok,
        binary_errors=binary_errors,
        hancom_status=hancom_status,
        hancom_ok=hancom_ok,
        hancom_errors=hancom_errors,
        hancom_artifacts=hancom_artifacts,
        text_snapshot=mutated_text,
        control_ids=mutated_control_ids,
        bindata_count=mutated_bindata_count,
        section_record_count=mutated_section_record_count,
        reopened_text_snapshot=reopened_text,
        reopened_control_ids=reopened_control_ids,
        reopened_bindata_count=reopened_bindata_count,
        reopened_section_record_count=reopened_section_record_count,
    )


def _default_picture_bytes() -> bytes:
    document = HwpDocument.blank()
    return document.binary_document().read_stream("BinData/BIN0001.bmp", decompress=False)


def _bundled_picture_bytes(stream_path: str) -> bytes:
    document = HwpDocument.blank()
    return document.binary_document().read_stream(stream_path, decompress=False)


def _matrix(prefix: str, rows: int, cols: int) -> list[list[str]]:
    return [[f"{prefix}{row + 1}{col + 1}" for col in range(cols)] for row in range(rows)]


def _append_paragraph_case(name: str, text: str) -> HwpStabilityCase:
    return HwpStabilityCase(
        name=name,
        mutate=lambda document, value=text: document.append_paragraph(value, section_index=1),
        expected_texts=(text,),
        min_section_record_delta=2,
    )


def _replace_case(name: str, paragraph_text: str, old: str, new: str) -> HwpStabilityCase:
    expected = paragraph_text.replace(old, new)
    return HwpStabilityCase(
        name=name,
        mutate=lambda document, paragraph_text=paragraph_text, old=old, new=new: (
            document.append_paragraph(paragraph_text, section_index=1),
            document.replace_text_same_length(old, new, section_index=1),
        ),
        expected_texts=(expected,),
        min_section_record_delta=2,
    )


def _table_case(
    name: str,
    *,
    rows: int,
    cols: int,
    cell_texts: list[list[str]] | list[str],
    row_heights: list[int] | None = None,
    col_widths: list[int] | None = None,
    cell_spans: dict[tuple[int, int], tuple[int, int]] | None = None,
    cell_border_fill_ids: dict[tuple[int, int], int] | None = None,
    table_border_fill_id: int = 1,
    expected_texts: tuple[str, ...] | None = None,
) -> HwpStabilityCase:
    if expected_texts is None:
        flat_expected: list[str] = []
        if cell_texts and isinstance(cell_texts[0], list):
            for row in cell_texts:  # type: ignore[assignment]
                flat_expected.extend(value for value in row if value)
        else:
            flat_expected.extend(value for value in cell_texts if value)  # type: ignore[arg-type]
        expected_texts = tuple(flat_expected[:4] if len(flat_expected) > 4 else flat_expected)
    min_delta = 8 if rows * cols == 1 else max(12, 6 + 4 * rows)
    return HwpStabilityCase(
        name=name,
        mutate=lambda document,
        rows=rows,
        cols=cols,
        cell_texts=cell_texts,
        row_heights=row_heights,
        col_widths=col_widths,
        cell_spans=cell_spans,
        cell_border_fill_ids=cell_border_fill_ids,
        table_border_fill_id=table_border_fill_id: document.append_table(
            rows=rows,
            cols=cols,
            cell_texts=cell_texts,
            row_heights=row_heights,
            col_widths=col_widths,
            cell_spans=cell_spans,
            cell_border_fill_ids=cell_border_fill_ids,
            table_border_fill_id=table_border_fill_id,
        ),
        expected_texts=expected_texts,
        expected_control_ids=("tbl ",),
        min_section_record_delta=min_delta,
    )


def _hyperlink_case(
    name: str,
    *,
    url: str,
    text: str | None = None,
    metadata_fields: list[str | int] | None = None,
) -> HwpStabilityCase:
    display_text = text or url
    return HwpStabilityCase(
        name=name,
        mutate=lambda document, url=url, text=text, metadata_fields=metadata_fields: document.append_hyperlink(
            url,
            text=text,
            metadata_fields=metadata_fields,
        ),
        expected_texts=(display_text,),
        expected_control_ids=("%hlk",),
        min_section_record_delta=5,
    )


def _picture_case(
    name: str,
    *,
    pictures: list[tuple[bytes | None, str | None]],
    min_bindata_delta: int = 0,
) -> HwpStabilityCase:
    min_section_delta = 7 * len(pictures)
    return HwpStabilityCase(
        name=name,
        mutate=lambda document, pictures=pictures: [
            document.append_picture(image_bytes, extension=extension) for image_bytes, extension in pictures
        ],
        expected_control_ids=("gso ",),
        min_bindata_delta=min_bindata_delta,
        min_section_record_delta=min_section_delta,
    )


def _combo_case(
    name: str,
    *,
    expected_texts: tuple[str, ...],
    expected_control_ids: tuple[str, ...],
    min_section_record_delta: int,
    min_bindata_delta: int = 0,
    mutate: Callable[[HwpDocument], None],
) -> HwpStabilityCase:
    return HwpStabilityCase(
        name=name,
        mutate=mutate,
        expected_texts=expected_texts,
        expected_control_ids=expected_control_ids,
        min_bindata_delta=min_bindata_delta,
        min_section_record_delta=min_section_record_delta,
    )


def _cases() -> list[HwpStabilityCase]:
    bmp_bytes = _default_picture_bytes()
    png_bytes = _bundled_picture_bytes("BinData/BIN0004.png")
    jpg_bytes = _bundled_picture_bytes("BinData/BIN0002.jpg")

    cases: list[HwpStabilityCase] = []

    paragraph_texts = (
        "PURE-HWP-STABILITY-A",
        "PURE-HWP-STABILITY-B",
        "PURE-HWP-STABILITY-C",
        "PURE-HWP-STABILITY-D",
        "PURE-HWP-STABILITY-E",
        "PURE-HWP-STABILITY-F",
    )
    for index, text in enumerate(paragraph_texts, 1):
        cases.append(_append_paragraph_case(f"paragraph_append_{index:02d}", text))

    replace_specs = (
        ("ALPHA-1234", "ALPHA", "OMEGA"),
        ("BRAVO-5678", "BRAVO", "DELTA"),
        ("FOCUS-2024", "FOCUS", "DRIVE"),
        ("STATE-9090", "STATE", "PHASE"),
        ("POINT-6060", "POINT", "TRACE"),
        ("LEVEL-4040", "LEVEL", "MODEL"),
    )
    for index, (paragraph_text, old, new) in enumerate(replace_specs, 1):
        cases.append(_replace_case(f"paragraph_replace_{index:02d}", paragraph_text, old, new))

    basic_table_specs = (
        (1, 1, _matrix("TB11-", 1, 1)),
        (1, 2, _matrix("TB12-", 1, 2)),
        (2, 1, _matrix("TB21-", 2, 1)),
        (2, 2, _matrix("TB22-", 2, 2)),
        (2, 3, _matrix("TB23-", 2, 3)),
        (3, 2, _matrix("TB32-", 3, 2)),
    )
    for index, (rows, cols, matrix) in enumerate(basic_table_specs, 1):
        cases.append(_table_case(f"table_basic_matrix_{index:02d}", rows=rows, cols=cols, cell_texts=matrix))
        flat = [value for row in matrix for value in row]
        cases.append(_table_case(f"table_basic_flat_{index:02d}", rows=rows, cols=cols, cell_texts=flat))

    geometry_cases = (
        dict(name="table_geometry_01", rows=2, cols=3, cell_texts=[["G11", "G12", "G13"], ["G21", "G22", "G23"]], row_heights=[101, 202]),
        dict(name="table_geometry_02", rows=3, cols=2, cell_texts=[["G31", "G32"], ["G33", "G34"], ["G35", "G36"]], row_heights=[120, 240, 360]),
        dict(name="table_geometry_03", rows=2, cols=3, cell_texts=[["W11", "W12", "W13"], ["W21", "W22", "W23"]], col_widths=[1000, 1500, 2000]),
        dict(
            name="table_geometry_04",
            rows=3,
            cols=3,
            cell_texts=[["B11", "B12", "B13"], ["B21", "B22", "B23"], ["B31", "B32", "B33"]],
            cell_border_fill_ids={(0, 0): 7, (1, 1): 8, (2, 2): 9},
            table_border_fill_id=6,
        ),
        dict(
            name="table_geometry_05",
            rows=2,
            cols=2,
            cell_texts=[["F11", "F12"], ["F21", "F22"]],
            cell_border_fill_ids={(0, 0): 7, (0, 1): 8, (1, 0): 9, (1, 1): 10},
            table_border_fill_id=5,
        ),
        dict(
            name="table_geometry_06",
            rows=2,
            cols=3,
            cell_texts=[["S11", "", "S13"], ["S21", "S22", "S23"]],
            cell_spans={(0, 0): (1, 2)},
        ),
        dict(
            name="table_geometry_07",
            rows=3,
            cols=2,
            cell_texts=[["V11", "V12"], ["", "V22"], ["V31", "V32"]],
            cell_spans={(0, 0): (2, 1)},
        ),
        dict(
            name="table_geometry_08",
            rows=3,
            cols=3,
            cell_texts=[["X11", "", "X13"], ["X21", "X22", ""], ["X31", "", ""]],
            row_heights=[120, 240, 360],
            col_widths=[900, 1300, 1700],
            cell_spans={(0, 0): (1, 2), (1, 1): (2, 2)},
            cell_border_fill_ids={(0, 0): 7, (1, 1): 8},
            table_border_fill_id=6,
        ),
        dict(
            name="table_geometry_09",
            rows=2,
            cols=4,
            cell_texts=[["Y11", "Y12", "Y13", "Y14"], ["Y21", "Y22", "Y23", "Y24"]],
            col_widths=[700, 800, 900, 1000],
        ),
        dict(
            name="table_geometry_10",
            rows=4,
            cols=2,
            cell_texts=[["Z11", "Z12"], ["Z21", "Z22"], ["Z31", "Z32"], ["Z41", "Z42"]],
            row_heights=[111, 222, 333, 444],
        ),
        dict(
            name="table_geometry_11",
            rows=3,
            cols=3,
            cell_texts=[["M11", "M12", "M13"], ["M21", "", "M23"], ["M31", "", "M33"]],
            cell_spans={(1, 1): (2, 1)},
            cell_border_fill_ids={(1, 1): 11},
            table_border_fill_id=5,
        ),
        dict(
            name="table_geometry_12",
            rows=3,
            cols=4,
            cell_texts=[["N11", "", "N13", "N14"], ["N21", "N22", "N23", "N24"], ["N31", "N32", "N33", "N34"]],
            cell_spans={(0, 0): (1, 2)},
            row_heights=[100, 200, 300],
            col_widths=[600, 700, 800, 900],
        ),
    )
    for spec in geometry_cases:
        cases.append(_table_case(**spec))

    hyperlink_specs = (
        dict(name="hyperlink_01", url="https://example.com/hwp-01", text="HWP-LINK-01"),
        dict(name="hyperlink_02", url="https://example.com/hwp-02", text="HWP-LINK-02"),
        dict(name="hyperlink_03", url="https://example.com/hwp-03", text="HWP-LINK-03"),
        dict(name="hyperlink_04", url="mailto:test@example.com", text="MAIL-LINK"),
        dict(name="hyperlink_05", url="https://example.com/hwp-05", text="META-A", metadata_fields=[1, 0, "anchorA"]),
        dict(name="hyperlink_06", url="https://example.com/hwp-06", text="META-B", metadata_fields=[2, 1, "anchorB"]),
        dict(name="hyperlink_07", url="https://example.com/hwp-07", text="META-C", metadata_fields=[3, 2, "anchorC"]),
        dict(name="hyperlink_08", url="https://example.com/hwp-08", text="META-D", metadata_fields=[4, 3, "anchorD"]),
        dict(name="hyperlink_09", url="https://example.com/hwp-09?x=1", text="QUERY-LINK"),
        dict(name="hyperlink_10", url="https://example.com/hwp-10#frag", text="FRAG-LINK", metadata_fields=[5, 4, "frag"]),
        dict(name="hyperlink_11", url="ftp://example.com/file", text="FTP-LINK"),
        dict(name="hyperlink_12", url="https://example.com/hwp-12", text="META-E", metadata_fields=[9, 7, "bookmark"]),
    )
    for spec in hyperlink_specs:
        cases.append(_hyperlink_case(**spec))

    picture_specs = (
        dict(name="picture_01", pictures=[(None, None)], min_bindata_delta=0),
        dict(name="picture_02", pictures=[(None, None), (None, None)], min_bindata_delta=0),
        dict(name="picture_03", pictures=[(bmp_bytes, "bmp")], min_bindata_delta=1),
        dict(name="picture_04", pictures=[(png_bytes, "png")], min_bindata_delta=1),
        dict(name="picture_05", pictures=[(jpg_bytes, "jpg")], min_bindata_delta=1),
        dict(name="picture_06", pictures=[(None, None), (bmp_bytes, "bmp")], min_bindata_delta=1),
        dict(name="picture_07", pictures=[(bmp_bytes, "bmp"), (png_bytes, "png")], min_bindata_delta=2),
        dict(name="picture_08", pictures=[(None, None), (None, None), (jpg_bytes, "jpg")], min_bindata_delta=1),
    )
    for spec in picture_specs:
        cases.append(_picture_case(**spec))

    combo_cases = (
        _combo_case(
            "combo_01",
            expected_texts=("CPH1", "CPH-LINK-1"),
            expected_control_ids=("%hlk",),
            min_section_record_delta=7,
            mutate=lambda document: (
                document.append_paragraph("CPH1", section_index=1),
                document.append_hyperlink("https://example.com/cph1", text="CPH-LINK-1"),
            ),
        ),
        _combo_case(
            "combo_02",
            expected_texts=("CPH2", "CPH-LINK-2"),
            expected_control_ids=("%hlk",),
            min_section_record_delta=7,
            mutate=lambda document: (
                document.append_paragraph("CPH2", section_index=1),
                document.append_hyperlink("https://example.com/cph2", text="CPH-LINK-2", metadata_fields=[1, 2, "cph2"]),
            ),
        ),
        _combo_case(
            "combo_03",
            expected_texts=("CPT11", "CPT12", "CPTP"),
            expected_control_ids=("tbl ",),
            min_section_record_delta=12,
            mutate=lambda document: (
                document.append_paragraph("CPTP", section_index=1),
                document.append_table(rows=1, cols=2, cell_texts=[["CPT11", "CPT12"]]),
            ),
        ),
        _combo_case(
            "combo_04",
            expected_texts=("CPT21", "CPT22", "CPT23"),
            expected_control_ids=("tbl ",),
            min_section_record_delta=16,
            mutate=lambda document: document.append_table(
                rows=2,
                cols=2,
                cell_texts=[["CPT21", "CPT22"], ["CPT23", ""]],
                cell_spans={(1, 0): (1, 2)},
            ),
        ),
        _combo_case(
            "combo_05",
            expected_texts=("CHP-LINK-1",),
            expected_control_ids=("%hlk", "gso "),
            min_section_record_delta=12,
            min_bindata_delta=1,
            mutate=lambda document: (
                document.append_hyperlink("https://example.com/chp1", text="CHP-LINK-1"),
                document.append_picture(bmp_bytes, extension="bmp"),
            ),
        ),
        _combo_case(
            "combo_06",
            expected_texts=("CHP-LINK-2",),
            expected_control_ids=("%hlk", "gso "),
            min_section_record_delta=12,
            mutate=lambda document: (
                document.append_hyperlink("https://example.com/chp2", text="CHP-LINK-2", metadata_fields=[3, 4, "mix"]),
                document.append_picture(),
            ),
        ),
        _combo_case(
            "combo_07",
            expected_texts=("CTH11", "CTH12", "CTH-LINK-1"),
            expected_control_ids=("tbl ", "%hlk"),
            min_section_record_delta=14,
            mutate=lambda document: (
                document.append_table(rows=1, cols=2, cell_texts=[["CTH11", "CTH12"]]),
                document.append_hyperlink("https://example.com/cth1", text="CTH-LINK-1"),
            ),
        ),
        _combo_case(
            "combo_08",
            expected_texts=("CTH21", "CTH22", "CTH23", "CTH-LINK-2"),
            expected_control_ids=("tbl ", "%hlk"),
            min_section_record_delta=18,
            mutate=lambda document: (
                document.append_table(
                    rows=2,
                    cols=2,
                    cell_texts=[["CTH21", "CTH22"], ["CTH23", ""]],
                    cell_spans={(1, 0): (1, 2)},
                ),
                document.append_hyperlink("https://example.com/cth2", text="CTH-LINK-2", metadata_fields=[7, 8, "cth2"]),
            ),
        ),
        _combo_case(
            "combo_09",
            expected_texts=("CTP11", "CTP12"),
            expected_control_ids=("tbl ", "gso "),
            min_section_record_delta=15,
            mutate=lambda document: (
                document.append_table(rows=1, cols=2, cell_texts=[["CTP11", "CTP12"]]),
                document.append_picture(),
            ),
        ),
        _combo_case(
            "combo_10",
            expected_texts=("CTP21", "CTP22"),
            expected_control_ids=("tbl ", "gso "),
            min_section_record_delta=20,
            min_bindata_delta=1,
            mutate=lambda document: (
                document.append_table(rows=2, cols=1, cell_texts=[["CTP21"], ["CTP22"]], row_heights=[150, 250]),
                document.append_picture(png_bytes, extension="png"),
            ),
        ),
        _combo_case(
            "combo_11",
            expected_texts=("CPHP", "CPH-LINK-3"),
            expected_control_ids=("%hlk", "gso "),
            min_section_record_delta=14,
            mutate=lambda document: (
                document.append_paragraph("CPHP", section_index=1),
                document.append_hyperlink("https://example.com/cphp", text="CPH-LINK-3"),
                document.append_picture(),
            ),
        ),
        _combo_case(
            "combo_12",
            expected_texts=("CPHP2", "CPH-LINK-4"),
            expected_control_ids=("%hlk", "gso "),
            min_section_record_delta=14,
            min_bindata_delta=1,
            mutate=lambda document: (
                document.append_paragraph("CPHP2", section_index=1),
                document.append_hyperlink("https://example.com/cphp2", text="CPH-LINK-4", metadata_fields=[5, 6, "cphp2"]),
                document.append_picture(jpg_bytes, extension="jpg"),
            ),
        ),
        _combo_case(
            "combo_13",
            expected_texts=("FULL11", "FULL12", "FULL-LINK-1"),
            expected_control_ids=("tbl ", "%hlk", "gso "),
            min_section_record_delta=22,
            mutate=lambda document: (
                document.append_table(rows=1, cols=2, cell_texts=[["FULL11", "FULL12"]]),
                document.append_hyperlink("https://example.com/full1", text="FULL-LINK-1"),
                document.append_picture(),
            ),
        ),
        _combo_case(
            "combo_14",
            expected_texts=("FULL21", "FULL22", "FULL-LINK-2"),
            expected_control_ids=("tbl ", "%hlk", "gso "),
            min_section_record_delta=24,
            min_bindata_delta=1,
            mutate=lambda document: (
                document.append_table(
                    rows=2,
                    cols=2,
                    cell_texts=[["FULL21", "FULL22"], ["FULL23", "FULL24"]],
                    cell_border_fill_ids={(0, 0): 7, (1, 1): 8},
                ),
                document.append_hyperlink("https://example.com/full2", text="FULL-LINK-2", metadata_fields=[2, 3, "full2"]),
                document.append_picture(bmp_bytes, extension="bmp"),
            ),
        ),
        _combo_case(
            "combo_15",
            expected_texts=("FULL31", "FULL33", "FULL-LINK-3"),
            expected_control_ids=("tbl ", "%hlk", "gso "),
            min_section_record_delta=26,
            min_bindata_delta=1,
            mutate=lambda document: (
                document.append_table(
                    rows=2,
                    cols=3,
                    cell_texts=[["FULL31", "", "FULL33"], ["FULL34", "FULL35", "FULL36"]],
                    cell_spans={(0, 0): (1, 2)},
                    row_heights=[170, 270],
                    col_widths=[900, 1100, 1300],
                ),
                document.append_hyperlink("https://example.com/full3", text="FULL-LINK-3"),
                document.append_picture(png_bytes, extension="png"),
            ),
        ),
        _combo_case(
            "combo_16",
            expected_texts=("FULL41", "FULL42", "FULL-LINK-4", "FULL-P"),
            expected_control_ids=("tbl ", "%hlk", "gso "),
            min_section_record_delta=26,
            min_bindata_delta=1,
            mutate=lambda document: (
                document.append_paragraph("FULL-P", section_index=1),
                document.append_table(rows=1, cols=2, cell_texts=[["FULL41", "FULL42"]]),
                document.append_hyperlink("https://example.com/full4", text="FULL-LINK-4", metadata_fields=[8, 9, "full4"]),
                document.append_picture(jpg_bytes, extension="jpg"),
            ),
        ),
    )
    cases.extend(combo_cases)

    assert len(cases) == 72, f"Expected 72 HWP stability cases, got {len(cases)}"
    return cases


def run_hwp_stability_matrix(output_dir: str | Path, *, validate_with_hancom: bool = False) -> list[HwpStabilityCaseResult]:
    root = Path(output_dir).expanduser().resolve()
    root.mkdir(parents=True, exist_ok=True)
    results: list[HwpStabilityCaseResult] = []
    for case in _cases():
        case_dir = root / case.name
        case_dir.mkdir(parents=True, exist_ok=True)
        results.append(_run_case(case, case_dir, validate_with_hancom=validate_with_hancom))
    return results


def write_hwp_stability_report(results: list[HwpStabilityCaseResult], output_path: str | Path) -> Path:
    target = Path(output_path).expanduser().resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps([asdict(result) for result in results], ensure_ascii=False, indent=2), encoding="utf-8")
    return target
