from __future__ import annotations

from ._hancom_models import HancomBlock

# Imported after hancom_document finishes defining shared conversion helpers.
from .hancom_document import (
    NS,
    Memo,
    _build_hwpx_paragraph_block,
    _extract_hwpx_caption_text,
    _extract_hwpx_run_node,
    _extract_missing_hwpx_descendant_blocks,
    _hwpx_node_key,
    _hwpx_owned_inline_text,
    _mark_hwpx_source_paragraph,
    _memo_requires_hwpx_carrier,
    _merge_hwpx_memos_with_carriers,
    _sort_hwpx_blocks_by_source_order,
    _source_order_for_element,
    etree,
)


def _extract_hwpx_paragraph_blocks(section, paragraph: etree._Element, paragraph_index: int) -> list[HancomBlock]:
    blocks: list[HancomBlock] = []
    memo_blocks: list[Memo] = []
    memo_carriers: list[Memo] = []
    skipped_field_id: str | None = None
    seen_nodes: set[str] = set()

    for run in paragraph.xpath("./hp:run", namespaces=NS):
        for child in list(run):
            local_name = etree.QName(child).localname
            if local_name == "t":
                if skipped_field_id is None:
                    text = (child.text or "").replace("\r", "")
                    if text:
                        blocks.append(
                            _mark_hwpx_source_paragraph(
                                _build_hwpx_paragraph_block(paragraph, text, char_pr_id=run.get("charPrIDRef")),
                                paragraph_index,
                                order=_source_order_for_element(paragraph, child),
                            )
                        )
                continue

            candidates = list(child) if local_name == "ctrl" else [child]
            for node in candidates:
                node_local_name = etree.QName(node).localname
                if skipped_field_id is not None:
                    if node_local_name == "fieldEnd" and node.get("beginIDRef") == skipped_field_id:
                        skipped_field_id = None
                    continue

                extracted = _extract_hwpx_run_node(section, node)
                if extracted is None:
                    continue
                seen_nodes.add(_hwpx_node_key(node))
                caption_text = _extract_hwpx_caption_text(node)
                if caption_text:
                    blocks.append(
                        _mark_hwpx_source_paragraph(
                            _build_hwpx_paragraph_block(paragraph, caption_text, char_pr_id=run.get("charPrIDRef")),
                            paragraph_index,
                            order=_source_order_for_element(paragraph, node),
                        )
                    )
                extracted_blocks = extracted if isinstance(extracted, list) else [extracted]
                for block in extracted_blocks:
                    owned_text = _hwpx_owned_inline_text(block)
                    if owned_text:
                        blocks.append(
                            _mark_hwpx_source_paragraph(
                                _build_hwpx_paragraph_block(paragraph, owned_text, char_pr_id=run.get("charPrIDRef")),
                                paragraph_index,
                                order=_source_order_for_element(paragraph, node),
                            )
                        )
                    _mark_hwpx_source_paragraph(block, paragraph_index, order=_source_order_for_element(paragraph, node))
                    if isinstance(block, Memo) and _memo_requires_hwpx_carrier(block):
                        memo_carriers.append(block)
                    elif isinstance(block, Memo):
                        memo_blocks.append(block)
                    else:
                        blocks.append(block)

                if node_local_name == "fieldBegin":
                    skipped_field_id = node.get("id") or node.get("fieldid")

    if memo_blocks or memo_carriers:
        for memo in _merge_hwpx_memos_with_carriers(memo_blocks, memo_carriers):
            blocks.append(_mark_hwpx_source_paragraph(memo, paragraph_index))
    blocks.extend(_extract_missing_hwpx_descendant_blocks(section, paragraph, paragraph_index, seen_nodes))
    return _sort_hwpx_blocks_by_source_order(blocks)

_extract_hwpx_paragraph_blocks.__module__ = "jakal_hwpx.hancom_document"

