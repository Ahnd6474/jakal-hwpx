from __future__ import annotations

from pathlib import Path

from jakal_hwpx._release_gate import (
    load_release_contract,
    validate_contract_definition,
    validate_corpus_samples,
)


REPO_ROOT = Path(__file__).resolve().parents[1]


def test_release_contract_definition_matches_repo_configuration() -> None:
    contract = load_release_contract(REPO_ROOT)
    errors = validate_contract_definition(contract, REPO_ROOT)
    assert errors == []


def test_release_contract_corpus_samples_are_valid() -> None:
    contract = load_release_contract(REPO_ROOT)
    report = validate_corpus_samples(contract, REPO_ROOT)

    assert report.failure_count == 0
    assert report.missing_categories == []
    assert report.sample_count >= len(contract["required_corpus_categories"])
