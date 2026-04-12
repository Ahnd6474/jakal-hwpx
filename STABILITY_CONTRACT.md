# Stability Contract

This repository now treats stability as an explicit contract, not an implied claim.
The machine-readable source of truth is [`stability_contract.json`](./stability_contract.json).

## Runtime Scope

- Supported Python runtimes: CPython `3.11`, `3.12`, `3.13`
- The contract is enforced against the versions declared in [`pyproject.toml`](./pyproject.toml)

## Format Tiers

Do not treat every format path as the same fidelity level.

| Format path | Guarantee tier | What is blocked by the release gate |
| --- | --- | --- |
| `HWPX` | `strong_roundtrip` | open, edit, save, reopen, control preservation, sample-corpus round-trip, Hancom corpus smoke on the full release profile |
| `HWP` | `constrained_binary_editing` | open, preview/body text extraction, same-length text edits, profile-backed binary editing, reopen after save |
| `HWP <-> HWPX` | `semantic_bridge_roundtrip` | semantic conversion, bridge-backed edits, reopen after save, bridge stability matrix |

## Release-Blocking Success Criteria

These are the hard numerical checks encoded in [`stability_contract.json`](./stability_contract.json).

- `pytest` failures: `0`
- corpus validation failures: `0`
- HWPX stability failures: `0`
- HWP stability failures: `0`
- HWP bridge failures: `0`
- Hancom corpus smoke failures: `0`
- Hancom warning or recovery dialogs on the release corpus: `0`
- layout-signature regressions in the HWPX stability lab: `0`

The repo also records an advisory PDF visual-diff threshold of `0.5%` changed pixels, but that check is not yet wired into the blocking gate because the repository does not currently ship a renderer-backed diff pipeline.

## Corpus Requirements

The release contract requires at least one checked sample for each of these categories:

- `simple_document`
- `multi_section_form`
- `control_rich`
- `complex_layout`
- `hard_case`
- `legacy_binary`

Each sample in [`stability_contract.json`](./stability_contract.json) includes fixed expectations for:

- expected text
- expected title when relevant
- required controls or features
- minimum section count

That keeps the corpus from drifting into "we have samples" without knowing what each sample is supposed to prove.

## Unsupported Scope

The contract is only credible if the non-goals are written down first.

- Byte-identical `HWP <-> HWPX` conversion is not promised.
- Pure-Python HWP writing is not a universal full-fidelity writer for every legacy document.
- Arbitrary unsupported HWP binary controls are not guaranteed unless a donor/profile-backed path covers them.
- Pixel-identical layout parity with Hancom is not promised.
- PDF visual diff remains advisory until the repo has a first-class renderer-backed pipeline.

## Gate Profiles

Two gate profiles are implemented in [`scripts/check_release.py`](./scripts/check_release.py).

- `ci`
  - portable
  - runs on CI without Hancom
  - validates the contract, corpus, pytest suite, HWPX stability lab, HWP stability lab, bridge stability lab, and packaging
- `release`
  - for a Windows machine with Hancom installed
  - runs everything in `ci`
  - also requires Hancom corpus smoke with zero warning or recovery dialogs
  - rejects Hancom-related skips

## Commands

Portable gate:

```bash
python scripts/check_release.py --profile ci
```

Full release gate:

```powershell
python scripts/check_release.py --profile release
```
