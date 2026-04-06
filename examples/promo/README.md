# Promo Bundle

This folder is for promotion-ready conversion results.

Generated outputs are written to `examples/promo/output/` by:

```powershell
python examples/build_promo_bundle.py
```

Current bundle targets:

- `hwpx2pdf`: a table-heavy or visually complex HWPX sample
- `pdf2hwpx`: an irregular PDF sample with dense non-text structure

The generated folder includes:

- `promo_hwpx2pdf_complex_layout.pdf`
- `promo_pdf2hwpx_weird_template.hwpx`
- `promo_report.md`
- `promo_manifest.json`
