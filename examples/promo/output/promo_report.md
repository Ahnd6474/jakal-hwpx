# Promo Conversion Bundle

## hwpx2pdf
- Source: `math_template_2_formula_fixed.hwpx`
- Why this sample: table-heavy and visually mixed HWPX with pictures and shapes
- Output: `examples\promo\output\promo_hwpx2pdf_complex_layout.pdf`
- Source tables: 16
- Source pictures: 9
- Source shapes: 56
- Output pages: 2

## pdf2hwpx
- Source: `평가원 수학 양식 (2)-jakal_hwpx_수정.pdf`
- Why this sample: irregular PDF layout with many images and dense non-text structure
- Output: `examples\promo\output\promo_pdf2hwpx_weird_template.hwpx`
- Source pages: 20
- Source page 1 image placements: 11
- Output sections: 20
- Output pictures: 129

## Notes
- `pdf2hwpx` uses supplied OCR-like text blocks for the text layer and imports PDF images into HWPX picture objects.
- `hwpx2pdf` focuses on readable text plus simple rendering for tables, pictures, and shapes.