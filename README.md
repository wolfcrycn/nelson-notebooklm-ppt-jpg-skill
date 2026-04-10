# NotebookLM PPT/PDF Watermark Remover

Remove NotebookLM watermark from PPT/PDF exports and stitch all pages into a single high-resolution image.

## Features

- **PDF** - 300 DPI JPG rendering via PyMuPDF
- **PPTX** - PDF - JPG via LibreOffice (fallback: python-pptx)
- Template matching to locate the NotebookLM logo (bottom-right corner)
- OpenCV Telea inpainting to intelligently remove watermark
- Vertical stitching of all pages into one long image

## Known Parameters

| Parameter | Value | Note |
|-----------|-------|------|
| Export DPI | 300 | ~5334x3000px |
| Logo position | (4935, 2908)-(5334, 2952) | 399x44px |
| Clean area (+30%) | (4816, 2902) 518x56px | |
| Inpaint radius | 5px | |

## Usage

`ash
pip install pymupdf opencv-python numpy pillow
python scripts/process.py slides.pdf --skip-locate
`

## Output

- clean_01.jpg ~ clean_NN.jpg (per-page, 300 DPI)
- clean_all_stitched.jpg (all pages stitched)

## License

MIT
