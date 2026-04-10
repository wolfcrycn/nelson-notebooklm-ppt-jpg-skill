# NotebookLM PPT/PDF Watermark Remover

Remove NotebookLM watermark from PPT/PDF exports and stitch all pages into a single high-resolution image.

## Features

- **PDF** - 300 DPI JPG rendering via PyMuPDF
- **PPTX** - PDF - JPG via LibreOffice (fallback: python-pptx)
- Template matching to locate the NotebookLM logo (bottom-right corner)
- OpenCV Telea inpainting to intelligently remove watermark
- Vertical stitching of all pages into one long image

## Known Parameters (NotebookLM current version)

| Parameter | Value | Note |
|-----------|-------|------|
| Export DPI | 300 | ~5334x3000px |
| Logo position | (4935, 2908) - (5334, 2952) | 399x44px |
| Clean area (+30%) | (4816, 2902) 518x56px | Covers logo + edge residue |
| Inpaint radius | 5px | OpenCV Telea |
| Consistency | 11/12 pages | Page 12 may need manual adjustment |

## Usage

`ash
pip install pymupdf opencv-python numpy pillow

# Basic
python scripts/process.py slides.pdf

# Skip template matching
python scripts/process.py slides.pdf --skip-locate

# Custom clean area
python scripts/process.py slides.pdf --clean-area 4816 2902 518 56
`

## Output

| File | Description |
|------|-------------|
| clean_01.jpg ~ clean_NN.jpg | Per-page cleaned images (300 DPI) |
| clean_all_stitched.jpg | All pages vertically stitched |

## Troubleshooting

- **Logo residue**: Expand clean area: --clean-area 4700 2900 600 80
- **Low match score (< 0.9)**: Use --skip-locate with manual --clean-area
- **Large file (> 25MB)**: Send in batches of 4-6 pages

## Requirements

- Python 3.8+
- PyMuPDF, OpenCV, NumPy, Pillow
- (Recommended) LibreOffice for best PPTX rendering

## License

MIT
