"""PDFプレビュー生成（PyMuPDF）"""

import fitz  # PyMuPDF


def render_preview(
    pdf_path: str,
    page_num: int = 0,
    rotation: int = 0,
    max_size: int = 250,
) -> bytes:
    """サムネイル画像（PNG）を生成して返す。"""
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_num)
    rect = page.rect
    scale = min(max_size / rect.width, max_size / rect.height)
    mat = fitz.Matrix(scale, scale).prerotate(rotation)
    pix = page.get_pixmap(matrix=mat)
    img_data = pix.tobytes("png")
    pix = None
    doc.close()
    return img_data
