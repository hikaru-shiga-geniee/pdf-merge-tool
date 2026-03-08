"""PDF結合ロジック（PyMuPDF）"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

import fitz  # PyMuPDF


def merge_pdfs(
    files: list[dict],
    output_path: str,
    on_progress: Callable[[int, int], Any] | None = None,
) -> None:
    """
    PDFファイルを結合して保存する。

    files: [{"path": str, "rotation": int}, ...]
        rotation: 0, 90, 180, 270
    output_path: 出力先パス
    on_progress: callback(processed_pages, total_pages)
    """
    output_doc = fitz.open()

    # 総ページ数を先に計算
    total_pages = 0
    for f in files:
        doc = fitz.open(f["path"])
        total_pages += doc.page_count
        doc.close()

    processed = 0

    for file_info in files:
        src_doc = fitz.open(file_info["path"])
        for page_num in range(src_doc.page_count):
            output_doc.insert_pdf(src_doc, from_page=page_num, to_page=page_num)
            if file_info["rotation"] != 0:
                page = output_doc[-1]
                page.set_rotation(file_info["rotation"])
            processed += 1
            if on_progress:
                on_progress(processed, total_pages)
        src_doc.close()

    output_doc.save(output_path, deflate=True, garbage=4)
    output_doc.close()


def get_page_count(pdf_path: str) -> int:
    """PDFのページ数を返す。"""
    doc = fitz.open(pdf_path)
    count = doc.page_count
    doc.close()
    return count


def suggest_output_filename(paths: list[str]) -> str:
    """ファイル名の共通部分を抽出して出力ファイル名を提案する。"""
    import os

    if not paths:
        return "output.pdf"

    basenames = [os.path.splitext(os.path.basename(p))[0] for p in paths]

    if len(basenames) == 1:
        return basenames[0] + "_merged.pdf"

    # 共通プレフィックスを抽出
    prefix = os.path.commonprefix(basenames).rstrip("_- ")

    if prefix and len(prefix) >= 3:
        return prefix + "_merged.pdf"

    return "output.pdf"
