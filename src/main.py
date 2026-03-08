"""PDF結合ツール - メインGUI（CustomTkinter）"""

import os
import threading
import tkinter as tk
from io import BytesIO
from tkinter import filedialog, messagebox

import customtkinter as ctk
from PIL import Image, ImageTk

from pdf_merge import get_page_count, merge_pdfs, suggest_output_filename
from pdf_preview import render_preview

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

PREVIEW_MAX_SIZE = 250


class PdfFileEntry:
    """管理用: 1つのPDFファイルの状態"""

    def __init__(self, path: str):
        self.path = path
        self.filename = os.path.basename(path)
        self.rotation = 0  # 0, 90, 180, 270
        self.page_count = get_page_count(path)
        self.current_page = 0  # プレビュー用


class FileItemWidget(ctk.CTkFrame):
    """ファイルリストの1行分のウィジェット"""

    def __init__(self, master, entry: PdfFileEntry, index: int, app: "App"):
        super().__init__(master, fg_color="transparent")
        self.entry = entry
        self.index = index
        self.app = app
        self._preview_image = None  # GC防止用参照

        self._build_ui()
        self._load_preview()

    def _build_ui(self):
        # --- ファイル情報行 ---
        self._row = ctk.CTkFrame(self, fg_color=("gray95", "gray20"), corner_radius=8)
        self._row.pack(fill="x", padx=4, pady=(4, 0))
        row = self._row

        # ドラッグハンドル + 連番
        grip = ctk.CTkLabel(row, text="≡", width=24, font=("", 16), cursor="hand2")
        grip.pack(side="left", padx=(8, 0))
        ctk.CTkLabel(
            row, text=f"#{self.index + 1}", width=32, font=("", 13, "bold")
        ).pack(side="left", padx=(4, 0))

        # ドラッグ&ドロップ並べ替え用イベント（行全体 + ハンドル）
        for w in [row, grip]:
            w.bind("<ButtonPress-1>", self._on_drag_start)
            w.bind("<B1-Motion>", self._on_drag_motion)
            w.bind("<ButtonRelease-1>", self._on_drag_end)

        # ファイル名
        name_label = ctk.CTkLabel(
            row,
            text=self.entry.filename,
            font=("", 13),
            anchor="w",
        )
        name_label.pack(side="left", padx=(8, 4), fill="x", expand=True)
        # ツールチップ用にフルパスをバインド
        name_label.bind(
            "<Enter>",
            lambda e: self._show_tooltip(e, self.entry.path),
        )
        name_label.bind("<Leave>", lambda e: self._hide_tooltip())

        # ページ数
        ctk.CTkLabel(row, text=f"{self.entry.page_count}p", width=40, font=("", 13)).pack(
            side="left", padx=4
        )

        # 回転コントロール
        ctk.CTkButton(
            row, text="↺左", width=40, height=28, font=("", 12),
            command=lambda: self._rotate(-90),
        ).pack(side="left", padx=2)

        self._angle_label = ctk.CTkLabel(
            row, text=f"{self.entry.rotation}°", width=36, font=("", 12, "bold")
        )
        self._angle_label.pack(side="left")

        ctk.CTkButton(
            row, text="右↻", width=40, height=28, font=("", 12),
            command=lambda: self._rotate(90),
        ).pack(side="left", padx=2)

        # 移動ボタン
        ctk.CTkButton(
            row, text="▲", width=30, height=28, font=("", 12),
            command=lambda: self.app.move_file(self.index, -1),
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            row, text="▼", width=30, height=28, font=("", 12),
            command=lambda: self.app.move_file(self.index, 1),
        ).pack(side="left", padx=2)

        # 削除ボタン
        ctk.CTkButton(
            row,
            text="×",
            width=30,
            height=28,
            font=("", 14),
            fg_color="transparent",
            text_color="red",
            hover_color=("red", "darkred"),
            command=lambda: self.app.remove_file(self.index),
        ).pack(side="left", padx=(2, 8))

        # --- プレビュー領域 ---
        preview_frame = ctk.CTkFrame(self, fg_color=("gray90", "gray25"), corner_radius=0)
        preview_frame.pack(fill="x", padx=4, pady=(0, 4))

        self._preview_label = ctk.CTkLabel(preview_frame, text="")
        self._preview_label.pack(pady=8)

        # ページ送り
        nav = ctk.CTkFrame(preview_frame, fg_color="transparent")
        nav.pack(pady=(0, 8))

        self._prev_btn = ctk.CTkButton(
            nav, text="前へ", width=50, height=26, font=("", 12),
            command=lambda: self._change_page(-1),
        )
        self._prev_btn.pack(side="left", padx=4)

        self._page_label = ctk.CTkLabel(
            nav,
            text=f"{self.entry.current_page + 1} / {self.entry.page_count} ページ",
            font=("", 12),
        )
        self._page_label.pack(side="left", padx=8)

        self._next_btn = ctk.CTkButton(
            nav, text="次へ", width=50, height=26, font=("", 12),
            command=lambda: self._change_page(1),
        )
        self._next_btn.pack(side="left", padx=4)

        self._update_nav_buttons()

    # --- ドラッグ&ドロップ並べ替え ---
    def _on_drag_start(self, event):
        self.app._drag_source_index = self.index
        self.app._dragging = True
        self._row.configure(fg_color=("gray85", "gray30"))

    def _on_drag_motion(self, event):
        if not getattr(self.app, "_dragging", False):
            return
        # マウスのウィンドウ内座標からドロップ先を特定
        target_index = self.app._find_drop_target(event)
        self.app._highlight_drop_target(target_index)

    def _on_drag_end(self, event):
        if not getattr(self.app, "_dragging", False):
            return
        self.app._dragging = False
        target_index = self.app._find_drop_target(event)
        self.app._clear_drop_highlight()
        source_index = self.app._drag_source_index
        if target_index is not None and target_index != source_index:
            # リストの並べ替え
            item = self.app.files.pop(source_index)
            self.app.files.insert(target_index, item)
            self.app._rebuild_file_list()
        else:
            # ドラッグキャンセル: 元の色に戻す
            self._row.configure(fg_color=("gray95", "gray20"))

    def _rotate(self, delta: int):
        self.entry.rotation = (self.entry.rotation + delta) % 360
        self._angle_label.configure(text=f"{self.entry.rotation}°")
        self._load_preview()
        self.app.update_total_pages()

    def _change_page(self, delta: int):
        new_page = self.entry.current_page + delta
        if 0 <= new_page < self.entry.page_count:
            self.entry.current_page = new_page
            self._page_label.configure(
                text=f"{self.entry.current_page + 1} / {self.entry.page_count} ページ"
            )
            self._update_nav_buttons()
            self._load_preview()

    def _update_nav_buttons(self):
        self._prev_btn.configure(
            state="normal" if self.entry.current_page > 0 else "disabled"
        )
        self._next_btn.configure(
            state="normal"
            if self.entry.current_page < self.entry.page_count - 1
            else "disabled"
        )

    def _load_preview(self):
        """バックグラウンドスレッドでプレビューを生成して表示"""

        def _generate():
            try:
                img_data = render_preview(
                    self.entry.path,
                    self.entry.current_page,
                    self.entry.rotation,
                    PREVIEW_MAX_SIZE,
                )
                img = Image.open(BytesIO(img_data))
                # メインスレッドでUI更新
                self.after(0, lambda: self._set_preview_image(img))
            except Exception as e:
                self.after(
                    0,
                    lambda: self._preview_label.configure(
                        text=f"プレビュー生成エラー: {e}", image=None
                    ),
                )

        threading.Thread(target=_generate, daemon=True).start()

    def _set_preview_image(self, img: Image.Image):
        ctk_img = ctk.CTkImage(light_image=img, size=(img.width, img.height))
        self._preview_image = ctk_img  # GC防止
        self._preview_label.configure(image=ctk_img, text="")

    # ツールチップ
    _tooltip_win = None

    def _show_tooltip(self, event, text):
        self._hide_tooltip()
        x = event.x_root + 10
        y = event.y_root + 10
        tw = tk.Toplevel(self)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw, text=text, background="#ffffe0", relief="solid", borderwidth=1, font=("", 11)
        )
        label.pack()
        self._tooltip_win = tw

    def _hide_tooltip(self):
        if self._tooltip_win:
            self._tooltip_win.destroy()
            self._tooltip_win = None


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("PDF結合ツール")
        self.geometry("700x800")
        self.minsize(600, 500)

        self.files: list[PdfFileEntry] = []
        self._dragging = False
        self._drag_source_index: int = -1
        self._drop_highlight_index: int | None = None

        self._build_ui()

    def _build_ui(self):
        # --- メインスクロール領域 ---
        self._main_frame = ctk.CTkScrollableFrame(self, label_text="")
        self._main_frame.pack(fill="both", expand=True, padx=12, pady=12)

        # ドロップゾーン（クリックでファイル選択）
        self._dropzone = ctk.CTkFrame(
            self._main_frame,
            height=120,
            corner_radius=10,
            border_width=2,
            border_color="gray70",
            fg_color=("gray96", "gray18"),
        )
        self._dropzone.pack(fill="x", pady=(0, 12))
        self._dropzone.pack_propagate(False)

        dz_inner = ctk.CTkFrame(self._dropzone, fg_color="transparent")
        dz_inner.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(dz_inner, text="⬆", font=("", 36), text_color="gray60").pack()
        ctk.CTkLabel(
            dz_inner,
            text="PDFファイルをドラッグ&ドロップ",
            font=("", 16, "bold"),
        ).pack()
        self._dropzone_sub = ctk.CTkLabel(
            dz_inner,
            text="または クリックしてファイルを選択",
            font=("", 13),
            text_color="gray50",
        )
        self._dropzone_sub.pack()

        # ドロップゾーンクリックでファイル選択
        for widget in [self._dropzone, dz_inner]:
            widget.bind("<Button-1>", lambda e: self._select_files())
        # 子ウィジェットにもバインド
        for child in dz_inner.winfo_children():
            child.bind("<Button-1>", lambda e: self._select_files())

        # ファイルリスト見出し
        self._file_list_header = ctk.CTkLabel(
            self._main_frame,
            text="選択されたファイル:",
            font=("", 14, "bold"),
            anchor="w",
        )

        # ファイルリストコンテナ
        self._file_list_frame = ctk.CTkFrame(
            self._main_frame, fg_color="transparent"
        )

        # --- フッター ---
        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(fill="x", padx=12, pady=(0, 12))

        self._total_label = ctk.CTkLabel(
            footer, text="", font=("", 13), anchor="w"
        )
        self._total_label.pack(fill="x", pady=(0, 4))

        self._progress = ctk.CTkProgressBar(footer)
        self._progress.pack(fill="x", pady=(0, 4))
        self._progress.set(0)
        self._progress.pack_forget()  # 初期は非表示

        self._progress_label = ctk.CTkLabel(
            footer, text="", font=("", 12), anchor="w"
        )

        self._merge_btn = ctk.CTkButton(
            footer,
            text="結合して保存",
            font=("", 15, "bold"),
            height=40,
            command=self._start_merge,
        )
        self._merge_btn.pack(fill="x")

        # TkinterDnD サポート（利用可能な場合）
        self._setup_dnd()

    def _setup_dnd(self):
        """tkinterdnd2 が利用可能ならD&Dを有効化"""
        try:
            from tkinterdnd2 import DND_FILES, TkinterDnD

            # TkinterDnDをミックスイン的に有効化
            # CustomTkinter は Tk を継承しているので直接drop_target_registerを呼べないため
            # トップレベルにDnDを追加
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_dnd_drop)
        except (ImportError, Exception):
            # tkinterdnd2 が無い場合はD&D非対応（ファイル選択ダイアログのみ）
            pass

    def _on_dnd_drop(self, event):
        """D&Dでファイルがドロップされた時"""
        raw = event.data
        # Windows: {path with spaces} or path
        # macOS: path1 path2 ...
        paths = []
        if "{" in raw:
            import re

            paths = re.findall(r"\{([^}]+)\}", raw)
        else:
            paths = raw.split()

        pdf_paths = [p for p in paths if p.lower().endswith(".pdf")]
        if pdf_paths:
            self._add_files(pdf_paths)

    def _select_files(self):
        paths = filedialog.askopenfilenames(
            title="PDFファイルを選択",
            filetypes=[("PDFファイル", "*.pdf"), ("すべてのファイル", "*.*")],
        )
        if paths:
            self._add_files(list(paths))

    def _add_files(self, paths: list[str]):
        """ファイルを追加（ファイル名ソートで自動整列）- バックグラウンドで読み込み"""
        # 重複除外
        new_paths = [p for p in paths if not any(f.path == p for f in self.files)]
        if not new_paths:
            return

        # UIをロード中状態にする
        self._set_loading_state(True, f"読み込み中... 0 / {len(new_paths)} ファイル")

        def _load():
            new_entries = []
            errors = []
            for i, p in enumerate(new_paths):
                try:
                    entry = PdfFileEntry(p)
                    new_entries.append(entry)
                except Exception as e:
                    errors.append(f"{os.path.basename(p)}: {e}")
                self.after(
                    0,
                    lambda idx=i + 1: self._set_loading_state(
                        True,
                        f"読み込み中... {idx} / {len(new_paths)} ファイル",
                        idx / len(new_paths),
                    ),
                )
            self.after(0, lambda: self._on_files_loaded(new_entries, errors))

        threading.Thread(target=_load, daemon=True).start()

    def _set_loading_state(self, loading: bool, text: str = "", progress: float = -1):
        """ドロップゾーンとフッターの読み込み状態を切り替え"""
        if loading:
            self._dropzone.configure(border_color="gray50")
            self._dropzone_sub.configure(text=text)
            # フッターにプログレスバーを表示
            self._progress.pack(fill="x", pady=(0, 4))
            if progress >= 0:
                self._progress.configure(mode="determinate")
                self._progress.set(progress)
            else:
                self._progress.configure(mode="indeterminate")
                self._progress.set(0)
            self._progress_label.pack(fill="x", pady=(0, 4))
            self._progress_label.configure(text=text)
            self._merge_btn.configure(state="disabled")
        else:
            self._dropzone.configure(border_color="gray70")
            self._dropzone_sub.configure(text="または クリックしてファイルを選択")
            self._progress.pack_forget()
            self._progress_label.pack_forget()
            self._merge_btn.configure(state="normal")

    def _on_files_loaded(self, new_entries: list, errors: list[str]):
        """ファイル読み込み完了時のコールバック"""
        self._set_loading_state(False)

        if errors:
            messagebox.showerror(
                "エラー",
                f"以下のファイルを開けませんでした:\n" + "\n".join(errors),
            )

        if not new_entries:
            return

        self.files.extend(new_entries)
        self.files.sort(key=lambda f: f.filename)
        self._rebuild_file_list()

    def _rebuild_file_list(self):
        """ファイルリストUIを再構築"""
        # 既存ウィジェットを削除
        for child in self._file_list_frame.winfo_children():
            child.destroy()

        if self.files:
            self._file_list_header.pack(fill="x", pady=(0, 4))
            self._file_list_frame.pack(fill="x")

            for i, entry in enumerate(self.files):
                widget = FileItemWidget(self._file_list_frame, entry, i, self)
                widget.pack(fill="x", pady=2)
        else:
            self._file_list_header.pack_forget()
            self._file_list_frame.pack_forget()

        self.update_total_pages()

    def update_total_pages(self):
        if self.files:
            total = sum(f.page_count for f in self.files)
            self._total_label.configure(text=f"結合後: {total}ページ")
        else:
            self._total_label.configure(text="")

    def _find_drop_target(self, event) -> int | None:
        """マウス位置からドロップ先のインデックスを特定"""
        widgets = self._file_list_frame.winfo_children()
        if not widgets:
            return None
        # event座標をルートウィンドウ基準に変換
        mouse_y = event.widget.winfo_rooty() + event.y
        for i, w in enumerate(widgets):
            wy = w.winfo_rooty()
            wh = w.winfo_height()
            # ウィジェットの中央より上か下かで挿入位置を判定
            if mouse_y < wy + wh // 2:
                return i
        return len(widgets) - 1

    def _highlight_drop_target(self, target_index: int | None):
        """ドロップ先を視覚的にハイライト"""
        if target_index == self._drop_highlight_index:
            return
        self._clear_drop_highlight()
        self._drop_highlight_index = target_index
        if target_index is None:
            return
        widgets = self._file_list_frame.winfo_children()
        if 0 <= target_index < len(widgets):
            widget = widgets[target_index]
            if hasattr(widget, "_row") and target_index != self._drag_source_index:
                widget._row.configure(
                    border_width=2, border_color="#3b82f6"
                )

    def _clear_drop_highlight(self):
        """ドロップ先ハイライトをクリア"""
        if self._drop_highlight_index is not None:
            widgets = self._file_list_frame.winfo_children()
            if 0 <= self._drop_highlight_index < len(widgets):
                widget = widgets[self._drop_highlight_index]
                if hasattr(widget, "_row"):
                    widget._row.configure(border_width=0)
        self._drop_highlight_index = None

    def move_file(self, index: int, direction: int):
        new_index = index + direction
        if 0 <= new_index < len(self.files):
            self.files[index], self.files[new_index] = (
                self.files[new_index],
                self.files[index],
            )
            self._rebuild_file_list()

    def remove_file(self, index: int):
        if 0 <= index < len(self.files):
            self.files.pop(index)
            self._rebuild_file_list()

    def _start_merge(self):
        if not self.files:
            messagebox.showwarning("警告", "PDFファイルを追加してください。")
            return

        # 出力ファイル名を推測
        suggested = suggest_output_filename([f.path for f in self.files])

        output_path = filedialog.asksaveasfilename(
            title="保存先を選択",
            defaultextension=".pdf",
            initialfile=suggested,
            filetypes=[("PDFファイル", "*.pdf")],
        )
        if not output_path:
            return

        # UIをロック
        total_pages = sum(f.page_count for f in self.files)
        self._merge_btn.configure(state="disabled", text="結合中...")
        self._progress.pack(fill="x", pady=(0, 4))
        self._progress.configure(mode="determinate")
        self._progress.set(0)
        self._progress_label.pack(fill="x", pady=(0, 4))
        self._progress_label.configure(text=f"結合中... 0 / {total_pages} ページ (0%)")

        file_data = [{"path": f.path, "rotation": f.rotation} for f in self.files]

        def _do_merge():
            try:
                merge_pdfs(file_data, output_path, on_progress=self._on_merge_progress)
                # 保存完了フェーズ
                self.after(
                    0,
                    lambda: (
                        self._progress.set(1.0),
                        self._progress_label.configure(text="保存完了"),
                    ),
                )
                self.after(0, lambda: self._on_merge_complete(output_path))
            except Exception as e:
                self.after(
                    0,
                    lambda: self._on_merge_error(str(e)),
                )

        threading.Thread(target=_do_merge, daemon=True).start()

    def _on_merge_progress(self, processed: int, total: int):
        progress = processed / total if total > 0 else 0
        pct = int(progress * 100)
        self.after(
            0,
            lambda: (
                self._progress.set(progress),
                self._progress_label.configure(
                    text=f"結合中... {processed} / {total} ページ ({pct}%)"
                ),
            ),
        )

    def _on_merge_complete(self, output_path: str):
        self._merge_btn.configure(state="normal", text="結合して保存")
        self._progress.pack_forget()
        self._progress_label.pack_forget()
        messagebox.showinfo(
            "完了",
            f"PDFの結合が完了しました。\n{output_path}",
        )

    def _on_merge_error(self, error_msg: str):
        self._merge_btn.configure(state="normal", text="結合して保存")
        self._progress.pack_forget()
        self._progress_label.pack_forget()
        messagebox.showerror("エラー", f"結合中にエラーが発生しました:\n{error_msg}")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
