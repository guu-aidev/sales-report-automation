"""
売上レポート ランチャー

起動方法:
    gui_launcher.pyw をダブルクリック（またはpython gui_launcher.pyw）

依存: pip install customtkinter
"""
import os
import sys
import threading
from pathlib import Path
from tkinter import messagebox

import customtkinter as ctk

from app_config import load_config, save_config
from generate_report import (
    DB_PATH,
    build_base_workbook,
    get_stores_from_db,
    get_years_from_db,
    import_csv_to_db,
    update_store_sheet,
)

_BASE       = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent.parent
INPUT_DIR   = _BASE / "data" / "sample"
OUTPUT_PATH = _BASE / "output" / "monthly_report.xlsx"

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

_GRAY  = "#94a3b8"
_GREEN = "#16a34a"
_RED   = "#dc2626"
_W     = 268   # コンボボックス幅（ウィンドウ幅 340 - 余白 72）


class SalesReportApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("売上レポート")
        self.resizable(False, False)
        self._build_ui()
        threading.Thread(target=self._startup, daemon=True).start()

    # ── UI構築 ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.configure(fg_color="#f1f5f9")

        # ── カード ──
        card = ctk.CTkFrame(self, fg_color="white", corner_radius=16)
        card.pack(padx=20, pady=20)

        # ── ヘッダー ──
        header = ctk.CTkFrame(card, fg_color="#1d4ed8", corner_radius=12, width=_W)
        header.pack(padx=16, pady=(16, 0))
        header.pack_propagate(False)

        h_inner = ctk.CTkFrame(header, fg_color="transparent")
        h_inner.pack(fill="x", padx=20, pady=20)

        icon_bg = ctk.CTkFrame(h_inner, fg_color="#1e40af", corner_radius=8, width=40, height=40)
        icon_bg.pack(anchor="w")
        icon_bg.pack_propagate(False)
        for x, y in [(8, 8), (22, 8), (8, 22), (22, 22)]:
            ctk.CTkFrame(icon_bg, fg_color="white", width=10, height=10, corner_radius=2).place(x=x, y=y)

        ctk.CTkLabel(
            h_inner, text="売上レポート",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white", anchor="w",
        ).pack(fill="x", pady=(10, 0))
        ctk.CTkLabel(
            h_inner, text="ダッシュボード",
            font=ctk.CTkFont(size=12), text_color="#bfdbfe", anchor="w",
        ).pack(fill="x")

        # ── フォーム ──
        form = ctk.CTkFrame(card, fg_color="transparent")
        form.pack(padx=16, pady=(20, 0), anchor="w")

        ctk.CTkLabel(
            form, text="年度",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#374151", anchor="w", width=_W,
        ).pack(anchor="w", pady=(0, 6))
        self.year_cb = ctk.CTkComboBox(
            form, values=["--"], state="disabled", width=_W,
            command=self._on_year_change,
        )
        self.year_cb.pack()

        ctk.CTkLabel(
            form, text="店舗",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#374151", anchor="w", width=_W,
        ).pack(anchor="w", pady=(16, 6))
        self.store_cb = ctk.CTkComboBox(
            form, values=["--"], state="disabled", width=_W,
            command=self._on_store_change,
        )
        self.store_cb.pack()

        # ── ボタン（Excelを開く＝左・DB更新＝右） ──
        btn_frame = ctk.CTkFrame(card, fg_color="transparent")
        btn_frame.pack(padx=16, pady=(20, 0), anchor="w")
        btn_w = (_W - 12) // 2

        self.btn_open = ctk.CTkButton(
            btn_frame, text="Excelを開く", command=self._open_excel,
            state="disabled", width=btn_w,
        )
        self.btn_open.pack(side="left", padx=(0, 12))

        self.btn_db = ctk.CTkButton(
            btn_frame, text="DB更新", command=self._on_db_update,
            state="disabled", width=btn_w,
            fg_color="#f1f5f9", hover_color="#e2e8f0",
            text_color="#374151", border_width=1, border_color="#e2e8f0",
        )
        self.btn_db.pack(side="left")

        # ── AIコメント設定（控えめなサブボタン）──
        self.btn_ai = ctk.CTkButton(
            card, text="⚙  AIコメント設定", command=self._open_ai_settings,
            width=_W, height=28,
            fg_color="transparent", hover_color="#f1f5f9",
            text_color="#64748b", border_width=1, border_color="#e2e8f0",
        )
        self.btn_ai.pack(padx=16, pady=(12, 0))

        # ── ステータス ──
        ctk.CTkFrame(card, fg_color="#f1f5f9", height=1, corner_radius=0).pack(
            fill="x", pady=(20, 0),
        )
        status_row = ctk.CTkFrame(card, fg_color="transparent")
        status_row.pack(fill="x", padx=16, pady=(10, 16))

        self.dot_lbl = ctk.CTkLabel(
            status_row, text="●",
            font=ctk.CTkFont(size=10), text_color=_GRAY,
        )
        self.dot_lbl.pack(side="left")
        self.status_lbl = ctk.CTkLabel(
            status_row, text="初期化中...",
            font=ctk.CTkFont(size=12), text_color=_GRAY,
        )
        self.status_lbl.pack(side="left", padx=(6, 0))

    # ── 起動処理 ────────────────────────────────────────────────────────────

    def _startup(self):
        try:
            if DB_PATH.exists() and OUTPUT_PATH.exists():
                self._set_status("DBとExcelを確認しました。読み込み中...", _GRAY)
            else:
                self._set_status("初回セットアップ中（CSV → DB → Excel）...", _GRAY)
                OUTPUT_PATH.parent.mkdir(exist_ok=True)
                import_csv_to_db(INPUT_DIR, DB_PATH)
                year = get_years_from_db(DB_PATH)[-1]
                build_base_workbook(DB_PATH, year, OUTPUT_PATH)
            self.after(0, self._load_selectors)
        except Exception as e:
            self._set_status(f"エラー: {e}", _RED)

    # ── セレクター更新 ──────────────────────────────────────────────────────

    def _load_selectors(self):
        years = [str(y) for y in get_years_from_db(DB_PATH)]
        self.year_cb.configure(values=years, state="normal")
        if self.year_cb.get() not in years:
            self.year_cb.set(years[-1])
        self._refresh_store_list()
        self.btn_db.configure(state="normal")
        self.btn_open.configure(state="normal")
        self._set_status("準備完了", _GREEN)

    def _refresh_store_list(self):
        year   = int(self.year_cb.get())
        stores = get_stores_from_db(DB_PATH, year)
        self.store_cb.configure(values=stores, state="normal")
        if stores:
            self.store_cb.set(stores[0])

    # ── イベントハンドラ ────────────────────────────────────────────────────

    def _on_year_change(self, _value=None):
        self._refresh_store_list()
        threading.Thread(target=self._rebuild_year, daemon=True).start()

    def _rebuild_year(self):
        try:
            year  = int(self.year_cb.get())
            store = self.store_cb.get()
            self._set_controls("disabled")
            self._set_status(f"{year}年のダッシュボードを更新中...", _GRAY)
            build_base_workbook(DB_PATH, year, OUTPUT_PATH)
            if store:
                update_store_sheet(DB_PATH, year, OUTPUT_PATH, store)
            self._set_status(f"{year}年に切り替えました", _GREEN)
        except PermissionError:
            self._set_status("Excelを閉じてから再試行してください", _RED)
        except Exception as e:
            self._set_status(f"エラー: {e}", _RED)
        finally:
            self._set_controls("normal")

    def _on_store_change(self, _value=None):
        threading.Thread(target=self._refresh_store, daemon=True).start()

    def _refresh_store(self):
        try:
            year  = int(self.year_cb.get())
            store = self.store_cb.get()
            self._set_controls("disabled")
            self._set_status(f"{store} のデータを更新中...", _GRAY)
            update_store_sheet(DB_PATH, year, OUTPUT_PATH, store)
            self._set_status(f"店舗詳細シートを更新しました：{store}", _GREEN)
        except PermissionError:
            self._set_status("Excelを閉じてから再試行してください", _RED)
        except Exception as e:
            self._set_status(f"エラー: {e}", _RED)
        finally:
            self._set_controls("normal")

    def _on_db_update(self):
        threading.Thread(target=self._do_db_update, daemon=True).start()

    def _do_db_update(self):
        try:
            self._set_controls("disabled")
            self._set_status("CSVをDBに取り込み中...", _GRAY)
            import_csv_to_db(INPUT_DIR, DB_PATH)
            self.after(0, self._load_selectors)
        except FileNotFoundError:
            self._set_status(
                f"CSVが見つかりません。{INPUT_DIR} にCSVファイルを追加してください。", _RED,
            )
        except Exception as e:
            self._set_status(f"エラー: {e}", _RED)
        finally:
            self._set_controls("normal")

    def _open_excel(self):
        if OUTPUT_PATH.exists():
            os.startfile(OUTPUT_PATH)
        else:
            messagebox.showwarning("ファイルなし", "Excelファイルが見つかりません。")

    # ── AIコメント設定ダイアログ ────────────────────────────────────────────

    def _open_ai_settings(self):
        cfg = load_config()

        win = ctk.CTkToplevel(self)
        win.title("AIコメント設定")
        win.resizable(False, False)
        win.configure(fg_color="#f1f5f9")
        win.transient(self)
        win.after(100, win.grab_set)  # Toplevel表示後にモーダル化（Windowsでのちらつき回避）

        card = ctk.CTkFrame(win, fg_color="white", corner_radius=16)
        card.pack(padx=20, pady=20, fill="both", expand=True)

        ctk.CTkLabel(
            card, text="AIコメント設定",
            font=ctk.CTkFont(size=16, weight="bold"), text_color="#1e40af",
        ).pack(anchor="w", padx=20, pady=(18, 4))
        ctk.CTkLabel(
            card, text="KPIダッシュボードにAI生成コメントを追加します。",
            font=ctk.CTkFont(size=11), text_color="#64748b",
            wraplength=300, justify="left",
        ).pack(anchor="w", padx=20, pady=(0, 14))

        enabled_var = ctk.BooleanVar(value=bool(cfg.get("ai_enabled", False)))
        ctk.CTkCheckBox(
            card, text="AIインサイトを有効にする",
            variable=enabled_var, font=ctk.CTkFont(size=13),
        ).pack(anchor="w", padx=20)

        ctk.CTkLabel(
            card, text="Anthropic APIキー",
            font=ctk.CTkFont(size=12, weight="bold"), text_color="#374151",
        ).pack(anchor="w", padx=20, pady=(16, 4))
        key_entry = ctk.CTkEntry(card, width=300, show="●", placeholder_text="sk-ant-...")
        key_entry.pack(padx=20)
        if cfg.get("api_key"):
            key_entry.insert(0, cfg["api_key"])

        ctk.CTkLabel(
            card, text="キーは config.json に保存されます。\nモデル: claude-sonnet-4-6",
            font=ctk.CTkFont(size=10), text_color="#94a3b8", justify="left",
        ).pack(anchor="w", padx=20, pady=(8, 16))

        btn_row = ctk.CTkFrame(card, fg_color="transparent")
        btn_row.pack(padx=20, pady=(0, 18), anchor="e")

        def _save():
            save_config(enabled_var.get(), key_entry.get())
            win.destroy()
            year = self.year_cb.get()
            if year.isdigit():
                self._set_status("AI設定を保存しました。ダッシュボードを更新中...", _GRAY)
                threading.Thread(target=self._rebuild_year, daemon=True).start()
            else:
                self._set_status("AI設定を保存しました", _GREEN)

        ctk.CTkButton(
            btn_row, text="キャンセル", width=100, command=win.destroy,
            fg_color="#f1f5f9", hover_color="#e2e8f0",
            text_color="#374151", border_width=1, border_color="#e2e8f0",
        ).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_row, text="保存", width=100, command=_save).pack(side="left")

        # 親ウィンドウ中央に配置
        win.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - win.winfo_width()) // 2
        y = self.winfo_y() + (self.winfo_height() - win.winfo_height()) // 2
        win.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    # ── ユーティリティ ──────────────────────────────────────────────────────

    def _set_status(self, msg: str, color: str):
        self.after(0, lambda: (
            self.dot_lbl.configure(text_color=color),
            self.status_lbl.configure(text=msg, text_color=color),
        ))

    def _set_controls(self, state: str):
        cb_state = "normal" if state == "normal" else "disabled"
        def _apply():
            self.year_cb.configure(state=cb_state)
            self.store_cb.configure(state=cb_state)
            self.btn_db.configure(state=state)
            self.btn_open.configure(state=state)
        self.after(0, _apply)


def main():
    app = SalesReportApp()
    app.mainloop()


if __name__ == "__main__":
    main()
