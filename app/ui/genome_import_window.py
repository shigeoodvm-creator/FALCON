"""
FALCON2 - ゲノムデータ取り込みウィンドウ
CSVファイルからゲノム項目（DWP$以降）を認識し、個体のゲノム吸い込みイベントとして登録する。
常に個体あたり最新1件のみ（過去のゲノム吸い込みイベントは削除してから新規登録）。
"""

import csv
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Optional, List, Tuple
import json
import re
import logging

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine

logger = logging.getLogger(__name__)


def _extract_jpn10(raw: str) -> Optional[str]:
    """
    CSVのセル値からJPN10（10桁個体識別番号）を抽出する。
    'HOJPN001616614612' → 数字のみ取り出して10桁を返す（12桁なら後ろ10桁）。
    10桁の数字のみの場合はそのまま返す。
    """
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None
    digits = re.sub(r"\D", "", s)
    if len(digits) == 10:
        return digits
    if len(digits) > 10:
        return digits[-10:]
    return None


def _find_column_index(headers: List[str], keywords: List[str]) -> Optional[int]:
    """ヘッダー行からキーワードに部分一致する列のインデックスを返す。
    キーワードの並び順を優先する（先に列挙したキーワードに一致する列を返す）。
    Official ID / JPN10 のように10桁IDが入る列を、動物ID（4桁）より優先するため。"""
    for kw in keywords:
        for idx, h in enumerate(headers):
            if not h:
                continue
            h_str = str(h).strip()
            if kw in h_str:
                return idx
    return None


# ゲノム項目ではない列（個体メタデータなど）のヘッダーに含まれるキーワード。
# これらに該当する列はゲノム項目から除外し、残りをゲノム項目として扱う。
NON_GENOME_HEADER_KEYWORDS = (
    "動物ID", "個体ID", "Official ID", "JPN10", "個体識別", "動物名", "名前", "識別番号", "識別",
    "品種", "BRD", "Breed", "生年月日", "BTHD", "Birth", "誕生日", "生年月",
    "Sex", "性別", "Group", "群", "PEN", "Status", "状態",
    "ID",  # 列名が単に ID の場合（動物ID等と別の列のとき）
)
def _is_non_genome_header(header: str) -> bool:
    """ヘッダーがゲノム項目でない（個体ID・品種・生年月日・Sex等）とみなすか"""
    if not header or not isinstance(header, str):
        return True
    h = header.strip()
    for kw in NON_GENOME_HEADER_KEYWORDS:
        if kw in h:
            return True
        if kw.isascii() and kw.lower() in h.lower():
            return True
    return False


# CSV列名 → FALCON総合指標キー（レポートのGDWP$/GNM$/GTPI列と一致させる）
# これ以外は従来どおり "G" + 列名 とする（例: "PL" → "GPL"）
HEADER_TO_COMPOSITE_KEY = {
    "DWP$": "GDWP$",
    "NM$": "GNM$",
    "TPI": "GTPI",
    "GDWP$": "GDWP$",
    "GNM$": "GNM$",
    "GTPI": "GTPI",
    "GDPR": "GDPR",
}


class GenomeImportWindow:
    """ゲノムデータ取り込みウィンドウ"""

    def __init__(
        self,
        parent: tk.Tk,
        db_handler: DBHandler,
        rule_engine: RuleEngine,
        item_dict_path: Optional[Path] = None,
    ):
        self.parent = parent
        self.db = db_handler
        self.rule_engine = rule_engine
        self.item_dict_path = Path(item_dict_path) if item_dict_path else None

        self.window = tk.Toplevel(parent)
        self.window.title("ゲノムデータ取り込み")
        self.window.geometry("900x600")

        self.csv_path: Optional[Path] = None
        self.data_rows: List[Dict[str, Any]] = []
        self.genome_keys: List[str] = []  # FALCON用キー（G付き）
        self.genome_column_indices: List[int] = []  # ゲノム項目の列インデックス
        self._create_widgets()

    def _create_widgets(self):
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        file_frame = ttk.LabelFrame(main_frame, text="CSVファイル選択", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))

        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=70, state="readonly").pack(
            side=tk.LEFT, padx=(0, 10)
        )
        ttk.Button(file_frame, text="参照...", command=self._select_csv_file).pack(side=tk.LEFT)

        preview_frame = ttk.LabelFrame(main_frame, text="プレビュー", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        tree_frame = ttk.Frame(preview_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree = ttk.Treeview(
            tree_frame,
            columns=("jpn10", "cow_id", "status", "sample"),
            show="headings",
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set,
        )
        self.tree.heading("jpn10", text="JPN10（個体識別）")
        self.tree.heading("cow_id", text="ID")
        self.tree.heading("status", text="状態")
        self.tree.heading("sample", text="ゲノム項目サンプル")
        self.tree.column("jpn10", width=120)
        self.tree.column("cow_id", width=80)
        self.tree.column("status", width=120)
        self.tree.column("sample", width=280)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.config(command=self.tree.yview)
        scrollbar_x.config(command=self.tree.xview)

        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        self.info_label = ttk.Label(info_frame, text="CSVファイルを選択してください")
        self.info_label.pack(side=tk.LEFT)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        ttk.Button(button_frame, text="取り込み実行", command=self._execute_import).pack(
            side=tk.LEFT, padx=(0, 10)
        )
        ttk.Button(button_frame, text="閉じる", command=self.window.destroy).pack(side=tk.LEFT)

    def _select_csv_file(self):
        file_path = filedialog.askopenfilename(
            title="ゲノムCSVファイルを選択",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if file_path:
            self.csv_path = Path(file_path)
            self.file_path_var.set(str(self.csv_path))
            self._parse_csv()

    def _parse_csv(self):
        if not self.csv_path or not self.csv_path.exists():
            messagebox.showerror("エラー", "CSVファイルが見つかりません")
            return

        try:
            for enc in ("utf-8", "utf-8-sig", "shift_jis", "cp932"):
                try:
                    with open(self.csv_path, "r", encoding=enc, newline="") as f:
                        reader = csv.reader(f)
                        lines = [row for row in reader]
                    break
                except UnicodeDecodeError:
                    continue
            else:
                messagebox.showerror("エラー", "CSVの文字コードを判別できませんでした。")
                return
            if not lines:
                messagebox.showerror("エラー", "CSVが空です")
                return

            # 列数を揃える
            max_cols = max(len(row) for row in lines)
            for row in lines:
                while len(row) < max_cols:
                    row.append("")

            headers = [str(h).strip() for h in lines[0]]

            # JPN10と部分一致する列（FALCONのJPN10に対応する10桁個体識別列）
            # 動物ID（4桁）は別に取得するため、ここでは10桁が入る列を優先
            col_jpn10 = _find_column_index(
                headers,
                ["Official ID", "JPN10", "個体識別番号", "個体識別", "識別番号", "JPN", "動物名"],
            )
            if col_jpn10 is None:
                messagebox.showerror("エラー", "JPN10（または個体識別）に相当する列が見つかりません。")
                return

            # 4桁の個体ID列（動物ID・個体ID）があればこちらを表示・吸い込みに使用
            col_cow_id = _find_column_index(
                headers,
                ["動物ID", "個体ID", "ID"],
            )

            # 個体ID・品種・生年月日・Sex等を除外し、残りの列をゲノム項目とする
            self.genome_keys = []  # FALCON用キー（G付き）
            self.genome_column_indices: List[int] = []  # ゲノム項目の列インデックス
            for idx, h in enumerate(headers):
                name = (h or "").strip()
                if not name:
                    continue
                if _is_non_genome_header(name):
                    continue
                if idx == col_jpn10:
                    continue  # 個体識別列はゲノム値ではない
                if col_cow_id is not None and idx == col_cow_id:
                    continue  # 4桁ID列もゲノム値ではない
                falcon_key = HEADER_TO_COMPOSITE_KEY.get(name, "G" + name)
                self.genome_keys.append(falcon_key)
                self.genome_column_indices.append(idx)

            if not self.genome_keys:
                messagebox.showerror("エラー", "ゲノム項目が1件もありません。個体ID・品種・生年月日・Sex等以外の数値列を確認してください。")
                return

            data_rows = []
            for row in lines[1:]:
                jpn10_raw = row[col_jpn10] if col_jpn10 < len(row) else ""
                jpn10 = _extract_jpn10(jpn10_raw)
                if not jpn10:
                    continue

                # 4桁ID列（動物ID・個体ID）があればその値を使用、なければJPN10の下4桁
                cow_id = ""
                if col_cow_id is not None and col_cow_id < len(row):
                    raw_id = row[col_cow_id]
                    cow_id = str(raw_id).strip() if raw_id not in (None, "") else ""
                if not cow_id:
                    cow_id = jpn10[-4:] if len(jpn10) >= 4 else jpn10.zfill(4)

                genome_vals = {}
                for key, col_idx in zip(self.genome_keys, self.genome_column_indices):
                    if col_idx >= len(row):
                        continue
                    val = row[col_idx]
                    if val is None or (isinstance(val, str) and not val.strip()):
                        continue
                    s = str(val).strip()
                    # 数値の桁区切りカンマを除去（2,724 → 2724）
                    s_num = s.replace(",", "")
                    try:
                        if "." in s_num or "e" in s_num.lower():
                            genome_vals[key] = float(s_num)
                        else:
                            genome_vals[key] = int(float(s_num))
                    except (ValueError, TypeError):
                        genome_vals[key] = s

                data_rows.append({
                    "jpn10": jpn10,
                    "cow_id": cow_id,
                    "genome": genome_vals,
                })

            self.data_rows = data_rows
            self._ensure_genome_items_in_dictionary()
            self._update_preview()
            self.info_label.config(
                text=f"ゲノム項目: {', '.join(self.genome_keys[:5])}{'...' if len(self.genome_keys) > 5 else ''}　データ: {len(data_rows)}件"
            )

        except Exception as e:
            logger.error(f"ゲノムCSVパースエラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"CSVの読み込みに失敗しました: {e}")

    def _ensure_genome_items_in_dictionary(self):
        """項目辞書にゲノム項目（G付き）が無ければ追加する"""
        if not self.item_dict_path or not self.item_dict_path.exists():
            return
        try:
            with open(self.item_dict_path, "r", encoding="utf-8") as f:
                item_dict = json.load(f)
        except Exception as e:
            logger.warning(f"項目辞書読み込み失敗: {e}")
            return

        updated = False
        for key in self.genome_keys:
            if key in item_dict:
                continue
            item_dict[key] = {
                "type": "source",
                "origin": "source",
                "data_type": "float",
                "source": f"{RuleEngine.EVENT_GENOMIC}.{key}",
                "display_name": key,
                "description": f"ゲノム吸い込み（{key}）",
                "category": "GENOME",
            }
            updated = True

        if updated:
            try:
                with open(self.item_dict_path, "w", encoding="utf-8") as f:
                    json.dump(item_dict, f, ensure_ascii=False, indent=2)
                logger.info("項目辞書にゲノム項目を追加しました")
            except Exception as e:
                logger.warning(f"項目辞書保存失敗: {e}")

    def _update_preview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for row_data in self.data_rows:
            jpn10 = row_data.get("jpn10", "")
            cow_id = row_data.get("cow_id", "")
            genome = row_data.get("genome", {})

            cow = self.db.get_cow_by_id(cow_id)
            if not cow:
                try:
                    cows = self.db.get_cows_by_jpn10(jpn10)
                    if cows:
                        cow = cows[0]
                except Exception:
                    cow = None

            status = "FALCONにありません" if not cow else "吸い込み対象"
            sample = ", ".join(f"{k}={v}" for k, v in list(genome.items())[:3])
            if len(genome) > 3:
                sample += "..."

            self.tree.insert("", tk.END, values=(jpn10, cow_id, status, sample))

    def _execute_import(self):
        if not self.data_rows:
            messagebox.showwarning("警告", "取り込むデータがありません")
            return

        result = messagebox.askyesno(
            "確認",
            f"ゲノムデータを {len(self.data_rows)} 件取り込みますか？\n"
            "既存のゲノム吸い込みイベントは個体ごとに削除され、今回のデータで1件のみになります。",
        )
        if not result:
            return

        import_date = datetime.now().strftime("%Y-%m-%d")
        success_count = 0
        error_count = 0
        error_cows: List[Dict[str, Any]] = []

        for row_data in self.data_rows:
            jpn10 = row_data.get("jpn10")
            cow_id = row_data.get("cow_id")
            genome = row_data.get("genome", {})

            cow = self.db.get_cow_by_id(cow_id)
            if not cow:
                try:
                    cows = self.db.get_cows_by_jpn10(jpn10)
                    if cows:
                        cow = cows[0]
                except Exception:
                    cow = None

            if not cow:
                error_count += 1
                error_cows.append({
                    "cow_id": cow_id,
                    "jpn10": jpn10,
                    "reason": "マスタに個体がありません",
                })
                continue

            cow_auto_id = cow.get("auto_id")
            events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
            for ev in events:
                if ev.get("event_number") == RuleEngine.EVENT_GENOMIC:
                    self.db.delete_event(ev["id"], soft_delete=True)

            json_data = {k: v for k, v in genome.items() if v is not None}
            try:
                event_data = {
                    "cow_auto_id": cow_auto_id,
                    "event_number": RuleEngine.EVENT_GENOMIC,
                    "event_date": import_date,
                    "json_data": json_data if json_data else None,
                    "note": "ゲノムCSVから取り込み",
                }
                event_id = self.db.insert_event(event_data)
                self.rule_engine.on_event_added(event_id)
                self.rule_engine._recalculate_and_update_cow(cow_auto_id)
                success_count += 1
            except Exception as e:
                error_count += 1
                error_cows.append({
                    "cow_id": cow_id,
                    "jpn10": jpn10,
                    "reason": f"イベント作成エラー: {str(e)}",
                })
                logger.error(f"ゲノムイベント作成エラー: cow_id={cow_id}, e={e}", exc_info=True)

        result_msg = (
            f"取り込みが完了しました\n成功: {success_count}件\nエラー: {error_count}件"
        )
        if error_cows:
            result_msg += f"\n\nFALCONに存在しない個体が{len(error_cows)}件あります。詳細を表示しますか？"
            if messagebox.askyesno("完了", result_msg):
                self._show_error_cows_window(error_cows)
        else:
            messagebox.showinfo("完了", result_msg)

        self._update_preview()

    def _show_error_cows_window(self, error_cows: List[Dict[str, Any]]):
        """乳検と同様に、FALCONにない個体一覧を表示"""
        err_win = tk.Toplevel(self.window)
        err_win.title("FALCONに存在しない個体一覧")
        err_win.geometry("600x400")

        main_frame = ttk.Frame(err_win, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(
            main_frame,
            text=f"ゲノムCSVにはあるがFALCONに個体がなく吸い込めなかった一覧（{len(error_cows)}件）",
            font=("", 10, "bold"),
        ).pack(pady=(0, 10))

        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree = ttk.Treeview(
            tree_frame,
            columns=("cow_id", "jpn10", "reason"),
            show="headings",
            yscrollcommand=scroll_y.set,
        )
        tree.heading("cow_id", text="ID")
        tree.heading("jpn10", text="個体識別番号")
        tree.heading("reason", text="理由")
        tree.column("cow_id", width=100)
        tree.column("jpn10", width=150)
        tree.column("reason", width=300)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.config(command=tree.yview)

        for ec in error_cows:
            tree.insert("", tk.END, values=(
                ec.get("cow_id", ""),
                ec.get("jpn10", ""),
                ec.get("reason", ""),
            ))

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(
            button_frame,
            text="テキストファイルとして保存",
            command=lambda: self._save_error_cows_to_file(error_cows, err_win),
        ).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="閉じる", command=err_win.destroy).pack(side=tk.RIGHT)

    def _save_error_cows_to_file(self, error_cows: List[Dict[str, Any]], parent_window: tk.Toplevel):
        """エラー個体の一覧をテキストファイルとして保存（乳検吸い込みと同様）"""
        file_path = filedialog.asksaveasfilename(
            title="エラー個体一覧を保存",
            defaultextension=".txt",
            filetypes=[("テキストファイル", "*.txt"), ("すべてのファイル", "*.*")],
        )
        if not file_path:
            return
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write("=" * 60 + "\n")
                f.write("ゲノム吸い込み エラー個体一覧（FALCONに存在しない個体）\n")
                f.write("=" * 60 + "\n")
                f.write(f"生成日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"エラー件数: {len(error_cows)}件\n")
                f.write("=" * 60 + "\n\n")
                f.write(f"{'ID':<10} {'個体識別番号':<15} {'理由':<30}\n")
                f.write("-" * 60 + "\n")
                for ec in error_cows:
                    cow_id = ec.get("cow_id", "")
                    jpn10 = ec.get("jpn10", "")
                    reason = ec.get("reason", "")
                    f.write(f"{cow_id:<10} {jpn10:<15} {reason:<30}\n")
            messagebox.showinfo("保存完了", f"エラー個体一覧を保存しました:\n{file_path}")
        except Exception as e:
            logger.error(f"ファイル保存エラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"ファイルの保存に失敗しました:\n{e}")

    def show(self):
        self.window.focus_set()
