"""
FALCON2 - 農場薬品設定ウインドウ
農場で使用する薬品の製品名・単価を管理
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Dict, Any, Optional, List, Tuple
from datetime import datetime
import json
import logging

import constants


class DrugSettingsWindow:
    """農場薬品設定ウインドウ"""
    
    def __init__(self, parent: tk.Tk, farm_path: Path):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ（FarmSettingsWindow）
            farm_path: 農場フォルダのパス
        """
        self.parent = parent
        self.farm_path = Path(farm_path)
        self.settings_file = self.farm_path / "drug_settings.json"
        
        # 設定データ
        self.drugs: Dict[str, List[Dict[str, Any]]] = {
            "pg": [{"name": "", "unit_price": None, "updated_at": None} for _ in range(3)],
            "gnrh": [{"name": "", "unit_price": None, "updated_at": None} for _ in range(3)],
            "vaginal_insert": [{"name": "", "unit_price": None, "updated_at": None} for _ in range(3)],
            "other": [{"name": "", "unit_price": None, "updated_at": None} for _ in range(5)]
        }
        
        # 処置料データ（全体で1つ）
        self.treatment_fees: Dict[str, Dict[str, Any]] = {
            "ultrasound": {"fee": None, "updated_at": None},  # 超音波検査
            "injection": {"fee": None, "updated_at": None},  # 筋肉注射
            "vaginal_insert": {"fee": None, "updated_at": None},  # 膣内留置
            "drug_administration": {"fee": None, "updated_at": None}  # 薬治
        }
        
        # ウィンドウ作成
        self.window = tk.Toplevel(parent)
        self.window.title("農場薬品設定")
        self.window.geometry("900x700")  # 幅を800から900に拡大
        
        # 設定をロード
        self._load_settings()
        
        # UI作成
        self._create_widgets()
        
        # ウィンドウを中央に配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def _load_settings(self):
        """設定をロード"""
        if self.settings_file.exists():
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # 各カテゴリのデータをロード（存在しない場合はデフォルト）
                    for category in ["pg", "gnrh", "vaginal_insert", "other"]:
                        if category in data:
                            loaded_items = data[category]
                            # 既存のリストを更新（長さを維持）
                            max_len = len(self.drugs[category])
                            for i in range(max_len):
                                if i < len(loaded_items):
                                    self.drugs[category][i] = loaded_items[i]
                                else:
                                    self.drugs[category][i] = {"name": "", "unit_price": None, "updated_at": None}
                    
                    # 処置料データをロード（全体で1つ）
                    if "treatment_fees" in data:
                        # 旧形式（カテゴリごと）から新形式（全体で1つ）への移行対応
                        if isinstance(data["treatment_fees"], dict):
                            # 新形式の場合
                            for fee_type in ["ultrasound", "injection", "vaginal_insert", "drug_administration"]:
                                if fee_type in data["treatment_fees"]:
                                    self.treatment_fees[fee_type] = data["treatment_fees"][fee_type]
                        # 旧形式の場合は無視（新形式で再保存される）
                logging.info("農場薬品設定を読み込みました")
            except Exception as e:
                logging.error(f"農場薬品設定ファイル読み込みエラー: {e}")
                messagebox.showerror("エラー", f"設定ファイルの読み込みに失敗しました: {e}")
        else:
            # ファイルが存在しない場合はデフォルトデータのまま
            logging.info("農場薬品設定ファイルが存在しません")
    
    def _settings_payload(self) -> Dict[str, Any]:
        """現在の self.drugs / self.treatment_fees から JSON 用 dict を構築"""
        return {
            "pg": self.drugs["pg"],
            "gnrh": self.drugs["gnrh"],
            "vaginal_insert": self.drugs["vaginal_insert"],
            "other": self.drugs["other"],
            "treatment_fees": self.treatment_fees,
        }

    def _write_settings_to_farm_dir(self, farm_dir: Path) -> None:
        """指定農場フォルダに drug_settings.json を書き込む（例外は呼び出し元で捕捉）"""
        farm_dir = Path(farm_dir)
        farm_dir.mkdir(parents=True, exist_ok=True)
        path = farm_dir / "drug_settings.json"
        data = self._settings_payload()
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        logging.info("農場薬品設定を保存しました: %s", path)

    @staticmethod
    def _list_sibling_farm_directories(farm_path: Path) -> List[Path]:
        """同じ親フォルダ直下に farm.db がある農場フォルダ一覧（名前順）"""
        parent = Path(farm_path).parent
        out: List[Path] = []
        try:
            for item in parent.iterdir():
                if item.is_dir() and (item / "farm.db").exists():
                    out.append(item.resolve())
        except OSError as e:
            logging.warning("同階層農場の列挙に失敗: %s", e)
        return sorted(out, key=lambda p: p.name.lower())

    @staticmethod
    def _list_all_farm_directories_under_root(farms_root: Path) -> List[Path]:
        """FARMS_ROOT 以下で farm.db があるディレクトリをすべて（再帰）"""
        root = Path(farms_root)
        if not root.exists():
            return []
        found = set()
        try:
            for db in root.rglob("farm.db"):
                if db.is_file():
                    found.add(db.parent.resolve())
        except OSError as e:
            logging.warning("FARMS 配下の農場列挙に失敗: %s", e)
        return sorted(found, key=lambda p: str(p).lower())

    def _prompt_save_scope(self) -> Optional[str]:
        """
        保存範囲を選択。戻り値: \"current\" | \"siblings\" | \"all\" | None（キャンセル）
        """
        siblings = self._list_sibling_farm_directories(self.farm_path)
        all_farms = self._list_all_farm_directories_under_root(constants.FARMS_ROOT)

        dlg = tk.Toplevel(self.window)
        dlg.title("保存先の選択")
        dlg.transient(self.window)
        dlg.grab_set()
        dlg.resizable(False, False)

        var = tk.StringVar(value="current")
        result: Dict[str, Optional[str]] = {"v": None}

        f = ttk.Frame(dlg, padding=16)
        f.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            f,
            text="薬品設定（製品名・単価・処置料）をどこに保存するか選んでください。",
            font=("", 10),
        ).pack(anchor=tk.W, pady=(0, 8))

        ttk.Radiobutton(
            f,
            text=f"この農場のみ（{self.farm_path.name}）",
            variable=var,
            value="current",
        ).pack(anchor=tk.W)

        sib_note = f"同じフォルダ内の全農場（{len(siblings)} 件）\n    親フォルダ: {self.farm_path.parent}"
        rb_sib = ttk.Radiobutton(
            f,
            text=sib_note,
            variable=var,
            value="siblings",
        )
        rb_sib.pack(anchor=tk.W, pady=(8, 0))
        if len(siblings) <= 1:
            rb_sib.state(["disabled"])

        all_note = (
            f"FARMS ルート配下の全農場（{len(all_farms)} 件）\n"
            f"    ルート: {constants.FARMS_ROOT}"
        )
        rb_all = ttk.Radiobutton(f, text=all_note, variable=var, value="all")
        rb_all.pack(anchor=tk.W, pady=(8, 0))
        if len(all_farms) <= 1:
            rb_all.state(["disabled"])

        ttk.Label(
            f,
            text="※ 複数農場へ保存すると、各農場の drug_settings.json が同じ内容で上書きされます。",
            font=("", 8),
            foreground="#555",
        ).pack(anchor=tk.W, pady=(12, 0))

        bf = ttk.Frame(f)
        bf.pack(fill=tk.X, pady=(14, 0))

        def ok() -> None:
            result["v"] = var.get()
            dlg.destroy()

        def cancel() -> None:
            result["v"] = None
            dlg.destroy()

        ttk.Button(bf, text="OK", command=ok, width=10).pack(side=tk.RIGHT, padx=4)
        ttk.Button(bf, text="キャンセル", command=cancel, width=10).pack(side=tk.RIGHT)

        dlg.protocol("WM_DELETE_WINDOW", cancel)
        dlg.update_idletasks()
        w, h = dlg.winfo_width(), dlg.winfo_height()
        px = self.window.winfo_rootx() + (self.window.winfo_width() // 2) - (w // 2)
        py = self.window.winfo_rooty() + (self.window.winfo_height() // 2) - (h // 2)
        dlg.geometry(f"+{max(0, px)}+{max(0, py)}")

        self.window.wait_window(dlg)
        return result["v"]

    def _targets_for_scope(self, scope: str) -> List[Path]:
        if scope == "current":
            return [self.farm_path.resolve()]
        if scope == "siblings":
            return list(self._list_sibling_farm_directories(self.farm_path))
        if scope == "all":
            found = list(self._list_all_farm_directories_under_root(constants.FARMS_ROOT))
            cur = self.farm_path.resolve()
            paths = {p.resolve() for p in found}
            if cur not in paths:
                # FARMS ルート外で開いている場合も、現在農場へは必ず書けるようにする
                found.append(cur)
            return sorted(found, key=lambda p: str(p).lower())
        return []

    def _save_to_targets(self, targets: List[Path]) -> Tuple[int, List[str]]:
        """複数農場へ保存。成功数とエラーメッセージ一覧を返す。"""
        errors: List[str] = []
        ok = 0
        for fp in targets:
            try:
                self._write_settings_to_farm_dir(fp)
                ok += 1
            except Exception as e:
                logging.exception("農場薬品設定の保存エラー: %s", fp)
                errors.append(f"{fp.name}: {e}")
        return ok, errors

    def _create_widgets(self):
        """ウィジェットを作成"""
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # スクロール可能なフレーム
        canvas_frame = ttk.Frame(main_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        canvas = tk.Canvas(canvas_frame)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        scrollable_frame.bind("<Configure>", on_frame_configure)
        
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        def on_canvas_configure(event):
            canvas_width = event.width
            canvas.itemconfig(canvas_window, width=canvas_width)
        
        canvas.bind("<Configure>", on_canvas_configure)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # マウスホイールでスクロール（bind_allは使用しない：ウィンドウ閉鎖後もコールバックが残るため）
        def on_mousewheel(event):
            try:
                if canvas.winfo_exists():
                    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except tk.TclError:
                pass  # ウィジェット破棄後に呼ばれた場合は無視
        
        canvas.bind("<MouseWheel>", on_mousewheel)
        scrollable_frame.bind("<MouseWheel>", on_mousewheel)
        self.window.bind("<MouseWheel>", on_mousewheel)  # 子ウィジェット上でもスクロール可能に
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # ========== PG設定 ==========
        pg_frame = ttk.LabelFrame(scrollable_frame, text="PG（プロスタグランジン）", padding=10)
        pg_frame.pack(fill=tk.X, pady=5)
        
        self.pg_entries = []
        for i in range(3):
            item_frame = ttk.Frame(pg_frame)
            item_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(item_frame, text=f"PG{i+1}：", width=8).pack(side=tk.LEFT, padx=5)
            
            name_entry = ttk.Entry(item_frame, width=20)
            name_entry.pack(side=tk.LEFT, padx=5)
            
            ttk.Label(item_frame, text="単価（円/ml）：", width=12).pack(side=tk.LEFT, padx=5)
            
            price_entry = ttk.Entry(item_frame, width=10)
            price_entry.pack(side=tk.LEFT, padx=5)
            
            date_label = ttk.Label(item_frame, text="", width=18)  # 幅を12から18に拡大
            date_label.pack(side=tk.LEFT, padx=5)
            
            self.pg_entries.append({
                "name": name_entry,
                "price": price_entry,
                "date": date_label
            })
        
        # ========== GnRH設定 ==========
        gnrh_frame = ttk.LabelFrame(scrollable_frame, text="GnRH", padding=10)
        gnrh_frame.pack(fill=tk.X, pady=5)
        
        self.gnrh_entries = []
        for i in range(3):
            item_frame = ttk.Frame(gnrh_frame)
            item_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(item_frame, text=f"GnRH{i+1}：", width=8).pack(side=tk.LEFT, padx=5)
            
            name_entry = ttk.Entry(item_frame, width=20)
            name_entry.pack(side=tk.LEFT, padx=5)
            
            ttk.Label(item_frame, text="単価（円/ml）：", width=12).pack(side=tk.LEFT, padx=5)
            
            price_entry = ttk.Entry(item_frame, width=10)
            price_entry.pack(side=tk.LEFT, padx=5)
            
            date_label = ttk.Label(item_frame, text="", width=18)  # 幅を12から18に拡大
            date_label.pack(side=tk.LEFT, padx=5)
            
            self.gnrh_entries.append({
                "name": name_entry,
                "price": price_entry,
                "date": date_label
            })
        
        # ========== 膣内留置剤設定 ==========
        vaginal_frame = ttk.LabelFrame(scrollable_frame, text="膣内留置剤", padding=10)
        vaginal_frame.pack(fill=tk.X, pady=5)
        
        self.vaginal_entries = []
        for i in range(3):
            item_frame = ttk.Frame(vaginal_frame)
            item_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(item_frame, text=f"留置{i+1}：", width=8).pack(side=tk.LEFT, padx=5)
            
            name_entry = ttk.Entry(item_frame, width=20)
            name_entry.pack(side=tk.LEFT, padx=5)
            
            ttk.Label(item_frame, text="単価（円/ml）：", width=12).pack(side=tk.LEFT, padx=5)
            
            price_entry = ttk.Entry(item_frame, width=10)
            price_entry.pack(side=tk.LEFT, padx=5)
            
            date_label = ttk.Label(item_frame, text="", width=18)  # 幅を12から18に拡大
            date_label.pack(side=tk.LEFT, padx=5)
            
            self.vaginal_entries.append({
                "name": name_entry,
                "price": price_entry,
                "date": date_label
            })
        
        # ========== その他薬品設定 ==========
        other_frame = ttk.LabelFrame(scrollable_frame, text="その他薬品", padding=10)
        other_frame.pack(fill=tk.X, pady=5)
        
        self.other_entries = []
        for i in range(5):
            item_frame = ttk.Frame(other_frame)
            item_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(item_frame, text=f"その他{i+1}：", width=8).pack(side=tk.LEFT, padx=5)
            
            name_entry = ttk.Entry(item_frame, width=20)
            name_entry.pack(side=tk.LEFT, padx=5)
            
            ttk.Label(item_frame, text="単価（円/ml）：", width=12).pack(side=tk.LEFT, padx=5)
            
            price_entry = ttk.Entry(item_frame, width=10)
            price_entry.pack(side=tk.LEFT, padx=5)
            
            date_label = ttk.Label(item_frame, text="", width=18)  # 幅を12から18に拡大
            date_label.pack(side=tk.LEFT, padx=5)
            
            self.other_entries.append({
                "name": name_entry,
                "price": price_entry,
                "date": date_label
            })
        
        # その他薬品処置料（全体で1つ）
        self._create_treatment_fees_section(other_frame)
        
        # データを入力欄に反映
        self._load_data_to_widgets()
        
        # ボタンフレーム
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(
            button_frame,
            text="保存",
            command=self._on_save
        ).pack(side=tk.LEFT, padx=5)

        ttk.Label(
            button_frame,
            text="※ 保存時に、この農場のみ／同じフォルダの全農場／FARMS 配下の全農場を選べます。",
            font=("", 8),
            foreground="#555",
        ).pack(side=tk.LEFT, padx=(8, 0))
        
        ttk.Button(
            button_frame,
            text="閉じる",
            command=self._on_close
        ).pack(side=tk.RIGHT, padx=5)
    
    def _create_treatment_fees_section(self, parent_frame: ttk.Frame):
        """
        処置料セクションを作成（全体で1つ）
        
        Args:
            parent_frame: 親フレーム
        """
        # 処置料フレーム
        fees_frame = ttk.LabelFrame(parent_frame, text="処置料", padding=10)
        fees_frame.pack(fill=tk.X, pady=(10, 0))
        
        # 処置料の種類
        fee_types = [
            ("ultrasound", "超音波検査"),
            ("injection", "筋肉注射"),
            ("vaginal_insert", "膣内留置"),
            ("drug_administration", "薬治")
        ]
        
        # 処置料エントリを保存する辞書
        if not hasattr(self, 'treatment_fee_entries'):
            self.treatment_fee_entries = {}
        
        for fee_key, fee_label in fee_types:
            item_frame = ttk.Frame(fees_frame)
            item_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(item_frame, text=f"{fee_label}：", width=12).pack(side=tk.LEFT, padx=5)
            
            fee_entry = ttk.Entry(item_frame, width=10)
            fee_entry.pack(side=tk.LEFT, padx=5)
            
            ttk.Label(item_frame, text="円", width=3).pack(side=tk.LEFT, padx=2)
            
            date_label = ttk.Label(item_frame, text="", width=18)
            date_label.pack(side=tk.LEFT, padx=5)
            
            self.treatment_fee_entries[fee_key] = {
                "fee": fee_entry,
                "date": date_label
            }
    
    def _load_data_to_widgets(self):
        """データをウィジェットに反映"""
        # PG
        for i, entry in enumerate(self.pg_entries):
            drug = self.drugs["pg"][i]
            entry["name"].delete(0, tk.END)
            entry["name"].insert(0, drug.get("name", "") or "")
            entry["price"].delete(0, tk.END)
            if drug.get("unit_price") is not None:
                entry["price"].insert(0, str(drug.get("unit_price", "")))
            updated_at = drug.get("updated_at")
            if updated_at:
                entry["date"].config(text=f"更新: {updated_at}")
            else:
                entry["date"].config(text="")
        
        # 処置料データをウィジェットに反映
        self._load_treatment_fees_to_widgets()
        
        # GnRH
        for i, entry in enumerate(self.gnrh_entries):
            drug = self.drugs["gnrh"][i]
            entry["name"].delete(0, tk.END)
            entry["name"].insert(0, drug.get("name", "") or "")
            entry["price"].delete(0, tk.END)
            if drug.get("unit_price") is not None:
                entry["price"].insert(0, str(drug.get("unit_price", "")))
            updated_at = drug.get("updated_at")
            if updated_at:
                entry["date"].config(text=f"更新: {updated_at}")
            else:
                entry["date"].config(text="")
        
        # 膣内留置剤
        for i, entry in enumerate(self.vaginal_entries):
            drug = self.drugs["vaginal_insert"][i]
            entry["name"].delete(0, tk.END)
            entry["name"].insert(0, drug.get("name", "") or "")
            entry["price"].delete(0, tk.END)
            if drug.get("unit_price") is not None:
                entry["price"].insert(0, str(drug.get("unit_price", "")))
            updated_at = drug.get("updated_at")
            if updated_at:
                entry["date"].config(text=f"更新: {updated_at}")
            else:
                entry["date"].config(text="")
        
        # その他
        for i, entry in enumerate(self.other_entries):
            drug = self.drugs["other"][i]
            entry["name"].delete(0, tk.END)
            entry["name"].insert(0, drug.get("name", "") or "")
            entry["price"].delete(0, tk.END)
            if drug.get("unit_price") is not None:
                entry["price"].insert(0, str(drug.get("unit_price", "")))
            updated_at = drug.get("updated_at")
            if updated_at:
                entry["date"].config(text=f"更新: {updated_at}")
            else:
                entry["date"].config(text="")
    
    def _load_treatment_fees_to_widgets(self):
        """処置料データをウィジェットに反映（全体で1つ）"""
        if not hasattr(self, 'treatment_fee_entries'):
            return
        
        for fee_key, entry in self.treatment_fee_entries.items():
            fee_data = self.treatment_fees[fee_key]
            entry["fee"].delete(0, tk.END)
            if fee_data.get("fee") is not None:
                entry["fee"].insert(0, str(fee_data.get("fee", "")))
            updated_at = fee_data.get("updated_at")
            if updated_at:
                entry["date"].config(text=f"更新: {updated_at}")
            else:
                entry["date"].config(text="")
    
    def _sync_widgets_to_model(self) -> bool:
        """
        入力欄の内容を self.drugs / self.treatment_fees に反映する。
        検証エラー時は False。
        """
        today = datetime.now().strftime("%Y-%m-%d")

        # PG
        for i, entry in enumerate(self.pg_entries):
            name = entry["name"].get().strip()
            price_str = entry["price"].get().strip()
            price = None
            if price_str:
                try:
                    price = float(price_str)
                except ValueError:
                    messagebox.showerror("エラー", f"PG{i+1}の単価が正しい数値ではありません。")
                    return False
            item = self.drugs["pg"][i]
            changed = (item.get("name") != name or item.get("unit_price") != price)
            item["name"] = name
            item["unit_price"] = price
            if changed:
                item["updated_at"] = today
        
        # GnRH
        for i, entry in enumerate(self.gnrh_entries):
            name = entry["name"].get().strip()
            price_str = entry["price"].get().strip()
            price = None
            if price_str:
                try:
                    price = float(price_str)
                except ValueError:
                    messagebox.showerror("エラー", f"GnRH{i+1}の単価が正しい数値ではありません。")
                    return False
            item = self.drugs["gnrh"][i]
            changed = (item.get("name") != name or item.get("unit_price") != price)
            item["name"] = name
            item["unit_price"] = price
            if changed:
                item["updated_at"] = today
        
        # 膣内留置剤
        for i, entry in enumerate(self.vaginal_entries):
            name = entry["name"].get().strip()
            price_str = entry["price"].get().strip()
            price = None
            if price_str:
                try:
                    price = float(price_str)
                except ValueError:
                    messagebox.showerror("エラー", f"留置{i+1}の単価が正しい数値ではありません。")
                    return False
            item = self.drugs["vaginal_insert"][i]
            changed = (item.get("name") != name or item.get("unit_price") != price)
            item["name"] = name
            item["unit_price"] = price
            if changed:
                item["updated_at"] = today
        
        # その他
        for i, entry in enumerate(self.other_entries):
            name = entry["name"].get().strip()
            price_str = entry["price"].get().strip()
            price = None
            if price_str:
                try:
                    price = float(price_str)
                except ValueError:
                    messagebox.showerror("エラー", f"その他{i+1}の単価が正しい数値ではありません。")
                    return False
            item = self.drugs["other"][i]
            changed = (item.get("name") != name or item.get("unit_price") != price)
            item["name"] = name
            item["unit_price"] = price
            if changed:
                item["updated_at"] = today
        
        # 処置料の保存（変更があった項目のみ updated_at を更新）
        if hasattr(self, 'treatment_fee_entries'):
            fee_label_map = {
                "ultrasound": "超音波検査",
                "injection": "筋肉注射",
                "vaginal_insert": "膣内留置",
                "drug_administration": "薬治"
            }
            
            for fee_key, entry in self.treatment_fee_entries.items():
                fee_str = entry["fee"].get().strip()
                fee = None
                if fee_str:
                    try:
                        fee = float(fee_str)
                    except ValueError:
                        fee_label = fee_label_map.get(fee_key, fee_key)
                        messagebox.showerror("エラー", f"{fee_label}の料金が正しい数値ではありません。")
                        return False
                fee_data = self.treatment_fees[fee_key]
                changed = (fee_data.get("fee") != fee)
                fee_data["fee"] = fee
                if changed:
                    fee_data["updated_at"] = today

        return True

    def _on_save(self):
        """保存ボタン：入力を検証後、保存先（この農場のみ／同階層全件／FARMS 配下全件）を選択して保存"""
        if not self._sync_widgets_to_model():
            return

        scope = self._prompt_save_scope()
        if scope is None:
            return

        targets = self._targets_for_scope(scope)
        if not targets:
            messagebox.showwarning("保存", "保存対象の農場フォルダがありません。")
            return

        if len(targets) > 1:
            preview = "\n".join(f"・{p.name}" for p in targets[:25])
            if len(targets) > 25:
                preview += f"\n…他 {len(targets) - 25} 件"
            if not messagebox.askyesno(
                "確認",
                f"次の {len(targets)} 件の農場に、同じ薬品設定（drug_settings.json）を書き込みます。\n"
                f"よろしいですか？\n\n{preview}",
            ):
                return

        ok, errors = self._save_to_targets(targets)
        if ok == 0:
            msg = "保存に失敗しました。"
            if errors:
                msg += "\n\n" + "\n".join(errors[:8])
            messagebox.showerror("エラー", msg)
            return

        if errors:
            messagebox.showwarning(
                "一部失敗",
                f"{ok} 件は保存できましたが、以下は失敗しました:\n\n"
                + "\n".join(errors[:12]),
            )
        elif ok == 1:
            messagebox.showinfo("保存完了", "農場薬品設定を保存しました。")
        else:
            messagebox.showinfo("保存完了", f"{ok} 件の農場に薬品設定を保存しました。")

        self._load_data_to_widgets()

    def _on_close(self):
        """閉じるボタンをクリック"""
        self.window.destroy()
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()
