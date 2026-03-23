"""
FALCON2 - 農場薬品設定ウインドウ
農場で使用する薬品の製品名・単価を管理
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Dict, Any, Optional, List
from datetime import datetime
import json
import logging


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
    
    def _save_settings(self):
        """設定を保存（更新日は _on_save で変更があった項目のみ設定済み）"""
        data = {
            "pg": self.drugs["pg"],
            "gnrh": self.drugs["gnrh"],
            "vaginal_insert": self.drugs["vaginal_insert"],
            "other": self.drugs["other"],
            "treatment_fees": self.treatment_fees
        }
        
        try:
            self.farm_path.mkdir(parents=True, exist_ok=True)
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            logging.info("農場薬品設定を保存しました")
            return True
        except Exception as e:
            logging.error(f"農場薬品設定ファイル保存エラー: {e}")
            messagebox.showerror("エラー", f"設定ファイルの保存に失敗しました: {e}")
            return False
    
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
        
        ttk.Button(
            button_frame,
            text="閉じる",
            command=self._on_close
        ).pack(side=tk.LEFT, padx=5)
    
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
    
    def _on_save(self):
        """保存ボタンをクリック"""
        today = datetime.now().strftime("%Y-%m-%d")
        
        # 入力値をデータに反映（変更があった項目のみ updated_at を更新）
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
                    return
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
                    return
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
                    return
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
                    return
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
                        return
                fee_data = self.treatment_fees[fee_key]
                changed = (fee_data.get("fee") != fee)
                fee_data["fee"] = fee
                if changed:
                    fee_data["updated_at"] = today
        
        # 保存
        if self._save_settings():
            messagebox.showinfo("保存完了", "農場薬品設定を保存しました。")
            # 更新日付を表示に反映
            self._load_data_to_widgets()
    
    def _on_close(self):
        """閉じるボタンをクリック"""
        self.window.destroy()
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()
