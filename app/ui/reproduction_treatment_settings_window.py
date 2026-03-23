"""
FALCON2 - 繁殖処置設定ウインドウ（農場設定＞繁殖処置設定）
ここで設定するのは「処置」のみ（WPG, CIDR, GN, PG, E2 等）。子宮OK・NS 等は所見であり、
入力画面の子宮・右・左・その他は所見欄のため、この一覧には含めない。
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from typing import Dict, Any, Optional, List
import json
import logging


class ReproductionTreatmentSettingsWindow:
    """繁殖処置設定ウインドウ"""
    
    def __init__(self, parent: tk.Tk, farm_path: Path):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ（FarmSettingsWindow）
            farm_path: 農場フォルダのパス
        """
        self.parent = parent
        self.farm_path = Path(farm_path)
        self.settings_file = self.farm_path / "reproduction_treatment_settings.json"
        
        # 設定データ
        self.treatments: Dict[str, Dict[str, Any]] = {}  # {"1": {"code": "1", "name": "WPG", "protocols": [...]}}
        
        # 薬品設定を読み込む
        self.drug_settings_file = self.farm_path / "drug_settings.json"
        self.drugs: Dict[str, List[Dict[str, Any]]] = {
            "pg": [],
            "gnrh": [],
            "vaginal_insert": [],
            "other": []
        }
        self.treatment_fees: Dict[str, Dict[str, Any]] = {}
        self._load_drug_settings()
        
        # ウィンドウ作成
        self.window = tk.Toplevel(parent)
        self.window.title("繁殖処置設定")
        self.window.geometry("1000x800")  # ウィンドウサイズを拡大
        
        # 設定をロード
        self._load_settings()
        
        # UI作成
        self._create_widgets()
        
        # ウィンドウを中央に配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
        
        # 初期選択（最初の項目を選択しない）
        self._refresh_tree()
        # 初期選択は行わない（ユーザーが選択するまで待つ）
    
    def _load_settings(self):
        """設定をロード"""
        if self.settings_file.exists():
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    treatments_raw = data.get('treatments', {})
                    # キーを文字列に統一
                    self.treatments = {}
                    for key, value in treatments_raw.items():
                        self.treatments[str(key)] = value
                    logging.info(f"繁殖処置設定を読み込みました: {list(self.treatments.keys())}")
            except Exception as e:
                logging.error(f"繁殖処置設定ファイル読み込みエラー: {e}")
                messagebox.showerror("エラー", f"設定ファイルの読み込みに失敗しました: {e}")
                self.treatments = {}
        else:
            # ファイルが存在しない場合は空の初期データ
            self.treatments = {}
            logging.info("繁殖処置設定ファイルが存在しません")
    
    def _load_drug_settings(self):
        """薬品設定を読み込む"""
        if self.drug_settings_file.exists():
            try:
                with open(self.drug_settings_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # 薬品データをロード
                    for category in ["pg", "gnrh", "vaginal_insert", "other"]:
                        if category in data:
                            self.drugs[category] = data[category]
                    # 処置料データをロード
                    if "treatment_fees" in data:
                        self.treatment_fees = data["treatment_fees"]
                logging.info("薬品設定を読み込みました")
            except Exception as e:
                logging.error(f"薬品設定ファイル読み込みエラー: {e}")
                self.drugs = {
                    "pg": [],
                    "gnrh": [],
                    "vaginal_insert": [],
                    "other": []
                }
                self.treatment_fees = {}
        else:
            logging.info("薬品設定ファイルが存在しません")
    
    def _get_all_drugs_list(self) -> List[tuple]:
        """
        すべての薬品をリスト形式で取得（プルダウン用）
        
        Returns:
            [(category, index, name), ...] の形式
        """
        drugs_list = []
        category_labels = {
            "pg": "PG",
            "gnrh": "GnRH",
            "vaginal_insert": "膣内留置剤",
            "other": "その他"
        }
        
        for category in ["pg", "gnrh", "vaginal_insert", "other"]:
            for i, drug in enumerate(self.drugs[category]):
                name = drug.get("name", "").strip()
                if name:
                    label = f"{category_labels[category]}{i+1}: {name}"
                    drugs_list.append((category, i, label))
        
        return drugs_list
    
    def _update_drug_combos(self):
        """薬品プルダウンのリストを更新"""
        drugs_list = self._get_all_drugs_list()
        drug_values = [label for _, _, label in drugs_list]
        
        for entry in self.drug_entries:
            current_value = entry["drug"].get()
            entry["drug"]["values"] = [""] + drug_values
            # 現在の値を維持（存在する場合）
            if current_value and current_value in drug_values:
                entry["drug"].set(current_value)
            else:
                entry["drug"].set("")
    
    def _calculate_totals(self):
        """処置料合計、薬品量合計、総合計を計算して表示"""
        # 処置料合計
        treatment_fee_total = 0
        for fee_key, var in self.treatment_checkboxes.items():
            if var.get():
                fee_data = self.treatment_fees.get(fee_key, {})
                fee = fee_data.get("fee")
                if fee is not None:
                    treatment_fee_total += fee
        
        # 薬品量合計
        drug_total = 0
        for entry in self.drug_entries:
            drug_value = entry["drug"].get()
            volume_str = entry["volume"].get().strip()
            
            if drug_value and volume_str:
                try:
                    volume = float(volume_str)
                    # 薬品情報を取得
                    drugs_list = self._get_all_drugs_list()
                    for category, index, label in drugs_list:
                        if label == drug_value:
                            drug_data = self.drugs[category][index]
                            unit_price = drug_data.get("unit_price")
                            if unit_price is not None:
                                drug_total += unit_price * volume
                            break
                except ValueError:
                    pass
        
        # 総合計
        grand_total = treatment_fee_total + drug_total
        
        # 表示を更新
        self.treatment_fee_total_label.config(text=f"処置料合計: {treatment_fee_total:,.0f}円")
        self.drug_total_label.config(text=f"薬品量合計: {drug_total:,.0f}円")
        self.grand_total_label.config(text=f"総合計: {grand_total:,.0f}円")
    
    def _save_settings(self):
        """設定を保存"""
        data = {
            "treatments": self.treatments
        }
        
        try:
            self.farm_path.mkdir(parents=True, exist_ok=True)
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            logging.info("繁殖処置設定を保存しました")
            return True
        except Exception as e:
            logging.error(f"繁殖処置設定ファイル保存エラー: {e}")
            messagebox.showerror("エラー", f"設定ファイルの保存に失敗しました: {e}")
            return False
    
    def _create_widgets(self):
        """ウィジェットを作成"""
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左側：処置リスト
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        list_label = ttk.Label(left_frame, text="処置一覧", font=("", 10, "bold"))
        list_label.pack(pady=(0, 5))
        
        # Treeview
        tree_frame = ttk.Frame(left_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(tree_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.treatment_tree = ttk.Treeview(
            tree_frame,
            columns=("code", "name"),
            show="headings",
            yscrollcommand=scrollbar.set,
            height=20
        )
        self.treatment_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.treatment_tree.yview)
        
        self.treatment_tree.heading("code", text="コード")
        self.treatment_tree.heading("name", text="処置内容")
        self.treatment_tree.column("code", width=80)
        self.treatment_tree.column("name", width=200)
        
        # 選択イベント
        self.treatment_tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        # ダブルクリックイベント
        self.treatment_tree.bind("<Double-Button-1>", self._on_tree_double_click)
        
        # 右クリックメニュー
        tree_menu = tk.Menu(self.window, tearoff=0)
        tree_menu.add_command(label="編集", command=self._edit_treatment)
        tree_menu.add_command(label="削除", command=self._delete_treatment)
        self.treatment_tree.bind("<Button-3>", lambda e: self._show_context_menu(e, tree_menu))
        
        # ボタンフレーム
        button_frame = ttk.Frame(left_frame)
        button_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(
            button_frame,
            text="追加",
            command=self._add_treatment
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame,
            text="編集",
            command=self._edit_treatment
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame,
            text="削除",
            command=self._delete_treatment
        ).pack(side=tk.LEFT, padx=5)
        
        # 右側：処置詳細編集
        right_frame = ttk.LabelFrame(main_frame, text="処置詳細", padding=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 処置コード
        code_frame = ttk.Frame(right_frame)
        code_frame.pack(fill=tk.X, pady=5)
        ttk.Label(code_frame, text="処置コード:").pack(side=tk.LEFT, padx=(0, 5))
        self.code_entry = ttk.Entry(code_frame, width=20)
        self.code_entry.pack(side=tk.LEFT)
        
        # 処置内容
        name_frame = ttk.Frame(right_frame)
        name_frame.pack(fill=tk.X, pady=5)
        ttk.Label(name_frame, text="処置内容:").pack(side=tk.LEFT, padx=(0, 5))
        self.name_entry = ttk.Entry(name_frame, width=20)
        self.name_entry.pack(side=tk.LEFT)
        
        # プロトコール設定
        protocol_label = ttk.Label(right_frame, text="プロトコール（処置日からの日数と指示）", font=("", 9, "bold"))
        protocol_label.pack(pady=(10, 5))
        
        # プロトコール入力欄（6つ）
        self.protocol_entries = []
        for i in range(6):
            protocol_frame = ttk.Frame(right_frame)
            protocol_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(protocol_frame, text=f"{i+1}:").pack(side=tk.LEFT, padx=(0, 5))
            
            days_label = ttk.Label(protocol_frame, text="日数:")
            days_label.pack(side=tk.LEFT, padx=(0, 2))
            days_entry = ttk.Entry(protocol_frame, width=8)
            days_entry.pack(side=tk.LEFT, padx=(0, 10))
            
            # 最後（6番目）は「AI」、それ以外は「指示」
            if i == 5:
                instruction_label = ttk.Label(protocol_frame, text="AI:")
                instruction_label.pack(side=tk.LEFT, padx=(0, 2))
                instruction_entry = ttk.Entry(protocol_frame, width=20, state='readonly')
                instruction_entry.insert(0, "AI")  # デフォルトで「AI」を設定
                instruction_entry.pack(side=tk.LEFT)
            else:
                instruction_label = ttk.Label(protocol_frame, text="指示:")
                instruction_label.pack(side=tk.LEFT, padx=(0, 2))
                instruction_entry = ttk.Entry(protocol_frame, width=20)
                instruction_entry.pack(side=tk.LEFT)
            
            self.protocol_entries.append({
                "days": days_entry,
                "instruction": instruction_entry
            })
        
        # 処置チェックボックス
        treatment_label = ttk.Label(right_frame, text="処置", font=("", 9, "bold"))
        treatment_label.pack(pady=(10, 5))
        
        treatment_frame = ttk.Frame(right_frame)
        treatment_frame.pack(fill=tk.X, pady=5)
        
        self.treatment_checkboxes = {}
        treatment_types = [
            ("ultrasound", "超音波検査"),
            ("injection", "筋肉注射"),
            ("vaginal_insert", "膣内留置"),
            ("drug_administration", "薬治")
        ]
        
        for fee_key, fee_label in treatment_types:
            var = tk.BooleanVar()
            checkbox = ttk.Checkbutton(treatment_frame, text=fee_label, variable=var, command=self._calculate_totals)
            checkbox.pack(side=tk.LEFT, padx=5)
            self.treatment_checkboxes[fee_key] = var
        
        # 使用薬品
        drug_label = ttk.Label(right_frame, text="使用薬品（最大5つ）", font=("", 9, "bold"))
        drug_label.pack(pady=(10, 5))
        
        self.drug_entries = []
        for i in range(5):
            drug_frame = ttk.Frame(right_frame)
            drug_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(drug_frame, text=f"薬品{i+1}:").pack(side=tk.LEFT, padx=(0, 5))
            
            # 薬品選択プルダウン
            drug_combo = ttk.Combobox(drug_frame, width=25, state="readonly")
            drug_combo.pack(side=tk.LEFT, padx=5)
            
            ttk.Label(drug_frame, text="容量(ml):").pack(side=tk.LEFT, padx=5)
            
            # 容量入力
            volume_entry = ttk.Entry(drug_frame, width=10)
            volume_entry.pack(side=tk.LEFT, padx=5)
            
            # 変更時に合計を再計算
            drug_combo.bind("<<ComboboxSelected>>", lambda e, idx=i: self._calculate_totals())
            volume_entry.bind("<KeyRelease>", lambda e, idx=i: self._calculate_totals())
            
            self.drug_entries.append({
                "drug": drug_combo,
                "volume": volume_entry
            })
        
        # 合計表示
        total_frame = ttk.LabelFrame(right_frame, text="合計", padding=10)
        total_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.treatment_fee_total_label = ttk.Label(total_frame, text="処置料合計: 0円", font=("", 9))
        self.treatment_fee_total_label.pack(anchor=tk.W, pady=2)
        
        self.drug_total_label = ttk.Label(total_frame, text="薬品量合計: 0円", font=("", 9))
        self.drug_total_label.pack(anchor=tk.W, pady=2)
        
        self.grand_total_label = ttk.Label(total_frame, text="総合計: 0円", font=("", 10, "bold"))
        self.grand_total_label.pack(anchor=tk.W, pady=(5, 0))
        
        # 薬品リストを更新
        self._update_drug_combos()
        
        # 保存ボタン
        save_button = ttk.Button(
            right_frame,
            text="保存",
            command=self._save_current_treatment
        )
        save_button.pack(pady=(10, 0))
        
        # 現在選択中の処置コード
        self.current_treatment_code = None
        
        # OK / キャンセル
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(
            bottom_frame,
            text="OK",
            command=self._on_ok
        ).pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(
            bottom_frame,
            text="キャンセル",
            command=self._on_cancel
        ).pack(side=tk.RIGHT, padx=5)
    
    def _refresh_tree(self):
        """Treeviewを更新"""
        for item in self.treatment_tree.get_children():
            self.treatment_tree.delete(item)
        
        for code, treatment in sorted(self.treatments.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 999):
            name = treatment.get('name', '')
            self.treatment_tree.insert("", tk.END, values=(code, name), tags=(code,))
    
    def _on_tree_select(self, event):
        """Treeview選択時の処理"""
        selection = self.treatment_tree.selection()
        if selection:
            item_id = selection[0]
            item = self.treatment_tree.item(item_id)
            # tagsから取得を試みる、なければvaluesから取得
            code = None
            if item.get('tags') and len(item['tags']) > 0:
                code = item['tags'][0]
            elif item.get('values') and len(item['values']) > 0:
                code = item['values'][0]
            
            logging.info(f"_on_tree_select: item_id={item_id}, item={item}, code={code}")
            if code:
                self._select_treatment(code)
    
    def _on_tree_double_click(self, event):
        """Treeviewダブルクリック時の処理"""
        selection = self.treatment_tree.selection()
        if selection:
            item_id = selection[0]
            item = self.treatment_tree.item(item_id)
            # tagsから取得を試みる、なければvaluesから取得
            code = None
            if item.get('tags') and len(item['tags']) > 0:
                code = item['tags'][0]
            elif item.get('values') and len(item['values']) > 0:
                code = item['values'][0]
            
            logging.info(f"_on_tree_double_click: item_id={item_id}, item={item}, code={code}")
            if code:
                self._select_treatment(code)
    
    def _select_treatment(self, code: str):
        """処置を選択して詳細を表示"""
        if not code:
            logging.warning("_select_treatment: codeが空です")
            return
        
        # コードを文字列に統一
        code = str(code).strip()
        
        logging.info(f"_select_treatment: code={code} (type={type(code)}), treatments={list(self.treatments.keys())} (types={[type(k) for k in self.treatments.keys()]})")
        self.current_treatment_code = code
        treatment = self.treatments.get(code, {})
        
        if not treatment:
            # 型が異なる可能性があるので、すべてのキーを文字列に変換して再試行
            for key, value in self.treatments.items():
                if str(key) == code:
                    treatment = value
                    logging.info(f"_select_treatment: 型変換後に見つかりました: {key} -> {code}")
                    break
            
            if not treatment:
                logging.warning(f"_select_treatment: 処置コード '{code}' が見つかりません。利用可能なコード: {list(self.treatments.keys())}")
                return
        
        # コードと名称を設定
        self.code_entry.config(state='normal')
        self.code_entry.delete(0, tk.END)
        self.code_entry.insert(0, code)
        self.code_entry.config(state='readonly')  # コードは編集不可
        
        # 処置内容は編集可能にする
        self.name_entry.config(state='normal')
        self.name_entry.delete(0, tk.END)
        self.name_entry.insert(0, treatment.get('name', ''))
        
        # プロトコールを設定（編集可能にする）
        # まず位置情報を含む配列を試す（新形式）
        protocols_by_position = treatment.get('protocols_by_position', None)
        if protocols_by_position is not None and len(protocols_by_position) == 6:
            # 新形式：位置情報を含む配列
            protocols = protocols_by_position
        else:
            # 旧形式：順番に並んだリスト（後方互換性）
            protocols_compact = treatment.get('protocols', [])
            # 6つの位置に対応する配列に変換
            protocols = [None] * 6
            for i, protocol in enumerate(protocols_compact):
                if i < 6:
                    protocols[i] = protocol
        
        for i, entry_pair in enumerate(self.protocol_entries):
            # 日数入力欄を編集可能にする
            entry_pair['days'].config(state='normal')
            entry_pair['days'].delete(0, tk.END)
            
            # 6番目（最後）の項目は「AI」固定
            if i == 5:
                entry_pair['instruction'].config(state='readonly')
                entry_pair['instruction'].delete(0, tk.END)
                entry_pair['instruction'].insert(0, "AI")
            else:
                # 指示入力欄を編集可能にする
                entry_pair['instruction'].config(state='normal')
                entry_pair['instruction'].delete(0, tk.END)
            
            # 位置iに対応するプロトコールを設定
            if i < len(protocols) and protocols[i] is not None:
                protocol = protocols[i]
                entry_pair['days'].insert(0, str(protocol.get('days', '')))
                # 6番目以外は指示を設定、6番目は「AI」固定
                if i != 5:
                    entry_pair['instruction'].insert(0, protocol.get('instruction', ''))
        
        # 処置チェックボックスを設定
        treatments = treatment.get('treatments', [])
        for fee_key, var in self.treatment_checkboxes.items():
            var.set(fee_key in treatments)
        
        # 薬品データを設定
        drugs = treatment.get('drugs', [])
        for i, entry in enumerate(self.drug_entries):
            if i < len(drugs):
                drug_data = drugs[i]
                # 薬品名を取得してプルダウンに設定
                category = drug_data.get('category')
                index = drug_data.get('index')
                if category is not None and index is not None:
                    drugs_list = self._get_all_drugs_list()
                    for cat, idx, label in drugs_list:
                        if cat == category and idx == index:
                            entry["drug"].set(label)
                            break
                else:
                    entry["drug"].set("")
                
                # 容量を設定
                volume = drug_data.get('volume')
                entry["volume"].delete(0, tk.END)
                if volume is not None:
                    entry["volume"].insert(0, str(volume))
            else:
                entry["drug"].set("")
                entry["volume"].delete(0, tk.END)
        
        # 合計を再計算
        self._calculate_totals()
    
    def _clear_editing_area(self):
        """編集エリアをクリア"""
        self.current_treatment_code = None
        self.code_entry.config(state='normal')
        self.code_entry.delete(0, tk.END)
        self.name_entry.delete(0, tk.END)
        for i, entry_pair in enumerate(self.protocol_entries):
            entry_pair['days'].delete(0, tk.END)
            if i == 5:
                # 6番目は「AI」固定
                entry_pair['instruction'].config(state='readonly')
                entry_pair['instruction'].delete(0, tk.END)
                entry_pair['instruction'].insert(0, "AI")
            else:
                entry_pair['instruction'].config(state='normal')
                entry_pair['instruction'].delete(0, tk.END)
        
        # 処置チェックボックスをクリア
        for var in self.treatment_checkboxes.values():
            var.set(False)
        
        # 薬品をクリア
        for entry in self.drug_entries:
            entry["drug"].set("")
            entry["volume"].delete(0, tk.END)
        
        # 合計を再計算
        self._calculate_totals()
    
    def _show_context_menu(self, event, menu: tk.Menu):
        """右クリックメニューを表示"""
        item = event.widget.selection()[0] if event.widget.selection() else None
        if item:
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()
    
    def _add_treatment(self):
        """処置を追加"""
        dialog = tk.Toplevel(self.window)
        dialog.title("処置追加")
        dialog.geometry("300x100")
        
        ttk.Label(dialog, text="処置コード:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        code_entry = ttk.Entry(dialog, width=20)
        code_entry.grid(row=0, column=1, padx=5, pady=5)
        
        def on_ok():
            code = code_entry.get().strip()
            
            if not code:
                messagebox.showwarning("警告", "処置コードを入力してください")
                return
            
            if code in self.treatments:
                messagebox.showwarning("警告", "このコードは既に登録されています")
                return
            
            # 新規処置を作成
            self.treatments[code] = {
                "code": code,
                "name": "",
                "protocols": [],
                "treatments": [],
                "drugs": []
            }
            
            self._refresh_tree()
            self._select_treatment(code)
            dialog.destroy()
        
        ttk.Button(dialog, text="OK", command=on_ok).grid(row=1, column=0, columnspan=2, pady=10)
        
        # 中央に配置
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        code_entry.focus()
    
    def _edit_treatment(self):
        """処置を編集（選択を促す）"""
        selection = self.treatment_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "編集する処置を選択してください")
            return
        
        item_id = selection[0]
        item = self.treatment_tree.item(item_id)
        # tagsから取得を試みる、なければvaluesから取得
        code = None
        if item.get('tags') and len(item['tags']) > 0:
            code = item['tags'][0]
        elif item.get('values') and len(item['values']) > 0:
            code = item['values'][0]
        
        logging.info(f"_edit_treatment: item_id={item_id}, item={item}, code={code}")
        if code:
            self._select_treatment(code)
        else:
            logging.error(f"_edit_treatment: コードを取得できませんでした。item={item}")
    
    def _delete_treatment(self):
        """処置を削除"""
        selection = self.treatment_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "削除する処置を選択してください")
            return
        
        item = self.treatment_tree.item(selection[0])
        # tagsから取得を試みる、なければvaluesから取得
        code = None
        if item.get('tags') and len(item['tags']) > 0:
            code = item['tags'][0]
        elif item.get('values') and len(item['values']) > 0:
            code = item['values'][0]
        
        name = item['values'][1] if len(item['values']) > 1 else code
        
        if code and messagebox.askyesno("確認", f"処置コード「{code}: {name}」を削除しますか？"):
            del self.treatments[code]
            self._refresh_tree()
            self._clear_editing_area()
    
    def _save_current_treatment(self):
        """現在編集中の処置を保存"""
        if not self.current_treatment_code:
            messagebox.showwarning("警告", "保存する処置を選択してください")
            return
        
        code = self.code_entry.get().strip()
        name = self.name_entry.get().strip()
        
        if not code:
            messagebox.showwarning("警告", "処置コードを入力してください")
            return
        
        if not name:
            messagebox.showwarning("警告", "処置内容を入力してください")
            return
        
        # プロトコールを取得（6つの固定位置に対応）
        protocols = [None] * 6  # 6つの位置を初期化
        for i, entry_pair in enumerate(self.protocol_entries):
            days_str = entry_pair['days'].get().strip()
            
            # 6番目（最後）の項目は指示を「AI」に固定
            if i == 5:
                if days_str:
                    try:
                        days = int(days_str)
                        protocols[i] = {
                            "days": days,
                            "instruction": "AI"
                        }
                    except ValueError:
                        messagebox.showwarning("警告", "日数は数値で入力してください")
                        return
            else:
                instruction = entry_pair['instruction'].get().strip()
                if days_str and instruction:
                    try:
                        days = int(days_str)
                        protocols[i] = {
                            "days": days,
                            "instruction": instruction
                        }
                    except ValueError:
                        messagebox.showwarning("警告", "日数は数値で入力してください")
                        return
                elif days_str or instruction:
                    # 片方だけ入力されている場合は警告
                    messagebox.showwarning("警告", "日数と指示の両方を入力してください")
                    return
        
        # Noneの要素を削除して、順番に並べたリストに変換（後方互換性のため）
        protocols_compact = [p for p in protocols if p is not None]
        
        # 処置を取得
        treatments = []
        for fee_key, var in self.treatment_checkboxes.items():
            if var.get():
                treatments.append(fee_key)
        
        # 薬品を取得
        drugs = []
        drugs_list = self._get_all_drugs_list()
        for entry in self.drug_entries:
            drug_value = entry["drug"].get()
            volume_str = entry["volume"].get().strip()
            
            if drug_value and volume_str:
                try:
                    volume = float(volume_str)
                    # 薬品情報を取得
                    for category, index, label in drugs_list:
                        if label == drug_value:
                            drugs.append({
                                "category": category,
                                "index": index,
                                "volume": volume
                            })
                            break
                except ValueError:
                    messagebox.showwarning("警告", "容量は数値で入力してください")
                    return
        
        # 処置を更新
        self.treatments[code] = {
            "code": code,
            "name": name,
            "protocols": protocols_compact,  # 後方互換性のため、Noneを除いたリストを保存
            "protocols_by_position": protocols,  # 位置情報を含む完全な配列も保存
            "treatments": treatments,
            "drugs": drugs
        }
        
        # コードが変更された場合
        if code != self.current_treatment_code:
            if self.current_treatment_code in self.treatments:
                del self.treatments[self.current_treatment_code]
            self.current_treatment_code = code
        
        # ファイルに保存
        if self._save_settings():
            self._refresh_tree()
            # 保存後に再度選択して表示を更新
            self._select_treatment(code)
            messagebox.showinfo("完了", "処置を保存しました")
        else:
            messagebox.showerror("エラー", "処置の保存に失敗しました")
    
    def _on_ok(self):
        """OKボタンクリック時の処理"""
        if self._save_settings():
            messagebox.showinfo("完了", "設定を保存しました")
            self.window.destroy()
    
    def _on_cancel(self):
        """キャンセルボタンクリック時の処理"""
        self.window.destroy()
    
    def show(self):
        """ウィンドウを表示"""
        self.window.wait_window()
