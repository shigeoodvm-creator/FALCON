"""
FALCON2 - 導入ウィンドウ
外部からの導入イベント入力用（複数個体対応）
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, Any, Optional, List
from pathlib import Path
from datetime import datetime
import json
import logging

from db.db_handler import DBHandler
from modules.rule_engine import RuleEngine
from modules.app_settings_manager import get_app_settings_manager
from settings_manager import SettingsManager
from ui.date_picker_window import DatePickerWindow


class IntroductionWindow:
    """導入ウィンドウ（複数個体対応）"""
    
    def __init__(self, parent: tk.Widget, db_handler: DBHandler,
                 rule_engine: RuleEngine,
                 farm_path: Path,
                 event_dictionary_path: Optional[Path] = None):
        """
        初期化
        
        Args:
            parent: 親ウィジェット
            db_handler: DBHandler インスタンス
            rule_engine: RuleEngine インスタンス
            farm_path: 農場フォルダのパス
            event_dictionary_path: event_dictionary.json のパス
        """
        self.parent = parent
        self.db = db_handler
        self.rule_engine = rule_engine
        self.farm_path = Path(farm_path)
        self.event_dict_path = event_dictionary_path
        
        # 導入日（共通）
        self.introduction_date = datetime.now().strftime('%Y-%m-%d')
        
        # 個体リスト（複数個体対応）
        self.cow_list: List[Dict[str, Any]] = []
        
        # SettingsManagerを初期化
        self.settings_manager = SettingsManager(self.farm_path)
        
        # ウィンドウ作成
        self.window = tk.Toplevel(parent)
        self.window.title("導入")
        self.window.geometry("1000x900")
        self.window.minsize(800, 750)
        self.window.configure(bg="#f5f5f5")
        
        # UI作成
        self._create_widgets()
        
        # ウィンドウを中央に配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def _create_widgets(self):
        """ウィジェットを作成（イベント入力ウィンドウとデザイン統一）"""
        _df = "Meiryo UI"
        bg = "#f5f5f5"
        # イベント入力と同じスタイル
        try:
            _style = ttk.Style()
            _style.configure("EventInput.TLabelframe", borderwidth=0)
            _style.configure("EventInput.TLabelframe.Label", font=(_df, 10, "bold"))
            try:
                _input_font_size = max(11, get_app_settings_manager().get_font_size() + 3)
            except Exception:
                _input_font_size = 12
            _style.configure("EventInput.TEntry", font=(_df, _input_font_size))
            _style.configure("EventInput.TCombobox", font=(_df, _input_font_size))
            _style.configure("EventInput.Hint.TLabel", font=(_df, 8), foreground="#78909c")
        except tk.TclError:
            pass
        _ENTRY_WIDTH = 24
        _label_col_minsize = 100

        main_container = tk.Frame(self.window, bg=bg, padx=24, pady=16)
        main_container.pack(fill=tk.BOTH, expand=True)

        # ヘッダー（タイトル左・登録/閉じる右）
        header = tk.Frame(main_container, bg=bg, pady=12)
        header.pack(fill=tk.X)
        tk.Label(header, text="\U0001f4e5", font=(_df, 22), bg=bg, fg="#3949ab").pack(side=tk.LEFT, padx=(0, 10))
        title_frame = tk.Frame(header, bg=bg)
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(title_frame, text="導入", font=(_df, 16, "bold"), bg=bg, fg="#263238").pack(anchor=tk.W)
        tk.Label(
            title_frame,
            text="JPN10（個体識別番号・10桁）を先に入力するとIDが自動で設定されます。",
            font=(_df, 10),
            bg=bg,
            fg="#607d8b",
        ).pack(anchor=tk.W)
        header_right = tk.Frame(header, bg=bg)
        header_right.pack(side=tk.RIGHT, padx=(10, 0))
        ttk.Button(header_right, text="登録", command=self._save_all, width=8).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(header_right, text="閉じる", command=self.window.destroy, width=10).pack(side=tk.LEFT)

        content_area = tk.Frame(main_container, bg=bg)
        content_area.pack(fill=tk.BOTH, expand=True)

        # 導入日（カレンダーボタン付き）
        date_frame = ttk.LabelFrame(
            content_area, text="導入日", padding=(12, 10), style="EventInput.TLabelframe"
        )
        date_frame.pack(fill=tk.X, pady=(0, 14))
        date_frame.columnconfigure(0, minsize=_label_col_minsize)
        date_frame.columnconfigure(1, weight=0, minsize=280)
        ttk.Label(date_frame, text="導入日*:").grid(row=0, column=0, sticky=tk.W, padx=(5, 8), pady=(6, 2))
        date_row = ttk.Frame(date_frame)
        date_row.grid(row=0, column=1, sticky=tk.W, padx=(0, 5), pady=(6, 2))
        self.date_entry = ttk.Entry(date_row, width=_ENTRY_WIDTH, style="EventInput.TEntry")
        self.date_entry.pack(side=tk.LEFT, padx=(0, 4))
        self.date_entry.insert(0, self.introduction_date)
        ttk.Button(date_row, text="\U0001f4c5", width=3, command=lambda: self._open_date_picker(self.date_entry)).pack(side=tk.LEFT)
        ttk.Label(date_frame, text="YYYY-MM-DD または \U0001f4c5で選択", style="EventInput.Hint.TLabel").grid(
            row=1, column=1, sticky=tk.W, padx=(0, 5), pady=(0, 4)
        )

        # 個体入力（2列レイアウト・ブロック分け・日付はカレンダー付き・IDは自動専用）
        # expand=False で下の「登録予定の個体」が確実に見えるようにする
        input_frame = ttk.LabelFrame(
            content_area, text="個体入力", padding=(12, 10), style="EventInput.TLabelframe"
        )
        input_frame.pack(fill=tk.X, pady=(0, 10))
        fields_frame = ttk.Frame(input_frame)
        fields_frame.pack(fill=tk.X, pady=5)
        fields_frame.columnconfigure(0, minsize=_label_col_minsize)
        fields_frame.columnconfigure(1, weight=0, minsize=320)

        _row = 0
        # 識別
        ttk.Label(fields_frame, text="JPN10*:").grid(row=_row, column=0, sticky=tk.W, padx=(5, 8), pady=(4, 2))
        self.jpn10_entry = ttk.Entry(fields_frame, width=18, style="EventInput.TEntry")
        self.jpn10_entry.grid(row=_row, column=1, sticky=tk.W, padx=(0, 5), pady=(4, 2))
        self.jpn10_entry.bind("<KeyRelease>", self._on_jpn10_changed)
        _row += 1
        ttk.Label(fields_frame, text="ID（自動）:").grid(row=_row, column=0, sticky=tk.W, padx=(5, 8), pady=2)
        self.cow_id_entry = ttk.Entry(fields_frame, width=10, style="EventInput.TEntry")
        self.cow_id_entry.grid(row=_row, column=1, sticky=tk.W, padx=(0, 5), pady=2)
        self.cow_id_entry.config(state="disabled")  # JPN10入力で自動設定されるため編集不可
        _row += 1
        ttk.Label(fields_frame, text="JPN10（10桁）を先に入力するとIDが自動で設定されます", style="EventInput.Hint.TLabel").grid(
            row=_row, column=1, sticky=tk.W, padx=(0, 5), pady=(0, 4))
        _row += 1
        ttk.Label(fields_frame, text="品種:").grid(row=_row, column=0, sticky=tk.W, padx=(5, 8), pady=2)
        self._breed_options = ["ホルスタイン", "ジャージー", "その他"]
        self.breed_entry = ttk.Combobox(
            fields_frame, values=self._breed_options, state="readonly", width=14, style="EventInput.TCombobox"
        )
        self.breed_entry.grid(row=_row, column=1, sticky=tk.W, padx=(0, 5), pady=2)
        _row += 1
        ttk.Label(fields_frame, text="生年月日:").grid(row=_row, column=0, sticky=tk.W, padx=(5, 8), pady=2)
        birth_row = ttk.Frame(fields_frame)
        birth_row.grid(row=_row, column=1, sticky=tk.W, padx=(0, 5), pady=2)
        self.birth_date_entry = ttk.Entry(birth_row, width=14, style="EventInput.TEntry")
        self.birth_date_entry.pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(birth_row, text="\U0001f4c5", width=3, command=lambda: self._open_date_picker(self.birth_date_entry)).pack(side=tk.LEFT)
        _row += 1
        # 繁殖
        ttk.Label(fields_frame, text="産次:").grid(row=_row, column=0, sticky=tk.W, padx=(5, 8), pady=2)
        self.lact_entry = ttk.Entry(fields_frame, width=8, style="EventInput.TEntry")
        self.lact_entry.grid(row=_row, column=1, sticky=tk.W, padx=(0, 5), pady=2)
        _row += 1
        ttk.Label(fields_frame, text="分娩月日:").grid(row=_row, column=0, sticky=tk.W, padx=(5, 8), pady=2)
        clvd_row = ttk.Frame(fields_frame)
        clvd_row.grid(row=_row, column=1, sticky=tk.W, padx=(0, 5), pady=2)
        self.clvd_entry = ttk.Entry(clvd_row, width=14, style="EventInput.TEntry")
        self.clvd_entry.pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(clvd_row, text="\U0001f4c5", width=3, command=lambda: self._open_date_picker(self.clvd_entry)).pack(side=tk.LEFT)
        _row += 1
        ttk.Label(fields_frame, text="最終授精日:").grid(row=_row, column=0, sticky=tk.W, padx=(5, 8), pady=2)
        last_ai_row = ttk.Frame(fields_frame)
        last_ai_row.grid(row=_row, column=1, sticky=tk.W, padx=(0, 5), pady=2)
        self.last_ai_date_entry = ttk.Entry(last_ai_row, width=14, style="EventInput.TEntry")
        self.last_ai_date_entry.pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(last_ai_row, text="\U0001f4c5", width=3, command=lambda: self._open_date_picker(self.last_ai_date_entry)).pack(side=tk.LEFT)
        _row += 1
        ttk.Label(fields_frame, text="最終授精SIRE:").grid(row=_row, column=0, sticky=tk.W, padx=(5, 8), pady=2)
        self.last_ai_sire_entry = ttk.Entry(fields_frame, width=14, style="EventInput.TEntry")
        self.last_ai_sire_entry.grid(row=_row, column=1, sticky=tk.W, padx=(0, 5), pady=2)
        _row += 1
        ttk.Label(fields_frame, text="授精回数:").grid(row=_row, column=0, sticky=tk.W, padx=(5, 8), pady=2)
        self.ai_count_entry = ttk.Entry(fields_frame, width=8, style="EventInput.TEntry")
        self.ai_count_entry.grid(row=_row, column=1, sticky=tk.W, padx=(0, 5), pady=2)
        _row += 1
        # 状態
        ttk.Label(fields_frame, text="受胎の有無:").grid(row=_row, column=0, sticky=tk.W, padx=(5, 8), pady=2)
        pregnant_frame = ttk.Frame(fields_frame)
        pregnant_frame.grid(row=_row, column=1, sticky=tk.W, padx=(0, 5), pady=2)
        self.pregnant_var = tk.StringVar(value="")
        ttk.Radiobutton(pregnant_frame, text="妊娠鑑定待ち", variable=self.pregnant_var, value="waiting").pack(side=tk.LEFT)
        ttk.Radiobutton(pregnant_frame, text="受胎なし", variable=self.pregnant_var, value="").pack(side=tk.LEFT)
        ttk.Radiobutton(pregnant_frame, text="妊娠中", variable=self.pregnant_var, value="pregnant").pack(side=tk.LEFT)
        ttk.Radiobutton(pregnant_frame, text="乾乳中", variable=self.pregnant_var, value="dry").pack(side=tk.LEFT)
        _row += 1
        ttk.Label(fields_frame, text="PEN:").grid(row=_row, column=0, sticky=tk.W, padx=(5, 8), pady=2)
        self.pen_entry = ttk.Combobox(fields_frame, width=14, state="normal", style="EventInput.TCombobox")
        self.pen_entry.grid(row=_row, column=1, sticky=tk.W, padx=(0, 5), pady=2)
        self._load_pen_options()
        _row += 1
        ttk.Button(fields_frame, text="追加", command=self._add_cow).grid(row=_row, column=0, columnspan=2, pady=12)
        ttk.Label(
            fields_frame,
            text="※「追加」で一覧に載せます。確定するには上部の「登録」ボタンを押してください。",
            style="EventInput.Hint.TLabel",
        ).grid(row=_row + 1, column=0, columnspan=2, sticky=tk.W, padx=(5, 0), pady=(0, 8))

        # 登録予定の個体（必ず見えるように上部は expand しない）
        list_frame = ttk.LabelFrame(
            content_area, text="登録予定の個体", padding=(12, 10), style="EventInput.TLabelframe"
        )
        list_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        columns = ("ID", "JPN10", "品種", "生年月日", "産次", "分娩月日", "最終授精日", "最終授精SIRE", "授精回数", "受胎", "PEN")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=10)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=80)
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        ttk.Button(list_frame, text="選択を削除", command=self._delete_selected).pack(pady=5)
    
    def _open_date_picker(self, entry: ttk.Entry) -> None:
        """日付選択カレンダーを開き、選択した日付をEntryにセットする"""
        current = entry.get().strip()
        initial = current if current else None
        try:
            picker = DatePickerWindow(
                parent=self.window,  # type: ignore[arg-type]
                initial_date=initial,
                on_date_selected=lambda date_str: entry.delete(0, tk.END) or entry.insert(0, date_str),
            )
            picker.show()
        except Exception as e:
            logging.warning(f"DatePickerWindow open failed: {e}")
            messagebox.showwarning("日付選択", "カレンダーを開けませんでした。")

    def _load_pen_options(self):
        """PEN設定をロードしてプルダウンに設定"""
        try:
            pen_settings = self.settings_manager.load_pen_settings()
            # プルダウン用のリストを作成（コード: 名前 の形式）
            pen_options = []
            for code, name in sorted(pen_settings.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 999):
                pen_options.append(f"{code}: {name}")
            self.pen_entry['values'] = pen_options
        except Exception as e:
            logging.error(f"PEN設定の読み込みに失敗しました: {e}")
            self.pen_entry['values'] = []
    
    def _normalize_date(self, date_str: str) -> Optional[str]:
        """
        日付文字列を正規化（YYYY-MM-DD形式に統一）
        
        Args:
            date_str: 日付文字列（例: "2022-2-2", "2022-02-02"）
        
        Returns:
            正規化された日付文字列（YYYY-MM-DD形式）、無効な場合はNone
        """
        if not date_str or not date_str.strip():
            return None
        
        date_str = date_str.strip()
        
        # 既にYYYY-MM-DD形式の場合はそのまま返す
        try:
            dt = datetime.strptime(date_str, '%Y-%m-%d')
            return dt.strftime('%Y-%m-%d')
        except ValueError:
            pass
        
        # YYYY-M-D形式を試す
        try:
            dt = datetime.strptime(date_str, '%Y-%m-%d')
            return dt.strftime('%Y-%m-%d')
        except ValueError:
            pass
        
        # YYYY-M-D形式を手動で処理
        parts = date_str.split('-')
        if len(parts) == 3:
            try:
                year = int(parts[0])
                month = int(parts[1])
                day = int(parts[2])
                dt = datetime(year, month, day)
                return dt.strftime('%Y-%m-%d')
            except (ValueError, TypeError):
                pass
        
        return None
    
    def _on_jpn10_changed(self, event):
        """JPN10入力変更時の処理（IDを自動取得・クリア）"""
        jpn10_value = self.jpn10_entry.get().strip()
        if jpn10_value.isdigit() and len(jpn10_value) == 10:
            cow_id = jpn10_value[5:9].zfill(4)
            if not self.cow_id_entry.get().strip():
                self.cow_id_entry.config(state="normal")
                self.cow_id_entry.delete(0, tk.END)
                self.cow_id_entry.insert(0, cow_id)
                self.cow_id_entry.config(state="disabled")
        else:
            # 10桁でなくなったらIDをクリア
            if self.cow_id_entry.get().strip():
                self.cow_id_entry.config(state="normal")
                self.cow_id_entry.delete(0, tk.END)
                self.cow_id_entry.config(state="disabled")
    
    def _add_cow(self):
        """個体をリストに追加"""
        cow_id = self.cow_id_entry.get().strip()
        jpn10 = self.jpn10_entry.get().strip()
        birth_date = self.birth_date_entry.get().strip()
        breed = (self.breed_entry.get() or "").strip()
        lact = self.lact_entry.get().strip()
        clvd = self.clvd_entry.get().strip()
        last_ai_date = self.last_ai_date_entry.get().strip()
        last_ai_sire = self.last_ai_sire_entry.get().strip()
        ai_count = self.ai_count_entry.get().strip()
        pregnant = self.pregnant_var.get()
        pen_input = self.pen_entry.get().strip()
        
        # 日付を正規化
        if birth_date:
            birth_date = self._normalize_date(birth_date)
            if not birth_date:
                messagebox.showwarning("警告", "生年月日の形式が正しくありません（YYYY-MM-DD形式で入力してください）。")
                self.birth_date_entry.focus_set()
                return
        
        if clvd:
            clvd = self._normalize_date(clvd)
            if not clvd:
                messagebox.showwarning("警告", "分娩月日の形式が正しくありません（YYYY-MM-DD形式で入力してください）。")
                self.clvd_entry.focus_set()
                return
        
        if last_ai_date:
            last_ai_date = self._normalize_date(last_ai_date)
            if not last_ai_date:
                messagebox.showwarning("警告", "最終授精日の形式が正しくありません（YYYY-MM-DD形式で入力してください）。")
                self.last_ai_date_entry.focus_set()
                return
        
        # PEN入力の処理：プルダウン選択（"1: 搾乳1"）または直接入力（"1"）
        pen = None
        if pen_input:
            # "コード: 名前" 形式の場合はコード部分を抽出
            if ':' in pen_input:
                pen = pen_input.split(':')[0].strip()
            else:
                # 直接入力された場合はそのまま使用
                pen = pen_input.strip()
        
        # 必須項目チェック
        if not jpn10:
            messagebox.showwarning("警告", "JPN10を入力してください。")
            self.jpn10_entry.focus_set()
            return
        
        if not jpn10.isdigit() or len(jpn10) != 10:
            messagebox.showwarning("警告", "JPN10は10桁の数字で入力してください。")
            self.jpn10_entry.focus_set()
            return
        
        # IDが未入力の場合はJPN10から自動生成
        # JPN10の6-9桁目（1-indexed、0-indexedで5-9桁目）をIDとして使用
        # 例: 1630013859 -> 1385
        if not cow_id:
            cow_id = jpn10[5:9].zfill(4)
        
        # 産次が未入力の場合は0
        if not lact:
            lact = "0"
        
        # 受胎の有無に基づいてRCを決定（表示用、実際のRCはイベント作成時に決定される）
        rc = RuleEngine.RC_OPEN
        if pregnant == "pregnant":
            rc = RuleEngine.RC_PREGNANT
        elif pregnant == "dry":
            rc = RuleEngine.RC_DRY
        elif pregnant == "waiting" or (pregnant == "" and last_ai_date):
            # 妊娠鑑定待ちまたは受胎なしで授精歴がある場合は授精中
            rc = RuleEngine.RC_BRED
        
        cow_data = {
            'cow_id': cow_id,
            'jpn10': jpn10,
            'breed': breed if breed else None,
            'birth_date': birth_date,
            'lact': int(lact) if lact.isdigit() else 0,
            'clvd': clvd if clvd else None,
            'last_ai_date': last_ai_date if last_ai_date else None,
            'last_ai_sire': last_ai_sire if last_ai_sire else None,
            'ai_count': int(ai_count) if ai_count.isdigit() else 0,
            'pregnant_status': pregnant,  # 受胎の有無の状態を保存
            'rc': rc,
            'pen': pen if pen else None
        }
        
        self.cow_list.append(cow_data)
        self._update_tree()
        
        # 入力欄をクリア（IDとJPN10以外）
        self.birth_date_entry.delete(0, tk.END)
        self.breed_entry.set("")
        self.lact_entry.delete(0, tk.END)
        self.clvd_entry.delete(0, tk.END)
        self.last_ai_date_entry.delete(0, tk.END)
        self.last_ai_sire_entry.delete(0, tk.END)
        self.ai_count_entry.delete(0, tk.END)
        self.pregnant_var.set("")
        self.pen_entry.delete(0, tk.END)
    
    def _update_tree(self):
        """Treeviewを更新"""
        # 既存のアイテムを削除
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # PEN設定を取得（表示名変換用）
        try:
            pen_settings = self.settings_manager.load_pen_settings()
        except Exception:
            pen_settings = {}
        
        # 個体リストを表示
        for cow in self.cow_list:
            # PEN値を表示名に変換
            pen_value = cow.get('pen', '') or ''
            if pen_value and pen_value in pen_settings:
                pen_display = f"{pen_value}: {pen_settings[pen_value]}"
            else:
                pen_display = pen_value
            
            values = (
                cow.get('cow_id', ''),
                cow.get('jpn10', ''),
                cow.get('breed', '') or '',
                cow.get('birth_date', ''),
                str(cow.get('lact', 0)),
                cow.get('clvd', '') or '',
                cow.get('last_ai_date', '') or '',
                cow.get('last_ai_sire', '') or '',
                str(cow.get('ai_count', 0)),
                '妊娠鑑定待ち' if cow.get('pregnant_status') == 'waiting' else '妊娠中' if cow.get('rc') == RuleEngine.RC_PREGNANT else '乾乳中' if cow.get('rc') == RuleEngine.RC_DRY else '授精中' if cow.get('rc') == RuleEngine.RC_BRED else '空胎',
                pen_display
            )
            self.tree.insert("", tk.END, values=values)
    
    def _delete_selected(self):
        """選択された個体を削除"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("警告", "削除する個体を選択してください。")
            return
        
        for item in selection:
            index = self.tree.index(item)
            if 0 <= index < len(self.cow_list):
                del self.cow_list[index]
        
        self._update_tree()
    
    def _save_all(self):
        """全ての個体を登録"""
        if not self.cow_list:
            messagebox.showwarning("警告", "登録する個体がありません。")
            return
        
        introduction_date = self.date_entry.get().strip()
        if not introduction_date:
            messagebox.showwarning("警告", "導入日を入力してください。")
            return
        
        # 導入日を正規化
        introduction_date = self._normalize_date(introduction_date)
        if not introduction_date:
            messagebox.showerror("エラー", "導入日の形式が正しくありません（YYYY-MM-DD形式で入力してください）。")
            return
        
        saved_count = 0
        
        for cow_data in self.cow_list:
            try:
                # 既に同じJPN10の牛が存在するか確認
                existing_cow = self.db.get_cow_by_id(cow_data['cow_id'])
                if existing_cow and existing_cow.get('jpn10') == cow_data['jpn10']:
                    # 既存の牛が存在する場合、その牛に対して導入イベントを作成
                    cow_auto_id = existing_cow.get('auto_id')
                else:
                    # 新しい牛を登録
                    new_cow_data = {
                        'cow_id': cow_data['cow_id'],
                        'jpn10': cow_data['jpn10'],
                        'brd': cow_data.get('breed'),  # 品種
                        'bthd': cow_data.get('birth_date'),
                        'entr': introduction_date,  # 導入日
                        'lact': cow_data.get('lact', 0),
                        'clvd': cow_data.get('clvd'),
                        'rc': cow_data.get('rc', RuleEngine.RC_OPEN),
                        'pen': cow_data.get('pen'),
                        'frm': None
                    }
                    cow_auto_id = self.db.insert_cow(new_cow_data)
                
                # 導入イベントを作成
                intro_json = {
                    "birth_date": cow_data.get('birth_date'),
                    "lactation": cow_data.get('lact', 0),
                    "calving_date": cow_data.get('clvd'),
                    "last_ai_date": cow_data.get('last_ai_date'),
                    "reproduction_code": cow_data.get('rc'),
                    "pen": cow_data.get('pen'),
                    "source": "manual_external"  # 外部からの手動入力
                }
                
                intro_event = {
                    "cow_auto_id": cow_auto_id,
                    "event_number": RuleEngine.EVENT_IN,
                    "event_date": introduction_date,
                    "json_data": intro_json,
                    "note": "外部からの導入"
                }
                
                event_id = self.db.insert_event(intro_event)
                self.rule_engine.on_event_added(event_id)
                
                # 分娩月日が入力された場合、分娩イベントを作成（baseline_calving=True で産次加算しない／DIM計算用）
                clvd = cow_data.get('clvd')
                if clvd:
                    # 分娩日を正規化
                    clvd_normalized = self._normalize_date(clvd)
                    if clvd_normalized:
                        calv_event = {
                            "cow_auto_id": cow_auto_id,
                            "event_number": RuleEngine.EVENT_CALV,
                            "event_date": clvd_normalized,
                            "json_data": {"baseline_calving": True},
                            "note": "導入時の分娩（baseline）"
                        }
                        calv_event_id = self.db.insert_event(calv_event)
                        self.rule_engine.on_event_added(calv_event_id)
                
                # 受胎の有無に応じて追加イベントを作成
                pregnant_status = cow_data.get('pregnant_status', '')
                last_ai_date = cow_data.get('last_ai_date')
                last_ai_sire = cow_data.get('last_ai_sire')
                ai_count = cow_data.get('ai_count', 0)
                
                # 最終授精日が入力されている場合、AIイベントを作成
                # AIイベントは最終授精日の日付で作成する
                if last_ai_date:
                    # 最終授精日を正規化
                    last_ai_date_normalized = self._normalize_date(last_ai_date)
                    if last_ai_date_normalized:
                        # AIイベントを作成
                        ai_json = {
                            "ai_count": ai_count  # 授精回数をjson_dataに含める
                        }
                        # 最終授精SIREが入力されている場合は含める
                        if last_ai_sire:
                            ai_json["sire"] = last_ai_sire
                        
                        ai_event = {
                            "cow_auto_id": cow_auto_id,
                            "event_number": RuleEngine.EVENT_AI,
                            "event_date": last_ai_date_normalized,  # 最終授精日を使用
                            "json_data": ai_json,
                            "note": "導入時の最終授精"
                        }
                        ai_event_id = self.db.insert_event(ai_event)
                        self.rule_engine.on_event_added(ai_event_id)
                
                # 1) 妊娠鑑定待ちの場合：追加処理なし（AIイベントは上で作成済み）
                # 2) 受胎なしの場合：追加処理なし（AIイベントは上で作成済み）
                
                # 3) 妊娠中の場合：導入日で妊娠プラスイベントを作成
                if pregnant_status == "pregnant":
                    # 妊娠プラスイベントを作成（PDP2を使用）
                    preg_json = {
                        "twin": False  # 双子フラグ（デフォルトはFalse）
                    }
                    preg_event = {
                        "cow_auto_id": cow_auto_id,
                        "event_number": RuleEngine.EVENT_PDP2,  # 検診以外の妊娠プラス
                        "event_date": introduction_date,  # 導入日を使用
                        "json_data": preg_json,
                        "note": "導入時の妊娠確認"
                    }
                    preg_event_id = self.db.insert_event(preg_event)
                    self.rule_engine.on_event_added(preg_event_id)
                
                # 4) 乾乳中の場合：導入日で妊娠プラスイベントかつ乾乳イベントを作成
                elif pregnant_status == "dry":
                    # 妊娠プラスイベントを作成
                    preg_json = {
                        "twin": False
                    }
                    preg_event = {
                        "cow_auto_id": cow_auto_id,
                        "event_number": RuleEngine.EVENT_PDP2,
                        "event_date": introduction_date,  # 導入日を使用
                        "json_data": preg_json,
                        "note": "導入時の妊娠確認"
                    }
                    preg_event_id = self.db.insert_event(preg_event)
                    self.rule_engine.on_event_added(preg_event_id)
                    
                    # 乾乳イベントを作成
                    dry_event = {
                        "cow_auto_id": cow_auto_id,
                        "event_number": RuleEngine.EVENT_DRY,
                        "event_date": introduction_date,  # 導入日を使用
                        "json_data": {},
                        "note": "導入時の乾乳"
                    }
                    dry_event_id = self.db.insert_event(dry_event)
                    self.rule_engine.on_event_added(dry_event_id)
                
                # 授精回数が入力されている場合、BRED項目としてitem_valueテーブルに保存
                if ai_count > 0 and cow_auto_id is not None:
                    self.db.set_item_value(int(cow_auto_id), "BRED", ai_count)
                
                saved_count += 1
                
            except Exception as e:
                logging.error(f"個体登録エラー: {e}")
                messagebox.showerror("エラー", f"個体登録中にエラーが発生しました: {e}")
        
        if saved_count > 0:
            messagebox.showinfo("完了", f"{saved_count}件の個体を登録しました。")
            self.window.destroy()
    
    def show(self):
        """ウィンドウを表示"""
        self.window.deiconify()
        self.window.lift()
        self.window.focus_set()

