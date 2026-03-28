"""
FALCON2 - 導入ウィンドウ
外部からの導入イベント入力用（1頭ずつ即時登録）
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
    """導入ウィンドウ（1頭ずつ即時登録）"""

    def __init__(self, parent: tk.Widget, db_handler: DBHandler,
                 rule_engine: RuleEngine,
                 farm_path: Path,
                 event_dictionary_path: Optional[Path] = None,
                 introduction_date: Optional[str] = None,
                 initial_cows: Optional[List[Dict[str, Any]]] = None,
                 close_after_save: bool = False):
        self.parent = parent
        self.db = db_handler
        self.rule_engine = rule_engine
        self.farm_path = Path(farm_path)
        self.event_dict_path = event_dictionary_path
        # 保存後にウィンドウを閉じるかどうか（乳検吸い込みからの確認モードで使用）
        self.close_after_save = close_after_save

        # 導入日（共通）
        self.introduction_date = introduction_date or datetime.now().strftime('%Y-%m-%d')

        # SettingsManagerを初期化
        self.settings_manager = SettingsManager(self.farm_path)

        # ウィンドウ作成
        self.window = tk.Toplevel(parent)
        self.window.title("導入")
        self.window.geometry("800x720")
        self.window.minsize(700, 600)
        self.window.configure(bg="#f5f5f5")

        # UI作成
        self._create_widgets()

        # 初期入力がある場合はフォームに流し込む（自動保存しない）
        if initial_cows:
            self._prefill_form(initial_cows[0])

        # ウィンドウを中央に配置
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")

    def _prefill_form(self, cow_data: Dict[str, Any]) -> None:
        """個体情報をフォームに流し込む（ユーザー確認前の表示用）"""
        jpn10 = str(cow_data.get("jpn10", "") or "").strip()
        cow_id = str(cow_data.get("cow_id", "") or "").strip()
        if not cow_id and jpn10.isdigit() and len(jpn10) == 10:
            cow_id = jpn10[5:9].zfill(4)

        # JPN10
        self.jpn10_entry.delete(0, tk.END)
        self.jpn10_entry.insert(0, jpn10)

        # ID（自動設定）
        self.cow_id_entry.config(state="normal")
        self.cow_id_entry.delete(0, tk.END)
        self.cow_id_entry.insert(0, cow_id)
        self.cow_id_entry.config(state="disabled")

        # 品種
        breed = str(cow_data.get("breed", "") or "").strip()
        self.breed_entry.set(breed)

        # 生年月日
        birth_date = self._normalize_date(str(cow_data.get("birth_date", "") or "")) or ""
        self.birth_date_entry.delete(0, tk.END)
        self.birth_date_entry.insert(0, birth_date)

        # 産次
        lact = cow_data.get("lact")
        self.lact_entry.delete(0, tk.END)
        if lact is not None and str(lact).strip() != "":
            self.lact_entry.insert(0, str(lact))

        # 分娩月日
        clvd = self._normalize_date(str(cow_data.get("clvd", "") or "")) or ""
        self.clvd_entry.delete(0, tk.END)
        self.clvd_entry.insert(0, clvd)

        # 最終授精日
        last_ai_date = self._normalize_date(str(cow_data.get("last_ai_date", "") or "")) or ""
        self.last_ai_date_entry.delete(0, tk.END)
        self.last_ai_date_entry.insert(0, last_ai_date)

        # 最終授精SIRE
        self.last_ai_sire_entry.delete(0, tk.END)
        self.last_ai_sire_entry.insert(0, str(cow_data.get("last_ai_sire", "") or ""))

        # 授精回数
        ai_count = cow_data.get("ai_count")
        self.ai_count_entry.delete(0, tk.END)
        if ai_count:
            self.ai_count_entry.insert(0, str(ai_count))

        # 受胎の有無
        self.pregnant_var.set(str(cow_data.get("pregnant_status", "") or ""))

        # PEN
        pen = str(cow_data.get("pen", "") or "")
        self.pen_entry.delete(0, tk.END)
        self.pen_entry.insert(0, pen)

    def _create_widgets(self):
        """ウィジェットを作成"""
        _df = "Meiryo UI"
        bg = "#f5f5f5"
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
        _label_col_minsize = 110

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
        ttk.Button(header_right, text="登録", command=self._save_cow, width=8).pack(side=tk.LEFT, padx=(0, 5))
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

        # 個体入力
        input_frame = ttk.LabelFrame(
            content_area, text="個体入力", padding=(12, 10), style="EventInput.TLabelframe"
        )
        input_frame.pack(fill=tk.X, pady=(0, 10))
        fields_frame = ttk.Frame(input_frame)
        fields_frame.pack(fill=tk.X, pady=5)
        fields_frame.columnconfigure(0, minsize=_label_col_minsize)
        fields_frame.columnconfigure(1, weight=0, minsize=320)

        _row = 0
        ttk.Label(fields_frame, text="JPN10*:").grid(row=_row, column=0, sticky=tk.W, padx=(5, 8), pady=(4, 2))
        self.jpn10_entry = ttk.Entry(fields_frame, width=18, style="EventInput.TEntry")
        self.jpn10_entry.grid(row=_row, column=1, sticky=tk.W, padx=(0, 5), pady=(4, 2))
        self.jpn10_entry.bind("<KeyRelease>", self._on_jpn10_changed)
        _row += 1
        ttk.Label(fields_frame, text="ID（自動）:").grid(row=_row, column=0, sticky=tk.W, padx=(5, 8), pady=2)
        self.cow_id_entry = ttk.Entry(fields_frame, width=10, style="EventInput.TEntry")
        self.cow_id_entry.grid(row=_row, column=1, sticky=tk.W, padx=(0, 5), pady=2)
        self.cow_id_entry.config(state="disabled")
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

        # ステータスラベル（登録後の確認メッセージ）
        self.status_label = tk.Label(
            content_area, text="", font=(_df, 10), bg=bg, fg="#2e7d32"
        )
        self.status_label.pack(pady=(8, 0))

    def _open_date_picker(self, entry: ttk.Entry) -> None:
        """日付選択カレンダーを開き、選択した日付をEntryにセットする"""
        current = entry.get().strip()
        initial = current if current else None
        try:
            picker = DatePickerWindow(
                parent=self.window,
                initial_date=initial,
                on_date_selected=lambda date_str: entry.delete(0, tk.END) or entry.insert(0, date_str),
            )
            picker.show()
        except Exception as e:
            logging.warning(f"DatePickerWindow open failed: {e}")
            messagebox.showwarning("日付選択", "カレンダーを開けませんでした。", parent=self.window)

    def _safe_focus(self, widget: tk.Widget) -> None:
        """破棄済みウィジェットへ focus_set しない（messagebox 直後の TclError 防止）"""
        try:
            if self.window.winfo_exists() and widget.winfo_exists():
                widget.focus_set()
        except tk.TclError:
            pass

    def _load_pen_options(self):
        """PEN設定をロードしてプルダウンに設定"""
        try:
            pen_settings = self.settings_manager.load_pen_settings()
            pen_options = []
            for code, name in sorted(pen_settings.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 999):
                pen_options.append(f"{code}: {name}")
            self.pen_entry['values'] = pen_options
        except Exception as e:
            logging.error(f"PEN設定の読み込みに失敗しました: {e}")
            self.pen_entry['values'] = []

    def _normalize_date(self, date_str: str) -> Optional[str]:
        """日付文字列を YYYY-MM-DD に正規化。無効な場合は None。"""
        if not date_str or not date_str.strip():
            return None
        date_str = date_str.strip()
        parts = date_str.split('-')
        if len(parts) == 3:
            try:
                dt = datetime(int(parts[0]), int(parts[1]), int(parts[2]))
                return dt.strftime('%Y-%m-%d')
            except (ValueError, TypeError):
                pass
        return None

    def _on_jpn10_changed(self, event):
        """JPN10入力変更時にIDを自動設定"""
        jpn10_value = self.jpn10_entry.get().strip()
        if jpn10_value.isdigit() and len(jpn10_value) == 10:
            cow_id = jpn10_value[5:9].zfill(4)
            if not self.cow_id_entry.get().strip():
                self.cow_id_entry.config(state="normal")
                self.cow_id_entry.delete(0, tk.END)
                self.cow_id_entry.insert(0, cow_id)
                self.cow_id_entry.config(state="disabled")
        else:
            if self.cow_id_entry.get().strip():
                self.cow_id_entry.config(state="normal")
                self.cow_id_entry.delete(0, tk.END)
                self.cow_id_entry.config(state="disabled")

    def _collect_form(self) -> Optional[Dict[str, Any]]:
        """フォームから個体データを収集・バリデーション。エラー時は None を返す。"""
        jpn10 = self.jpn10_entry.get().strip()
        cow_id = self.cow_id_entry.get().strip()
        breed = (self.breed_entry.get() or "").strip()
        lact = self.lact_entry.get().strip()
        ai_count = self.ai_count_entry.get().strip()
        pregnant = self.pregnant_var.get()

        # JPN10 必須チェック
        if not jpn10:
            messagebox.showwarning("警告", "JPN10を入力してください。", parent=self.window)
            self._safe_focus(self.jpn10_entry)
            return None
        if not jpn10.isdigit() or len(jpn10) != 10:
            messagebox.showwarning("警告", "JPN10は10桁の数字で入力してください。", parent=self.window)
            self._safe_focus(self.jpn10_entry)
            return None

        # 日付バリデーション
        birth_date_raw = self.birth_date_entry.get().strip()
        birth_date = None
        if birth_date_raw:
            birth_date = self._normalize_date(birth_date_raw)
            if not birth_date:
                messagebox.showwarning("警告", "生年月日の形式が正しくありません（YYYY-MM-DD形式）。", parent=self.window)
                self._safe_focus(self.birth_date_entry)
                return None

        clvd_raw = self.clvd_entry.get().strip()
        clvd = None
        if clvd_raw:
            clvd = self._normalize_date(clvd_raw)
            if not clvd:
                messagebox.showwarning("警告", "分娩月日の形式が正しくありません（YYYY-MM-DD形式）。", parent=self.window)
                self._safe_focus(self.clvd_entry)
                return None

        last_ai_date_raw = self.last_ai_date_entry.get().strip()
        last_ai_date = None
        if last_ai_date_raw:
            last_ai_date = self._normalize_date(last_ai_date_raw)
            if not last_ai_date:
                messagebox.showwarning("警告", "最終授精日の形式が正しくありません（YYYY-MM-DD形式）。", parent=self.window)
                self._safe_focus(self.last_ai_date_entry)
                return None

        # IDが未入力の場合はJPN10から自動生成
        if not cow_id:
            cow_id = jpn10[5:9].zfill(4)

        # PEN
        pen_input = self.pen_entry.get().strip()
        pen = None
        if pen_input:
            pen = pen_input.split(':')[0].strip() if ':' in pen_input else pen_input

        # RC決定
        rc = RuleEngine.RC_OPEN
        if pregnant == "pregnant":
            rc = RuleEngine.RC_PREGNANT
        elif pregnant == "dry":
            rc = RuleEngine.RC_DRY
        elif pregnant == "waiting" or (pregnant == "" and last_ai_date):
            rc = RuleEngine.RC_BRED

        return {
            'cow_id': cow_id,
            'jpn10': jpn10,
            'breed': breed or None,
            'birth_date': birth_date,
            'lact': int(lact) if lact.isdigit() else 0,
            'clvd': clvd,
            'last_ai_date': last_ai_date,
            'last_ai_sire': self.last_ai_sire_entry.get().strip() or None,
            'ai_count': int(ai_count) if ai_count.isdigit() else 0,
            'pregnant_status': pregnant,
            'rc': rc,
            'pen': pen,
        }

    def _save_cow(self):
        """フォームの内容を即時DB登録し、フォームをクリアして次の入力へ"""
        # 導入日バリデーション
        introduction_date_raw = self.date_entry.get().strip()
        if not introduction_date_raw:
            messagebox.showwarning("警告", "導入日を入力してください。", parent=self.window)
            return
        introduction_date = self._normalize_date(introduction_date_raw)
        if not introduction_date:
            messagebox.showerror("エラー", "導入日の形式が正しくありません（YYYY-MM-DD形式）。", parent=self.window)
            return

        # 未来日付チェック
        try:
            from datetime import date as _date
            intro_dt = _date.fromisoformat(introduction_date)
            if intro_dt > _date.today():
                proceed = messagebox.askyesno(
                    "確認",
                    f"導入日（{introduction_date}）が本日より先の日付です。\nこのまま登録しますか？",
                    parent=self.window,
                )
                if not proceed:
                    return
        except (ValueError, TypeError):
            pass

        cow_data = self._collect_form()
        if cow_data is None:
            return

        try:
            self._do_save(cow_data, introduction_date)
        except Exception as e:
            logging.error(f"個体登録エラー: {e}")
            messagebox.showerror("エラー", f"登録中にエラーが発生しました: {e}", parent=self.window)
            return

        registered_id = cow_data['cow_id']
        if self.close_after_save:
            # 乳検吸い込みなど確認モード：登録後にウィンドウを閉じる
            self.window.destroy()
        else:
            # 通常モード：フォームをクリアして次の入力へ
            self._clear_form()
            self.status_label.config(text=f"✓ ID {registered_id} を登録しました")
            self._safe_focus(self.jpn10_entry)

    def _do_save(self, cow_data: Dict[str, Any], introduction_date: str) -> None:
        """1頭分のデータをDBに保存する（内部共通処理）"""
        # 既存個体確認
        existing_cow = self.db.get_cow_by_id(cow_data['cow_id'])
        if existing_cow and existing_cow.get('jpn10') == cow_data['jpn10']:
            cow_auto_id = existing_cow.get('auto_id')
        else:
            new_cow_data = {
                'cow_id': cow_data['cow_id'],
                'jpn10': cow_data['jpn10'],
                'brd': cow_data.get('breed'),
                'bthd': cow_data.get('birth_date'),
                'entr': introduction_date,
                'lact': cow_data.get('lact', 0),
                'clvd': cow_data.get('clvd'),
                'rc': cow_data.get('rc', RuleEngine.RC_OPEN),
                'pen': cow_data.get('pen'),
                'frm': None
            }
            cow_auto_id = self.db.insert_cow(new_cow_data)

        # 導入イベント
        intro_event = {
            "cow_auto_id": cow_auto_id,
            "event_number": RuleEngine.EVENT_IN,
            "event_date": introduction_date,
            "json_data": {
                "birth_date": cow_data.get('birth_date'),
                "lactation": cow_data.get('lact', 0),
                "calving_date": cow_data.get('clvd'),
                "last_ai_date": cow_data.get('last_ai_date'),
                "reproduction_code": cow_data.get('rc'),
                "pen": cow_data.get('pen'),
                "source": "manual_external"
            },
            "note": "外部からの導入"
        }
        event_id = self.db.insert_event(intro_event)
        self.rule_engine.on_event_added(event_id)

        # 分娩イベント（baseline）
        clvd = cow_data.get('clvd')
        if clvd:
            calv_event_id = self.db.insert_event({
                "cow_auto_id": cow_auto_id,
                "event_number": RuleEngine.EVENT_CALV,
                "event_date": clvd,
                "json_data": {"baseline_calving": True},
                "note": "導入時の分娩（baseline）"
            })
            self.rule_engine.on_event_added(calv_event_id)

        # 最終授精イベント
        last_ai_date = cow_data.get('last_ai_date')
        if last_ai_date:
            ai_json: Dict[str, Any] = {"ai_count": cow_data.get('ai_count', 0)}
            if cow_data.get('last_ai_sire'):
                ai_json["sire"] = cow_data['last_ai_sire']
            ai_event_id = self.db.insert_event({
                "cow_auto_id": cow_auto_id,
                "event_number": RuleEngine.EVENT_AI,
                "event_date": last_ai_date,
                "json_data": ai_json,
                "note": "導入時の最終授精"
            })
            self.rule_engine.on_event_added(ai_event_id)

        pregnant_status = cow_data.get('pregnant_status', '')

        # 妊娠中 or 乾乳中：妊娠プラスイベント
        if pregnant_status in ("pregnant", "dry"):
            preg_event_id = self.db.insert_event({
                "cow_auto_id": cow_auto_id,
                "event_number": RuleEngine.EVENT_PDP2,
                "event_date": introduction_date,
                "json_data": {"twin": False},
                "note": "導入時の妊娠確認"
            })
            self.rule_engine.on_event_added(preg_event_id)

        # 乾乳中：乾乳イベント
        if pregnant_status == "dry":
            dry_event_id = self.db.insert_event({
                "cow_auto_id": cow_auto_id,
                "event_number": RuleEngine.EVENT_DRY,
                "event_date": introduction_date,
                "json_data": {},
                "note": "導入時の乾乳"
            })
            self.rule_engine.on_event_added(dry_event_id)

        # 授精回数をitem_valueに保存
        ai_count = cow_data.get('ai_count', 0)
        if ai_count > 0 and cow_auto_id is not None:
            self.db.set_item_value(int(cow_auto_id), "BRED", ai_count)

    def _clear_form(self):
        """個体入力フォームを全クリア"""
        self.jpn10_entry.delete(0, tk.END)
        self.cow_id_entry.config(state="normal")
        self.cow_id_entry.delete(0, tk.END)
        self.cow_id_entry.config(state="disabled")
        self.breed_entry.set("")
        self.birth_date_entry.delete(0, tk.END)
        self.lact_entry.delete(0, tk.END)
        self.clvd_entry.delete(0, tk.END)
        self.last_ai_date_entry.delete(0, tk.END)
        self.last_ai_sire_entry.delete(0, tk.END)
        self.ai_count_entry.delete(0, tk.END)
        self.pregnant_var.set("")
        self.pen_entry.delete(0, tk.END)

    def show(self):
        """ウィンドウを表示"""
        self.window.deiconify()
        self.window.lift()
        try:
            if self.window.winfo_exists():
                self.window.focus_set()
        except tk.TclError:
            pass
