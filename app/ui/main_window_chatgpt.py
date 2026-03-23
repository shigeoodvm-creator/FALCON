"""
FALCON2 - ChatGPT 連携・分析実行 Mixin
MainWindow から分離した Mixin クラス。
MainWindow のみが継承することを前提とし、self.* 属性は MainWindow.__init__ で初期化される。
"""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import TYPE_CHECKING, Optional, Literal, List, Dict, Any, Tuple, Union
from pathlib import Path
from datetime import datetime, timedelta, date
import json
import logging
import re
import io
import threading
import webbrowser
import tempfile
import html

try:
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

try:
    import numpy as np
    NUMPY_AVAILABLE = True
except ImportError:
    NUMPY_AVAILABLE = False

if TYPE_CHECKING:
    from db.db_handler import DBHandler
    from modules.formula_engine import FormulaEngine
    from modules.rule_engine import RuleEngine

logger = logging.getLogger(__name__)

class ChatGPTMixin:
    """Mixin: FALCON2 - ChatGPT 連携・分析実行 Mixin"""
    def _ask_chatgpt(self, user_prompt: str):
        """
        ChatGPTに質問を送信（非推奨：ルールエンジンで処理できない場合のみ）
        
        注意：このメソッドは使用されません（ルールエンジンで未対応の場合は「未対応」として返す）
        
        Args:
            user_prompt: ユーザーの質問
        """
        # AIは使用しない（ルールエンジンで処理できない場合は未対応として返す）
        logging.warning(f"[ルールエンジン] _ask_chatgptが呼ばれましたが、AIは使用しません: \"{user_prompt}\"")
        self.add_message(role="system", text="未対応のクエリです。")
    
    def _run_analysis_in_thread(self, user_prompt: str, job_id: int):
        """
        バックグラウンドスレッドで分析処理を実行（非推奨：使用されていません）
        
        注意：このメソッドは使用されません（_run_db_aggregation_in_threadを使用）
        
        Args:
            user_prompt: ユーザーの質問
            job_id: ジョブID
        """
        # このメソッドは使用されていません
        logging.warning(f"[ルールエンジン] _run_analysis_in_threadが呼ばれましたが、使用されていません: job_id={job_id}")
        if job_id == self.current_job_id:
            self.analysis_running = False
    
    def _handle_analysis_result(self, user_prompt: str, calculated_data: Any, job_id: int):
        """
        分析結果を処理（メインスレッドで実行）
        
        Args:
            user_prompt: ユーザーの質問
            calculated_data: 計算済みデータ
            job_id: ジョブID
        """
        try:
            # ジョブIDが古い場合は無視（キャンセルされたジョブ）
            if job_id != self.current_job_id:
                logging.info(f"[分析] 古いジョブの結果を無視: job_id={job_id}, current_job_id={self.current_job_id}")
                # 古いジョブの場合はロック解除は不要（現在のジョブが処理する）
                return
            
            if calculated_data is None:
                # エラーメッセージは表示しない（次の処理に進む）
                matched_items = self._match_items_from_query(user_prompt)
                if not matched_items:
                    logging.info(f"[分析] 該当する分析処理が見つかりませんでした（次の処理に進む）: query=\"{user_prompt}\"")
                return
            
            # イベント抽出系クエリの結果（文字列）を表示
            if self._is_event_extraction_query(user_prompt) and isinstance(calculated_data, str):
                self.add_message(role="ai", text=calculated_data)
                logging.info(f"[分析] イベント抽出結果表示: job_id={job_id}")
                return
            
            # 数値系クエリの場合は数値のみを直接表示
            if self._is_numeric_query(user_prompt):
                if isinstance(calculated_data, str):
                    self.add_message(role="ai", text=calculated_data)
                return
            
            # 散布図クエリの場合はグラフを表示（ChatGPTを経由しない）
            if "散布図" in user_prompt and isinstance(calculated_data, dict):
                if 'type' not in calculated_data or calculated_data.get('type') != 'table_report':
                    try:
                        self._display_scatter_plot(calculated_data)
                    except Exception as e:
                        logging.error(f"[分析] 散布図表示エラー: {e}")
                        import traceback
                        traceback.print_exc()
                        self.add_message(role="system", text="散布図の表示に失敗しました。")
                    return
            
            # 表レポートの場合はReportTableWindowを開く（ChatGPTを経由しない）
            if isinstance(calculated_data, dict) and calculated_data.get('type') == 'table_report':
                try:
                    self._display_table_report(calculated_data)
                except Exception as e:
                    logging.error(f"[分析] 表レポート表示エラー: {e}")
                    import traceback
                    traceback.print_exc()
                    self.add_message(role="system", text="表レポートの表示に失敗しました。")
                return
            
            # その他の表・CSV形式の場合はそのまま表示
            if isinstance(calculated_data, str) and ("表" in user_prompt or "csv" in user_prompt.lower()):
                self.add_message(role="ai", text=calculated_data)
                return
            
            # その他の分析クエリは未対応として返す（AIは使用しない）
            logging.info(f"[分析] 未対応の分析クエリ: job_id={job_id}, query=\"{user_prompt}\"")
            self.add_message(role="system", text="未対応のクエリです。")
        
        finally:
            # 分析ロックを必ず解除（成功・失敗・例外のいずれの場合でも）
            if job_id == self.current_job_id:
                self.analysis_running = False
                logging.info(f"[分析] 分析ロック解除: job_id={job_id}")
    
    def _handle_analysis_error(self, job_id: int, error_message: str):
        """
        分析エラーを処理（メインスレッドで実行）
        
        Args:
            job_id: ジョブID
            error_message: エラーメッセージ
        """
        try:
            # ジョブIDが古い場合は無視
            if job_id != self.current_job_id:
                logging.info(f"[分析] 古いジョブのエラーを無視: job_id={job_id}, current_job_id={self.current_job_id}")
                return
            
            logging.error(f"[分析] 分析処理例外発生: job_id={job_id}, error={error_message}")
            # エラーメッセージは表示しない（内部エラー）
            self.add_message(role="system", text="処理できませんでした。")
        
        finally:
            # 分析ロックを必ず解除（成功・失敗・例外のいずれの場合でも）
            if job_id == self.current_job_id:
                self.analysis_running = False
                logging.info(f"[分析] 分析ロック解除（エラー時）: job_id={job_id}")
    
    def _find_event_by_name(self, search_str: str) -> Optional[int]:
        """
        イベント名（alias または name_jp）でイベント番号を検索
        
        Args:
            search_str: 検索文字列（大文字小文字を区別しない）
        
        Returns:
            見つかったイベント番号、見つからない場合はNone
        """
        if not self.event_dict_path or not self.event_dict_path.exists():
            return None
        
        try:
            # イベント辞書を読み込む
            with open(self.event_dict_path, 'r', encoding='utf-8') as f:
                event_dictionary = json.load(f)
            
            search_lower = search_str.lower().strip()
            
            # イベント辞書を検索（alias または name_jp で前方一致）
            for event_num_str, event_data in event_dictionary.items():
                # deprecated のイベントは除外
                if event_data.get('deprecated', False):
                    continue
                
                # alias で検索
                alias = event_data.get('alias', '')
                if alias and alias.lower().startswith(search_lower):
                    return int(event_num_str)
                
                # name_jp で検索
                name_jp = event_data.get('name_jp', '')
                if name_jp and name_jp.lower().startswith(search_lower):
                    return int(event_num_str)
            
            return None
        except Exception as e:
            logging.error(f"イベント検索エラー: {e}")
            return None
    
    def _open_event_input_for_event(self, event_number: int):
        """
        指定されたイベントに固定されたイベント入力ウィンドウを開く
        
        Args:
            event_number: イベント番号
        """
        if self.event_dict_path is None:
            messagebox.showerror(
                "エラー",
                "event_dictionary.json が見つかりません"
            )
            return
        
        # EventInputWindow を生成または再利用して表示（該当イベントに固定）
        event_input_window = EventInputWindow.open_or_focus(
            parent=self.root,
            db_handler=self.db,
            rule_engine=self.rule_engine,
            event_dictionary_path=self.event_dict_path,
            cow_auto_id=None,  # 汎用モード
            on_saved=self._on_event_saved,
            farm_path=self.farm_path,
            allowed_event_numbers=[event_number],  # 該当イベントのみ
            default_event_number=event_number  # デフォルトで選択
        )
        event_input_window.show()
    
    def _jump_to_cow_card(self, cow_id: str):
        """
        個体カードウィンドウを開く
        
        Args:
            cow_id: 4桁の牛ID（例: "0980"）
        """
        from ui.cow_card_window import CowCardWindow
        
        # 4桁IDで検索（複数件取得可能）
        cows = self.db.get_cows_by_id(cow_id)
        
        if not cows:
            # 見つからない場合は、正規化されたIDで検索を試みる
            normalized_id = cow_id.lstrip('0')  # 左ゼロを除去
            cow = self.db.get_cow_by_normalized_id(normalized_id)
            if cow:
                cows = [cow]
        
        if not cows:
            # 見つからない場合のUX
            messagebox.showinfo(
                "検索結果",
                f"ID {cow_id} の個体は見つかりませんでした"
            )
            self._add_chat_message("システム", f"ID {cow_id} の個体は見つかりませんでした")
            return
        
        # 1件のみの場合はそのまま表示
        if len(cows) == 1:
            cow = cows[0]
            cow_auto_id = cow.get('auto_id')
            if cow_auto_id:
                # 個体カードウィンドウを開く
                cow_card_window = CowCardWindow(
                    parent=self.root,
                    db_handler=self.db,
                    formula_engine=self.formula_engine,
                    rule_engine=self.rule_engine,
                    event_dictionary_path=self.event_dict_path,
                    item_dictionary_path=self.item_dict_path,
                    cow_auto_id=cow_auto_id
                )
                cow_card_window.show()
                # 個体カードウィンドウを最前面に出す（ブラウザの後ろに隠れないようにする）
                try:
                    win = cow_card_window.window
                    win.deiconify()
                    win.lift()
                    win.focus_force()
                    win.attributes("-topmost", True)
                    win.after(200, lambda w=win: w.attributes("-topmost", False))
                except Exception:
                    pass
                self._add_chat_message("システム", f"個体カードを表示しました: {cow.get('cow_id', cow_id)}")
                # チャット履歴に個体情報とイベント履歴を表示（色付き）
                self._display_cow_info_in_chat(cow_auto_id)
            return
        
        # 複数件見つかった場合は選択ダイアログを表示
        selected_cow = self._show_cow_selection_dialog(cows, cow_id)
        if selected_cow:
            cow_auto_id = selected_cow.get('auto_id')
            if cow_auto_id:
                # 個体カードウィンドウを開く
                cow_card_window = CowCardWindow(
                    parent=self.root,
                    db_handler=self.db,
                    formula_engine=self.formula_engine,
                    rule_engine=self.rule_engine,
                    event_dictionary_path=self.event_dict_path,
                    item_dictionary_path=self.item_dict_path,
                    cow_auto_id=cow_auto_id
                )
                cow_card_window.show()
                # 個体カードウィンドウを最前面に出す
                try:
                    win = cow_card_window.window
                    win.deiconify()
                    win.lift()
                    win.focus_force()
                    win.attributes("-topmost", True)
                    win.after(200, lambda w=win: w.attributes("-topmost", False))
                except Exception:
                    pass
                self._add_chat_message("システム", f"個体カードを表示しました: {selected_cow.get('cow_id', cow_id)} (個体識別番号: {selected_cow.get('jpn10', '')})")
                # チャット履歴に個体情報とイベント履歴を表示（色付き）
                self._display_cow_info_in_chat(cow_auto_id)
    
    def _show_cow_selection_dialog(self, cows: List[Dict[str, Any]], cow_id: str) -> Optional[Dict[str, Any]]:
        """
        4桁IDが重複している場合の個体選択ダイアログ
        
        Args:
            cows: 検索結果の牛リスト
            cow_id: 検索した4桁ID
        
        Returns:
            選択された牛の情報、キャンセルの場合はNone
        """
        # 選択ダイアログウィンドウを作成
        dialog = tk.Toplevel(self.root)
        dialog.title("個体選択")
        dialog.geometry("500x400")
        
        # メッセージ
        message = f"拡大4桁ID {cow_id} に該当する個体が複数見つかりました。\n個体識別番号から選択してください。"
        ttk.Label(dialog, text=message, wraplength=450).pack(pady=10)
        
        # リストボックス
        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # スクロールバー付きリストボックス
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=("Courier", 10))
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        # 選択結果を保持
        selected_cow = None
        
        # リストボックスに個体情報を追加
        for cow in cows:
            jpn10 = cow.get('jpn10', '')
            brd = cow.get('brd', '')
            bthd = cow.get('bthd', '')
            pen = cow.get('pen', '')
            
            # 表示形式: 個体識別番号 | 品種 | 生年月日 | 群
            display_text = f"{jpn10:12s} | {brd:10s} | {bthd:10s} | {pen:10s}"
            listbox.insert(tk.END, display_text)
        
        # リストボックスの選択イベント
        def on_select(event):
            nonlocal selected_cow
            selection = listbox.curselection()
            if selection:
                idx = selection[0]
                selected_cow = cows[idx]
        
        listbox.bind('<<ListboxSelect>>', on_select)
        listbox.bind('<Double-Button-1>', lambda e: dialog.destroy())
        
        # ヘッダー行を追加（読みやすさのため）
        header_frame = ttk.Frame(dialog)
        header_frame.pack(fill=tk.X, padx=10)
        header_text = f"{'個体識別番号':12s} | {'品種':10s} | {'生年月日':10s} | {'群':10s}"
        ttk.Label(header_frame, text=header_text, font=("Courier", 10, "bold")).pack()
        
        # ボタンフレーム
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        def on_ok():
            nonlocal selected_cow
            selection = listbox.curselection()
            if selection:
                idx = selection[0]
                selected_cow = cows[idx]
                dialog.destroy()
            else:
                messagebox.showwarning("警告", "個体を選択してください。")
        
        def on_cancel():
            nonlocal selected_cow
            selected_cow = None
            dialog.destroy()
        
        ttk.Button(button_frame, text="選択", command=on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="キャンセル", command=on_cancel).pack(side=tk.LEFT, padx=5)
        
        # ダイアログを中央に配置
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # モーダルダイアログとして表示
        dialog.wait_window()
        
        return selected_cow
    
