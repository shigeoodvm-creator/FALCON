"""
FALCON2 - 個体カードウィンドウ
個体カードを別ウィンドウとして表示
"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Callable
from pathlib import Path

from db.db_handler import DBHandler
from modules.formula_engine import FormulaEngine
from modules.rule_engine import RuleEngine
from ui.cow_card import CowCard
from ui.cow_history_window import CowHistoryWindow


class CowCardWindow:
    """個体カードウィンドウ（Toplevel）"""
    
    def __init__(self, parent: tk.Tk, db_handler: DBHandler,
                 formula_engine: FormulaEngine,
                 rule_engine: RuleEngine,
                 event_dictionary_path: Optional[Path] = None,
                 item_dictionary_path: Optional[Path] = None,
                 cow_auto_id: Optional[int] = None):
        """
        初期化
        
        Args:
            parent: 親ウィンドウ
            db_handler: DBHandler インスタンス
            formula_engine: FormulaEngine インスタンス
            rule_engine: RuleEngine インスタンス
            event_dictionary_path: event_dictionary.json のパス
            item_dictionary_path: item_dictionary.json のパス（オプション）
            cow_auto_id: 表示する牛の auto_id（オプション、後でload_cowで設定可能）
        """
        self.db = db_handler
        self.formula_engine = formula_engine
        self.rule_engine = rule_engine
        self.event_dict_path = event_dictionary_path
        self.item_dict_path = item_dictionary_path
        self.history_window: Optional[CowHistoryWindow] = None
        
        # ウィンドウ作成（繁殖検診フローなどと同一デザイン）
        _df = "Meiryo UI"
        bg = "#f5f5f5"
        # 除籍（売却・死亡）表示用の色
        self._bg_normal = bg
        self._bg_disposed = "#ffebee"   # 薄い赤背景（視認性・スマートさのバランス）
        self._fg_disposed = "#b71c1c"   # 濃い赤文字
        btn_secondary_bd = "#b0bec5"

        self.window = tk.Toplevel(parent)
        self.window.title("個体カード")
        self.window.configure(bg=bg)
        screen_width = parent.winfo_screenwidth()
        screen_height = parent.winfo_screenheight()
        # 初期サイズは画面より少し小さくして、最大化ボタンを有効な状態にする
        init_w = min(1520, max(1100, screen_width - 120))
        init_h = min(900, max(700, screen_height - 120))
        self.window.geometry(f"{init_w}x{init_h}")
        self.window.minsize(1100, 700)
        self.window.resizable(True, True)
        self.window.protocol("WM_DELETE_WINDOW", self._on_close)
        
        # ========== 左右分割（PanedWindow）==========
        paned = ttk.PanedWindow(self.window, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)
        
        # 左ペイン：個体カード用
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=1)
        
        # ========== ヘッダー（1行：タイトル + 表示中ID強調 + 個体IDナビ + 削除） ==========
        header = tk.Frame(left_frame, bg=bg, pady=12, padx=24)
        header.pack(fill=tk.X)
        header.columnconfigure(1, weight=1)  # 中央の余白でタイトルとナビを両端寄せ
        
        tk.Label(header, text="\U0001f4cb", font=(_df, 22), bg=bg, fg="#3949ab").grid(row=0, column=0, padx=(0, 10), sticky="w")
        self.title_block = tk.Frame(header, bg=bg)
        self.title_block.grid(row=0, column=1, sticky="w", padx=(0, 16))
        self.card_title_label = tk.Label(self.title_block, text="個体カード", font=(_df, 16, "bold"), bg=bg, fg="#263238")
        self.card_title_label.pack(side=tk.LEFT, padx=(0, 8))
        self.title_cow_id_label = tk.Label(self.title_block, text="", font=(_df, 20, "bold"), bg=bg, fg="#1565c0")
        self.title_cow_id_label.pack(side=tk.LEFT, padx=(0, 4))
        self.title_jpn10_label = tk.Label(self.title_block, text="", font=(_df, 10), bg=bg, fg="#607d8b")
        self.title_jpn10_label.pack(side=tk.LEFT)
        # 除籍時のみ表示するバッジ（売却・死亡など牛群にいない個体用）
        self.disposed_badge = tk.Label(
            self.title_block, text="除籍", font=(_df, 9, "bold"),
            bg=bg, fg="#c62828",
            padx=6, pady=2
        )
        # 初期は非表示（load_cowで除籍時のみpack）
        
        nav_inner = tk.Frame(header, bg=bg)
        nav_inner.grid(row=0, column=2, sticky="e")
        tk.Label(nav_inner, text="個体ID:", font=(_df, 10), bg=bg, fg="#263238").pack(side=tk.LEFT, padx=(0, 4))
        prev_btn = tk.Button(nav_inner, text="◀", font=(_df, 10), width=2,
                             bg="#fafafa", fg="#546e7a", relief=tk.FLAT,
                             highlightbackground=btn_secondary_bd, highlightthickness=1,
                             cursor="hand2", command=self._on_prev_cow)
        prev_btn.pack(side=tk.LEFT, padx=(0, 2))
        self.cow_id_entry = tk.Entry(nav_inner, width=8, font=(_df, 11), relief=tk.SOLID, borderwidth=1)
        self.cow_id_entry.pack(side=tk.LEFT, padx=2)
        self.cow_id_entry.bind('<Return>', lambda e: self._on_switch_cow())
        next_btn = tk.Button(nav_inner, text="▶", font=(_df, 10), width=2,
                             bg="#fafafa", fg="#546e7a", relief=tk.FLAT,
                             highlightbackground=btn_secondary_bd, highlightthickness=1,
                             cursor="hand2", command=self._on_next_cow)
        next_btn.pack(side=tk.LEFT, padx=(2, 12))
        
        delete_btn = tk.Button(header, text="削除", font=(_df, 10), width=8,
                               bg="#fafafa", fg="#546e7a", relief=tk.FLAT,
                               highlightbackground=btn_secondary_bd, highlightthickness=1,
                               cursor="hand2", command=self._on_delete_cow)
        delete_btn.grid(row=0, column=3, sticky="e")
        
        self.current_cow_auto_id = None
        
        # CowCardを作成（左ペイン用）
        # イベント履歴は右ペインの CowHistoryWindow のみ（CowCard 内に二重表示しない）
        self.cow_card = CowCard(
            parent=left_frame,
            db_handler=db_handler,
            formula_engine=formula_engine,
            rule_engine=rule_engine,
            event_dictionary_path=event_dictionary_path,
            item_dictionary_path=item_dictionary_path,
            show_event_history=False,
            on_remote_event_history_refresh=self._notify_embedded_history_refresh,
        )
        cow_card_widget = self.cow_card.get_widget()
        cow_card_widget.pack(fill=tk.BOTH, expand=True, padx=24, pady=(0, 10))
        
        # 右ペイン：イベント履歴（同じウィンドウ内に埋め込み、ドラッグで一緒に移動）
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=1)
        self.history_window = CowHistoryWindow(
            parent=right_frame,
            cow_card=self.cow_card,
            db_handler=self.db,
        )
        
        # 牛の情報を読み込んで表示（cow_auto_idが指定されている場合）
        if cow_auto_id:
            self.load_cow(cow_auto_id)

    def _notify_embedded_history_refresh(self, cow_auto_id: int) -> None:
        """CowCard 側の履歴表示が無いとき、右ペインの CowHistoryWindow を同期する"""
        hw = getattr(self, "history_window", None)
        if hw and hw.window.winfo_exists():
            hw.load_cow(cow_auto_id)
    
    def load_cow(self, cow_auto_id: int):
        """
        牛の情報を読み込んで表示
        
        Args:
            cow_auto_id: 牛の auto_id
        """
        try:
            cow = self.db.get_cow_by_auto_id(cow_auto_id)
            if not cow:
                messagebox.showerror("エラー", f"個体ID {cow_auto_id} の個体が見つかりませんでした")
                return
            
            # 現在の個体IDを保持
            self.current_cow_auto_id = cow_auto_id
            
            # ウィンドウタイトルを更新
            cow_id = cow.get('cow_id', '')
            jpn10 = cow.get('jpn10', '')
            if jpn10:
                self.window.title(f"個体カード - {cow_id} ({jpn10})")
            else:
                self.window.title(f"個体カード - {cow_id}")
            
            # ヘッダーに表示中の個体IDを強調表示（4桁は大きく、10桁は【】で通常サイズ）
            self.title_cow_id_label.config(text=cow_id)
            self.title_jpn10_label.config(text=f"【{jpn10}】" if jpn10 else "")
            
            # 個体ID入力欄に現在のIDを表示（4桁形式）
            self.cow_id_entry.delete(0, tk.END)
            self.cow_id_entry.insert(0, cow_id)
            
            # 売却・死亡など牛群に存在しない場合はヘッダーを視認しやすく強調（薄赤背景＋除籍バッジ）
            is_disposed = self._is_cow_disposed(cow_auto_id)
            if is_disposed:
                self.title_block.config(bg=self._bg_disposed)
                self.card_title_label.config(bg=self._bg_disposed)
                self.title_cow_id_label.config(bg=self._bg_disposed, fg=self._fg_disposed)
                self.title_jpn10_label.config(bg=self._bg_disposed, fg="#5d4037")
                self.cow_id_entry.config(bg=self._bg_disposed, fg=self._fg_disposed)
                self.disposed_badge.config(bg=self._bg_disposed)
                self.disposed_badge.pack(side=tk.LEFT, padx=(8, 0))
            else:
                self.title_block.config(bg=self._bg_normal)
                self.card_title_label.config(bg=self._bg_normal)
                self.title_cow_id_label.config(bg=self._bg_normal, fg="#1565c0")
                self.title_jpn10_label.config(bg=self._bg_normal, fg="#607d8b")
                self.cow_id_entry.config(bg="white", fg="#263238")
                self.disposed_badge.pack_forget()
            
            # CowCardに牛の情報を読み込む（内部で _notify_embedded_history_refresh により右ペインも更新）
            self.cow_card.load_cow(cow_auto_id)
            
        except Exception as e:
            import traceback
            print(f"ERROR: load_cow で例外が発生しました: {e}")
            traceback.print_exc()
            messagebox.showerror("エラー", f"個体情報の読み込みに失敗しました: {e}")
    
    def _is_cow_disposed(self, cow_auto_id: int) -> bool:
        """
        個体が売却または死亡・淘汰により牛群に存在しないかをチェック
        
        Args:
            cow_auto_id: 牛の auto_id
            
        Returns:
            売却または死亡・淘汰されている場合 True
        """
        events = self.db.get_events_by_cow(cow_auto_id, include_deleted=False)
        for event in events:
            event_number = event.get('event_number')
            if event_number in (RuleEngine.EVENT_SOLD, RuleEngine.EVENT_DEAD):
                return True
        return False
    
    def _on_prev_cow(self):
        """前の個体ボタンがクリックされたときの処理"""
        if self.current_cow_auto_id is None:
            return
        
        try:
            # 全個体を取得（cow_id順）
            all_cows = self.db.get_all_cows()
            if not all_cows:
                return
            
            # 現在の個体の位置を探す
            current_index = None
            for i, cow in enumerate(all_cows):
                if cow.get('auto_id') == self.current_cow_auto_id:
                    current_index = i
                    break
            
            if current_index is None:
                return
            
            # 前の個体を取得（循環しない）
            if current_index > 0:
                prev_cow = all_cows[current_index - 1]
                prev_auto_id = prev_cow.get('auto_id')
                if prev_auto_id:
                    self.load_cow(prev_auto_id)
        
        except Exception as e:
            import traceback
            print(f"ERROR: _on_prev_cow で例外が発生しました: {e}")
            traceback.print_exc()
    
    def _on_next_cow(self):
        """次の個体ボタンがクリックされたときの処理"""
        if self.current_cow_auto_id is None:
            return
        
        try:
            # 全個体を取得（cow_id順）
            all_cows = self.db.get_all_cows()
            if not all_cows:
                return
            
            # 現在の個体の位置を探す
            current_index = None
            for i, cow in enumerate(all_cows):
                if cow.get('auto_id') == self.current_cow_auto_id:
                    current_index = i
                    break
            
            if current_index is None:
                return
            
            # 次の個体を取得（循環しない）
            if current_index < len(all_cows) - 1:
                next_cow = all_cows[current_index + 1]
                next_auto_id = next_cow.get('auto_id')
                if next_auto_id:
                    self.load_cow(next_auto_id)
        
        except Exception as e:
            import traceback
            print(f"ERROR: _on_next_cow で例外が発生しました: {e}")
            traceback.print_exc()
    
    def _on_switch_cow(self):
        """個体切り替えボタンがクリックされたときの処理"""
        try:
            # 個体ID入力欄から値を取得
            cow_id_input = self.cow_id_entry.get().strip()
            if not cow_id_input:
                messagebox.showwarning("警告", "個体IDを入力してください")
                return
            
            # 4桁にゼロパディング（例: 980 → 0980）
            padded_id = cow_id_input.zfill(4)
            
            # 4桁IDで検索（複数件取得可能）
            cows = self.db.get_cows_by_id(padded_id)
            
            if not cows:
                # 見つからない場合は、正規化されたIDで検索を試みる
                normalized_id = cow_id_input.lstrip('0')  # 左ゼロを除去
                cow = self.db.get_cow_by_normalized_id(normalized_id)
                if cow:
                    cows = [cow]
            
            if not cows:
                messagebox.showerror("エラー", f"ID {cow_id_input} の個体は見つかりませんでした")
                return
            
            # 1件のみの場合はそのまま表示
            if len(cows) == 1:
                cow = cows[0]
                cow_auto_id = cow.get('auto_id')
                if cow_auto_id:
                    self.load_cow(cow_auto_id)
                return
            
            # 複数件見つかった場合は選択ダイアログを表示
            selected_cow = self._show_cow_selection_dialog(cows, padded_id)
            if selected_cow:
                cow_auto_id = selected_cow.get('auto_id')
                if cow_auto_id:
                    self.load_cow(cow_auto_id)
        
        except Exception as e:
            import traceback
            print(f"ERROR: _on_switch_cow で例外が発生しました: {e}")
            traceback.print_exc()
            messagebox.showerror("エラー", f"個体の切り替えに失敗しました: {e}")
    
    def _show_cow_selection_dialog(self, cows, cow_id: str):
        """4桁IDが重複している場合の個体選択ダイアログ"""
        from typing import List, Dict, Any
        
        dialog = tk.Toplevel(self.window)
        dialog.title("個体選択")
        dialog.geometry("400x300")
        
        ttk.Label(dialog, text=f"拡大4桁ID {cow_id} に該当する個体が複数見つかりました。\n個体識別番号から選択してください。").pack(pady=10)
        
        listbox = tk.Listbox(dialog, height=10)
        listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        for cow in cows:
            jpn10 = cow.get('jpn10', '')
            display_text = f"{jpn10}"
            listbox.insert(tk.END, display_text)
        
        selected_cow = None
        
        def on_ok():
            nonlocal selected_cow
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning("警告", "個体を選択してください。")
                return
            idx = sel[0]
            selected_cow = cows[idx]
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        listbox.bind('<Double-Button-1>', lambda e: on_ok())
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        dialog.wait_window()
        return selected_cow
    
    def _on_delete_cow(self):
        """個体削除ボタンがクリックされたときの処理"""
        if self.current_cow_auto_id is None:
            messagebox.showwarning("警告", "削除する個体が選択されていません。")
            return
        
        # 個体情報を取得
        cow = self.db.get_cow_by_auto_id(self.current_cow_auto_id)
        if not cow:
            messagebox.showerror("エラー", "個体が見つかりません")
            return
        
        cow_id = cow.get('cow_id', '')
        jpn10 = cow.get('jpn10', '')
        
        # 確認ダイアログ
        result = messagebox.askyesno(
            "確認",
            f"個体を削除しますか？\n\n"
            f"管理番号: {cow_id}\n"
            f"個体識別番号: {jpn10}\n\n"
            f"※ この操作は取り消せません。\n"
            f"※ 関連するイベントも削除されます。"
        )
        
        if result:
            try:
                # 個体を削除
                self.db.delete_cow(self.current_cow_auto_id)
                messagebox.showinfo("完了", "個体を削除しました")
                
                # ウィンドウを閉じる
                self.window.destroy()
                
            except Exception as e:
                import logging
                logging.error(f"個体削除エラー: {e}")
                messagebox.showerror("エラー", f"削除に失敗しました:\n{e}")
    
    def _on_close(self):
        """ウィンドウクローズ時の処理"""
        # 別ウィンドウで履歴を開いている場合のみ閉じる（埋め込み時は右ペインなので不要）
        if self.history_window and isinstance(self.history_window.window, tk.Toplevel) and self.history_window.window.winfo_exists():
            self.history_window.window.destroy()
        self.window.destroy()
    
    def show(self):
        """ウィンドウを表示"""
        self.window.deiconify()

        # ダッシュボード裏（HTML内）から起動されるケースでは、
        # 単なる focus_set() だけだとフォーカスが奪われず裏に回ることがある。
        # いったん topmost にして前面表示し、その後解除する。
        try:
            self.window.attributes('-topmost', True)
        except tk.TclError:
            pass

        self.window.lift()

        # focus_force がある環境では強制フォーカスを試す
        try:
            self.window.focus_force()
        except Exception:
            self.window.focus_set()

        def _unset_topmost():
            try:
                self.window.attributes('-topmost', False)
            except tk.TclError:
                pass
            # topmost解除直後に背面化するケースがあるため、もう一度前面化を試みる
            try:
                self.window.lift()
                self.window.focus_force()
            except Exception:
                try:
                    self.window.focus_set()
                except Exception:
                    pass

        try:
            # 描画・別ウィンドウ更新が落ち着くまで少し長めに保持
            self.window.after(1200, _unset_topmost)
        except Exception:
            _unset_topmost()


