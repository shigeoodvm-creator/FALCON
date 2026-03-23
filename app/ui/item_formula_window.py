"""
FALCON2 - 計算式確認・編集ウィンドウ
calc/event 項目の formula / source を表示・更新する
"""

import json
import logging
import os
from pathlib import Path
from typing import Dict, Any, Optional, Callable

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

logger = logging.getLogger(__name__)


class ItemFormulaWindow:
    """計算式確認・編集ウィンドウ"""

    def __init__(
        self,
        parent: tk.Tk,
        item_key: str,
        item_data: Dict[str, Any],
        item_dictionary_path: Path,
        on_saved: Optional[Callable[[], None]] = None,
    ):
        self.parent = parent
        self.item_key = item_key
        self.item_data = item_data.copy()
        self.item_dictionary_path = item_dictionary_path
        self.on_saved = on_saved

        self.origin = item_data.get("origin") or item_data.get("type") or "calc"

        self.window = tk.Toplevel(parent)
        self.window.title(f"計算式を確認 - {item_key}")
        self.window.geometry("760x560")

        self._create_widgets()
        self._populate_fields()

        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        main_frame = ttk.Frame(self.window, padding=12)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text=f"項目キー: {self.item_key}", font=("", 11, "bold")).pack(
            anchor=tk.W, pady=(0, 6)
        )
        label = (
            self.item_data.get("label")
            or self.item_data.get("display_name")
            or self.item_data.get("name_jp")
            or self.item_key
        )
        ttk.Label(main_frame, text=f"表示名: {label}", font=("", 10)).pack(
            anchor=tk.W, pady=(0, 4)
        )
        ttk.Label(main_frame, text=f"origin: {self.origin}", font=("", 10)).pack(
            anchor=tk.W, pady=(0, 10)
        )

        formula_frame = ttk.LabelFrame(main_frame, text="計算式 / source", padding=10)
        formula_frame.pack(fill=tk.BOTH, expand=True)

        self.formula_text = scrolledtext.ScrolledText(formula_frame, height=12)
        self.formula_text.pack(fill=tk.BOTH, expand=True)

        btn_frame = ttk.Frame(formula_frame)
        btn_frame.pack(fill=tk.X, pady=6)
        if self.origin == "calc":
            ttk.Button(
                btn_frame,
                text="AIに計算式を作ってもらう",
                command=self._generate_formula_with_ai,
                width=24,
            ).pack(side=tk.LEFT, padx=4)

        save_frame = ttk.Frame(main_frame)
        save_frame.pack(fill=tk.X, pady=8)
        ttk.Button(save_frame, text="保存", command=self._on_save, width=12).pack(
            side=tk.LEFT, padx=5
        )
        ttk.Button(save_frame, text="キャンセル", command=self.window.destroy, width=12).pack(
            side=tk.LEFT, padx=5
        )

    def _populate_fields(self):
        if self.origin == "event":
            source = self.item_data.get("source") or ""
            self.formula_text.insert(1.0, source)
            self.formula_text.config(state=tk.DISABLED)
        else:
            formula = self.item_data.get("formula") or ""
            self.formula_text.insert(1.0, formula)

    def _on_save(self):
        content = self.formula_text.get(1.0, tk.END).strip()
        try:
            with open(self.item_dictionary_path, "r", encoding="utf-8") as f:
                item_dict = json.load(f)
        except Exception as e:
            messagebox.showerror("エラー", f"item_dictionary.json の読み込みに失敗しました: {e}")
            return

        current = item_dict.get(self.item_key, {}).copy()
        if self.origin == "event":
            current["source"] = content
        else:
            current["formula"] = content
            current["origin"] = current.get("origin") or "calc"
            current["type"] = current.get("type") or "calc"

        item_dict[self.item_key] = current
        try:
            # フォルダが存在しない場合は作成
            if self.item_dictionary_path:
                self.item_dictionary_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.item_dictionary_path, "w", encoding="utf-8") as f:
                json.dump(item_dict, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("完了", "計算式を保存しました")
            if self.on_saved:
                self.on_saved()
            self.window.destroy()
        except Exception as e:
            messagebox.showerror("エラー", f"保存に失敗しました: {e}")

    def _generate_formula_with_ai(self):
        """
        OpenAI API を利用して計算式を生成（環境変数 OPENAI_API_KEY が必要）。
        APIキーやライブラリが無い場合はプロンプトをクリップボードへコピーする。
        """
        prompt = self._build_ai_prompt()

        try:
            import openai
        except Exception:
            self._fallback_ai_prompt(prompt)
            return

        api_key = os.environ.get("OPENAI_API_KEY")
        if not api_key:
            self._fallback_ai_prompt(prompt)
            return

        openai.api_key = api_key
        try:
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You generate a single Python expression for FormulaEngine."},
                    {"role": "user", "content": prompt},
                ],
                temperature=0.2,
                max_tokens=256,
            )
            text = response["choices"][0]["message"]["content"].strip()
            # コードブロックが返る場合を考慮して抽出
            if "```" in text:
                parts = text.split("```")
                if len(parts) >= 3:
                    text = parts[1]
            self.formula_text.delete(1.0, tk.END)
            self.formula_text.insert(1.0, text.strip())
            messagebox.showinfo("完了", "AIが計算式を提案しました。内容を確認してください。")
        except Exception as e:
            logger.error(f"AI生成に失敗: {e}")
            self._fallback_ai_prompt(prompt, error=str(e))

    def _build_ai_prompt(self) -> str:
        description = self.item_data.get("description") or ""
        label = (
            self.item_data.get("label")
            or self.item_data.get("display_name")
            or self.item_data.get("name_jp")
            or self.item_key
        )
        lines = [
            "FormulaEngine 用のPython式を1行で返してください。",
            f"項目キー: {self.item_key}",
            f"表示名: {label}",
            f"説明: {description}",
            "使用可能な変数: cow_id, auto_id, cow, events, result",
            "使用禁止: 個体識別番号（JPN10）などの外部ID",
            "安全な組込み関数のみ使用してください（min, max, sum, len, abs, round）。",
            "返答は式のみを出力し、文脈説明は不要です。",
        ]
        return "\n".join(lines)

    def _fallback_ai_prompt(self, prompt: str, error: Optional[str] = None):
        """API利用不可時のフォールバック（プロンプトをコピー）"""
        try:
            self.window.clipboard_clear()
            self.window.clipboard_append(prompt)
        except Exception:
            pass
        msg = "AI API を利用できません。プロンプトをクリップボードにコピーしました。"
        if error:
            msg += f"\n詳細: {error}"
        messagebox.showinfo("情報", msg)

    def show(self):
        self.window.wait_window()







































