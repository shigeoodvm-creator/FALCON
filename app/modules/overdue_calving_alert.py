"""
FALCON2 - 分娩過期牛 月次アラート
分娩予定日から15日以上経過した牛を月1回リストアップして確認を促す。
"""

import json
import logging
import tkinter as tk
from tkinter import ttk
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import List, Dict, Any, Optional

from modules.long_open_alert import show_event_dialog

logger = logging.getLogger(__name__)

_RC_PREGNANT = 5
_RC_DRY      = 6

_EV_CALV = 202
_EV_SOLD = 205

_CHECK_FILE      = "overdue_calving_check.json"
_THRESHOLD_DAYS  = 15   # 分娩予定日を何日超過したら警告するか


# ---------------------------------------------------------------------------
# チェック実行判定
# ---------------------------------------------------------------------------

def _check_file_path(farm_path: Path) -> Path:
    return farm_path / _CHECK_FILE


def should_run_this_month(farm_path: Path) -> bool:
    fp = _check_file_path(farm_path)
    if not fp.exists():
        return True
    try:
        data = json.loads(fp.read_text(encoding="utf-8"))
        return data.get("last_check_ym", "") != date.today().strftime("%Y-%m")
    except Exception:
        return True


def record_checked(farm_path: Path) -> None:
    fp = _check_file_path(farm_path)
    try:
        fp.write_text(
            json.dumps({"last_check_ym": date.today().strftime("%Y-%m")},
                       ensure_ascii=False, indent=2),
            encoding="utf-8"
        )
    except Exception as e:
        logger.warning(f"分娩過期チェック記録エラー: {e}")


# ---------------------------------------------------------------------------
# 対象牛の抽出
# ---------------------------------------------------------------------------

def find_overdue_calving_cows(db, formula_engine) -> List[Dict[str, Any]]:
    """
    分娩予定日から THRESHOLD_DAYS 日以上経過した牛を抽出する。

    対象: rc が 妊娠(5) または 乾乳(6) の牛で、
          計算された分娩予定日が今日 - THRESHOLD_DAYS 日より前。
    """
    today = date.today()
    cutoff = today - timedelta(days=_THRESHOLD_DAYS)
    results = []

    try:
        all_cows = db.get_all_cows()
    except Exception as e:
        logger.error(f"全牛取得エラー: {e}")
        return []

    for cow in all_cows:
        rc = cow.get("rc")
        if rc not in (_RC_PREGNANT, _RC_DRY):
            continue

        auto_id = cow.get("auto_id")
        if not auto_id:
            continue

        # 分娩予定日をイベント履歴から計算
        try:
            events = db.get_events_by_cow(auto_id, include_deleted=False)
            due_date_str = formula_engine._calculate_due_date(events, cow.get("clvd"))
        except Exception as e:
            logger.debug(f"分娩予定日計算エラー (auto_id={auto_id}): {e}")
            continue

        if not due_date_str:
            continue

        try:
            due_date = datetime.strptime(due_date_str, "%Y-%m-%d").date()
        except (ValueError, TypeError):
            continue

        if due_date > cutoff:
            continue  # まだ閾値未満

        overdue_days = (today - due_date).days

        results.append({
            "cow_id":       cow.get("cow_id", ""),
            "jpn10":        cow.get("jpn10", ""),
            "lact":         cow.get("lact", ""),
            "clvd":         cow.get("clvd", ""),
            "due_date":     due_date_str,
            "overdue_days": overdue_days,
            "rc":           rc,
            "rc_label":     "妊娠" if rc == _RC_PREGNANT else "乾乳",
            "auto_id":      auto_id,
        })

    results.sort(key=lambda r: r["overdue_days"], reverse=True)
    return results


# ---------------------------------------------------------------------------
# アラートウィンドウ
# ---------------------------------------------------------------------------

def show_alert_window(root: tk.Tk, farm_path: Path, db,
                      formula_engine=None, rule_engine=None,
                      force: bool = False) -> None:
    """分娩過期牛アラートウィンドウを表示。
    force=True の場合は月次制限をバイパスし、チェック日も更新しない。
    """
    if not force and not should_run_this_month(farm_path):
        return

    if formula_engine is None:
        logger.warning("formula_engine が未指定のため分娩過期チェックをスキップ")
        return

    cows = find_overdue_calving_cows(db, formula_engine)

    if not force:
        record_checked(farm_path)

    if not cows:
        if force:
            from tkinter import messagebox
            messagebox.showinfo(
                "分娩超過チェック",
                f"分娩予定日から {_THRESHOLD_DAYS} 日以上経過した牛はいません。",
                parent=root
            )
        else:
            logger.info("分娩過期牛なし（%d日未満）", _THRESHOLD_DAYS)
        return

    # ---- ウィンドウ ----
    win = tk.Toplevel(root)
    win.title("分娩超過牛の確認" if force else "⚠ 分娩超過牛の確認（月次チェック）")
    win.resizable(True, True)
    win.grab_set()

    root.update_idletasks()
    w, h = 780, 520
    sx = root.winfo_x() + max(0, (root.winfo_width()  - w) // 2)
    sy = root.winfo_y() + max(0, (root.winfo_height() - h) // 2)
    win.geometry(f"{w}x{h}+{sx}+{sy}")

    # ---- 説明 ----
    msg_frame = tk.Frame(win, bg="white", bd=2, relief=tk.SOLID,
                         highlightbackground="#E6A817", highlightthickness=2)
    msg_frame.pack(fill=tk.X, padx=12, pady=(12, 4))

    remaining_var = tk.StringVar()

    def _update_header():
        n = len(tree.get_children())
        remaining_var.set(
            f"⚠  分娩予定日から {_THRESHOLD_DAYS} 日以上経過した牛が {n} 頭います" if n > 0
            else "✓  すべての確認が完了しました"
        )

    tk.Label(msg_frame, textvariable=remaining_var, bg="white", fg="#856404",
             font=("Meiryo UI", 11, "bold"), anchor=tk.W).pack(
        fill=tk.X, padx=10, pady=(6, 2))
    tk.Label(
        msg_frame,
        text=(
            "分娩が未登録のまま、または売却・死亡が登録されていない可能性があります。\n"
            "牛群の在籍・状態を確認し、1頭ずつ処理してください。"
        ),
        bg="white", fg="#856404", font=("Meiryo UI", 9),
        justify=tk.LEFT, anchor=tk.W
    ).pack(fill=tk.X, padx=10, pady=(0, 6))

    # ---- ヒント ----
    tk.Label(
        win,
        text="行を選択 → 下のボタンで処理　／　ダブルクリック → 個体カードを開く　／　右クリック → JPN10コピー",
        font=("Meiryo UI", 8), foreground="gray", anchor=tk.W
    ).pack(fill=tk.X, padx=14, pady=(2, 0))

    # ---- テーブル ----
    tree_frame = ttk.Frame(win)
    tree_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)

    cols = ("cow_id", "jpn10", "lact", "clvd", "due_date", "overdue_days", "rc_label")
    col_defs = {
        "cow_id":       ("ID",        60,  tk.CENTER),
        "jpn10":        ("個体識別番号", 150, tk.W),
        "lact":         ("産次",       50,  tk.CENTER),
        "clvd":         ("前回分娩日",  110, tk.CENTER),
        "due_date":     ("分娩予定日",  110, tk.CENTER),
        "overdue_days": ("超過日数",    80,  tk.CENTER),
        "rc_label":     ("繁殖区分",    80,  tk.CENTER),
    }

    vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
    tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                        yscrollcommand=vsb.set, selectmode="browse")
    vsb.config(command=tree.yview)
    vsb.pack(side=tk.RIGHT, fill=tk.Y)
    tree.pack(fill=tk.BOTH, expand=True)

    iid_to_cow: Dict[str, Dict] = {}

    for col in cols:
        lbl, width, anchor = col_defs[col]
        tree.heading(col, text=lbl)
        tree.column(col, width=width, anchor=anchor)

    for cow in cows:
        iid = str(cow["auto_id"])
        tree.insert("", tk.END, iid=iid, values=(
            cow["cow_id"],
            cow["jpn10"],
            cow["lact"] if cow["lact"] is not None else "-",
            cow["clvd"] or "-",
            cow["due_date"],
            f'{cow["overdue_days"]}日超過',
            cow["rc_label"],
        ))
        iid_to_cow[iid] = cow

    _update_header()

    # ---- ヘルパー ----
    def _selected_iid() -> Optional[str]:
        sel = tree.selection()
        return sel[0] if sel else None

    def _selected_cow() -> Optional[Dict]:
        iid = _selected_iid()
        return iid_to_cow.get(iid) if iid else None

    def _remove_row(iid: str):
        next_iid = tree.next(iid) or tree.prev(iid)
        tree.delete(iid)
        iid_to_cow.pop(iid, None)
        if next_iid and tree.exists(next_iid):
            tree.selection_set(next_iid)
        _update_header()
        _refresh_buttons()
        if not tree.get_children():
            close_btn.focus_set()

    def _refresh_buttons():
        state = "normal" if _selected_iid() else "disabled"
        calv_btn.config(state=state)
        sold_btn.config(state=state)
        skip_btn.config(state=state)
        card_btn.config(state=state)

    # ---- アクション ----
    def _do_calv():
        cow = _selected_cow()
        if not cow:
            return
        iid = str(cow["auto_id"])
        try:
            from constants import CONFIG_DEFAULT_DIR
            event_dict_path = CONFIG_DEFAULT_DIR / "event_dictionary.json"

            def _on_calv_saved(_cow_auto_id: int):
                if win.winfo_exists():
                    win.after(0, lambda: _remove_row(iid))

            from ui.event_input import EventInputWindow
            event_window = EventInputWindow.open_or_focus(
                parent=win,
                db_handler=db,
                rule_engine=rule_engine,
                event_dictionary_path=event_dict_path if event_dict_path.exists() else None,
                cow_auto_id=cow["auto_id"],
                on_saved=_on_calv_saved,
                farm_path=farm_path,
                default_event_number=_EV_CALV,
                allowed_event_numbers=[_EV_CALV],
            )
            event_window.show()
        except Exception as e:
            logger.error(f"分娩入力ウィンドウを開けませんでした: {e}", exc_info=True)

    def _do_sold():
        cow = _selected_cow()
        if not cow:
            return
        iid = str(cow["auto_id"])
        show_event_dialog(win, cow, _EV_SOLD, db, rule_engine,
                          on_success=lambda: _remove_row(iid))

    def _do_skip():
        iid = _selected_iid()
        if iid:
            _remove_row(iid)

    def _open_card(cow: Optional[Dict] = None):
        if cow is None:
            cow = _selected_cow()
        if not cow or not cow.get("auto_id"):
            return
        try:
            from constants import CONFIG_DEFAULT_DIR
            event_dict_path = CONFIG_DEFAULT_DIR / "event_dictionary.json"
            item_dict_path  = CONFIG_DEFAULT_DIR / "item_dictionary.json"
            from ui.cow_card_window import CowCardWindow
            cw = CowCardWindow(
                parent=root,
                db_handler=db,
                formula_engine=formula_engine,
                rule_engine=rule_engine,
                event_dictionary_path=event_dict_path if event_dict_path.exists() else None,
                item_dictionary_path=item_dict_path  if item_dict_path.exists()  else None,
                cow_auto_id=cow["auto_id"]
            )
            cw.show()
        except Exception as e:
            logger.error(f"個体カードを開けませんでした: {e}", exc_info=True)

    def _copy_jpn10(cow: Optional[Dict] = None):
        if cow is None:
            cow = _selected_cow()
        if not cow or not cow.get("jpn10"):
            return
        jpn10 = cow["jpn10"]
        win.clipboard_clear()
        win.clipboard_append(jpn10)
        status_var.set(f"コピーしました: {jpn10}")
        win.after(2500, lambda: status_var.set(""))

    tree.bind("<<TreeviewSelect>>", lambda _: _refresh_buttons())
    tree.bind("<Double-Button-1>", lambda _: _open_card())

    ctx = tk.Menu(win, tearoff=0)

    def _show_ctx(event):
        iid = tree.identify_row(event.y)
        if not iid:
            return
        tree.selection_set(iid)
        cow = iid_to_cow.get(iid)
        if not cow:
            return
        jpn10 = cow.get("jpn10", "")
        ctx.delete(0, tk.END)
        ctx.add_command(
            label=f"JPN10「{jpn10}」をコピー" if jpn10 else "JPN10をコピー（番号なし）",
            command=lambda c=cow: _copy_jpn10(c),
            state="normal" if jpn10 else "disabled"
        )
        ctx.add_separator()
        ctx.add_command(label="個体カードを開く", command=lambda c=cow: _open_card(c))

    tree.bind("<Button-3>", _show_ctx)

    # ---- アクションボタン ----
    action_frame = ttk.LabelFrame(win, text="選択した牛の処理")
    action_frame.pack(fill=tk.X, padx=12, pady=(2, 4))

    calv_btn = ttk.Button(action_frame, text="分娩を登録", width=14,
                          command=_do_calv, state="disabled")
    calv_btn.pack(side=tk.LEFT, padx=(10, 6), pady=6)

    sold_btn = ttk.Button(action_frame, text="↓ 売却", width=12,
                          command=_do_sold, state="disabled")
    sold_btn.pack(side=tk.LEFT, padx=(0, 6), pady=6)

    skip_btn = ttk.Button(action_frame, text="✔ そのまま", width=12,
                          command=_do_skip, state="disabled")
    skip_btn.pack(side=tk.LEFT, padx=(0, 6), pady=6)

    ttk.Separator(action_frame, orient=tk.VERTICAL).pack(
        side=tk.LEFT, fill=tk.Y, padx=6, pady=4)

    card_btn = ttk.Button(action_frame, text="個体カードを開く",
                          command=_open_card, state="disabled")
    card_btn.pack(side=tk.LEFT, padx=(0, 10), pady=6)

    # ---- ステータス + 閉じる ----
    status_var = tk.StringVar()
    bottom = ttk.Frame(win)
    bottom.pack(fill=tk.X, padx=12, pady=(0, 10))

    tk.Label(bottom, textvariable=status_var,
             font=("Meiryo UI", 8), foreground="#1B5E20", anchor=tk.W).pack(side=tk.LEFT)
    if not force:
        tk.Label(bottom, text="※ このチェックは月に1回自動表示されます",
                 font=("Meiryo UI", 8), foreground="gray").pack(side=tk.LEFT, padx=(8, 0))

    close_btn = ttk.Button(bottom, text="閉じる", command=win.destroy)
    close_btn.pack(side=tk.RIGHT)
