"""
FALCON2 - 長期空胎牛 月次アラート
空胎300日以上の牛を月1回リストアップして確認を促す。
"""

import json
import logging
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, date
from pathlib import Path
from typing import List, Dict, Any, Optional

logger = logging.getLogger(__name__)

# 繁殖コード（RuleEngineと同値）
_RC_STOPPED   = 1
_RC_FRESH     = 2
_RC_BRED      = 3
_RC_OPEN      = 4
_RC_PREGNANT  = 5
_RC_DRY       = 6

_CHECK_FILE     = "long_open_check.json"
_THRESHOLD_DAYS = 300

# イベント番号
_EV_CALV = 202
_EV_SOLD = 205
_EV_DEAD = 206

_RC_LABELS = {
    _RC_STOPPED:  "繁殖停止",
    _RC_FRESH:    "新分娩",
    _RC_BRED:     "授精済み",
    _RC_OPEN:     "空胎",
    _RC_PREGNANT: "妊娠",
    _RC_DRY:      "乾乳",
}


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
        logger.warning(f"長期空胎チェック記録エラー: {e}")


# ---------------------------------------------------------------------------
# 対象牛の抽出
# ---------------------------------------------------------------------------

def find_long_open_cows(db) -> List[Dict[str, Any]]:
    today = date.today()
    results = []
    try:
        all_cows = db.get_all_cows()
    except Exception as e:
        logger.error(f"全牛取得エラー: {e}")
        return []

    for cow in all_cows:
        rc = cow.get("rc")
        if rc not in (_RC_FRESH, _RC_BRED, _RC_OPEN):
            continue
        clvd_str = cow.get("clvd")
        if not clvd_str:
            continue
        try:
            clvd_date = datetime.strptime(clvd_str, "%Y-%m-%d").date()
        except (ValueError, TypeError):
            continue
        dim = (today - clvd_date).days
        if dim < _THRESHOLD_DAYS:
            continue
        results.append({
            "cow_id":   cow.get("cow_id", ""),
            "jpn10":    cow.get("jpn10", ""),
            "lact":     cow.get("lact", ""),
            "clvd":     clvd_str,
            "dim":      dim,
            "rc":       rc,
            "rc_label": _RC_LABELS.get(rc, str(rc)),
            "auto_id":  cow.get("auto_id"),
        })

    results.sort(key=lambda r: r["dim"], reverse=True)
    return results


# ---------------------------------------------------------------------------
# イベント登録ダイアログ（死亡廃用 / 売却）
# ---------------------------------------------------------------------------

def show_event_dialog(parent: tk.Toplevel, cow: Dict, event_number: int,
                      db, rule_engine, on_success) -> None:
    """
    分娩（202）・売却（205）・死亡廃用（206）のイベント入力ダイアログ。
    登録完了後に on_success() を呼び出す。
    """
    _meta = {
        _EV_CALV: ("分娩の登録",    "#E3F2FD", "分娩日:"),
        _EV_SOLD: ("売却の登録",    "#E8F5E9", "売却日:"),
        _EV_DEAD: ("死亡廃用の登録", "#FFEBEE", "死亡廃用日:"),
    }
    title, color, label = _meta.get(event_number, ("イベント登録", "#F5F5F5", "日付:"))

    dlg = tk.Toplevel(parent)
    dlg.title(title)
    dlg.resizable(False, False)
    dlg.grab_set()
    dlg.transient(parent)

    # 中央配置
    dlg.update_idletasks()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    dw, dh = 380, 230
    dx = parent.winfo_x() + (pw - dw) // 2
    dy = parent.winfo_y() + (ph - dh) // 2
    dlg.geometry(f"{dw}x{dh}+{dx}+{dy}")

    # ヘッダー
    hdr = tk.Frame(dlg, bg=color, bd=1, relief=tk.SOLID)
    hdr.pack(fill=tk.X, padx=10, pady=(10, 6))
    cow_label = f"ID: {cow['cow_id']}  /  {cow['jpn10']}"
    tk.Label(hdr, text=cow_label, bg=color,
             font=("Meiryo UI", 10, "bold"), anchor=tk.W).pack(
        fill=tk.X, padx=8, pady=4)

    # フォーム
    form = ttk.Frame(dlg)
    form.pack(fill=tk.X, padx=14, pady=4)

    ttk.Label(form, text=label, font=("Meiryo UI", 9)).grid(
        row=0, column=0, sticky=tk.W, pady=4, padx=(0, 8))
    date_var = tk.StringVar(value=date.today().strftime("%Y-%m-%d"))
    date_entry = ttk.Entry(form, textvariable=date_var, width=14,
                           font=("Meiryo UI", 9))
    date_entry.grid(row=0, column=1, sticky=tk.W, pady=4)

    ttk.Label(form, text="備考:", font=("Meiryo UI", 9)).grid(
        row=1, column=0, sticky=tk.W, pady=4, padx=(0, 8))
    note_entry = ttk.Entry(form, width=28, font=("Meiryo UI", 9))
    note_entry.grid(row=1, column=1, sticky=tk.EW, pady=4)
    form.columnconfigure(1, weight=1)

    err_var = tk.StringVar()
    tk.Label(dlg, textvariable=err_var, fg="red",
             font=("Meiryo UI", 8)).pack(padx=14, anchor=tk.W)

    # ボタン
    def _register():
        ev_date = date_var.get().strip()
        try:
            datetime.strptime(ev_date, "%Y-%m-%d")
        except ValueError:
            err_var.set("日付は YYYY-MM-DD 形式で入力してください")
            return
        try:
            event_data = {
                "cow_auto_id":  cow["auto_id"],
                "event_number": event_number,
                "event_date":   ev_date,
                "json_data":    None,
                "note":         note_entry.get().strip() or None,
            }
            event_id = db.insert_event(event_data)
            if rule_engine:
                rule_engine.on_event_added(event_id)
            dlg.destroy()
            on_success()
        except Exception as e:
            logger.error(f"イベント登録エラー: {e}", exc_info=True)
            err_var.set(f"登録に失敗しました: {e}")

    btn_frame = ttk.Frame(dlg)
    btn_frame.pack(fill=tk.X, padx=14, pady=(4, 10))
    ttk.Button(btn_frame, text="登録", command=_register).pack(side=tk.RIGHT, padx=(6, 0))
    ttk.Button(btn_frame, text="キャンセル", command=dlg.destroy).pack(side=tk.RIGHT)

    date_entry.focus_set()
    date_entry.select_range(0, tk.END)
    dlg.bind("<Return>", lambda _: _register())


# ---------------------------------------------------------------------------
# アラートウィンドウ
# ---------------------------------------------------------------------------

def show_alert_window(root: tk.Tk, farm_path: Path, db,
                      formula_engine=None, rule_engine=None) -> None:
    """長期空胎牛アラートウィンドウを表示（月1回）"""
    if not should_run_this_month(farm_path):
        return

    cows = find_long_open_cows(db)
    record_checked(farm_path)

    if not cows:
        logger.info("長期空胎牛なし（%d日未満）", _THRESHOLD_DAYS)
        return

    # ---- ウィンドウ ----
    win = tk.Toplevel(root)
    win.title("⚠ 長期空胎牛の確認（月次チェック）")
    win.resizable(True, True)
    win.grab_set()

    root.update_idletasks()
    w, h = 740, 520
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
            f"⚠  空胎 {_THRESHOLD_DAYS} 日以上の牛が {n} 頭います" if n > 0
            else "✓  すべての確認が完了しました"
        )

    tk.Label(msg_frame, textvariable=remaining_var, bg="white", fg="#856404",
             font=("Meiryo UI", 11, "bold"), anchor=tk.W).pack(
        fill=tk.X, padx=10, pady=(6, 2))
    tk.Label(
        msg_frame,
        text=(
            "売却・死亡・繁殖停止などのイベントが未登録の可能性があります。\n"
            "牛群の在籍を確認し、1頭ずつ処理してください。"
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

    cols = ("cow_id", "jpn10", "lact", "clvd", "dim", "rc_label")
    col_defs = {
        "cow_id":   ("ID",        60,  tk.CENTER),
        "jpn10":    ("個体識別番号", 150, tk.W),
        "lact":     ("産次",       50,  tk.CENTER),
        "clvd":     ("最終分娩日",  110, tk.CENTER),
        "dim":      ("空胎日数",    80,  tk.CENTER),
        "rc_label": ("繁殖区分",    90,  tk.CENTER),
    }

    vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
    tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                        yscrollcommand=vsb.set, selectmode="browse")
    vsb.config(command=tree.yview)
    vsb.pack(side=tk.RIGHT, fill=tk.Y)
    tree.pack(fill=tk.BOTH, expand=True)

    for col in cols:
        lbl, width, anchor = col_defs[col]
        tree.heading(col, text=lbl)
        tree.column(col, width=width, anchor=anchor)

    iid_to_cow: Dict[str, Dict] = {}

    for cow in cows:
        iid = str(cow["auto_id"])
        tree.insert("", tk.END, iid=iid, values=(
            cow["cow_id"],
            cow["jpn10"],
            cow["lact"] if cow["lact"] is not None else "-",
            cow["clvd"],
            f'{cow["dim"]}日',
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
        """行を削除し、次の行を自動選択してボタン状態を更新"""
        next_iid = tree.next(iid) or tree.prev(iid)
        tree.delete(iid)
        iid_to_cow.pop(iid, None)
        if next_iid and tree.exists(next_iid):
            tree.selection_set(next_iid)
        _update_header()
        _refresh_buttons()
        # 残り0頭になったら閉じるボタンにフォーカス
        if not tree.get_children():
            close_btn.focus_set()

    def _refresh_buttons():
        state = "normal" if _selected_iid() else "disabled"
        dead_btn.config(state=state)
        sold_btn.config(state=state)
        skip_btn.config(state=state)
        card_btn.config(state=state)

    # ---- アクション ----
    def _do_dead():
        cow = _selected_cow()
        if not cow:
            return
        iid = str(cow["auto_id"])
        show_event_dialog(win, cow, _EV_DEAD, db, rule_engine,
                          on_success=lambda: _remove_row(iid))

    def _do_sold():
        cow = _selected_cow()
        if not cow:
            return
        iid = str(cow["auto_id"])
        show_event_dialog(win, cow, _EV_SOLD, db, rule_engine,
                          on_success=lambda: _remove_row(iid))

    def _do_skip():
        """まだいる：イベント登録なしでリストから除外"""
        iid = _selected_iid()
        if iid:
            _remove_row(iid)

    def _open_card(cow: Optional[Dict] = None):
        if cow is None:
            cow = _selected_cow()
        if not cow or not cow.get("auto_id"):
            return
        try:
            falcon_root = Path(__file__).parent.parent.parent
            config_default = falcon_root / "config_default"
            event_dict_path = config_default / "event_dictionary.json"
            item_dict_path  = config_default / "item_dictionary.json"
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

    # ---- 選択変化でボタン状態更新 ----
    tree.bind("<<TreeviewSelect>>", lambda _: _refresh_buttons())

    # ---- ダブルクリック → 個体カード ----
    tree.bind("<Double-Button-1>", lambda _: _open_card())

    # ---- 右クリックメニュー ----
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

    dead_btn = ttk.Button(action_frame, text="↓ 死亡廃用", width=14, command=_do_dead, state="disabled")
    dead_btn.pack(side=tk.LEFT, padx=(10, 6), pady=6)

    sold_btn = ttk.Button(action_frame, text="↓ 売却", width=12, command=_do_sold, state="disabled")
    sold_btn.pack(side=tk.LEFT, padx=(0, 6), pady=6)

    skip_btn = ttk.Button(action_frame, text="✔ まだいる", width=12, command=_do_skip, state="disabled")
    skip_btn.pack(side=tk.LEFT, padx=(0, 6), pady=6)

    ttk.Separator(action_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=4)

    card_btn = ttk.Button(action_frame, text="個体カードを開く", command=_open_card, state="disabled")
    card_btn.pack(side=tk.LEFT, padx=(0, 10), pady=6)

    # ---- ステータス + 閉じる ----
    status_var = tk.StringVar()
    bottom = ttk.Frame(win)
    bottom.pack(fill=tk.X, padx=12, pady=(0, 10))

    tk.Label(bottom, textvariable=status_var,
             font=("Meiryo UI", 8), foreground="#1B5E20", anchor=tk.W).pack(side=tk.LEFT)
    tk.Label(bottom, text="※ このチェックは月に1回表示されます",
             font=("Meiryo UI", 8), foreground="gray").pack(side=tk.LEFT, padx=(8, 0))

    close_btn = ttk.Button(bottom, text="閉じる", command=win.destroy)
    close_btn.pack(side=tk.RIGHT)
