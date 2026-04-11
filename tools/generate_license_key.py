"""
FALCON2 - ライセンスキー生成ツール（開発者専用・配布しないこと）

使い方:
  python tools/generate_license_key.py

  または引数指定:
  python tools/generate_license_key.py --type vet --expiry 2027-03-31 --id VET001
"""

import argparse
import base64
import hashlib
import hmac
import json
import sys
from datetime import date, timedelta

# ─── license_manager.py と必ず同じ値を使うこと ───
_SECRET = b"FALCON2-REPLACE-THIS-WITH-A-LONG-RANDOM-SECRET-2026"

# 端末課金・農場数無制限。価格は顧客ごとに交渉で決定する（参考価格）
CUSTOMER_TYPES = {
    "nosai": {"label": "NOSAI関連獣医師",      "price_ref": "〜20,000円/年"},
    "vet":   {"label": "開業獣医師",           "price_ref": "〜50,000円/年"},
    "org":   {"label": "JA・乳検等関連団体",   "price_ref": "〜30,000円/年"},
}


def generate_key(customer_type: str, expiry: date, unique_id: str) -> str:
    payload = {
        "type":   customer_type,
        "expiry": expiry.isoformat(),
        "id":     unique_id,
    }
    payload_bytes = json.dumps(payload, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    payload_b64   = base64.b32encode(payload_bytes).decode().rstrip("=")
    checksum      = hmac.new(_SECRET, payload_bytes, hashlib.sha256).hexdigest()[:6].upper()
    return f"FALCON-{payload_b64}-{checksum}"


def interactive():
    print("=" * 60)
    print("  FALCON2 ライセンスキー生成ツール（開発者専用）")
    print("=" * 60)
    print()

    # 顧客種別
    print("顧客種別を選択してください:")
    keys = list(CUSTOMER_TYPES.keys())
    for i, k in enumerate(keys, 1):
        info = CUSTOMER_TYPES[k]
        print(f"  {i}. {k:8s} {info['label']:20s} 参考: {info['price_ref']}")
    print()

    while True:
        try:
            choice = int(input("番号を入力 > ").strip())
            if 1 <= choice <= len(keys):
                ctype = keys[choice - 1]
                break
        except ValueError:
            pass
        print(f"1〜{len(keys)}の番号を入力してください")

    # 有効期限
    print()
    default_expiry = date(date.today().year + 1, 3, 31)  # デフォルト: 来年3月末
    expiry_str = input(f"有効期限（YYYY-MM-DD）[{default_expiry}] > ").strip()
    if not expiry_str:
        expiry = default_expiry
    else:
        try:
            expiry = date.fromisoformat(expiry_str)
        except ValueError:
            print("日付の形式が正しくありません。デフォルトを使用します。")
            expiry = default_expiry

    # ユニークID
    print()
    uid = input("顧客ID（例: VET001, JA-TOKACHI） > ").strip().upper()
    if not uid:
        uid = f"{ctype.upper()}-{date.today().strftime('%Y%m')}"

    # 生成
    key = generate_key(ctype, expiry, uid)

    print()
    print("=" * 60)
    print("  生成されたライセンスキー")
    print("=" * 60)
    print(f"  キー      : {key}")
    print(f"  顧客種別  : {CUSTOMER_TYPES[ctype]['label']}")
    print(f"  有効期限  : {expiry}")
    print(f"  顧客ID    : {uid}")
    print(f"  参考価格  : {CUSTOMER_TYPES[ctype]['price_ref']}")
    print("=" * 60)
    print()

    # 発行ログ
    log_entry = {
        "issued_at": date.today().isoformat(),
        "type":      ctype,
        "expiry":    expiry.isoformat(),
        "id":        uid,
        "key":       key,
    }
    import os
    log_path = os.path.join(os.path.dirname(__file__), "issued_keys.jsonl")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(json.dumps(log_entry, ensure_ascii=False) + "\n")
    print(f"  発行ログ: {log_path}")
    print()


def from_args():
    parser = argparse.ArgumentParser(description="FALCON2 ライセンスキー生成")
    parser.add_argument("--type",   required=True, choices=CUSTOMER_TYPES.keys())
    parser.add_argument("--expiry", required=True, help="YYYY-MM-DD")
    parser.add_argument("--id",     required=True, help="顧客ID")
    args = parser.parse_args()

    expiry = date.fromisoformat(args.expiry)
    key    = generate_key(args.type, expiry, args.id.upper())
    print(key)


if __name__ == "__main__":
    if len(sys.argv) > 1:
        from_args()
    else:
        interactive()
