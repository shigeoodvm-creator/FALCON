"""
FALCON2 - ライセンス管理モジュール

【試用期間】30日間
【記録場所】Windowsレジストリ + AppData暗号化ファイル（2箇所で改ざん検知）
【ライセンスキー形式】FALCON-{base32_payload}-{hmac_6chars}
"""

import hashlib
import hmac as _hmac
import json
import logging
import os
import winreg
import base64
from datetime import date, datetime
from pathlib import Path
from typing import Optional, Tuple, Dict, Any

logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
#  ビルド時に必ず変更すること（非公開・バイナリに埋め込まれる）
# ─────────────────────────────────────────────
_SECRET = b"e628d09dd3f7390f46d5cdd21a2043ff2c3bb9a3e61fc93927e7938115fa208654565e60f9169a34"

# AppData の保存先
_APPDATA_DIR  = Path(os.environ.get("APPDATA", "C:/Users/Public")) / "FALCON2"
_LICENSE_FILE = _APPDATA_DIR / "license.dat"
_TRIAL_FILE   = _APPDATA_DIR / ".init"

# Windowsレジストリ
_REG_PATH  = r"Software\FALCON2"
_REG_VALUE = "InstallDate"

TRIAL_DAYS = 30


# ─────────────────────────────────────────────
#  ステータス定数
# ─────────────────────────────────────────────

class LicenseStatus:
    TRIAL_ACTIVE    = "trial_active"     # 試用期間中
    TRIAL_EXPIRED   = "trial_expired"    # 試用期限切れ（未購入）
    LICENSED        = "licensed"         # 有効なライセンス
    LICENSE_EXPIRED = "license_expired"  # ライセンス期限切れ（更新が必要）
    LICENSE_INVALID = "license_invalid"  # 不正なライセンスファイル
    MACHINE_MISMATCH = "machine_mismatch" # 別端末のライセンス

# 顧客種別の定義（端末課金・農場数無制限）
CUSTOMER_TYPES: Dict[str, Dict[str, Any]] = {
    "nosai":      {"label": "NOSAI関連獣医師"},
    "vet":        {"label": "開業獣医師"},
    "org":        {"label": "JA・乳検等関連団体"},
    "individual": {"label": "個人農家"},
}


# ─────────────────────────────────────────────
#  ライセンス情報を保持するデータクラス
# ─────────────────────────────────────────────

class LicenseInfo:
    def __init__(self):
        self.status: str = LicenseStatus.TRIAL_ACTIVE
        self.customer_type: str = ""
        self.customer_label: str = ""
        self.farm_limit: int = 999  # 端末課金のため常に無制限
        self.expiry_date: Optional[date] = None
        self.trial_remaining_days: int = TRIAL_DAYS
        self.license_key: str = ""
        self.unique_id: str = ""

    @property
    def is_usable(self) -> bool:
        return self.status in (LicenseStatus.TRIAL_ACTIVE, LicenseStatus.LICENSED)

    @property
    def trial_warning(self) -> bool:
        """試用残り7日以内は警告バナーを出す"""
        return self.status == LicenseStatus.TRIAL_ACTIVE and self.trial_remaining_days <= 7

    def summary(self) -> str:
        if self.status == LicenseStatus.LICENSED:
            return f"ライセンス済み（{self.customer_label}）　有効期限: {self.expiry_date}"
        if self.status == LicenseStatus.TRIAL_ACTIVE:
            return f"試用版　残り {self.trial_remaining_days} 日"
        if self.status == LicenseStatus.TRIAL_EXPIRED:
            return "試用期間が終了しました。ライセンスを購入してください。"
        if self.status == LicenseStatus.LICENSE_EXPIRED:
            return f"ライセンスの有効期限が切れています（{self.expiry_date}）。更新してください。"
        if self.status == LicenseStatus.LICENSE_INVALID:
            return "ライセンスファイルが不正です。再アクティベーションが必要です。"
        if self.status == LicenseStatus.MACHINE_MISMATCH:
            return "このライセンスは別の端末用です。サポートにお問い合わせください。"
        return ""


# ─────────────────────────────────────────────
#  ライセンスマネージャー本体
# ─────────────────────────────────────────────

class LicenseManager:

    def __init__(self):
        _APPDATA_DIR.mkdir(parents=True, exist_ok=True)

    # ── マシンID ───────────────────────────────

    def _get_machine_id(self) -> str:
        """
        Windows MachineGuid を取得してSHA256ハッシュ(32文字)を返す。
        取得失敗時はホスト名でフォールバック。
        """
        try:
            key = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Cryptography"
            )
            guid, _ = winreg.QueryValueEx(key, "MachineGuid")
            winreg.CloseKey(key)
            return hashlib.sha256(guid.encode("utf-8")).hexdigest()[:32]
        except Exception as e:
            logger.warning(f"MachineGuid取得失敗、ホスト名で代替: {e}")
            import socket
            return hashlib.sha256(socket.gethostname().encode("utf-8")).hexdigest()[:32]

    # ── 起動時チェック ──────────────────────────

    def check(self) -> LicenseInfo:
        """起動時に呼び出す。LicenseInfo を返す。"""
        info = LicenseInfo()

        # 1. ライセンスファイルが存在する場合
        if _LICENSE_FILE.exists():
            ok, data = self._validate_license_file()
            if not ok:
                info.status = LicenseStatus.LICENSE_INVALID
                return info

            # マシンID検証
            stored_mid = data.get("machine_id")
            if stored_mid is None:
                # machine_id未記録（旧形式）→ 再アクティベーションを要求
                info.status = LicenseStatus.LICENSE_INVALID
                return info
            if stored_mid != self._get_machine_id():
                info.status = LicenseStatus.MACHINE_MISMATCH
                return info

            expiry = date.fromisoformat(data["expiry"])
            ctype  = data.get("type", "standard")
            cinfo  = CUSTOMER_TYPES.get(ctype, {})

            info.status         = LicenseStatus.LICENSED if expiry >= date.today() else LicenseStatus.LICENSE_EXPIRED
            info.customer_type  = ctype
            info.customer_label = cinfo.get("label", ctype)
            info.farm_limit     = 999  # 端末課金・農場数無制限
            info.expiry_date    = expiry
            info.license_key    = data.get("key", "")
            info.unique_id      = data.get("id", "")
            return info

        # 2. 試用期間チェック
        first_launch = self._get_first_launch_date()
        if first_launch is None:
            first_launch = date.today()
            self._save_first_launch_date(first_launch)

        elapsed  = (date.today() - first_launch).days
        remaining = max(0, TRIAL_DAYS - elapsed)

        info.trial_remaining_days = remaining
        info.status = LicenseStatus.TRIAL_ACTIVE if remaining > 0 else LicenseStatus.TRIAL_EXPIRED
        return info

    # ── アクティベーション ─────────────────────

    def activate(self, key: str) -> Tuple[bool, str]:
        """
        ライセンスキーを検証し、有効であれば保存する。
        Returns: (成功したか, メッセージ)
        """
        ok, data, msg = self._parse_key(key.strip())
        if not ok:
            return False, msg

        expiry = date.fromisoformat(data["expiry"])
        if expiry < date.today():
            return False, f"このライセンスキーは有効期限切れです（{expiry}）"

        data["key"] = key.strip().upper()
        data["machine_id"] = self._get_machine_id()
        self._save_license(data)
        ctype = data.get("type", "")
        label = CUSTOMER_TYPES.get(ctype, {}).get("label", ctype)
        return True, f"ライセンスを有効化しました\n顧客種別: {label}\n有効期限: {expiry}"

    def deactivate(self):
        """ライセンスを削除（試用期間には戻らない）"""
        try:
            _LICENSE_FILE.unlink(missing_ok=True)
        except Exception as e:
            logger.warning(f"ライセンス削除エラー: {e}")

    # ── ライセンスキーのパース・検証 ──────────

    def _parse_key(self, key: str) -> Tuple[bool, dict, str]:
        """
        キー形式: FALCON-{BASE32_PAYLOAD}-{HMAC6}
        payload(JSON): {"type":"vet_pro","farm_limit":20,"expiry":"2027-03-31","id":"VET001"}
        """
        try:
            parts = key.upper().split("-")
            if len(parts) < 3 or parts[0] != "FALCON":
                return False, {}, "キーの形式が正しくありません（FALCON-XXXX-XXXXの形式）"

            # 末尾がチェックサム（6文字）、中間がペイロード
            checksum    = parts[-1]
            payload_b64 = "-".join(parts[1:-1])

            # Base32デコード（パディング補完）
            pad = (8 - len(payload_b64) % 8) % 8
            payload_bytes = base64.b32decode(payload_b64 + "=" * pad)

            # HMAC検証
            expected = _hmac.new(_SECRET, payload_bytes, hashlib.sha256).hexdigest()[:6].upper()
            if not _hmac.compare_digest(checksum, expected):
                return False, {}, "ライセンスキーが無効です（認証失敗）"

            data = json.loads(payload_bytes.decode("utf-8"))
            return True, data, "OK"

        except Exception as e:
            logger.debug(f"キーパースエラー: {e}")
            return False, {}, "ライセンスキーの解析に失敗しました"

    # ── ライセンスファイルの読み書き ──────────

    def _save_license(self, data: dict):
        raw = json.dumps(data, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        sig = _hmac.new(_SECRET, raw, hashlib.sha256).digest()
        _LICENSE_FILE.write_bytes(raw + sig)

    def _validate_license_file(self) -> Tuple[bool, dict]:
        try:
            raw = _LICENSE_FILE.read_bytes()
            body, sig = raw[:-32], raw[-32:]
            expected = _hmac.new(_SECRET, body, hashlib.sha256).digest()
            if not _hmac.compare_digest(sig, expected):
                logger.warning("ライセンスファイルの署名検証失敗（改ざん検知）")
                return False, {}
            return True, json.loads(body.decode("utf-8"))
        except Exception as e:
            logger.warning(f"ライセンスファイル読み込みエラー: {e}")
            return False, {}

    # ── 試用期間の記録（2箇所） ────────────────

    def _get_first_launch_date(self) -> Optional[date]:
        file_date = self._read_trial_file()
        reg_date  = self._read_registry_date()
        dates = [d for d in [file_date, reg_date] if d is not None]
        if not dates:
            return None
        return min(dates)  # より古い日付を採用（改ざんで新しくされても最古を使う）

    def _save_first_launch_date(self, d: date):
        self._write_trial_file(d)
        self._write_registry_date(d)

    def _read_trial_file(self) -> Optional[date]:
        try:
            if not _TRIAL_FILE.exists():
                return None
            raw = _TRIAL_FILE.read_bytes()
            body, sig = raw[:-32], raw[-32:]
            expected = _hmac.new(_SECRET, body, hashlib.sha256).digest()
            if not _hmac.compare_digest(sig, expected):
                return None
            return date.fromisoformat(body.decode("utf-8"))
        except Exception:
            return None

    def _write_trial_file(self, d: date):
        try:
            body = d.isoformat().encode("utf-8")
            sig  = _hmac.new(_SECRET, body, hashlib.sha256).digest()
            _TRIAL_FILE.write_bytes(body + sig)
        except Exception as e:
            logger.warning(f"試用期間ファイル書き込みエラー: {e}")

    def _read_registry_date(self) -> Optional[date]:
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REG_PATH)
            val, _ = winreg.QueryValueEx(key, _REG_VALUE)
            winreg.CloseKey(key)
            return date.fromisoformat(val)
        except Exception:
            return None

    def _write_registry_date(self, d: date):
        try:
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, _REG_PATH)
            winreg.SetValueEx(key, _REG_VALUE, 0, winreg.REG_SZ, d.isoformat())
            winreg.CloseKey(key)
        except Exception as e:
            logger.warning(f"レジストリ書き込みエラー: {e}")


# ─────────────────────────────────────────────
#  シングルトン
# ─────────────────────────────────────────────

_manager: Optional[LicenseManager] = None

def get_license_manager() -> LicenseManager:
    global _manager
    if _manager is None:
        _manager = LicenseManager()
    return _manager
