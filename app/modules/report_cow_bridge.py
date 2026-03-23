"""
FALCON2 - レポート（HTML）から個体カードを開くためのブリッジ

ブラウザで開いたレポート内の Plotly 散布図でポイントをクリックすると、
このモジュールが待ち受けているローカルHTTPサーバに cow_id が送られ、
メインウィンドウがキューをポーリングして個体カードを開く。
"""

import logging
import queue
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
from typing import Optional
from urllib.parse import urlparse, parse_qs

logger = logging.getLogger(__name__)

# レポートの JS からアクセスするポート（固定）
DEFAULT_PORT = 51985

# キュー：ブラウザから届いた cow_id を入れる
_cow_open_queue: queue.Queue = queue.Queue()


def get_pending_cow_ids():
    """キューに溜まった cow_id をすべて取り出して返す（メインスレッドから呼ぶ）"""
    ids = []
    try:
        while True:
            ids.append(_cow_open_queue.get_nowait())
    except queue.Empty:
        pass
    return ids


def _run_server(port: int):
    """HTTPサーバを起動（スレッドで実行）"""
    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            parsed = urlparse(self.path)
            if parsed.path == "/open_cow":
                qs = parse_qs(parsed.query)
                cow_ids = qs.get("cow_id", [])
                for cid in cow_ids:
                    if cid and str(cid).strip():
                        _cow_open_queue.put(str(cid).strip())
                self.send_response(204)
                self.end_headers()
                return
            self.send_response(404)
            self.end_headers()

        def log_message(self, format, *args):
            logger.debug("report_cow_bridge: %s", args[0] if args else "")

    try:
        server = HTTPServer(("127.0.0.1", port), Handler)
        server.serve_forever()
    except OSError as e:
        logger.warning("report_cow_bridge: サーバ起動失敗 (port=%s): %s", port, e)
    except Exception as e:
        logger.warning("report_cow_bridge: %s", e)


_server_thread: Optional[threading.Thread] = None


def start_server(port: int = DEFAULT_PORT) -> bool:
    """ブリッジ用HTTPサーバをバックグラウンドスレッドで起動する。既に起動済みなら何もしない。"""
    global _server_thread
    if _server_thread is not None and _server_thread.is_alive():
        return True
    try:
        _server_thread = threading.Thread(target=_run_server, args=(port,), daemon=True)
        _server_thread.start()
        logger.info("report_cow_bridge: 待ち受け開始 port=%s", port)
        return True
    except Exception as e:
        logger.warning("report_cow_bridge: 起動失敗: %s", e)
        return False
