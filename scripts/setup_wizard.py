#!/usr/bin/env python3
"""Loopback-only setup wizard for mail-invoice-register."""
from __future__ import annotations
import argparse
import json
import secrets
import sys
import threading
import webbrowser
from dataclasses import asdict
from datetime import date, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))
from download_and_register import DEFAULT_KEYWORDS, run_job, run_local_folder_job, test_connection, validate_config  # noqa: E402
from buyer_organizer import organize  # noqa: E402

HOST = "127.0.0.1"
TOKEN = secrets.token_urlsafe(24)
HTML = (ROOT / "assets" / "wizard.html").read_text(encoding="utf-8")


def safe_error(exc: Exception) -> str:
    text = str(exc)
    return text.replace("IMAP_PASSWORD", "应用专用密码").replace("password", "密码")[:500]


def parse_payload(handler: BaseHTTPRequestHandler) -> dict:
    if handler.headers.get("X-Setup-Token") != TOKEN:
        raise PermissionError("页面令牌无效，请刷新本地向导")
    size = int(handler.headers.get("Content-Length", "0"))
    if size < 2 or size > 100_000:
        raise ValueError("请求大小异常")
    return json.loads(handler.rfile.read(size).decode("utf-8"))


def config_from(data: dict) -> dict[str, str]:
    return validate_config({
        "IMAP_HOST": data.get("host", ""),
        "IMAP_PORT": data.get("port", "993"),
        "IMAP_USERNAME": data.get("username", ""),
        "IMAP_PASSWORD": data.get("password", ""),
        "IMAP_MAILBOX": data.get("mailbox", "INBOX"),
        "IMAP_USE_SSL": "true" if data.get("ssl", True) else "false",
    })


class Handler(BaseHTTPRequestHandler):
    server_version = "InvoiceSetup/2"
    def log_message(self, *_args):
        return

    def _allowed_host(self) -> bool:
        host = self.headers.get("Host", "").split(":", 1)[0].strip("[]")
        return host in {"127.0.0.1", "localhost", "::1"}

    def send_bytes(self, code: int, body: bytes, content_type: str):
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Content-Security-Policy", "default-src 'self'; style-src 'self' 'unsafe-inline'; script-src 'self' 'unsafe-inline'; connect-src 'self'; frame-ancestors 'none'")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if not self._allowed_host():
            self.send_error(403)
            return
        path = urlparse(self.path).path
        if path == "/":
            today = date.today()
            page = HTML.replace("__TOKEN__", TOKEN).replace("__SINCE__", str(today - timedelta(days=30))).replace("__UNTIL__", str(today))
            self.send_bytes(200, page.encode("utf-8"), "text/html; charset=utf-8")
        elif path == "/health":
            self.send_bytes(200, b'{"ok":true}', "application/json")
        else:
            self.send_error(404)

    def do_POST(self):
        if not self._allowed_host():
            self.send_error(403)
            return
        try:
            data = parse_payload(self)
            action = urlparse(self.path).path
            source_mode = data.get("source_mode", "email")
            if action == "/api/test":
                if source_mode != "email":
                    raise ValueError("本地文件夹模式不需要测试邮箱连接")
                result = test_connection(config_from(data))
            elif action in {"/api/preview", "/api/run"}:
                output = Path(data.get("output_dir", "")).expanduser()
                if not str(output).strip():
                    raise ValueError("请选择输出目录")
                if source_mode == "local":
                    input_dir = Path(data.get("input_dir", "")).expanduser()
                    if not input_dir.is_dir():
                        raise ValueError("请选择有效的发票输入文件夹")
                    candidates = [p for p in input_dir.iterdir() if p.is_file() and (p.suffix.lower() == ".pdf" if data.get("pdf_only") else p.suffix.lower() in {".pdf", ".ofd", ".xml", ".jpg", ".jpeg", ".png", ".tif", ".tiff"})]
                    if action.endswith("preview"):
                        result = {"ok": True, "mode": "preview", "summary": {"candidates": len(candidates), "source_mode": "local_folder"}, "items": [{"filename": p.name, "type": p.suffix.lower()} for p in sorted(candidates)[:200]], "truncated": len(candidates) > 200}
                    else:
                        records, summary = run_local_folder_job(input_dir, output, bool(data.get("pdf_only")))
                        result = {"ok": True, "mode": "run", "summary": summary}
                else:
                    config = config_from(data)
                    since, until = date.fromisoformat(data["since"]), date.fromisoformat(data["until"])
                    keywords = tuple(x.strip() for x in data.get("keywords", "").split(",") if x.strip()) or DEFAULT_KEYWORDS
                    records, summary = run_job(config, since, until, output, keywords, bool(data.get("include_all_pdfs")), bool(data.get("pdf_only")), action == "/api/preview")
                    result = {"ok": True, "mode": "preview" if action.endswith("preview") else "run", "summary": summary}
                    if action.endswith("preview"):
                        result["items"] = [{"date": r.mail_date, "subject": r.subject, "sender": r.sender, "filename": r.filename, "type": r.suffix} for r in records[:200]]
                        result["truncated"] = len(records) > 200
                if action.endswith("run") and (data.get("organize_files") or data.get("buyer_sheets")):
                    workbook = Path(summary["output_file"])
                    rules_text = str(data.get("rules_path", "")).strip()
                    rules_path = Path(rules_text).expanduser() if rules_text else None
                    if rules_path and not rules_path.is_file():
                        raise ValueError("购方规则文件不存在")
                    buyer_result = organize(workbook, output / "attachments", output, rules_path, bool(data.get("organize_files")), bool(data.get("buyer_sheets")))
                    summary["buyer_enhancement"] = buyer_result
                    (output / "run-summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
                    result["summary"] = summary
            elif action == "/api/shutdown":
                result = {"ok": True, "message": "本地向导已关闭。浏览器中的密码不会保存。"}
                threading.Thread(target=self.server.shutdown, daemon=True).start()
            else:
                raise ValueError("未知操作")
            self.send_bytes(200, json.dumps(result, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")
        except PermissionError as exc:
            self.send_bytes(403, json.dumps({"ok": False, "error": safe_error(exc)}, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")
        except Exception as exc:
            self.send_bytes(400, json.dumps({"ok": False, "error": safe_error(exc)}, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="邮箱发票自动登记：本地配置向导")
    parser.add_argument("--port", type=int, default=0, help="本地端口；默认自动选择")
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()
    server = ThreadingHTTPServer((HOST, args.port), Handler)
    url = f"http://{HOST}:{server.server_port}/"
    print(f"本地向导已启动：{url}")
    print("仅监听本机；关闭页面后请按 Ctrl+C，或点击页面里的‘安全退出’。")
    if not args.no_browser:
        threading.Timer(0.4, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
