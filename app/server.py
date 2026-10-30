#!/usr/bin/env python3
"""座位轮转助手 · 本地离线服务端（纯 Python 标准库）。

启动：python3 app/server.py  → 自动打开 http://127.0.0.1:<port>
只监听 127.0.0.1，不联网。
"""

from __future__ import annotations

import argparse
import base64
import datetime as _dt
import os
import json
import logging
import mimetypes
import re
import shutil
import subprocess
import sys
import threading
import time
import traceback
import uuid
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse

# 兼容两种运行方式：
# 1) 源码运行：app/server.py  → 程序在 app/，数据在上一级
# 2) PyInstaller 免安装包：exe 所在目录为 BASE_DIR，网页资源打包在 _MEIPASS/app/web
FROZEN = bool(getattr(sys, "frozen", False))
if FROZEN:
    APP_DIR = Path(sys.executable).resolve().parent
    BASE_DIR = APP_DIR
    WEB_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR)) / "app" / "web"
else:
    APP_DIR = Path(__file__).resolve().parent
    BASE_DIR = APP_DIR.parent
    WEB_DIR = APP_DIR / "web"

# 测试时可用环境变量把数据/导出目录指到临时目录，避免污染用户目录
DATA_DIR = Path(os.environ.get("SEAT_DATA_DIR") or (BASE_DIR / "数据"))
EXPORT_DIR = Path(os.environ.get("SEAT_EXPORT_DIR") or (BASE_DIR / "导出结果"))
SESSION_DIR = DATA_DIR / "sessions"
LOG_DIR = DATA_DIR / "logs"
SAMPLE_DIR = BASE_DIR / "示例座位表"
MAX_BODY = 30 * 1024 * 1024

for directory in (DATA_DIR, SESSION_DIR, EXPORT_DIR, LOG_DIR, WEB_DIR):
    directory.mkdir(parents=True, exist_ok=True)

sys.path.insert(0, str(APP_DIR))

from seatlib import detect as detect_mod  # noqa: E402
from seatlib import naming, period, rotation, xlsx_export, xlsx_io  # noqa: E402

LOGGER = logging.getLogger("seat-rotator")

MIME_OVERRIDES = {
    ".js": "text/javascript; charset=utf-8",
    ".mjs": "text/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".html": "text/html; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".svg": "image/svg+xml",
    ".png": "image/png",
    ".ico": "image/x-icon",
}
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


# ---------------------------------------------------------------- 会话

class SessionStore:
    """导入的工作簿：内存缓存 + 磁盘副本（进程重启后仍可下载/导出）。"""

    def __init__(self):
        self._lock = threading.Lock()
        self._cache = {}

    def _session_dir(self, workbook_id):
        return SESSION_DIR / workbook_id

    def create(self, data: bytes, filename: str) -> dict:
        workbook_id = "wb_" + uuid.uuid4().hex[:10]
        directory = self._session_dir(workbook_id)
        directory.mkdir(parents=True, exist_ok=True)
        (directory / "original.xlsx").write_bytes(data)
        (directory / "meta.json").write_text(
            json.dumps({"filename": filename, "created": time.time()}, ensure_ascii=False),
            encoding="utf-8")
        sheets = xlsx_io.list_sheets(data)
        with self._lock:
            self._cache[workbook_id] = {
                "data": data, "filename": filename, "sheets": sheets,
                "sheet_cache": {}, "detect_cache": {},
            }
        return {"workbook_id": workbook_id, "filename": filename,
                "sheets": sheets, "warnings": []}

    def _load(self, workbook_id: str):
        if not re.fullmatch(r"wb_[0-9a-f]{6,32}", workbook_id or ""):
            raise ValueError("工作簿标识不合法，请重新导入")
        with self._lock:
            if workbook_id in self._cache:
                return self._cache[workbook_id]
        directory = self._session_dir(workbook_id)
        original = directory / "original.xlsx"
        if not original.is_file():
            raise KeyError("工作簿不存在或已失效，请重新导入")
        data = original.read_bytes()
        meta = {}
        meta_file = directory / "meta.json"
        if meta_file.is_file():
            try:
                meta = json.loads(meta_file.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                meta = {}
        session = {"data": data, "filename": meta.get("filename") or "座位表.xlsx",
                   "sheets": xlsx_io.list_sheets(data), "sheet_cache": {}, "detect_cache": {}}
        with self._lock:
            self._cache[workbook_id] = session
        return session

    def sheets(self, workbook_id):
        return self._load(workbook_id)["sheets"]

    def filename(self, workbook_id):
        return self._load(workbook_id)["filename"]

    def sheet(self, workbook_id: str, index: int):
        session = self._load(workbook_id)
        with self._lock:
            if index in session["sheet_cache"]:
                return session["sheet_cache"][index]
        sheet = xlsx_io.read_sheet(session["data"], index)
        with self._lock:
            session["sheet_cache"][index] = sheet
        return sheet

    def detect(self, workbook_id: str, index: int):
        session = self._load(workbook_id)
        with self._lock:
            if index in session["detect_cache"]:
                return session["detect_cache"][index]
        result = detect_mod.detect_sheet(self.sheet(workbook_id, index))
        with self._lock:
            session["detect_cache"][index] = result
        return result

    def original(self, workbook_id: str) -> bytes:
        return self._load(workbook_id)["data"]


STORE = SessionStore()


# ---------------------------------------------------------------- 工具

def _col_letter(col: int) -> str:
    letters = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _to_a1(value) -> str | None:
    """把 [r, c] / "r,c" / "C3"（可带 $）统一成 A1 引用。"""
    if value is None:
        return None
    if isinstance(value, (list, tuple)) and len(value) == 2:
        return f"{_col_letter(int(value[1]))}{int(value[0])}"
    text = str(value).strip().replace("$", "").upper()
    if re.fullmatch(r"[A-Z]{1,3}[0-9]{1,7}", text):
        return text
    if re.fullmatch(r"[0-9]+,[0-9]+", text):
        r, c = text.split(",")
        return f"{_col_letter(int(c))}{int(r)}"
    raise ValueError(f"无法识别的单元格引用：{value!r}")


def _as_bool(value) -> bool:
    """宽容地把 JSON 值转成布尔；非法值按 False 处理。"""
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    if isinstance(value, str):
        return value.strip().lower() in ("1", "true", "yes", "on", "是", "开")
    return False


def _clamp_int(value, default, low, high):
    try:
        number = int(value)
    except (TypeError, ValueError):
        number = default
    return max(low, min(high, number))


def _filename_date_plans(analysis, rounds, cycle_weeks):
    """没有「座位周期」表头时，用文件名里的日期推断逐轮周期。"""
    dates = analysis.get("dates") or []
    if not dates:
        return [{"round": i + 1, "start": None, "end": None, "text": ""}
                for i in range(rounds)]
    base_start = dates[0]["date"]
    base_end = dates[1]["date"] if len(dates) >= 2 else None
    cycle_days = cycle_weeks * 7
    plans = []
    for index in range(rounds):
        shift = cycle_days * (index + 1)  # 原文件名代表当前轮，新文件 = 下一轮
        start = base_start + _dt.timedelta(days=shift)
        end = (base_end + _dt.timedelta(days=shift)) if base_end else (
            start + _dt.timedelta(days=cycle_days - 1))
        plans.append({"round": index + 1, "start": start, "end": end, "text": ""})
    return plans


def _open_in_file_manager(path, reveal=False) -> bool:
    """跨平台打开文件夹/在文件管理器中定位文件。失败返回 False。"""
    target = Path(path)
    try:
        if sys.platform == "darwin":
            cmd = ["open", "-R", str(target)] if reveal else ["open", str(target)]
            subprocess.Popen(cmd)
        elif os.name == "nt":
            if reveal:
                subprocess.Popen(["explorer", f"/select,{target}"])
            else:
                os.startfile(str(target))  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", str(target.parent if reveal else target)])
        return True
    except (OSError, AttributeError) as exc:
        LOGGER.warning("无法打开文件管理器：%s", exc)
        return False


def _safe_filename(name: str, fallback: str = "座位表") -> str:
    name = re.sub(r"[\\/:*?\"<>|\r\n\t]+", "_", str(name or "")).strip().strip(".")
    return (name or fallback)[:120]


def _within(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def _json_bytes(payload) -> bytes:
    return json.dumps(payload, ensure_ascii=False).encode("utf-8")


# ---------------------------------------------------------------- HTTP

class Handler(BaseHTTPRequestHandler):
    server_version = "SeatRotator/1.0"

    # 静默默认访问日志，改写入文件
    def log_message(self, fmt, *args):  # noqa: A003
        LOGGER.info("%s - %s", self.address_string(), fmt % args)

    # ---- 基础响应
    def _send(self, status, body: bytes, content_type="application/json; charset=utf-8",
              extra_headers=None):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for key, value in (extra_headers or {}).items():
            self.send_header(key, value)
        self.end_headers()
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _json(self, payload, status=200):
        self._send(status, _json_bytes(payload))

    def _error(self, status, message):
        LOGGER.warning("请求失败 %s: %s", status, message)
        self._json({"error": str(message)}, status)

    def _read_json(self):
        try:
            length = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            length = 0
        if length <= 0:
            return {}
        if length > MAX_BODY:
            raise ValueError("上传内容过大（超过 30MB），请检查文件")
        raw = self.rfile.read(length)
        try:
            return json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ValueError("请求体不是合法 JSON") from exc

    # ---- 路由
    def do_GET(self):  # noqa: N802
        parsed = urlparse(self.path)
        path = unquote(parsed.path)
        try:
            if path in ("/favicon.ico", "/favicon.png"):
                return self._send(204, b"", "image/x-icon")
            if path == "/api/health":
                return self._json({"ok": True, "app": "seat-rotator", "version": 1})
            if path == "/api/sheet":
                query = parse_qs(parsed.query)
                workbook_id = (query.get("workbook_id") or [""])[0]
                index = int((query.get("index") or ["0"])[0])
                return self._handle_sheet(workbook_id, index)
            if path == "/api/download":
                query = parse_qs(parsed.query)
                return self._handle_download((query.get("path") or [""])[0])
            return self._serve_static(path)
        except ValueError as exc:
            return self._error(400, exc)
        except KeyError as exc:
            return self._error(404, exc.args[0] if exc.args else "文件不存在")
        except FileNotFoundError:
            return self._error(404, "文件不存在")
        except Exception:  # noqa: BLE001
            LOGGER.error("GET %s 异常\n%s", self.path, traceback.format_exc())
            return self._error(500, "服务器内部错误，请查看终端日志")

    def do_POST(self):  # noqa: N802
        path = unquote(urlparse(self.path).path)
        routes = {
            "/api/import": self._handle_import,
            "/api/detect": self._handle_detect,
            "/api/regroup": self._handle_regroup,
            "/api/rotate": self._handle_rotate,
            "/api/period": self._handle_period,
            "/api/export": self._handle_export,
            "/api/reveal": self._handle_reveal,
        }
        handler = routes.get(path)
        try:
            if path == "/api/sheet":  # 兼容 POST 方式取 sheet
                body = self._read_json()
                return self._handle_sheet(body.get("workbook_id", ""), int(body.get("index", 0)))
            if handler is None:
                return self._error(404, "接口不存在")
            body = self._read_json()
            return handler(body)
        except ValueError as exc:
            return self._error(400, exc)
        except KeyError as exc:
            return self._error(404, exc.args[0] if exc.args else "资源不存在")
        except FileNotFoundError:
            return self._error(404, "文件不存在")
        except Exception:  # noqa: BLE001
            LOGGER.error("POST %s 异常\n%s", self.path, traceback.format_exc())
            return self._error(500, "服务器内部错误，请查看终端日志")

    # ---- 静态资源
    def _serve_static(self, path: str):
        if path in ("/", ""):
            path = "/index.html"
        relative = path.lstrip("/")
        candidate = (WEB_DIR / relative).resolve()
        if not _within(candidate, WEB_DIR) or not candidate.is_file():
            return self._error(404, "页面不存在")
        mime = MIME_OVERRIDES.get(candidate.suffix.lower()) or \
            mimetypes.guess_type(str(candidate))[0] or "application/octet-stream"
        self._send(200, candidate.read_bytes(), mime)

    # ---- API 实现
    def _handle_import(self, body):
        filename = str(body.get("filename") or "座位表.xlsx")
        data = body.get("data") or ""
        if isinstance(data, str) and data.startswith("data:"):
            data = data.split(",", 1)[-1]
        try:
            binary = base64.b64decode(data, validate=False)
        except Exception as exc:  # noqa: BLE001
            raise ValueError("文件内容解码失败，请重新选择文件") from exc
        if not binary:
            raise ValueError("没有收到文件内容")
        if len(binary) > MAX_BODY:
            raise ValueError("文件过大（超过 30MB）")
        try:
            result = STORE.create(binary, filename)
        except ValueError as exc:
            raise ValueError("不是有效的 .xlsx 文件，请在 Excel/WPS 里另存为 .xlsx 后重试") from exc
        LOGGER.info("导入工作簿 %s (%s, %d bytes)", result["workbook_id"], filename, len(binary))
        return self._json(result)

    def _handle_sheet(self, workbook_id, index):
        sheet = STORE.sheet(workbook_id, index)
        return self._json({"workbook_id": workbook_id, "sheet": sheet})

    def _handle_detect(self, body):
        workbook_id = str(body.get("workbook_id") or "")
        index = int(body.get("index") or 0)
        result = STORE.detect(workbook_id, index)
        result = dict(result)
        result["workbook_id"] = workbook_id
        result["index"] = index
        return self._json(result)

    def _handle_regroup(self, body):
        result = detect_mod.regroup(body.get("seats") or [], body.get("group_labels") or [])
        return self._json(result)

    def _handle_rotate(self, body):
        seats = body.get("seats") or []
        groups = body.get("groups") or []
        podium = body.get("podium") or {}
        rules = body.get("rules") or {}
        result = rotation.rotate(seats, groups, podium, rules)
        return self._json(result)

    def _handle_period(self, body):
        workbook_id = str(body.get("workbook_id") or "")
        index = int(body.get("index") or 0)
        rounds = _clamp_int(body.get("rounds"), 1, 1, 12)
        cycle_weeks = _clamp_int(body.get("cycle_weeks"), 2, 1, 8)
        start_date = body.get("start_date")
        sheet = STORE.sheet(workbook_id, index)
        found = period.find_period(sheet)
        stem = naming.strip_extension(STORE.filename(workbook_id))
        analysis = naming.analyze_filename(stem)

        if found:
            period_plans = period.plan(found, rounds=rounds, start_date=start_date,
                                       cycle_weeks=cycle_weeks)
            plans = [{
                "round": item["round"],
                "start": _dt.date.fromisoformat(item["start"]),
                "end": _dt.date.fromisoformat(item["end"]),
                "text": item["text"],
            } for item in period_plans]
            period_public = period.public_period(found)
        else:
            period_public = None
            plans = _filename_date_plans(analysis, rounds, cycle_weeks)

        filenames = naming.suggest_filenames(stem, plans, rounds=rounds)
        rounds_payload = []
        for item, filename in zip(plans, filenames):
            rounds_payload.append({
                "round": item["round"],
                "start": item["start"].isoformat() if item["start"] else None,
                "end": item["end"].isoformat() if item["end"] else None,
                "text": item.get("text") or "",
                "filename": filename,
            })

        return self._json({
            "period": period_public,
            "cycle_weeks": cycle_weeks,
            "naming": naming.public_analysis(analysis),
            "rounds": rounds_payload,
        })

    def _handle_export(self, body):
        workbook_id = str(body.get("workbook_id") or "")
        index = int(body.get("index") or 0)
        filename_base = _safe_filename(body.get("filename_base") or "座位表")
        move_marks = _as_bool(body.get("move_marks", False))
        rounds = body.get("rounds") or []
        if not isinstance(rounds, list) or not rounds:
            raise ValueError("没有可导出的轮转结果")
        data = STORE.original(workbook_id)

        files = []
        for position, round_item in enumerate(rounds, start=1):
            moves_raw = (round_item or {}).get("moves") or []
            moves = []
            for move in moves_raw:
                to_ref = _to_a1((move or {}).get("to"))
                from_ref = _to_a1((move or {}).get("from"))
                if not to_ref or not from_ref or to_ref == from_ref:
                    continue
                moves.append({"to": to_ref, "from": from_ref})
            custom_name = naming.strip_extension((round_item or {}).get("filename") or "")
            if custom_name:
                out_path = EXPORT_DIR / f"{_safe_filename(custom_name)}.xlsx"
            else:
                suffix = "" if len(rounds) == 1 else (
                    f"_第{int((round_item or {}).get('round') or position)}轮")
                out_path = EXPORT_DIR / f"{filename_base}{suffix}.xlsx"
            counter = 2
            base_stem, suffix_ext = out_path.stem, out_path.suffix
            while out_path.exists():
                out_path = EXPORT_DIR / f"{base_stem}_{counter}{suffix_ext}"
                counter += 1
            text_updates = []
            for item in (round_item or {}).get("header_updates") or []:
                try:
                    to_ref = _to_a1((item or {}).get("to"))
                except ValueError:
                    continue
                text_value = (item or {}).get("text")
                if to_ref and text_value is not None:
                    text_updates.append({"to": to_ref, "text": str(text_value)})
            export_kwargs = {}
            if move_marks:
                export_kwargs["move_marks"] = True
            if text_updates:
                export_kwargs["text_updates"] = text_updates
            report = xlsx_export.export_workbook(data, index, moves, str(out_path), **export_kwargs)
            try:
                relative = str(out_path.relative_to(BASE_DIR))
            except ValueError:
                relative = str(out_path)  # 测试环境导出到临时目录
            files.append({
                "name": out_path.name,
                "path": relative,
                "round": (round_item or {}).get("round", position),
                "applied": report.get("applied", len(moves)),
                "marks_moved": report.get("marks_moved", 0),
                "text_updates_applied": report.get("text_updates_applied", 0),
                "skipped": report.get("skipped", []),
                "download_url": "/api/download?path=" + quote(out_path.name),
            })
            LOGGER.info("导出 %s（应用 %s 处，标记跟随 %s 处）",
                        out_path.name, report.get("applied"), report.get("marks_moved", 0))

        if body.get("open_folder"):
            _open_in_file_manager(EXPORT_DIR)

        try:
            folder = str(EXPORT_DIR.relative_to(BASE_DIR))
        except ValueError:
            folder = str(EXPORT_DIR)
        return self._json({"files": files, "folder": folder, "move_marks": move_marks})

    def _handle_download(self, raw_path):
        if not raw_path:
            raise ValueError("缺少文件路径")
        cleaned = raw_path.strip().lstrip("/")
        candidates = [EXPORT_DIR / cleaned]
        if cleaned.startswith("导出结果/"):
            candidates.append(BASE_DIR / cleaned)
        # 允许下载示例座位表（便于核对）
        candidates.append(SAMPLE_DIR / cleaned)
        for candidate in candidates:
            resolved = candidate.resolve()
            if not resolved.is_file():
                continue
            allowed = (_within(resolved, EXPORT_DIR) or _within(resolved, SAMPLE_DIR))
            if not allowed:
                raise ValueError("不允许下载该路径")
            body = resolved.read_bytes()
            disposition = "attachment; filename*=UTF-8''" + quote(resolved.name)
            self._send(200, body, XLSX_MIME, {"Content-Disposition": disposition})
            return
        raise FileNotFoundError("导出文件不存在，可能已被移动或删除")

    def _handle_reveal(self, body):
        raw_path = str(body.get("path") or "").strip()
        if not raw_path:
            raise ValueError("缺少文件路径")
        candidate = (BASE_DIR / raw_path).resolve()
        if not candidate.exists():
            candidate = (EXPORT_DIR / Path(raw_path).name).resolve()
        if not (_within(candidate, EXPORT_DIR) or _within(candidate, SAMPLE_DIR)):
            raise ValueError("不允许打开该路径")
        if not _open_in_file_manager(candidate, reveal=True):
            raise ValueError("无法自动打开文件所在文件夹，请手动打开「导出结果」目录")
        return self._json({"ok": True})


# ---------------------------------------------------------------- 启动

def _build_server(host: str, port: int):
    class Server(ThreadingHTTPServer):
        daemon_threads = True
        allow_reuse_address = True

    return Server((host, port), Handler)


def _start_logging():
    LOGGER.setLevel(logging.INFO)
    LOGGER.propagate = False
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    file_handler = logging.FileHandler(LOG_DIR / "app.log", encoding="utf-8")
    file_handler.setFormatter(formatter)
    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(logging.Formatter("%(message)s"))
    LOGGER.handlers = [file_handler, stream_handler]


def main(argv=None):
    parser = argparse.ArgumentParser(description="座位轮转助手 · 本地服务")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=0, help="0 表示自动选择 8765~8799")
    parser.add_argument("--no-browser", action="store_true", help="不自动打开浏览器")
    args = parser.parse_args(argv)

    _start_logging()

    candidates = [args.port] if args.port else list(range(8765, 8800))
    server = None
    for port in candidates:
        try:
            server = _build_server(args.host, port)
            break
        except OSError:
            continue
    if server is None:
        print("找不到可用端口（8765~8799 全被占用），请关闭其他程序后重试。")
        return 1

    url = f"http://{args.host}:{server.server_address[1]}/"
    (DATA_DIR / "last_url.txt").write_text(url, encoding="utf-8")
    LOGGER.info("=" * 56)
    LOGGER.info("座位轮转助手已启动：%s", url)
    LOGGER.info("关闭此终端窗口即可停止服务；导出的 Excel 在「导出结果」文件夹。")
    LOGGER.info("=" * 56)

    if not args.no_browser:
        threading.Thread(target=lambda: (time.sleep(0.6), webbrowser.open(url)),
                         daemon=True).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        LOGGER.info("已停止服务")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
