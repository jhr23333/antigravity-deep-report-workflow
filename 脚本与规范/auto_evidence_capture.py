# -*- coding: utf-8 -*-
"""
国金证券研报工作底稿——全自动原始一手红框标注截图引擎 (auto_evidence_capture.py)

核心升级（逐数核验 + 防歧义找页 + Sidecar 元数据落盘 + 严禁凑图）：
1. 防歧义 PDF 精准找页（find_verified_pdf_page）：
   - 强制要求同一页同时命中 must_all 中的全部实体/数字（严禁只命中 1 个词就抓页）；
   - 内置常见歧义短词上下文过滤（如搜“大陆”时自动剔除“中国大陆/大陆地区”干扰，杜绝误把地理名词当成客户“大陆汽车”）。
2. PDF 精准定位与红框标注截图（capture_pdf_with_redbox）：
   - 强制执行 `must_contain` 数字/关键词硬校验：若该 PDF 页缺少任一必需数字/实体，直接返回 False，绝不生成错配截图；
   - 默认要求至少绘制出 1 个真实红框（require_redbox=True），禁止生成无红框白板截图；
   - 每次截屏成功自动在同目录生成 `<图片名>.meta.json` 溯源元数据（记录来源、页码、已验证词表、页面归一化全文），供 `build_workpaper.py` 做自动化“原子事实覆盖率”审计！
3. 网页 Chrome DevTools Protocol (CDP) 精准 DOM 滚动红框截图（capture_web_with_dom_redbox）：
   - 自动打开网页并跟随跳转，获取最终真实文章页 URL（绝不留首页域名）；
   - 强制校验网页全文包含 `must_contain` 全部目标数字，未命中直接拒绝凑图；
   - 自动将命中段落/表格行滚入视口中央并直接绘制 3px 醒目 DOM 红框，并同步输出 `.meta.json` 溯源元数据。
"""

import base64
import datetime
import io
import json
import os
import random
import socket
import struct
import subprocess
import tempfile
import time
import urllib.request
from typing import Dict, List, Optional, Tuple
from urllib.parse import urlparse

import fitz
from PIL import Image, ImageDraw, ImageFont


CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]

# 容易在年报/招股书中产生语义歧义的短词及其常见干扰词（防误把“中国大陆”当“大陆汽车”）
DEFAULT_AMBIGUOUS_CONTEXTS: Dict[str, List[str]] = {
    "大陆": ["中国大陆", "大陆地区", "大陆市场", "大陆厂商", "大陆产能", "大陆基地", "大陆境内"],
    "通用": ["通用服务器", "通用型", "通用设备", "通用照明", "通用计算", "通用准则"],
    "现代": ["现代化", "现代服务", "现代企业", "现代制造"],
    "理想": ["理想状态", "较为理想", "不理想", "理想情况"],
}


def _normalize_text(text: str) -> str:
    """去除所有空白换行字符，统一用于跨行断句的严格子串匹配。"""
    return "".join((text or "").split())


def _write_sidecar_meta(
    out_png: str,
    capture_type: str,
    source_ref: str,
    must_contain: Optional[List[str]],
    matched_box_count: int,
    page_text_norm: str,
    extra: Optional[Dict] = None,
) -> None:
    """
    在每张生成的佐证截图旁同步落盘 `<图片名>.meta.json` 溯源元数据。
    该文件会被 build_workpaper.py 自动读取，用来核对每张图是否真实覆盖了左图/左表的全部原子事实与数字。
    """
    meta_path = f"{out_png}.meta.json"
    payload = {
        "image_file": os.path.basename(out_png),
        "capture_type": capture_type,
        "source_ref": source_ref,
        "verified_tokens": list(must_contain or []),
        "matched_box_count": int(matched_box_count),
        "captured_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        # 保留足够长的页面归一化文本（上限 30,000 字），便于下游底稿审计脚本做二次事实回检
        "page_text_norm": page_text_norm[:30000],
    }
    if extra:
        payload.update(extra)
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def _get_cn_font(size: int = 19):
    """加载 Windows 常见中文字体（微软雅黑/黑体），用于顶部来源标注栏。"""
    for fp in (
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\msyhbd.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
        r"C:\Windows\Fonts\simkai.ttf",
    ):
        if os.path.exists(fp):
            try:
                return ImageFont.truetype(fp, size)
            except Exception:
                pass
    return ImageFont.load_default()


def _add_provenance_banner(img: Image.Image, banner_text: str) -> Image.Image:
    """在截图顶部拼接国金深蓝来源/页码/精确网址栏，确保质控核查一目了然。"""
    if not banner_text:
        return img
    w, h = img.size
    banner_h = 44
    canvas = Image.new("RGB", (w, h + banner_h), (255, 255, 255))
    draw = ImageDraw.Draw(canvas)
    draw.rectangle([0, 0, w, banner_h], fill=(0, 72, 112))
    font = _get_cn_font(17)
    draw.text((14, 11), banner_text[:128], fill=(255, 255, 255), font=font)
    canvas.paste(img, (0, banner_h))
    return canvas


def find_headless_browser() -> Optional[str]:
    """查找本机已安装的 Chrome 或 Edge 浏览器路径。"""
    for exe in CHROME_CANDIDATES:
        if os.path.exists(exe):
            return exe
    return None


# ==================== 一、防歧义 PDF 精准找页 + 红框标注截图 ====================

def find_verified_pdf_page(
    pdf_path: str,
    must_all: List[str],
    exclude_context: Optional[Dict[str, List[str]]] = None,
    start_page: int = 1,
    end_page: Optional[int] = None,
) -> List[int]:
    """
    在 PDF 中查找【同时包含 must_all 全部关键词/实体/数字】的 1-indexed 页码列表。
    内置歧义词干扰剔除机制（如检查“大陆”前先剔除“中国大陆”等地理词），杜绝错页命中。
    """
    if not os.path.exists(pdf_path):
        print(f"[PDF-SEARCH-FAIL] 文件不存在: {pdf_path}")
        return []

    merged_excludes: Dict[str, List[str]] = {k: list(v) for k, v in DEFAULT_AMBIGUOUS_CONTEXTS.items()}
    if exclude_context:
        for k, v in exclude_context.items():
            merged_excludes[k] = list(v)

    for tok in must_all:
        if tok in DEFAULT_AMBIGUOUS_CONTEXTS:
            print(
                f"[PDF-SEARCH-WARN] 检索词「{tok}」为高歧义短词，已启用上下文过滤 {merged_excludes[tok]}；"
                f"建议优先传入全称（如「大陆汽车」）。"
            )

    doc = fitz.open(pdf_path)
    total = len(doc)
    lo = max(1, start_page)
    hi = min(total, end_page or total)
    hits: List[int] = []

    for pno in range(lo, hi + 1):
        norm_txt = _normalize_text(doc[pno - 1].get_text())
        all_ok = True
        for tok in must_all:
            ntok = _normalize_text(tok)
            if not ntok:
                continue
            check_txt = norm_txt
            # 若该词存在已知歧义干扰词，先剔除干扰词再判断真实出现
            if tok in merged_excludes:
                for noise in merged_excludes[tok]:
                    check_txt = check_txt.replace(_normalize_text(noise), "")
            if ntok not in check_txt:
                all_ok = False
                break
        if all_ok:
            hits.append(pno)

    doc.close()
    return hits


def capture_pdf_with_redbox(
    pdf_path: str,
    page_num: int,
    keywords: List[str],
    out_png: str,
    banner_text: str = "",
    must_contain: Optional[List[str]] = None,
    dpi: int = 160,
    fallback_box: Optional[Tuple[float, float, float, float]] = None,
    crop_y_range: Optional[Tuple[float, float]] = None,
    require_redbox: bool = True,
    exclude_context: Optional[Dict[str, List[str]]] = None,
) -> bool:
    """
    从正式 PDF 指定页（1-indexed）截取高清图片，并在关键词所在原文处自动画红框。
    硬校验：
    1. 若传入 must_contain，强制校验页面文本必须包含全部目标数字/短语（含歧义词过滤），缺一即拒截；
    2. require_redbox=True 时，若未匹配到任何红框坐标且未提供 fallback_box，直接拒绝输出无框白板图；
    3. 截屏成功后自动落盘 `<out_png>.meta.json`，供底稿构建脚本审计事实覆盖率。
    """
    if not os.path.exists(pdf_path):
        print(f"[PDF-FAIL] 文件不存在: {pdf_path}")
        return False

    doc = fitz.open(pdf_path)
    if page_num < 1 or page_num > len(doc):
        print(f"[PDF-FAIL] 页码越界: {page_num} / {len(doc)}")
        doc.close()
        return False

    page = doc[page_num - 1]
    page_text = page.get_text()
    norm_page_text = _normalize_text(page_text)

    merged_excludes: Dict[str, List[str]] = {k: list(v) for k, v in DEFAULT_AMBIGUOUS_CONTEXTS.items()}
    if exclude_context:
        for k, v in exclude_context.items():
            merged_excludes[k] = list(v)

    # 逐数核验：去换行空格后确认所有必需数字/关键表述真实存在于该页
    if must_contain:
        missing = []
        for tok in must_contain:
            ntok = _normalize_text(tok)
            check_txt = norm_page_text
            if tok in merged_excludes:
                for noise in merged_excludes[tok]:
                    check_txt = check_txt.replace(_normalize_text(noise), "")
            if ntok not in check_txt:
                missing.append(tok)
        if missing:
            print(f"[PDF-REJECT] {os.path.basename(pdf_path)} P.{page_num} 缺少必需内容 {missing}，拒绝凑图！")
            doc.close()
            return False
    else:
        print(f"[PDF-WARN] {os.path.basename(pdf_path)} P.{page_num} 未传 must_contain，建议始终传入核心数字/实体以防错页。")

    pw, ph = page.rect.width, page.rect.height
    rects: List[fitz.Rect] = []
    for kw in keywords:
        if kw:
            rects.extend(page.search_for(kw))

    shape = page.new_shape()
    drawn_boxes = 0
    if rects:
        rects_sorted = sorted(rects, key=lambda r: (r.y0, r.x0))
        clusters: List[fitz.Rect] = []
        for r in rects_sorted:
            if not clusters or r.y0 - clusters[-1].y1 > 22:
                clusters.append(fitz.Rect(r))
            else:
                clusters[-1] |= r
        for c in clusters:
            padded = fitz.Rect(
                max(18, c.x0 - 8),
                max(18, c.y0 - 5),
                min(pw - 18, c.x1 + 8),
                min(ph - 18, c.y1 + 5),
            )
            shape.draw_rect(padded)
            shape.finish(color=(0.85, 0.18, 0.13), width=2.4)
            drawn_boxes += 1
    elif fallback_box:
        x1, y1, x2, y2 = fallback_box
        padded = fitz.Rect(x1 * pw, y1 * ph, x2 * pw, y2 * ph)
        shape.draw_rect(padded)
        shape.finish(color=(0.85, 0.18, 0.13), width=2.6)
        drawn_boxes = 1
    elif require_redbox:
        print(f"[PDF-REJECT] {os.path.basename(pdf_path)} P.{page_num} 未定位到任何红框坐标（keywords={keywords}），拒绝输出无框图！")
        doc.close()
        return False

    shape.commit()

    clip = fitz.Rect(0, crop_y_range[0] * ph, pw, crop_y_range[1] * ph) if crop_y_range else None
    pix = page.get_pixmap(dpi=dpi, clip=clip)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    doc.close()

    img = _add_provenance_banner(img, banner_text)
    os.makedirs(os.path.dirname(os.path.abspath(out_png)), exist_ok=True)
    img.save(out_png, format="PNG", optimize=True)

    # 自动生成溯源元数据 sidecar，供 build_workpaper.py 审计覆盖率
    _write_sidecar_meta(
        out_png=out_png,
        capture_type="pdf",
        source_ref=f"{os.path.basename(pdf_path)}#P.{page_num}",
        must_contain=must_contain,
        matched_box_count=drawn_boxes,
        page_text_norm=norm_page_text,
        extra={"pdf_path": os.path.abspath(pdf_path), "page_num": page_num, "keywords": keywords},
    )
    return True


# ==================== 二、轻量纯标准库 CDP WebSocket 客户端 ====================

class _MiniCDP:
    """极简纯标准库 Chrome DevTools Protocol 客户端，用于 DOM 文本核验、滚动定位与元素红框标注。"""

    def __init__(self, ws_url: str, timeout: float = 25.0):
        parsed = urlparse(ws_url)
        host = parsed.hostname or "127.0.0.1"
        port = parsed.port or 80
        path = parsed.path or "/"
        if parsed.query:
            path += "?" + parsed.query

        self.sock = socket.create_connection((host, port), timeout=timeout)
        key = base64.b64encode(os.urandom(16)).decode("ascii")
        req = (
            f"GET {path} HTTP/1.1\r\n"
            f"Host: {host}:{port}\r\n"
            "Upgrade: websocket\r\n"
            "Connection: Upgrade\r\n"
            f"Sec-WebSocket-Key: {key}\r\n"
            "Sec-WebSocket-Version: 13\r\n\r\n"
        )
        self.sock.sendall(req.encode("ascii"))
        resp = b""
        while b"\r\n\r\n" not in resp:
            chunk = self.sock.recv(4096)
            if not chunk:
                break
            resp += chunk
        self._buf = resp.split(b"\r\n\r\n", 1)[1] if b"\r\n\r\n" in resp else b""
        self._msg_id = 0

    def _recv_exact(self, n: int) -> bytes:
        while len(self._buf) < n:
            chunk = self.sock.recv(65536)
            if not chunk:
                raise ConnectionError("CDP 连接提前关闭")
            self._buf += chunk
        out, self._buf = self._buf[:n], self._buf[n:]
        return out

    def _send_text(self, text: str) -> None:
        payload = text.encode("utf-8")
        mask = os.urandom(4)
        header = bytearray([0x81])
        ln = len(payload)
        if ln < 126:
            header.append(0x80 | ln)
        elif ln < 65536:
            header.append(0x80 | 126)
            header.extend(struct.pack("!H", ln))
        else:
            header.append(0x80 | 127)
            header.extend(struct.pack("!Q", ln))
        header.extend(mask)
        masked = bytes(b ^ mask[i % 4] for i, b in enumerate(payload))
        self.sock.sendall(bytes(header) + masked)

    def _recv_frame(self) -> Tuple[int, bytes]:
        b1, b2 = self._recv_exact(2)
        opcode = b1 & 0x0F
        ln = b2 & 0x7F
        if ln == 126:
            ln = struct.unpack("!H", self._recv_exact(2))[0]
        elif ln == 127:
            ln = struct.unpack("!Q", self._recv_exact(8))[0]
        if b2 & 0x80:
            mask = self._recv_exact(4)
            raw = self._recv_exact(ln)
            data = bytes(b ^ mask[i % 4] for i, b in enumerate(raw))
        else:
            data = self._recv_exact(ln)
        return opcode, data

    def call(self, method: str, params: Optional[dict] = None) -> dict:
        self._msg_id += 1
        cid = self._msg_id
        self._send_text(json.dumps({"id": cid, "method": method, "params": params or {}}))
        while True:
            op, data = self._recv_frame()
            if op == 8:
                raise ConnectionError("CDP 已关闭")
            if op == 9:
                continue
            if op != 1:
                continue
            msg = json.loads(data.decode("utf-8", errors="ignore"))
            if msg.get("id") == cid:
                return msg.get("result", {})

    def close(self) -> None:
        try:
            self.sock.close()
        except Exception:
            pass


def capture_web_with_dom_redbox(
    url: str,
    must_contain: List[str],
    out_png: str,
    banner_prefix: str = "",
    viewport: Tuple[int, int] = (1280, 960),
    wait_sec: float = 5.0,
) -> Tuple[bool, str, str]:
    """
    通过本机 Chrome DevTools Protocol (CDP) 打开权威公开网页：
    1. 跟随跳转获取最终真实文章 URL；
    2. 校验页面全文是否真实包含 must_contain 中全部目标数字（缺一即拒截，严禁凑图）；
    3. 在 DOM 中精准定位含目标数字的段落/表格行，自动滚入视口中央并直接给 DOM 元素套 3px 醒目红框；
    4. 截屏并叠加含真实文章 URL 的国金来源顶栏，同步落盘 `.meta.json` 溯源元数据。

    返回：(是否成功, 最终真实文章URL, 诊断信息)
    """
    if not must_contain:
        return False, url, "网页截屏必须传入非空 must_contain 核验数字列表，严禁无校验盲截"

    browser = find_headless_browser()
    if not browser:
        return False, url, "未找到 Chrome/Edge"

    port = random.randint(19200, 19899)
    user_dir = tempfile.mkdtemp(prefix="cdp_prof_")
    proc = subprocess.Popen(
        [
            browser,
            "--headless=new",
            "--disable-gpu",
            "--no-first-run",
            "--no-default-browser-check",
            f"--remote-debugging-port={port}",
            f"--user-data-dir={user_dir}",
            f"--window-size={viewport[0]},{viewport[1]}",
            "about:blank",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    cdp = None
    try:
        # 等待 Chrome 调试端口就绪
        ws_url = None
        for _ in range(30):
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{port}/json", timeout=2) as r:
                    tabs = json.loads(r.read().decode("utf-8"))
                    for t in tabs:
                        if t.get("type") == "page" and t.get("webSocketDebuggerUrl"):
                            ws_url = t["webSocketDebuggerUrl"]
                            break
                if ws_url:
                    break
            except Exception:
                time.sleep(0.25)

        if not ws_url:
            return False, url, "CDP 端口启动超时"

        cdp = _MiniCDP(ws_url, timeout=25.0)
        cdp.call("Page.enable")
        cdp.call("Runtime.enable")
        cdp.call(
            "Emulation.setDeviceMetricsOverride",
            {"width": viewport[0], "height": viewport[1], "deviceScaleFactor": 1, "mobile": False},
        )
        cdp.call("Page.navigate", {"url": url})
        time.sleep(wait_sec)

        # 获取最终真实 URL 与页面文本，并执行 DOM 精准红框与滚动
        js_tokens = json.dumps(must_contain, ensure_ascii=False)
        js_code = f"""
        (() => {{
            const tokens = {js_tokens};
            const href = window.location.href || '';
            const bodyText = (document.body && document.body.innerText) ? document.body.innerText : '';
            const normBody = bodyText.replace(/\\s+/g, '');
            const missing = tokens.filter(t => !normBody.includes(t.replace(/\\s+/g, '')));
            if (missing.length > 0) {{
                return JSON.stringify({{ok: false, href, missing, sample: bodyText.slice(0, 120)}});
            }}

            // 关闭常见 Cookie/弹窗遮挡层，防止盖住正文
            document.querySelectorAll('[id*="cookie" i],[class*="cookie" i],[class*="consent" i],[class*="popup" i]').forEach(el => {{
                if (el.offsetHeight < 260) el.style.display = 'none';
            }});

            // 查找包含任意核心数字的最紧致块级元素（p / tr / li / h1-h3 / div）并画红框
            const candidates = Array.from(document.querySelectorAll('p, tr, li, h1, h2, h3, article div'));
            const matched = [];
            for (const el of candidates) {{
                const txt = (el.innerText || '').replace(/\\s+/g, '');
                if (!txt || txt.length > 650) continue;
                const hit = tokens.some(t => txt.includes(t.replace(/\\s+/g, '')));
                if (hit) {{
                    const hasMatchedChild = Array.from(el.querySelectorAll('p, tr, li')).some(ch => {{
                        const ct = (ch.innerText || '').replace(/\\s+/g, '');
                        return ct && ct.length <= 650 && tokens.some(t => ct.includes(t.replace(/\\s+/g, '')));
                    }});
                    if (!hasMatchedChild) matched.push(el);
                }}
            }}

            if (matched.length > 0) {{
                matched[0].scrollIntoView({{behavior: 'instant', block: 'center'}});
                matched.slice(0, 4).forEach(el => {{
                    el.style.outline = '3px solid #D92D20';
                    el.style.outlineOffset = '3px';
                    el.style.backgroundColor = 'rgba(254, 242, 242, 0.35)';
                }});
            }}
            return JSON.stringify({{
                ok: true,
                href,
                matchedCount: matched.length,
                normBody: normBody.slice(0, 30000)
            }});
        }})()
        """
        eval_res = cdp.call("Runtime.evaluate", {"expression": js_code, "returnByValue": True})
        raw_val = eval_res.get("result", {}).get("value", "{}")
        info = json.loads(raw_val)
        final_url = info.get("href", url)

        if not info.get("ok"):
            return False, final_url, f"页面未命中必需数字 {info.get('missing')}（已拒绝凑图）"

        # 滚完后等待 0.6s 确保重绘完成再截屏
        time.sleep(0.6)
        shot_res = cdp.call("Page.captureScreenshot", {"format": "png"})
        b64 = shot_res.get("data", "")
        if not b64:
            return False, final_url, "CDP 截屏返回空"

        img = Image.open(io.BytesIO(base64.b64decode(b64))).convert("RGB")
        banner = f"{banner_prefix} | {final_url}" if banner_prefix else f"原始公开网页：{final_url}"
        img = _add_provenance_banner(img, banner)
        os.makedirs(os.path.dirname(os.path.abspath(out_png)), exist_ok=True)
        img.save(out_png, format="PNG", optimize=True)

        # 自动写入溯源元数据 sidecar
        _write_sidecar_meta(
            out_png=out_png,
            capture_type="web",
            source_ref=final_url,
            must_contain=must_contain,
            matched_box_count=int(info.get("matchedCount", 0)),
            page_text_norm=info.get("normBody", ""),
            extra={"requested_url": url, "final_url": final_url},
        )
        return True, final_url, f"精确命中并红框标注 {info.get('matchedCount', 0)} 处段落"
    except Exception as e:
        return False, url, f"CDP 异常: {e}"
    finally:
        if cdp:
            cdp.close()
        try:
            proc.terminate()
            proc.wait(timeout=3)
        except Exception:
            try:
                proc.kill()
            except Exception:
                pass


def render_notice_text_card(
    title: str,
    source_meta: str,
    body_lines: List[str],
    highlight_lines: List[int],
    out_png: str,
    must_contain: Optional[List[str]] = None,
) -> bool:
    """
    仅用于上交所 e互动/投资者关系活动记录表等无直接公开 PDF 链接的正式交易所披露原文卡片渲染（带红框标注）。
    渲染成功后同样生成 `.meta.json` 溯源元数据。
    """
    width = 1100
    font_title = _get_cn_font(24)
    font_meta = _get_cn_font(16)
    font_body = _get_cn_font(19)

    wrapped: List[Tuple[str, bool]] = []
    for idx, raw in enumerate(body_lines):
        is_hl = idx in highlight_lines
        line = ""
        for ch in raw:
            if len(line.encode("gbk", errors="ignore")) >= 86:
                wrapped.append((line, is_hl))
                line = ch
            else:
                line += ch
        if line:
            wrapped.append((line, is_hl))

    line_h = 30
    total_h = 130 + len(wrapped) * line_h + 40
    img = Image.new("RGB", (width, total_h), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    draw.rectangle([0, 0, width, 44], fill=(0, 72, 112))
    draw.text((20, 10), source_meta[:110], fill=(255, 255, 255), font=font_meta)
    draw.text((24, 62), title, fill=(15, 23, 42), font=font_title)
    draw.line([(24, 102), (width - 24, 102)], fill=(0, 72, 112), width=2)

    y = 118
    hl_start = None
    box_count = 0
    for idx, (txt, is_hl) in enumerate(wrapped):
        draw.text((34, y), txt, fill=(30, 41, 59), font=font_body)
        if is_hl and hl_start is None:
            hl_start = y - 5
        if (not is_hl or idx == len(wrapped) - 1) and hl_start is not None:
            hl_end = (y + line_h) if is_hl else (y - 2)
            draw.rectangle([22, hl_start, width - 22, hl_end], outline="#D92D20", width=3)
            box_count += 1
            hl_start = None
        y += line_h

    os.makedirs(os.path.dirname(os.path.abspath(out_png)), exist_ok=True)
    img.save(out_png, format="PNG", optimize=True)

    full_norm = _normalize_text(title + "".join(body_lines))
    _write_sidecar_meta(
        out_png=out_png,
        capture_type="notice_card",
        source_ref=source_meta,
        must_contain=must_contain or [],
        matched_box_count=box_count,
        page_text_norm=full_norm,
    )
    return True


def run_manifest_capture(
    manifest_path: str,
    out_dir: Optional[str] = None,
    overwrite: bool = False,
    write_back: bool = False,
) -> Tuple[int, int, int]:
    """
    通用批量截图驱动器：直接读取 06_证据清单.json 中各条目的 `capture_spec`（单个 dict）
    或 `capture_specs`（多图 list），自动批量执行 PDF 防歧义红框截图与 Chrome CDP 网页红框截图，
    彻底消除“每家公司都要另写临时截图胶水脚本”的问题。

    单条 capture_spec 标准结构示例：
    - PDF 截图：
      {
        "type": "pdf",
        "out_file": "fig02_hdi_hkex.png",
        "pdf_path": "00_原始资料/pdfs/港股招股书.pdf",   # 相对 manifest 所在目录或绝对路径
        "page_num": 140,                                 # 可选；不填则按 must_all 自动防歧义找页
        "must_all": ["43.8%", "HDI"],                    # 自动找页必需词
        "must_contain": ["43.8%"],                       # 截图页硬校验数字
        "keywords": ["43.8%", "HDI"],                    # 红框高亮词（缺省复用 must_contain）
        "banner": "超颖电子港股聆讯稿 P.140"
      }
    - 网页截图：
      {
        "type": "web",
        "out_file": "fig15_trendforce_capex.png",
        "url": "https://www.trendforce.com/presscenter/news/...",
        "must_contain": ["957", "106%"],
        "banner": "TrendForce 官方新闻稿"
      }

    返回：(成功数, 跳过数, 失败数)
    """
    manifest_dir = os.path.dirname(os.path.abspath(manifest_path))
    target_dir = os.path.abspath(out_dir or os.path.join(manifest_dir, "08_底稿证据"))
    os.makedirs(target_dir, exist_ok=True)

    with open(manifest_path, "r", encoding="utf-8") as f:
        manifest = json.load(f)

    ok_cnt = 0
    skip_cnt = 0
    fail_cnt = 0
    modified = False

    for sec in manifest.get("sections", []):
        for item in sec.get("items", []):
            if item.get("is_database_exempt"):
                continue
            specs = []
            if isinstance(item.get("capture_specs"), list):
                specs = item["capture_specs"]
            elif isinstance(item.get("capture_spec"), dict):
                specs = [item["capture_spec"]]
            if not specs:
                continue

            item_label = item.get("label") or item.get("title") or "未命名条目"
            resolved_entries = []

            for idx, sp in enumerate(specs, 1):
                ctype = str(sp.get("type", "")).lower().strip()
                out_file = sp.get("out_file") or f"{item.get('id', 'ev')}_{idx}.png"
                out_png = out_file if os.path.isabs(out_file) else os.path.join(target_dir, out_file)
                meta_path = f"{out_png}.meta.json"

                if (not overwrite) and os.path.exists(out_png) and os.path.exists(meta_path):
                    skip_cnt += 1
                    print(f"  [SKIP] {item_label[:26]} -> {os.path.basename(out_png)}（已存在含 .meta.json）")
                    resolved_entries.append((out_png, sp.get("url", "")))
                    continue

                if ctype == "pdf":
                    raw_pdf = sp.get("pdf_path", "")
                    pdf_path = raw_pdf if os.path.isabs(raw_pdf) else os.path.join(manifest_dir, raw_pdf)
                    must_contain = list(sp.get("must_contain") or sp.get("must_all") or [])
                    keywords = list(sp.get("keywords") or must_contain)
                    page_num = sp.get("page_num")
                    if page_num is None:
                        must_all = list(sp.get("must_all") or must_contain)
                        page_num = find_verified_pdf_page(pdf_path, must_all)
                    if not page_num:
                        fail_cnt += 1
                        print(f"  [FAIL] {item_label[:26]} -> PDF 未找到同时命中 {sp.get('must_all') or must_contain} 的页面")
                        continue
                    banner = sp.get("banner") or f"{os.path.basename(pdf_path)} P.{page_num}"
                    ok = capture_pdf_with_redbox(
                        pdf_path=pdf_path,
                        page_num=int(page_num),
                        keywords=keywords,
                        out_png=out_png,
                        banner_text=banner,
                        must_contain=must_contain,
                        require_redbox=True,
                    )
                    if ok:
                        ok_cnt += 1
                        print(f"  [OK-PDF] {item_label[:26]} -> {os.path.basename(out_png)} (P.{page_num})")
                        resolved_entries.append((out_png, ""))
                    else:
                        fail_cnt += 1
                        print(f"  [FAIL-PDF] {item_label[:26]} -> P.{page_num} 校验或红框失败")

                elif ctype == "web":
                    url = sp.get("url", "").strip()
                    must_contain = list(sp.get("must_contain") or [])
                    banner = sp.get("banner") or item_label[:40]
                    ok, final_url, msg = capture_web_with_dom_redbox(
                        url=url,
                        must_contain=must_contain,
                        out_png=out_png,
                        banner_prefix=banner,
                    )
                    if ok:
                        ok_cnt += 1
                        print(f"  [OK-WEB] {item_label[:26]} -> {os.path.basename(out_png)}")
                        resolved_entries.append((out_png, final_url))
                    else:
                        fail_cnt += 1
                        print(f"  [FAIL-WEB] {item_label[:26]} -> {msg}")
                else:
                    fail_cnt += 1
                    print(f"  [FAIL] {item_label[:26]} -> 未知 capture_spec.type='{ctype}'")

            if write_back and resolved_entries:
                rel_paths = [os.path.relpath(p, manifest_dir).replace("\\", "/") for p, _ in resolved_entries]
                if len(rel_paths) == 1:
                    item["evidence_screenshot"] = rel_paths[0]
                    if resolved_entries[0][1]:
                        item["source_url"] = resolved_entries[0][1]
                else:
                    item["evidence_screenshots"] = [
                        {"path": rp, "url": u} if u else {"path": rp}
                        for rp, (_, u) in zip(rel_paths, resolved_entries)
                    ]
                modified = True

    if write_back and modified:
        with open(manifest_path, "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)
        print(f"[INFO] 已自动将最新截图路径回写至 {manifest_path}")

    print(f"[SUMMARY] 批量截图完成：成功 {ok_cnt} 张 | 跳过 {skip_cnt} 张 | 失败 {fail_cnt} 张")
    return ok_cnt, skip_cnt, fail_cnt


if __name__ == "__main__":
    import argparse
    import sys

    parser = argparse.ArgumentParser(description="国金证券研报底稿——通用批量红框截图 CLI")
    parser.add_argument("--manifest", required=True, help="06_证据清单.json 路径")
    parser.add_argument("--out-dir", default=None, help="截图输出目录（默认与 manifest 同级的 08_底稿证据/）")
    parser.add_argument("--overwrite", action="store_true", help="强制覆盖已有含 .meta.json 的截图")
    parser.add_argument("--write-back", action="store_true", help="截图完成后自动把相对路径与最终 URL 回写进清单")
    args = parser.parse_args()

    _, _, fails = run_manifest_capture(
        manifest_path=args.manifest,
        out_dir=args.out_dir,
        overwrite=args.overwrite,
        write_back=args.write_back,
    )
    sys.exit(1 if fails > 0 else 0)
