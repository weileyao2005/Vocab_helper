# -*- coding: utf-8 -*-
"""
扇贝生词助手 —— 背单词自动解释浮窗

用法:双击「启动.vbs」(无控制台运行)
- 首次运行:引导框选两个区域(① 黄条提示区 ② 单词区)
- 主屏背单词,点"不认识"后黄条出现 → 自动收集生词到队列
- 攒几个词后点「解释这N个词」,AI 流式输出词根级讲解
- 下方输入框可追问 / 补充自己的理解

窗口建议放在第二块屏幕。
"""
import ctypes
import datetime
import difflib
import json
import os
import queue as queue_mod
import re
import threading
import time
import traceback
from pathlib import Path

# ---- DPI 感知:必须在创建 Tk 窗口之前设置,保证 Tk 坐标 = 截图像素 ----
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PER_MONITOR_AWARE
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageGrab
import requests

import sys

if getattr(sys, "frozen", False):
    # 打包版:数据存到用户目录(AppData),exe 文件保持无状态 ——
    # 无论怎么转发 exe,都不会带出使用者的记录 / API Key
    APP_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "扇贝生词助手"
    try:
        APP_DIR.mkdir(parents=True, exist_ok=True)
    except Exception:
        APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent
CONFIG_PATH = APP_DIR / "config.json"
API_PATH = APP_DIR / "api.json"            # API 配置(每位使用者自己的,勿随包分发)
HISTORY_PATH = APP_DIR / "history.jsonl"
WORDS_PATH = APP_DIR / "words.json"        # 单词累计统计(出现次数/解释次数)
QUEUE_PATH = APP_DIR / "queue.json"        # 待解释队列持久化(重启不丢)
LOGS_DIR = APP_DIR / "logs"                # 每日记录(生词 + 完整对话)
LOG_PATH = APP_DIR / "vocab_helper.log"
SETTINGS_PATH = Path(os.environ.get("USERPROFILE", str(Path.home()))) / ".claude" / "settings.json"

DEFAULT_FLAG = "稍后将继续安排这个单词的学习"
FLAG_COVERAGE = 0.55          # 标志句最长公共块覆盖率阈值
NOISE_TOKENS = {"uk", "us", "tab", "esc", "cet"}
DEFAULT_MODEL = "deepseek-flash"

# OCR:Windows 系统自带 OCR(winrt)——常驻系统服务,单次 5~90ms,CPU 占用极低
_winocr_local = threading.local()


def ocr_win(img):
    """Windows 系统 OCR(线程内懒加载)。
    返回 (words, lines):words=[(height, top, text, 1.0)], lines=[(text, top, height)]"""
    state = getattr(_winocr_local, "state", None)
    if state is None:
        import asyncio
        from winrt.windows.media.ocr import OcrEngine
        from winrt.windows.globalization import Language
        engine = OcrEngine.try_create_from_language(Language("zh-Hans-CN"))
        if engine is None:
            engine = OcrEngine.try_create_from_user_profile_languages()
        state = (asyncio.new_event_loop(), engine)
        _winocr_local.state = state
    loop, engine = state
    return loop.run_until_complete(_win_ocr_async(engine, img))


async def _win_ocr_async(engine, img):
    import io
    from winrt.windows.storage.streams import InMemoryRandomAccessStream, DataWriter
    from winrt.windows.graphics.imaging import BitmapDecoder
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream)
    writer.write_bytes(buf.getvalue())
    await writer.store_async()
    await writer.flush_async()
    stream.seek(0)
    decoder = await BitmapDecoder.create_async(stream)
    bmp = await decoder.get_software_bitmap_async()
    res = await engine.recognize_async(bmp)
    words, lines = [], []
    for line in res.lines:
        lines.append(line.text)
        for w in line.words:
            r = w.bounding_rect
            words.append((r.height, r.y, w.text, 1.0))
    return words, lines

SYSTEM_PROMPT = """你是一位 GRE 词汇老师。学生正在用背单词软件刷 GRE 词表,遇到没记住的词就发给你。
背景:软件自带释义太简略。他基础六级、目标 GRE,核心需求是"为什么这个词是这个意思"(词根逻辑)以及"怎么记住"。

对每个词,按下面结构讲:
**【词根故事】** 用 4~6 句讲一段迷你词源故事:词根的出处(哪门语言、字面义)→ 历史/使用场景 → 语义一步步演变到今天。要有画面感和可读性,但只讲与理解词义直接相关的部分。
**【家族】** 3~5 个同根或高相关的 GRE 词(词 + 极简释义);如果存在形近但不同源的"假亲戚",必须点破
**【例句】** 1~2 句(体现固定搭配)

通用要求:
- 不要跑偏:不铺垫、不总结、不写与记词无关的延伸
- 篇幅约 600 字/词,以"词根故事讲透"为最优先,其余部分从简
- 若学生给了个人联想或疑问(如"跟XX像""我以为是YY的意思"),优先用 1~3 句话破案:对在哪、错在哪、真实关系
- 生词列表是屏幕 OCR 自动采集,偶有识别错(缺字母/粘连/截断);若某词看起来不是完整合理的英文单词,先推断最可能的词并注明(如「识别修正:bbergast → flabbergast」),再按推断词讲解
- 全程中文讲解、关键处加粗;排版适配窄窗口;逐词分节,标题用 `# 1. 单词`"""


# ============================================================ 基础工具

def log(msg):
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now():%H:%M:%S}] {msg}\n")
    except Exception:
        pass


def load_config():
    if CONFIG_PATH.exists():
        try:
            cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            return cfg if isinstance(cfg, dict) else {}
        except Exception:
            log("config.json 读取失败,使用默认配置")
    return {}


def save_config(cfg):
    try:
        CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        log(f"配置保存失败: {e}")


# ---------------- API 配置(api.json 优先;开发者本机可用环境变量/Claude 配置兜底) ----------------

DEFAULT_BASE_URL = "https://api.deepseek.com/anthropic"


def load_api_config():
    cfg = {"api_key": "", "base_url": DEFAULT_BASE_URL, "model": DEFAULT_MODEL}
    if API_PATH.exists():
        try:
            data = json.loads(API_PATH.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                cfg.update({k: v for k, v in data.items() if v})
        except Exception:
            log("api.json 读取失败")
    if not cfg.get("api_key"):
        # 兜底:环境变量 / 本机 Claude Code 配置(仅供开发者本机)
        key = os.environ.get("ANTHROPIC_AUTH_TOKEN")
        base = os.environ.get("ANTHROPIC_BASE_URL")
        if not key and SETTINGS_PATH.exists():
            try:
                env = json.loads(SETTINGS_PATH.read_text(encoding="utf-8")).get("env", {})
                key = key or env.get("ANTHROPIC_AUTH_TOKEN")
                base = base or env.get("ANTHROPIC_BASE_URL")
            except Exception:
                pass
        if key:
            cfg["api_key"] = key
        if base:
            cfg["base_url"] = base
    return cfg


def save_api_config(cfg):
    try:
        API_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        log(f"api.json 保存失败: {e}")


def today_explained_count():
    n = 0
    today = datetime.date.today().isoformat()
    if HISTORY_PATH.exists():
        try:
            for line in HISTORY_PATH.read_text(encoding="utf-8").splitlines():
                try:
                    if json.loads(line).get("ts", "").startswith(today):
                        n += 1
                except Exception:
                    pass
        except Exception:
            pass
    return n


def append_history(words):
    try:
        with open(HISTORY_PATH, "a", encoding="utf-8") as f:
            for w in words:
                f.write(json.dumps({"ts": datetime.datetime.now().isoformat(timespec="seconds"),
                                    "word": w}, ensure_ascii=False) + "\n")
    except Exception as e:
        log(f"历史写入失败: {e}")


# ---------------- 单词统计库(累计出现次数/解释次数) ----------------

def load_word_stats():
    """words.json: {word: {count, first_seen, last_seen, explained}}
    首次运行若已有 history.jsonl,则从中导入历史。"""
    if WORDS_PATH.exists():
        try:
            d = json.loads(WORDS_PATH.read_text(encoding="utf-8"))
            return d if isinstance(d, dict) else {}
        except Exception:
            log("words.json 读取失败,重建")
    stats = {}
    if HISTORY_PATH.exists():
        try:
            for line in HISTORY_PATH.read_text(encoding="utf-8").splitlines():
                try:
                    w = json.loads(line).get("word")
                    if w:
                        st = stats.setdefault(w, {"count": 0, "first_seen": None,
                                                  "last_seen": None, "explained": 0})
                        st["count"] += 1
                        st["explained"] += 1
                except Exception:
                    pass
        except Exception:
            pass
    return stats


def save_word_stats(stats):
    try:
        WORDS_PATH.write_text(json.dumps(stats, ensure_ascii=False, indent=1), encoding="utf-8")
    except Exception as e:
        log(f"单词库保存失败: {e}")


def append_daily(text):
    """追加到 logs/YYYY-MM-DD.md"""
    try:
        LOGS_DIR.mkdir(exist_ok=True)
        p = LOGS_DIR / f"{datetime.date.today().isoformat()}.md"
        fresh = not p.exists()
        with open(p, "a", encoding="utf-8") as f:
            if fresh:
                f.write(f"# 扇贝生词记录 · {datetime.date.today().isoformat()}\n\n")
            f.write(text)
    except Exception as e:
        log(f"每日记录写入失败: {e}")


# ============================================================ OCR 解析

def grab_region(bbox):
    return ImageGrab.grab(bbox=tuple(bbox), all_screens=True)


def ocr_blocks(img):
    """兼容旧接口:Windows OCR 返回词级块 [(height, top, text, conf), ...]"""
    words, _ = ocr_win(img)
    return words


def pick_word(blocks):
    """从识别块中选主词:字号最大(高度最大)的合法英文 token"""
    best = None
    for h, top, text, conf in blocks:
        t = (text or "").strip().strip(".,;:!?()[]{}'\"")
        if not re.fullmatch(r"[A-Za-z][A-Za-z'\-]{2,17}", t):
            continue
        if t.isupper() or conf < 0.5:
            continue
        key = (h, -top)
        if best is None or key > best[0]:
            best = (key, t.lower())
    return best[1] if best else None


def _norm_cn(s):
    return re.sub(r"[\s\W_]+", "", s or "")


def match_flag(text, target=DEFAULT_FLAG, coverage=FLAG_COVERAGE):
    """OCR 文本是否包含标志句(容忍识别错字)。
    ① 整句包含;② 最长公共块覆盖 ≥55%;③ 特征词命中 ≥3 个;
    ④ 字符集合覆盖率 ≥50% 且含强特征词("继续/安排/稍后"类)。"""
    t = _norm_cn(text)
    tgt = _norm_cn(target)
    if not t or not tgt:
        return False
    if tgt in t:
        return True
    m = difflib.SequenceMatcher(None, tgt, t).find_longest_match(0, len(tgt), 0, len(t))
    if m.size >= coverage * len(tgt):
        return True
    kws = ("稍后", "继续", "安排", "单词", "学习")
    if sum(1 for k in kws if k in t) >= 3:
        return True
    strong = ("继续", "安排", "排汶", "稍", "継")
    if any(s in t for s in strong):
        cov = len(set(t) & set(tgt)) / max(len(set(tgt)), 1)
        if cov >= 0.5:
            return True
    return False


# ============================================================ 检测线程

class Detector(threading.Thread):
    """单区域检测(Windows 系统 OCR,单次 5~90ms):
    每轮识别整块检测区 → 找"稍后将继续安排"黄条行 + 字号最大的英文词 → 入队"""

    def __init__(self, app):
        super().__init__(daemon=True)
        self.app = app
        self.paused = False
        self.stopped = False
        self._flag_on = False
        self._cur_word = None
        self._last_queued = None
        self._word_ts = 0.0        # 当前词最近一次变化的时间(用于"词稳定"判定)
        self._primed = False       # 启动后第一轮只记录现状,不把屏幕上残留的词重复入队
        self._last_p_text = ""     # 最近一次识别原文(入队日志留证)

    # ---------- 循环 ----------
    def run(self):
        while not self.stopped:
            cfg = self.app.cfg
            rb = cfg.get("region_bbox")
            if self.paused or not rb:
                time.sleep(0.3)
                continue
            t0 = time.time()
            try:
                self._tick(cfg, rb)
            except Exception:
                log("检测异常: " + traceback.format_exc().replace("\n", " ")[:300])
            # 动态间歇:保证"检测+处理"总周期接近 poll_interval
            interval = float(cfg.get("poll_interval", 0.15))
            time.sleep(max(0.02, interval - (time.time() - t0)))

    def _tick(self, cfg, rb):
        now = time.time()

        # ---- 一次识别整块区域:黄条行文本与单词词块都在结果里 ----
        words, lines = ocr_win(grab_region(rb))
        target = cfg.get("flag_sentence") or DEFAULT_FLAG
        flag_now = any(match_flag(t, target) for t in lines)
        self._flag_on = flag_now
        self._last_p_text = " | ".join(lines)[:100]

        # 黄条显示时提取主词(字号最大的英文词)
        if flag_now:
            word = pick_word(words)
            if word and word != self._cur_word:
                self._cur_word = word
                self._word_ts = now        # 词刚发生变化,记录时间

        # ---- 入队判定 ----
        w = self._cur_word
        if not self._primed:
            self._primed = True
            # 启动后头一轮:屏幕上若残留着"词库里已有"的旧词,不重复入队
            if w and w in self.app.word_stats:
                self._last_queued = w
            return

        # 条件:黄条在显示 + 有词 + 未入队过 + 词已稳定 ≥0.3s(不是刚切换的过渡帧)
        if (flag_now and w and w != self._last_queued
                and now - self._word_ts >= 0.3):
            self._last_queued = w
            self.app.q.put(("enqueue", w))
            log(f"入队: {w} | 识别: {self._last_p_text[:50]!r}")


# ============================================================ AI 层

class Chatter:
    """封装 DeepSeek(Anthropic 兼容端点)流式对话"""

    def __init__(self, model=None):
        cfg = load_api_config()
        self.key = cfg.get("api_key", "")
        self.base = cfg.get("base_url", DEFAULT_BASE_URL)
        self.model = model or cfg.get("model") or DEFAULT_MODEL
        self.messages = []

    @property
    def ready(self):
        return bool(self.key and self.base)

    def chat(self, user_text, on_delta, on_think=None):
        """同步流式请求;返回助手全文。messages 历史自动维护。
        on_think(len):模型思维链增量(仅用于界面提示"思考中")"""
        if not self.ready:
            raise RuntimeError("未找到 API 配置(ANTHROPIC_AUTH_TOKEN / BASE_URL)")
        self.messages.append({"role": "user", "content": user_text})
        payload = {
            "model": self.model,
            "max_tokens": 8000,
            "system": SYSTEM_PROMPT,
            "messages": self.messages,
            "stream": True,
            # 关闭思考模式:单词讲解是纯知识任务,无需推理链
            # 实测首字延迟 17.2s → 0.8s,总耗时 20.9s → 7.6s
            "thinking": {"type": "disabled"},
        }
        headers = {
            "Authorization": "Bearer " + self.key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        }
        parts = []
        try:
            self._stream_request(headers, payload, parts, on_delta, on_think)
        except Exception:
            # 请求失败时回滚本轮 user 消息,保持历史一致
            if self.messages and self.messages[-1].get("role") == "user":
                self.messages.pop()
            raise
        text = "".join(parts)
        self.messages.append({"role": "assistant", "content": text})
        return text

    def _stream_request(self, headers, payload, parts, on_delta, on_think=None):
        # read timeout 90s:只要服务端持续有数据(含思维链)就不会触发;
        # 完全卡死(90s 无任何事件)则尽快报错,而不是让用户干等
        with requests.post(self.base.rstrip("/") + "/v1/messages", headers=headers,
                           json=payload, stream=True, timeout=(15, 90)) as r:
            if r.status_code != 200:
                body = r.text[:300]
                raise RuntimeError(f"API 错误 {r.status_code}: {body}")
            for raw in r.iter_lines(decode_unicode=True):
                if not raw or not raw.startswith("data:"):
                    continue
                data = raw[5:].strip()
                if not data or data == "[DONE]":
                    continue
                try:
                    evt = json.loads(data)
                except Exception:
                    continue
                etype = evt.get("type")
                if etype == "content_block_delta":
                    delta = evt.get("delta") or {}
                    if delta.get("type") == "text_delta" and delta.get("text"):
                        parts.append(delta["text"])
                        on_delta(delta["text"])
                    elif delta.get("type") == "thinking_delta" and on_think and delta.get("text"):
                        on_think(len(delta["text"]))
                elif etype == "message_stop":
                    break

    def reset(self):
        self.messages = []


# ============================================================ 区域框选

class RegionSelector:
    """全屏半透明拖拽框选;返回 (x1,y1,x2,y2) 物理像素;取消返回 None"""

    def __init__(self, root, hint):
        self.root = root
        self.hint = hint
        self.result = None
        u = ctypes.windll.user32
        self.vx = u.GetSystemMetrics(76)   # 虚拟屏起点
        self.vy = u.GetSystemMetrics(77)
        self.vw = u.GetSystemMetrics(78)   # 虚拟屏尺寸
        self.vh = u.GetSystemMetrics(79)
        self.main_w = u.GetSystemMetrics(0)

    def select(self):
        top = tk.Toplevel(self.root)
        top.overrideredirect(True)
        top.geometry(f"{self.vw}x{self.vh}{self.vx:+d}{self.vy:+d}")
        top.attributes("-topmost", True)
        top.attributes("-alpha", 0.35)
        top.configure(bg="#0d1b2a")
        cv = tk.Canvas(top, bg="#0d1b2a", highlightthickness=0, cursor="crosshair")
        cv.pack(fill="both", expand=True)
        cx = max(200, self.main_w // 2 - self.vx)
        cv.create_text(cx, 90, fill="#ffffff", justify="center", width=self.vw - 240,
                       font=("Microsoft YaHei", 19, "bold"),
                       text=self.hint + "\n\n按住鼠标左键拖动框选;按 Esc 取消")
        st = {}

        def on_press(e):
            st["x0"], st["y0"] = e.x, e.y
            st["rect"] = cv.create_rectangle(e.x, e.y, e.x, e.y, outline="#00e0a0", width=2)

        def on_move(e):
            if "rect" in st:
                cv.coords(st["rect"], st["x0"], st["y0"], e.x, e.y)

        def on_release(e):
            if "x0" not in st:
                return
            x1, y1 = min(st["x0"], e.x), min(st["y0"], e.y)
            x2, y2 = max(st["x0"], e.x), max(st["y0"], e.y)
            if x2 - x1 >= 10 and y2 - y1 >= 8:
                self.result = (self.vx + x1, self.vy + y1, self.vx + x2, self.vy + y2)
            top.destroy()

        cv.bind("<ButtonPress-1>", on_press)
        cv.bind("<B1-Motion>", on_move)
        cv.bind("<ButtonRelease-1>", on_release)
        top.bind("<Escape>", lambda e: top.destroy())
        top.focus_force()
        try:
            top.grab_set()
        except Exception:
            pass
        self.root.wait_window(top)
        return self.result


def format_bbox(bbox):
    if not bbox:
        return "未设置"
    x1, y1, x2, y2 = bbox
    return f"({x1},{y1})-({x2},{y2}) {x2 - x1}×{y2 - y1}"


# ============================================================ 主界面

class App:
    def __init__(self):
        self.cfg = load_config()
        self._migrate_config()
        self.q = queue_mod.Queue()
        self.queue_items = self.load_queue()   # 待解释队列(启动时恢复):[{"word","repeat","joined"}]
        self._answer_started = False
        self._answer_anchor = None             # 本轮提问位置(Text 索引),用于"回到本轮问题"
        self.word_stats = load_word_stats()   # 单词累计统计
        self.batch_words = []          # 当前批次正在解释的词(写记录用)
        self.chatter = Chatter(model=self.cfg.get("model"))
        self.ai_busy = False
        self.md_buf = ""               # 流式中未渲染的尾部
        self.today_count = today_explained_count()
        self._geo_dirty = False        # 窗口尺寸变化待保存
        self._geo_ts = 0.0
        self.detector = Detector(self)

        self.root = tk.Tk()
        self.root.title("扇贝生词助手")
        self._build_ui()
        if self.queue_items:           # 恢复上次未解释的队列
            self._refresh_queue()
        self.detector.start()

        if not self.cfg.get("region_bbox"):
            self.root.after(400, self.first_run_wizard)
        elif not load_api_config().get("api_key"):
            self.root.after(400, self.open_settings)
            self._refresh_status("请先填写 API Key(点右上角「设置」)")

        self.root.after(120, self._pump)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.bind("<Configure>", self._on_configure)

    # ---------------- UI 构建 ----------------
    def _build_ui(self):
        r = self.root
        r.title("扇贝生词助手")
        geo = self.cfg.get("win_geometry")
        if not geo:
            u = ctypes.windll.user32
            vx, vw = u.GetSystemMetrics(76), u.GetSystemMetrics(78)
            geo = f"470x780+{vx + vw // 2 + 40}+60"
        r.geometry(geo)
        r.minsize(420, 560)
        r.attributes("-topmost", bool(self.cfg.get("topmost", True)))

        pad = dict(padx=8, pady=3)

        # 顶部控制行
        top = ttk.Frame(r)
        top.pack(fill="x", **pad)
        self.dot = tk.Label(top, text="●", fg="#0a0", font=("Segoe UI", 12))
        self.dot.pack(side="left")
        self.det_state = tk.Label(top, text="检测中")
        self.det_state.pack(side="left", padx=(0, 8))
        self.btn_pause = ttk.Button(top, text="暂停检测", width=9, command=self.toggle_pause)
        self.btn_pause.pack(side="left")
        self.btn_top = ttk.Button(top, text="置顶", width=5, command=self.toggle_top)
        self.btn_top.pack(side="right")
        ttk.Button(top, text="记录", width=5, command=self.open_logs).pack(side="right", padx=(0, 4))
        ttk.Button(top, text="设置", width=5, command=self.open_settings).pack(side="right", padx=(0, 4))
        self.var_top = tk.BooleanVar(value=bool(self.cfg.get("topmost", True)))

        # 区域设置行
        zones = ttk.Frame(r)
        zones.pack(fill="x", **pad)
        ttk.Button(zones, text="重选检测区", command=self.reselect).pack(side="left")
        self.zone_label = tk.Label(zones, text="", fg="#666", anchor="w")
        self.zone_label.pack(side="left", fill="x", expand=True)
        self._refresh_zone_label()

        ttk.Separator(r).pack(fill="x", pady=4)

        # 队列区
        qhead = ttk.Frame(r)
        qhead.pack(fill="x", **pad)
        self.q_label = tk.Label(qhead, text="待解释 (0):", font=("Microsoft YaHei", 10, "bold"))
        self.q_label.pack(side="left")
        ttk.Button(qhead, text="+ 补抓当前词", width=12, command=self.grab_current_word).pack(side="right")

        self.q_frame = ttk.Frame(r)
        self.q_frame.pack(fill="x", **pad)

        # 补充说明 + 解释按钮(多行输入,随文字自动增高)
        ttk.Label(r, text="补充说明(可选,会一起发给 AI):").pack(anchor="w", padx=8)
        self.extra_box = tk.Text(r, height=2, wrap="word", font=("Microsoft YaHei", 10),
                                 relief="solid", borderwidth=1, undo=True)
        self.extra_box.pack(fill="x", **pad)
        self.extra_box.bind("<KeyRelease>", lambda e: self._autosize(self.extra_box, 2, 5))
        self.btn_explain = ttk.Button(r, text="解释这 0 个词", command=self.on_explain)
        self.btn_explain.pack(fill="x", **pad)

        ttk.Separator(r).pack(fill="x", pady=4)

        # 解释区工具条
        txt_tools = ttk.Frame(r)
        txt_tools.pack(fill="x", padx=8)
        ttk.Button(txt_tools, text="↩ 回到本轮问题", width=15,
                   command=self.goto_answer_start).pack(side="right")

        # 解释区
        mid = ttk.Frame(r)
        mid.pack(fill="both", expand=True, padx=8)
        self.text = tk.Text(mid, wrap="word", font=("Microsoft YaHei", 10),
                            state="disabled", bg="#fbfbfb", relief="solid", borderwidth=1)
        sb = ttk.Scrollbar(mid, command=self.text.yview)
        self.text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.text.pack(side="left", fill="both", expand=True)
        self.text.tag_configure("h1", font=("Microsoft YaHei", 12, "bold"), spacing1=8, spacing3=4)
        self.text.tag_configure("h2", font=("Microsoft YaHei", 11, "bold"), spacing1=6, spacing3=3)
        self.text.tag_configure("bold", font=("Microsoft YaHei", 10, "bold"))
        self.text.tag_configure("bullet", lmargin1=16, lmargin2=28)
        self.text.tag_configure("mono", font=("Consolas", 9), foreground="#334")
        self.text.tag_configure("sys", foreground="#888")

        # 追问:标签行(按钮放右侧)+ 输入框独占一行,与其它框左右对齐
        ask_head = ttk.Frame(r)
        ask_head.pack(fill="x", **pad)
        ttk.Label(ask_head, text="追问:").pack(side="left")
        ttk.Button(ask_head, text="新对话", width=7, command=self.new_chat).pack(side="right")
        self.btn_ask = ttk.Button(ask_head, text="发送", width=6, command=self.on_ask)
        self.btn_ask.pack(side="right", padx=(0, 4))
        self.ask_box = tk.Text(r, height=2, wrap="word", font=("Microsoft YaHei", 10),
                               relief="solid", borderwidth=1, undo=True)
        self.ask_box.pack(fill="x", **pad)
        self.ask_box.bind("<Return>", self._on_ask_enter)
        self.ask_box.bind("<Shift-Return>", self._on_ask_newline)
        self.ask_box.bind("<KeyRelease>", lambda e: self._autosize(self.ask_box, 2, 6))

        # 状态栏
        self.status = tk.Label(r, text="", anchor="w", fg="#555")
        self.status.pack(fill="x", padx=8, pady=(0, 6))
        self._refresh_status("就绪")

    def _migrate_config(self):
        """旧版双区域(prompt_bbox+word_bbox)自动合并为新版单检测区,无需用户重框"""
        if "region_bbox" in self.cfg:
            return
        pb, wb = self.cfg.get("prompt_bbox"), self.cfg.get("word_bbox")
        if pb and wb:
            self.cfg["region_bbox"] = [min(pb[0], wb[0]), min(pb[1], wb[1]),
                                       max(pb[2], wb[2]), max(pb[3], wb[3])]
            save_config(self.cfg)
            log(f"配置迁移: 双区域合并为 region_bbox={self.cfg['region_bbox']}")

    # ---------------- 辅助刷新 ----------------
    def _refresh_zone_label(self):
        rb = self.cfg.get("region_bbox")
        if rb:
            self.zone_label.config(text=f"检测区 {format_bbox(rb)}")
        else:
            self.zone_label.config(text="检测区未设置")

    def _refresh_queue(self):
        for w in self.q_frame.winfo_children():
            w.destroy()
        # Tk 在"最后一个子控件被销毁"时不会重算 Frame 高度,队列空时强制收缩
        self.q_frame.configure(height=1 if not self.queue_items else 0)
        active = [it for it in self.queue_items if it["joined"]]
        self.q_label.config(text=f"待解释 ({len(active)}):")
        for i, it in enumerate(self.queue_items, 1):
            row = ttk.Frame(self.q_frame)
            row.pack(fill="x", pady=1)
            ttk.Label(row, text=f"{i}. {it['word']}", font=("Consolas", 11)).pack(side="left")
            if it["repeat"]:
                mark = f"已出现过 {it['repeat']} 次" + (" · 本次再解释" if it["joined"] else "")
                tk.Label(row, text=mark, fg="#c60", font=("Microsoft YaHei", 9)).pack(side="left", padx=5)
                if not it["joined"]:
                    ttk.Button(row, text="再解释", width=6,
                               command=lambda w=it["word"]: self.join_again(w)).pack(side="right")
            ttk.Button(row, text="×", width=3,
                       command=lambda w=it["word"]: self.remove_word(w)).pack(side="right")
        n = len(active)
        self.btn_explain.config(text=f"解释这 {n} 个词" if n else "解释(队列为空)")

    def _refresh_status(self, extra=None):
        base = f"今日已解释 {self.today_count} 词"
        if extra:
            base += f"  |  {extra}"
        self.status.config(text=base)

    def _det_ui_state(self):
        if not self.cfg.get("region_bbox"):
            self.dot.config(fg="#aaa")
            self.det_state.config(text="未设置区域")
            self.btn_pause.config(text="暂停检测")
        elif self.detector.paused:
            self.dot.config(fg="#d80")
            self.det_state.config(text="已暂停")
            self.btn_pause.config(text="继续检测")
        else:
            self.dot.config(fg="#0a0")
            self.det_state.config(text="检测中")
            self.btn_pause.config(text="暂停检测")

    # ---------------- 队列操作 ----------------
    def load_queue(self):
        """启动时恢复当天未解释的队列(重启不丢)"""
        if QUEUE_PATH.exists():
            try:
                d = json.loads(QUEUE_PATH.read_text(encoding="utf-8"))
                if d.get("date") == datetime.date.today().isoformat() and isinstance(d.get("items"), list):
                    return [it for it in d["items"]
                            if isinstance(it, dict) and it.get("word")]
            except Exception:
                log("queue.json 读取失败")
        return []

    def save_queue(self):
        try:
            QUEUE_PATH.write_text(json.dumps(
                {"date": datetime.date.today().isoformat(), "items": self.queue_items},
                ensure_ascii=False, indent=1), encoding="utf-8")
        except Exception as e:
            log(f"队列保存失败: {e}")

    def remove_word(self, word):
        self.queue_items = [it for it in self.queue_items if it["word"] != word]
        self.save_queue()
        self._refresh_queue()

    def join_again(self, word):
        """重复词 → 加入本次解释"""
        for it in self.queue_items:
            if it["word"] == word:
                it["joined"] = True
        self.save_queue()
        self._refresh_queue()

    def grab_current_word(self):
        """手动补抓:立即识别检测区,把当前词加入队列"""
        rb = self.cfg.get("region_bbox")
        if not rb:
            self._refresh_status("请先框选检测区")
            return
        try:
            word = pick_word(ocr_blocks(grab_region(rb)))
        except Exception as e:
            self._refresh_status(f"抓取失败: {e}")
            return
        if word:
            self.add_word(word, manual=True)
        else:
            self._refresh_status("没识别到单词")

    def add_word(self, word, manual=False):
        if any(it["word"] == word for it in self.queue_items):
            return
        # ---- 更新累计统计 ----
        st = self.word_stats.setdefault(word, {"count": 0, "first_seen": None,
                                               "last_seen": None, "explained": 0})
        prev = st["count"]                       # 此前出现过几次
        now = datetime.datetime.now().isoformat(timespec="seconds")
        st["count"] += 1
        if not st["first_seen"]:
            st["first_seen"] = now
        st["last_seen"] = now
        save_word_stats(self.word_stats)
        # ---- 记入当日记录 ----
        mark = f"(已出现过 {prev} 次)" if prev > 0 else ""
        append_daily(f"- [{datetime.datetime.now():%H:%M}] {word} {mark}\n")
        # ---- 入队:重复词默认不参与本批解释 ----
        self.queue_items.append({"word": word, "repeat": prev, "joined": prev == 0})
        self.save_queue()
        self._refresh_queue()
        if not manual:
            self._refresh_status(f"检测到生词: {word}" + (f"(第 {prev + 1} 次出现)" if prev else ""))

    # ---------------- 解释 / 追问 ----------------
    def on_explain(self):
        words = [it["word"] for it in self.queue_items if it["joined"]]
        extra = self.extra_box.get("1.0", "end-1c").strip()
        if not words and not extra:
            self._refresh_status("队列为空")
            return
        if self.ai_busy:
            return
        if words:
            lines = "\n".join(f"{i + 1}. {w}" for i, w in enumerate(words))
            msg = f"今天背单词遇到的生词:\n{lines}"
            if extra:
                msg += f"\n\n补充说明:{extra}"
        else:
            msg = extra
        self.queue_items = []
        self.save_queue()
        self.extra_box.delete("1.0", "end")
        self.extra_box.config(height=2)
        self._refresh_queue()
        append_history(words)
        self.today_count += len(words)
        self._refresh_status()
        self.start_ai(msg, is_new_batch=True, batch_words=words)

    def on_ask(self):
        text = self.ask_box.get("1.0", "end-1c").strip()
        if not text or self.ai_busy:
            return
        self.ask_box.delete("1.0", "end")
        self.ask_box.config(height=2)
        self.start_ai(text, is_new_batch=False)

    def start_ai(self, user_text, is_new_batch=False, batch_words=None):
        self.ai_busy = True
        self._answer_started = False
        self.batch_words = list(batch_words or [])
        self.btn_explain.config(state="disabled")
        self.btn_ask.config(state="disabled")
        # 记录本轮问题位置,讲解完成后自动滚回这里
        try:
            self._answer_anchor = self.text.index("end-1c")
        except Exception:
            self._answer_anchor = None
        if is_new_batch:
            self._append_sys("\n" + "─" * 46 + f"\n▷ 你: {user_text}\n\n")
        else:
            self._append_sys(f"\n▷ 你: {user_text}\n\n")
        self._refresh_status("AI 正在思考…")
        threading.Thread(target=self._ai_worker, args=(user_text,), daemon=True).start()

    def goto_answer_start(self):
        """滚动回到本轮提问的位置"""
        if self._answer_anchor:
            try:
                self.text.see(self._answer_anchor)
            except Exception:
                pass

    def _ai_worker(self, user_text):
        think = {"n": 0, "t": 0.0}

        def on_think(n):
            think["n"] += n
            now = time.time()
            if now - think["t"] > 0.4:          # 节流:每 0.4s 最多推一次界面
                think["t"] = now
                self.q.put(("think", think["n"]))

        try:
            full = self.chatter.chat(user_text, lambda s: self.q.put(("delta", s)), on_think)
            self.q.put(("ai_full", (user_text, full)))
        except Exception as e:
            self.q.put(("ai_error", str(e)))
        finally:
            self.q.put(("ai_done", None))

    def _save_chat_record(self, user_text, full):
        """完整对话落盘到 logs/日期.md,并累计各词的解释次数"""
        ts = datetime.datetime.now().strftime("%H:%M")
        words = self.batch_words
        self.batch_words = []
        if words:
            head = f"\n## [{ts}] 解释 {len(words)} 个词:{'、'.join(words)}\n"
            for w in words:
                st = self.word_stats.get(w)
                if st:
                    st["explained"] = st.get("explained", 0) + 1
            save_word_stats(self.word_stats)
        else:
            head = f"\n## [{ts}] 追问\n"
        quoted = "\n".join("> " + ln for ln in user_text.splitlines())
        append_daily(head + f"\n**你:**\n{quoted}\n\n**AI:**\n{full}\n\n---\n")

    def new_chat(self):
        if self.ai_busy:
            return
        self.chatter.reset()
        self.md_buf = ""
        self._answer_anchor = None
        self.text.config(state="normal")
        self.text.delete("1.0", "end")
        self.text.config(state="disabled")
        self._refresh_status("已开新对话")

    def _on_ask_enter(self, event):
        self.on_ask()
        return "break"   # 阻止默认换行,Enter 直接发送

    def _on_ask_newline(self, event):
        """Shift+Enter:插入换行(而非发送)"""
        self.ask_box.insert("insert", "\n")
        self._autosize(self.ask_box, 2, 6)
        return "break"

    def _autosize(self, box, min_h=2, max_h=6):
        """多行输入框随内容行数自动增高"""
        try:
            box.update_idletasks()
            idx = box.count("1.0", "end", "displaylines")
            lines = int(idx[0]) if idx else 1
        except Exception:
            lines = box.get("1.0", "end-1c").count("\n") + 1
        h = max(min_h, min(lines, max_h))
        if h != int(box.cget("height")):
            box.config(height=h)

    def open_logs(self):
        try:
            LOGS_DIR.mkdir(exist_ok=True)
            os.startfile(str(LOGS_DIR))
        except Exception as e:
            self._refresh_status(f"打开记录失败: {e}")

    def open_settings(self):
        """API 设置对话框(每位使用者填自己的 Key)"""
        cfg = load_api_config()
        dlg = tk.Toplevel(self.root)
        dlg.title("API 设置")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        body = ttk.Frame(dlg, padding=12)
        body.pack(fill="both", expand=True)

        ttk.Label(body, text="API Key(在 DeepSeek 官网申请):").grid(row=0, column=0, sticky="w", pady=4)
        e_key = ttk.Entry(body, width=48, show="*")
        e_key.grid(row=1, column=0, sticky="we", pady=(0, 8))
        e_key.insert(0, cfg.get("api_key", ""))

        ttk.Label(body, text="接口地址:").grid(row=2, column=0, sticky="w", pady=4)
        e_base = ttk.Entry(body, width=48)
        e_base.grid(row=3, column=0, sticky="we", pady=(0, 8))
        e_base.insert(0, cfg.get("base_url", DEFAULT_BASE_URL))

        ttk.Label(body, text="模型名:").grid(row=4, column=0, sticky="w", pady=4)
        e_model = ttk.Entry(body, width=48)
        e_model.grid(row=5, column=0, sticky="we", pady=(0, 4))
        e_model.insert(0, cfg.get("model") or DEFAULT_MODEL)

        ttk.Label(body, text="Key 只保存在本程序目录的 api.json 中,不会随程序分发",
                  foreground="#888").grid(row=6, column=0, sticky="w", pady=(0, 10))

        def do_save():
            save_api_config({
                "api_key": e_key.get().strip(),
                "base_url": e_base.get().strip() or DEFAULT_BASE_URL,
                "model": e_model.get().strip() or DEFAULT_MODEL,
            })
            old_msgs = self.chatter.messages
            self.chatter = Chatter()
            self.chatter.messages = old_msgs
            self._refresh_status("API 设置已保存")
            dlg.destroy()

        btns = ttk.Frame(body)
        btns.grid(row=7, column=0, sticky="e")
        ttk.Button(btns, text="取消", command=dlg.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="保存", command=do_save).pack(side="right")
        dlg.grab_set()

    # ---------------- 文本渲染(轻量 Markdown) ----------------
    def _append_sys(self, s):
        self.text.config(state="normal")
        self.text.insert("end", s, ("sys",))
        self.text.see("end")
        self.text.config(state="disabled")

    def _append_md_line(self, line):
        t = self.text
        stripped = line.strip()
        if not stripped:
            t.insert("end", "\n")
            return
        if stripped.startswith("#"):
            body = stripped.lstrip("#").strip()
            tag = "h1" if len(stripped) - len(stripped.lstrip("#")) <= 2 else "h2"
            t.insert("end", body + "\n", (tag,))
            return
        if re.fullmatch(r"[|\-\s:]+", stripped):   # markdown 表格分隔行,丢弃
            return
        if stripped.startswith("|"):               # 表格内容行,等宽原样显示
            t.insert("end", stripped + "\n", ("mono",))
            return
        is_bullet = stripped.startswith(("- ", "* ", "· ", "• "))
        body = stripped[2:] if is_bullet else stripped
        if is_bullet:
            t.insert("end", "  • ", ("bullet",))
        for seg in re.split(r"(\*\*.+?\*\*|`.+?`)", body):
            if not seg:
                continue
            if len(seg) > 4 and seg.startswith("**") and seg.endswith("**"):
                t.insert("end", seg[2:-2], ("bold",))
            elif len(seg) > 2 and seg.startswith("`") and seg.endswith("`"):
                t.insert("end", seg[1:-1], ("mono",))
            else:
                t.insert("end", seg)
        t.insert("end", "\n")

    def _flush_md(self, force=False):
        """把缓冲区中完整的行渲染出来;force 时连残行一起"""
        if not self.md_buf:
            return
        if force:
            chunk, self.md_buf = self.md_buf, ""
            lines = chunk.split("\n")
            if lines and lines[-1] == "":
                lines = lines[:-1]
        else:
            idx = self.md_buf.rfind("\n")
            if idx < 0:
                return
            chunk, self.md_buf = self.md_buf[:idx + 1], self.md_buf[idx + 1:]
            lines = chunk.split("\n")[:-1]
        self.text.config(state="normal")
        for line in lines:
            self._append_md_line(line)
        self.text.see("end")
        self.text.config(state="disabled")

    # ---------------- 事件泵 ----------------
    def _pump(self):
        try:
            while True:
                evt, payload = self.q.get_nowait()
                if evt == "enqueue":
                    self.add_word(payload)
                elif evt == "delta":
                    if not self._answer_started:
                        self._answer_started = True
                        self._refresh_status("AI 正在回答…")
                    self.md_buf += payload
                    self._flush_md()
                elif evt == "think":
                    if not self._answer_started:
                        self._refresh_status(f"AI 正在思考…({payload} 字)")
                elif evt == "ai_full":
                    self._save_chat_record(*payload)
                elif evt == "ai_done":
                    self._flush_md(force=True)
                    self.ai_busy = False
                    self.btn_explain.config(state="normal")
                    self.btn_ask.config(state="normal")
                    self._refresh_status("完成")
                    # 输出完毕,自动滚回本轮提问处,方便从头阅读
                    self.goto_answer_start()
                elif evt == "ai_error":
                    # 把本批词放回队列,避免失败后丢失
                    for w in self.batch_words:
                        if not any(it["word"] == w for it in self.queue_items):
                            prev = max(self.word_stats.get(w, {}).get("count", 1) - 1, 0)
                            self.queue_items.append({"word": w, "repeat": prev, "joined": True})
                    self.batch_words = []
                    self.save_queue()
                    self._refresh_queue()
                    self._append_sys(f"\n[出错] {payload}(词已放回队列,可以重试)\n")
                    self._refresh_status("AI 请求失败")
                elif evt == "log":
                    self._refresh_status(payload)
        except queue_mod.Empty:
            pass
        if self._geo_dirty and time.time() - self._geo_ts > 3:
            save_config(self.cfg)
            self._geo_dirty = False
        self._det_ui_state()
        self.root.after(120, self._pump)

    # ---------------- 区域设置 ----------------
    def first_run_wizard(self):
        self._do_reselect("请拖框选中【检测区】:包含【黄色提示条】和【单词】那一块即可;"
                          "框大一点没关系(包含释义、例句也不影响识别)")
        self._refresh_status("区域设置完成,开始检测")
        if not load_api_config().get("api_key"):
            self.open_settings()

    def reselect(self):
        self._do_reselect("请拖框选中检测区(包含黄色提示条和单词那一块;框大一点没关系)")

    def _do_reselect(self, hint):
        sel = RegionSelector(self.root, hint)
        bbox = sel.select()
        if bbox:
            self.cfg["region_bbox"] = list(bbox)
            save_config(self.cfg)
            self._refresh_zone_label()
            log(f"设置 region_bbox = {bbox}")
            # 重置检测器状态,避免旧区域缓存干扰
            self.detector._primed = False
            self.detector._cur_word = None
            self.detector._flag_on = False

    # ---------------- 其它 ----------------
    def toggle_pause(self):
        self.detector.paused = not self.detector.paused
        self._refresh_status("检测已暂停" if self.detector.paused else "检测继续")

    def toggle_top(self):
        cur = not bool(self.var_top.get())
        self.var_top.set(cur)
        self.root.attributes("-topmost", cur)
        self.cfg["topmost"] = cur
        save_config(self.cfg)

    def _on_configure(self, e):
        """窗口被移动/缩放时记录 geometry(3 秒防抖后落盘)"""
        if e.widget is self.root:
            try:
                self.cfg["win_geometry"] = self.root.geometry()
                self._geo_dirty = True
                self._geo_ts = time.time()
            except Exception:
                pass

    def _on_close(self):
        try:
            self.cfg["win_geometry"] = self.root.geometry()
            save_config(self.cfg)
        except Exception:
            pass
        self.detector.stopped = True
        self.root.destroy()


_INSTANCE_SOCK = None


def already_running():
    """单实例保护:占用本地端口,占不到说明已有实例在运行"""
    global _INSTANCE_SOCK
    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", 51987))
    except OSError:
        s.close()
        return True
    _INSTANCE_SOCK = s
    return False


def main():
    if already_running():
        try:
            r = tk.Tk()
            r.withdraw()
            from tkinter import messagebox as _mb
            _mb.showinfo("扇贝生词助手", "程序已经在运行了,窗口可能在另一块屏幕上。")
            r.destroy()
        except Exception:
            pass
        return
    app = App()
    app.root.mainloop()


if __name__ == "__main__":
    main()
