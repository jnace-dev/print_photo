#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
照片打印工具
==========
类似 Windows 系统自带的照片打印，支持选择照片、选择打印机、纸张、布局、张数，
并新增【单面 / 双面】打印选项。
依赖: pywin32、Pillow（在函数内延迟导入）；拖拽功能可选依赖 tkinterdnd2，
      缺失时自动降级且不影响文件对话框添加照片；tkinter 为 Python 标准库。
打包: 手工命令参考 pyinstaller --onefile --windowed；
      实际构建入口 safe_build() 使用「照片打印工具.spec」+ --clean +
      --workpath/--distpath build_tmp，需单独调用，未接入 main()。
DPI:  main() 在创建 Tk 根窗口前声明 Windows DPI 感知，随后按 96 DPI 基准
      换算缩放系数；所有硬编码像素经 px()/pxp() 换算，字体交给 tk scaling，
      弹窗遮罩截图按感知状态换算屏幕坐标，并监听跨屏 DPI 变化重排 UI。
"""
import os
import sys
import threading
import time
import traceback
import ctypes
from ctypes import wintypes
from collections import OrderedDict
# === 标准库 GUI 导入（tkinter 是 Python 标准库） ===
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------
# 双面模式: 1=单面, 2=长边翻页(书籍), 3=短边翻页(日历)
DUPLEX_NONE = 1
DUPLEX_VERTICAL = 2
DUPLEX_HORIZONTAL = 3
# 布局: (行, 列, 每页张数)
LAYOUTS = {
    "全页 (1张/页)": (1, 1, 1),
    "2张/页": (1, 2, 2),
    "4张/页": (2, 2, 4),
    "9张/页": (3, 3, 9),
    "16张/页": (4, 4, 16),
}
# 纸张: 名称 -> (类型, 值)  类型: 'size'=DMPAPER 常量, 'custom'=(宽, 高) 单位 0.1mm
PAPER_SIZES = {
    "A4": ("size", 9),          # DMPAPER_A4
    "Letter": ("size", 1),      # DMPAPER_LETTER
    "A5": ("size", 11),         # DMPAPER_A5
    "4x6 照片纸": ("custom", (1016, 1524)),
    "5x7 照片纸": ("custom", (1270, 1778)),
}
# 自动旋转阈值: 图片与纸张宽高比的「相对偏差」 |r_img-r_paper|/r_paper
# 超过此值时单向旋转 90 度；实现不比较旋转前后谁更贴合，个别极端比例旋转后反而更偏。
ROTATION_THRESHOLD = 0.5
# 支持的图片扩展名: 文件对话框与拖拽过滤共用
SUPPORTED_IMAGE_EXTENSIONS = (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif")
# 纸张物理尺寸映射(宽, 高) 单位mm, 用于纸张宽高比、预览绘制与打印时的自动旋转判定
PAPER_DIMENSIONS = {
    "A4": (210, 297),
    "Letter": (215.9, 279.4),
    "A5": (148, 210),
    "4x6 照片纸": (101.6, 152.4),
    "5x7 照片纸": (127, 177.8),
}
# 以下「像素」常量均为 100%(96 DPI) 设计基准值，运行时统一由 px()/pxp() 换算；
# BUTTON_WIDTH / Combobox width 等「字符单位」不换算，它们随字体自动缩放。
# 照片列表右侧间距(像素): 文件名文字右边缘到滚动条左边缘的最小可见间距。
# 滚动条与 canvas 平级、常驻列表右侧且始终可见(带上下箭头，不再置灰/隐藏)，
# 其左边缘即 canvas 右边界；interior 宽度跟随 canvas，所以文字右侧只留此间距，
# 不再叠加滚动条宽度。canvas 宽度变化时按新宽度重算并同步到复用的条目。
RIGHT_LIST_PADDING = 6
# 照片列表左侧间距(像素): 列表容器左边缘与预览区之间的可见间距
LEFT_LIST_PADDING = 6
# 照片列表右侧边缘间距(像素): 列表容器右边缘与窗口边缘之间的可见间距
RIGHT_LIST_EDGE_PADDING = 6
# 列表图片最大尺寸(像素): 每个条目显示图片的最大像素尺寸, 超出部分等比缩放
LIST_IMAGE_MAX_SIZE = 60
# 缩略图与文件名文字之间的间距(像素); grid 的 padx 与列宽计算共用此值
THUMBNAIL_TEXT_GAP = 6
# 列表缩略图列固定宽度(像素): 由「缩略图最大尺寸 + 文字间距」推导，避免魔法数；
# 用于计算文件名可用宽度；文件名左边缘对齐靠条目内的 grid 列宽实现。
THUMBNAIL_COLUMN_WIDTH = LIST_IMAGE_MAX_SIZE + THUMBNAIL_TEXT_GAP
CONFIRM_IMAGE_MAX_SIZE = 300
# 确认弹窗遮罩(平滑): 在「主窗口内部」抓一份客户区快照，整体调暗后覆盖回去。
# 快照本身保留原来的 UI 内容，只是整体变暗，因此视觉上是均匀半透明黑层、没有
# 扫描线；遮罩是 root 的直接子控件(Label)，不依赖 -alpha / 窗口合成器，关闭时
# 销毁 Label 即恢复真实控件，不会有「不透明黑层糊住主窗口」的风险。
# 代价: 弹窗期间背景为静态快照(内容不实时刷新)；截图坐标受多显示器 / DPI 缩放
# 影响，Pillow 用 all_screens=True 兼容多屏，尺寸不符时按比例缩回客户区。
# 抓图失败时自动降级为「不显示遮罩」，弹窗与业务功能不受影响。
DIALOG_SCRIM_DARKEN = 0.34
# ImageGrab 失败或窗口不可见时的降级色: 不用纯黑，避免偶尔残留变糊屏
DIALOG_SCRIM_FALLBACK_BG = "#A8A8A8"
SELECTED_BG_COLOR = "#DDEBF7"
UNSELECTED_BG_COLOR = "#FFFFFF"
FILE_NAME_COLOR = "#333333"
# spec 6.1.1: 工具栏与翻页条的所有 ttk.Button 共用此宽度(控件级不再传 width)。
# 按钮文字需容纳「上一页 (PgUp)」/「下一页 (PgDn)」：中文按约 2 字符宽计，
# 3*2+1+6=13 字符，Segoe UI 9 下实测约 90px；原 12 字符(84px)会被文字撑开，
# 导致翻页按钮比纯中文的工具栏按钮宽、破坏上一轮统一的宽高。故放大到 16 字符。
# ttk.Button 没有 height 选项，高度由 TButton 的垂直 padding 统一放大；
# 因此工具栏、翻页条、关闭按钮放大后仍保持彼此宽高一致。
# BUTTON_WIDTH 是「字符宽度」而非像素：字体随 tk scaling 放大后按钮自动变宽，
# 不能再乘 DPI 系数，否则会双重放大。
BUTTON_WIDTH = 16
# TButton 内边距(水平, 垂直): 像素值，_apply_win10_style 里经 pxp() 按 DPI 换算；
# 垂直方向放大按钮高度，水平留白避免文字贴边
BUTTON_PADDING = (8, 6)
# 渐变按钮初始尺寸(像素, 100% 基准): tk.Button 用像素尺寸，无法与 ttk 的
# 「字符宽度」严格相等；创建时经 px() 换算，打印 / 删除确定 / 清空确定三个
# 渐变按钮随后统一由 _align_gradient_button() 运行时测量 btn_add 对齐到相同
# 宽高（测量值本身已是当前 DPI 下的物理像素，天然跟随缩放）。
PRINT_BTN_GRADIENT_WIDTH = 132
PRINT_BTN_GRADIENT_HEIGHT = 32
# 翻页条容器高度(像素, 100% 基准): pack_propagate(False) 下必须容纳放大后的
# 按钮，否则按钮底部被裁切；取值略大于 TButton 自然高度 + 垂直 padding，
# 运行时经 px() 按 DPI 换算。
PAGINATION_BAR_HEIGHT = 40
# spec 6.1.3: UI 构造 + deiconify 耗时上限(秒)；不含解释器启动、Tk() 创建
# 以及 after(0) 延迟执行的默认值/样式/图标/打印机刷新，非「启动完成」全耗时。
STARTUP_TIME_LIMIT = 0.1
# 性能指标上限(秒)
PREVIEW_REDRAW_TIME_LIMIT = 0.5
LIST_REFRESH_TIME_LIMIT = 2.0
PAGE_COMPOSE_TIME_LIMIT = 1.0
# 所有耗时日志(PerfMeasure、拖放添加、启动)的小数位数由 MEASURE_PRECISION 推导
MEASURE_PRECISION = 0.001
TIME_LOG_DIGITS = (len(str(MEASURE_PRECISION).split(".")[-1])
                   if 0 < MEASURE_PRECISION < 1 else 0)
# exe 体积指标(字节 / 百分比)
EXE_SIZE_LIMIT = 52428800
CURRENT_EXE_SIZE_BASELINE = 178330724
VOLUME_REDUCTION_RATIO = 70.0
# 缩减比例日志的小数位数
VOLUME_MEASURE_PRECISION = 1
# 图片缓存容量上限(条目数): 按用途分级，避免整页大图长时间驻留内存。
# 页面合成图按打印机分辨率可达数 MB/张，上限最小；预览次之；缩略图最小可放宽。
IMAGE_CACHE_MAX_PAGE = 16
IMAGE_CACHE_MAX_PREVIEW = 64
IMAGE_CACHE_MAX_THUMB = 128
# spec 6.1.2: TCombobox 内边距 (水平, 垂直)，像素值，运行时经 pxp() 换算
COMBOBOX_PADDING = (4, 4)
# 应用标题与日志文件路径
APP_TITLE = "照片打印工具"
LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "print_log.txt")
# ---------------------------------------------------------------------------
# DPI 缩放适配
# ---------------------------------------------------------------------------
# Windows 对「未声明 DPI 感知」的进程做位图拉伸，界面会发糊；必须在创建任何
# Tk 窗口之前声明感知。声明成功后 Tk 的几何坐标即物理像素，字体由 tk scaling
# 按点数自动放大；所有硬编码像素再经 px() 换算，二者配合才不会双重放大。
DPI_BASE = 96.0                # 100% 缩放对应的 DPI
DPI_SCALE_MIN = 1.0            # 低于 100% 不缩小，避免界面过小
DPI_SCALE_MAX = 3.0            # 上限防御异常 DPI 读数
_DPI_SCALE = 1.0               # 当前缩放系数，由 read_dpi_scale() 更新
_DPI_AWARENESS = "unavailable" # enable_dpi_awareness() 的实际结果
def enable_dpi_awareness():
    """在创建 Tk 根窗口【之前】调用；返回实际采用的感知模式。
    优先级: per-monitor v2(Win10 1703+) > per-monitor(Win8.1+) >
    system(Vista+) > unavailable(非 Windows / 全部失败，保持系统拉伸)。
    SetProcessDpiAwarenessContext / SetProcessDpiAwareness 每个进程只能调用一次，
    若已被宿主(如 pythonw 清单)设置过会返回失败，此时逐级回退即可，不视为错误。
    """
    global _DPI_AWARENESS
    if sys.platform != "win32":
        _DPI_AWARENESS = "unavailable"
        return _DPI_AWARENESS
    # per-monitor v2: DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = (void*)-4
    try:
        fn = ctypes.windll.user32.SetProcessDpiAwarenessContext
        fn.argtypes = [ctypes.c_void_p]
        fn.restype = ctypes.c_bool
        if fn(ctypes.c_void_p(-4)):
            _DPI_AWARENESS = "per-monitor-v2"
            return _DPI_AWARENESS
    except Exception:
        pass
    # per-monitor v1: PROCESS_PER_MONITOR_DPI_AWARE = 2
    try:
        fn = ctypes.windll.shcore.SetProcessDpiAwareness
        fn.argtypes = [ctypes.c_int]
        fn.restype = ctypes.c_int
        if fn(2) == 0:            # S_OK；E_ACCESSDENIED 表示已设置过，继续回退
            _DPI_AWARENESS = "per-monitor"
            return _DPI_AWARENESS
    except Exception:
        pass
    # 系统级感知(单屏 DPI)，最老但最广泛可用
    try:
        ctypes.windll.user32.SetProcessDPIAware()
        _DPI_AWARENESS = "system"
        return _DPI_AWARENESS
    except Exception:
        _DPI_AWARENESS = "unavailable"
        return _DPI_AWARENESS
def is_dpi_aware():
    """进程是否成功声明 DPI 感知；决定弹窗截图的屏幕坐标是否要乘缩放系数。"""
    return _DPI_AWARENESS != "unavailable"
def _query_dpi(root=None):
    """读取当前窗口/屏幕 DPI，失败返回 96。优先窗口级(支持多显示器不同 DPI)。"""
    if sys.platform != "win32":
        return DPI_BASE
    try:
        user32 = ctypes.windll.user32
        if root is not None:
            try:
                hwnd = int(root.winfo_id())
                fn = user32.GetDpiForWindow
                fn.argtypes = [ctypes.c_void_p]
                fn.restype = ctypes.c_uint
                dpi = fn(ctypes.c_void_p(hwnd))
                if dpi and dpi > 0:
                    return float(dpi)
            except Exception:
                pass
        # 回退: 主屏 DC 的 LOGPIXELSX(88)
        user32.GetDC.argtypes = [ctypes.c_void_p]
        user32.GetDC.restype = ctypes.c_void_p
        user32.ReleaseDC.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
        hdc = user32.GetDC(None)
        try:
            gdi32 = ctypes.windll.gdi32
            gdi32.GetDeviceCaps.argtypes = [ctypes.c_void_p, ctypes.c_int]
            gdi32.GetDeviceCaps.restype = ctypes.c_int
            dpi = gdi32.GetDeviceCaps(hdc, 88)
        finally:
            user32.ReleaseDC(None, hdc)
        if dpi and dpi > 0:
            return float(dpi)
    except Exception:
        pass
    return DPI_BASE
def read_dpi_scale(root=None):
    """查询 DPI 并更新全局缩放系数；返回夹逼到 [MIN, MAX] 的系数。"""
    global _DPI_SCALE
    s = _query_dpi(root) / DPI_BASE
    _DPI_SCALE = max(DPI_SCALE_MIN, min(DPI_SCALE_MAX, s))
    return _DPI_SCALE
def dpi_scale():
    """当前 DPI 缩放系数(相对 100%)。"""
    return _DPI_SCALE
def px(value):
    """把 100% 设计稿的像素换算为当前 DPI 下的像素；最小 1 防止 0 尺寸控件。"""
    return max(1, int(round(float(value) * _DPI_SCALE)))
def pxp(*values):
    """元组版换算，保留 0（间距元组里 0 表示「无间距」，不能被放大成 1）。"""
    return tuple(max(0, int(round(float(v) * _DPI_SCALE))) for v in values)
def apply_tk_scaling(root, scale=None):
    """让 Tk 的「点 -> 像素」换算跟随 DPI，字体因此自动放大，无需改字号。
    标准 100% 下 tk scaling = 96/72 ≈ 1.333；只在与期望值不一致时才写入，
    避免覆盖 Tk 自身已完成的 DPI 初始化。返回生效后的 scaling 值。
    """
    s = _DPI_SCALE if scale is None else scale
    expected = DPI_BASE * s / 72.0
    try:
        current = float(root.tk.call("tk", "scaling"))
        if abs(current - expected) > 0.01:
            root.tk.call("tk", "scaling", expected)
            log("DPI 适配: tk scaling %.3f -> %.3f (缩放 %.2f)"
                % (current, expected, s))
        return float(root.tk.call("tk", "scaling"))
    except Exception:
        return expected
# ---------------------------------------------------------------------------
# 资源路径 / 窗口图标
# ---------------------------------------------------------------------------
def resource_path(rel):
    # 兼容 PyInstaller 打包后的 _MEIPASS 目录，未打包时回退到脚本所在目录
    base = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, rel)
def set_window_icon(win):
    # 打包后尝试用 exe 自身作为图标资源，未打包时用 app.ico；
    # 只调用一次 iconbitmap，且异常静默忽略，不保证大小图标都设置成功。
    ico = sys.executable if getattr(sys, "frozen", False) else resource_path("app.ico")
    try:
        win.iconbitmap(ico)          # Windows 下 Tk 支持传 .exe/.ico 路径
    except Exception:
        pass
# ---------------------------------------------------------------------------
# 日志
# ---------------------------------------------------------------------------
def log(msg):
    # 将日志消息追加写入日志文件，异常时静默忽略
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass
# ---------------------------------------------------------------------------
# 拖拽文件解析与过滤工具函数
# ---------------------------------------------------------------------------
def parse_dnd_file_paths(data):
    """解析 tkinterdnd2 event.data 字符串为文件路径列表。"""
    if not data:
        return []
    paths = []
    current = ""
    in_braces = False
    for ch in data:
        if ch == "{":
            in_braces = True
            current = ""
        elif ch == "}":
            if current:
                paths.append(current)
            current = ""
            in_braces = False
        elif ch.isspace() and not in_braces:
            if current:
                paths.append(current)
            current = ""
        else:
            current += ch
    if current:
        paths.append(current)
    return paths
def filter_image_paths(paths):
    """按 SUPPORTED_IMAGE_EXTENSIONS 的扩展名过滤，仅保留图片路径；不校验文件内容。"""
    result = []
    for p in paths:
        ext = os.path.splitext(p)[1].lower()
        if ext in SUPPORTED_IMAGE_EXTENSIONS:
            result.append(p)
    return result
# ---------------------------------------------------------------------------
# 图片缓存层
# ---------------------------------------------------------------------------
class ImageCache:
    """以 (路径, 目标尺寸, paper_ratio, allow_upscale) 为键缓存缩放后的 PIL.Image。
    采用 LRU：命中时移到末尾，超过 max_size 时淘汰最久未用项，
    避免反复添加/删除照片后解码结果无上限驻留内存。
    """
    def __init__(self, max_size=64):
        self._cache = OrderedDict()
        self.max_size = max(1, int(max_size))
    def get_thumbnail(self, path, size, allow_upscale=True):
        return self._get(path, size, to_rgb=True, allow_upscale=allow_upscale)
    def get_preview_image(self, path, size, paper_ratio=0.0):
        return self._get(path, size, to_rgb=True, paper_ratio=paper_ratio)
    def get_fitted_image(self, path, box, paper_ratio=0.0):
        return self._get(path, box, to_rgb=True, paper_ratio=paper_ratio)
    def invalidate(self):
        self._cache.clear()
    def _get(self, path, box, to_rgb, paper_ratio=0.0, allow_upscale=True):
        key = (path, tuple(box), paper_ratio, allow_upscale)
        if key in self._cache:
            self._cache.move_to_end(key)
            return self._cache[key]
        img = self._decode_and_fit(path, box, to_rgb, paper_ratio, allow_upscale)
        if img is not None:
            self._cache[key] = img
            while len(self._cache) > self.max_size:
                self._cache.popitem(last=False)
        return img
    def _decode_and_fit(self, path, box, to_rgb, paper_ratio=0.0, allow_upscale=True):
        from PIL import Image
        try:
            im = Image.open(path)
            im.draft("RGB", tuple(box))
            w, h = im.size
            bw, bh = box
            if w <= 0 or h <= 0:
                return None
            # 自动旋转: |图片宽高比-纸张宽高比|/纸张宽高比 超过 ROTATION_THRESHOLD
            # 时单向旋转 90 度（不比较旋转前后谁更贴合，极端比例旋转后可能更偏离）
            if paper_ratio > 0:
                img_ratio = w / h
                if abs(img_ratio - paper_ratio) / paper_ratio > ROTATION_THRESHOLD:
                    im = im.transpose(Image.ROTATE_90)
                    w, h = im.size
            scale = min(bw / w, bh / h)
            if not allow_upscale:
                scale = min(scale, 1.0)
            new_w = max(1, int(round(w * scale)))
            new_h = max(1, int(round(h * scale)))
            if new_w != w or new_h != h:
                try:
                    resample = Image.LANCZOS
                except AttributeError:
                    resample = Image.ANTIALIAS
                im = im.resize((new_w, new_h), resample)
            if to_rgb and im.mode != "RGB":
                im = im.convert("RGB")
            return im
        except Exception:
            return None
# ---------------------------------------------------------------------------
# 文本换行截断器
# ---------------------------------------------------------------------------
class TextWrapTruncator:
    @staticmethod
    def wrap(text, max_width, font, max_lines=3):
        if not text:
            return text
        all_lines = []
        current = ""
        for ch in text:
            candidate = current + ch
            if font.measure(candidate) <= max_width:
                current = candidate
            else:
                if current:
                    all_lines.append(current)
                    current = ch
                else:
                    all_lines.append(ch)
                    current = ""
        if current:
            all_lines.append(current)
        if len(all_lines) <= max_lines:
            return "\n".join(all_lines)
        result = all_lines[:max_lines]
        last = result[max_lines - 1]
        ellipsis = "..."
        while last and font.measure(last + ellipsis) > max_width:
            last = last[:-1]
        if not last:
            last = ellipsis
        else:
            last = last + ellipsis
        result[max_lines - 1] = last
        return "\n".join(result)
# ---------------------------------------------------------------------------
# 自定义照片列表组件
# ---------------------------------------------------------------------------
class PhotoListWidget:
    def __init__(self, parent, entries, on_select, on_delete=None, right_padding=RIGHT_LIST_PADDING):
        self.parent = parent
        self.entries = entries
        self.on_select = on_select
        self.on_delete = on_delete
        # 传入的是 100% 基准值，按当前 DPI 换算后供 _render/_create_list_item 使用
        self.right_padding = px(right_padding)
        self.selected = set()
        self._item_frames = []
        self._build()
    def _build(self):
        self.canvas = tk.Canvas(self.parent, bg="#FFFFFF", highlightthickness=0)
        # 「可视的滚动条」：原生 tk.Scrollbar 自带上下箭头与滑块，配浅灰凹槽+
        # 细边框，内容不足时也不隐藏、不置灰，用户始终能看到它。
        # ttk.Scrollbar 在 vista 主题下 disabled 时会淡到几乎看不见，故不用。
        # command 统一走 _yview，所有滚动入口都被顶边 clamp 约束。
        self.vsb = tk.Scrollbar(self.parent, orient=tk.VERTICAL,
                                command=self._yview,
                                width=px(16), bd=px(1), relief="sunken",
                                troughcolor="#F0F0F0",
                                activebackground="#C8DCF0",
                                elementborderwidth=px(1),
                                highlightthickness=0)
        self._scroll_pending = False
        self.interior = tk.Frame(self.canvas, bg="#FFFFFF")
        self.interior_id = self.canvas.create_window((0, 0), window=self.interior,
                                                      anchor="nw")
        self.interior.bind("<Configure>", self._on_interior_configure)
        # interior 内的条目用 pack 布局，不使用 grid 列配置；宽度改由 canvas 的
        # <Configure> 事件同步，保证 text_width/right_pad 按当前可视宽度计算。
        self.canvas.configure(yscrollcommand=self._vsb_set)
        # 【关键】滚动条必须先 pack：tk.Canvas 默认请求宽度约 378px，若先 pack
        # canvas，pack 会把滚动条压缩到 6px、几乎看不见 —— 这正是「没有可视
        # 滚动条」的根因。先让滚动条拿到完整 18px，再让 canvas 占剩余空间并
        # expand；滚动条常驻且不随内容增删隐藏，列表宽度不会抖动。
        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        # 滚轮覆盖 canvas、interior、条目、子控件以及滚动条本身，
        # 鼠标停在滚动条轨道上时同样能滚。
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.interior.bind("<MouseWheel>", self._on_mousewheel)
        self.vsb.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Delete>", self._on_delete_key)
        self.interior.bind("<Delete>", self._on_delete_key)
        self._render()
    def _vsb_set(self, first, last):
        # tk.Scrollbar 没有 state()：不再禁用/置灰，保证滚动条始终可见。
        # 只把 Tk 给的比例规范化到 [0,1] 后回填滑块；内容不足时 first=0,last=1，
        # 滑块占满轨道，且 _can_scroll() 会拦截所有滚动，拖不动但依然看得见。
        try:
            f = max(0.0, min(1.0, float(first)))
            l = max(f, min(1.0, float(last)))
        except (TypeError, ValueError):
            f, l = 0.0, 1.0
        self.vsb.set(f, l)
    def _yview(self, *args):
        """所有滚动的唯一入口（滚轮、滚动条拖动/箭头、代码调用都经过这里）。
        顶边不允许出现负偏移，内容不足时强制回到顶部 —— 第一张图片永远贴在
        列表顶部，不会被「向前」滚动推出空白，也不会离开列表。
        """
        result = self.canvas.yview(*args)
        if not self._can_scroll() or self.canvas.yview()[0] < 0:
            self.canvas.yview_moveto(0)
        return result
    def _can_scroll(self):
        """滚动区域是否真的高于可视区；内容不足时返回 False，禁止一切滚动。"""
        try:
            sr = str(self.canvas.cget("scrollregion")).split()
            if len(sr) < 4:
                return False
            total = float(sr[3]) - float(sr[1])
            vh = self.canvas.winfo_height()
            return vh > 0 and total > vh + 1
        except Exception:
            return False
    def _on_interior_configure(self, event):
        self._update_scrollregion()
    def _update_scrollregion(self, viewport_height=None):
        """按规范化区域立即刷新，并排一次 idle 刷新消除几何布局滞后。"""
        if not self.canvas.winfo_exists():
            return
        self._apply_scrollregion(viewport_height)
        # <Configure> 回调里 Tk 几何可能尚未完成；idle 后按最终布局再刷一次。
        # 用标志防抖，避免连续 resize/增删时堆积大量重复回调。
        if not self._scroll_pending:
            self._scroll_pending = True
            self.canvas.after_idle(self._idle_scrollregion)
    def _idle_scrollregion(self):
        self._scroll_pending = False
        self._apply_scrollregion()
    def _apply_scrollregion(self, viewport_height=None):
        """规范化 scrollregion：顶边 0、宽度=canvas 可视宽、高度=max(内容, 视口)。
        - 顶边固定 0：第一张图片永远贴顶，不可能被向上滚出顶部空白；
        - 宽度固定为可视宽度：长文件名 Label 的请求宽度曾让 bbox("all")
          横向撑到 900+px，这里不再采用 bbox 的横向值；
        - 高度取 max(内容, 视口)：内容不足时区域恰好一屏，完全滚不动；
          内容溢出时只能向下滚，且不会 overscroll。
        """
        try:
            canvas = self.canvas
            if not canvas.winfo_exists():
                return
            vw = max(1, canvas.winfo_width())
            vh = max(1, canvas.winfo_height())
            if viewport_height and viewport_height > vh:
                vh = int(viewport_height)
            bbox = canvas.bbox("all")
            content_h = float(bbox[3]) - float(bbox[1]) if bbox else 0.0
            canvas.configure(scrollregion=(0, 0, vw, int(max(content_h, vh))))
            # 内容不足或删除后残留偏移，一律拉回顶部
            if not self._can_scroll():
                canvas.yview_moveto(0)
        except Exception:
            pass
    def _on_canvas_configure(self, event):
        # 让 interior 跟随 canvas 宽度，并按新宽度重排文字区域
        try:
            self.canvas.itemconfigure(self.interior_id, width=event.width)
        except Exception:
            pass
        # 传入真实视口高度，避免 <Configure> 阶段 winfo_height() 还是旧值时
        # 把 scrollregion 设得比视口小/大，造成短暂「可向前滚动」的窗口。
        self._render(event.width, event.height)
    def refresh(self, entries):
        self.entries = entries
        self.selected.clear()
        self._render()
        # 列表重建后回到顶部：第一张永远从列表顶部开始，不沿用旧的滚动位置。
        # _on_canvas_configure 走的是 _render 而非 refresh，故 resize 不会打断
        # 用户当前的滚动位置。
        self.canvas.yview_moveto(0)
    def _on_mousewheel(self, event):
        if self.canvas.winfo_exists():
            # 内容较少时完全不响应滚轮：这是「只能向后滚动」的边界 ——
            # 没有可向后看的内容，就既不向前也不向后滚，第一张固定在顶部。
            if not self._can_scroll():
                self.canvas.yview_moveto(0)
                return "break"
            delta = getattr(event, "delta", 0)
            # Windows 常规鼠标 delta 为 ±120 的倍数；高分辨率触控板/鼠标可能
            # 给出更小的值，此时按符号至少滚 1 格，避免「滚轮完全没反应」。
            if abs(delta) >= 120:
                steps = int(-delta / 120)
            else:
                steps = -1 if delta > 0 else (1 if delta < 0 else 0)
            if steps:
                self._yview("scroll", steps, "units")
        return "break"
    def _on_delete_key(self, event):
        if self.on_delete:
            self.on_delete()
        return "break"
    def set_selection(self, indices):
        valid = set(i for i in indices if 0 <= i < len(self.entries))
        self.selected = valid
        self._update_visual()
        # 新增/批量选中后把第一张选中项滚入视野：否则一次拖入几十张照片时
        # 选中项落在列表底部，视图还停在顶部，看起来像「列表没有滚动条」。
        if valid:
            self.ensure_visible(min(valid))
        self._fire_callback()
    def ensure_visible(self, index):
        """把指定索引的条目滚动到可视区内；不改变选中状态。"""
        if index < 0 or index >= len(self._item_frames):
            return
        try:
            self.canvas.update_idletasks()
            sr = str(self.canvas.cget("scrollregion")).split()
            if len(sr) < 4:
                return
            total_h = float(sr[3]) - float(sr[1])
            vh = self.canvas.winfo_height()
            if total_h <= 0 or vh <= 0 or total_h <= vh:
                return
            item = self._item_frames[index]
            top = item.winfo_y()
            bottom = top + item.winfo_height()
            first = float(self.canvas.yview()[0])
            view_top = first * total_h
            view_bottom = view_top + vh
            if top < view_top:
                frac = top / total_h
            elif bottom > view_bottom:
                frac = (bottom - vh) / total_h
            else:
                return
            # 夹逼到 [0, 1 - 视口/内容]，再由 _yview 做顶边保护；
            # 避免 yview_moveto(1.0) 这类越界值让内容整体上移露出空白。
            max_frac = max(0.0, 1.0 - vh / total_h)
            self._yview("moveto", max(0.0, min(max_frac, frac)))
        except Exception:
            pass
    def _create_list_item(self, i, name, thumb, text_width, right_pad, font):
        display_name = "%d. %s" % (i + 1, name)
        wrapped = TextWrapTruncator.wrap(display_name, text_width, font, 3)
        item = tk.Frame(self.interior, bg=UNSELECTED_BG_COLOR, height=px(60))
        # 上边距归零：第一张图片严格贴在列表顶边（winfo_y()==0），
        # 不会因 pack 的 pady 产生 1px 位移，也便于 ensure_visible 精确定位；
        # pxp 保留 0，缩放后仍是 0。
        item.pack(fill=tk.X, pady=pxp(0, 1))
        item.pack_propagate(False)
        item.grid_columnconfigure(0, minsize=px(LIST_IMAGE_MAX_SIZE))
        item.grid_columnconfigure(1, weight=1)
        item.grid_rowconfigure(0, minsize=px(LIST_IMAGE_MAX_SIZE))
        thumb_area = tk.Frame(item, bg=UNSELECTED_BG_COLOR,
                              width=px(LIST_IMAGE_MAX_SIZE),
                              height=px(LIST_IMAGE_MAX_SIZE))
        thumb_area.pack_propagate(False)
        thumb_area.grid(row=0, column=0, sticky="nsew",
                        padx=pxp(0, THUMBNAIL_TEXT_GAP))
        if thumb is not None:
            tl = tk.Label(thumb_area, image=thumb, bg=UNSELECTED_BG_COLOR)
        else:
            tl = tk.Label(thumb_area, text="", bg=UNSELECTED_BG_COLOR)
        tl.pack(expand=True)
        # wrap() 已按 max_width 插入 \n 并截断到 3 行；不再设 wraplength，
        # 避免 Tk 二次换行突破截断行数。text_width 只用于 wrap() 计算。
        nl = tk.Label(item, text=wrapped, bg=UNSELECTED_BG_COLOR, fg=FILE_NAME_COLOR,
                      justify=tk.LEFT)
        nl.grid(row=0, column=1, sticky="nsew", padx=(0, right_pad))
        item._thumb_label = tl
        item._name_label = nl
        item._thumb_area = thumb_area
        item.bind("<Button-1>", lambda e, idx=i: self._on_click(idx, e))
        item.bind("<Shift-Button-1>", lambda e, idx=i: self._on_shift_click(idx, e))
        item.bind("<Control-Button-1>", lambda e, idx=i: self._on_ctrl_click(idx, e))
        item.bind("<MouseWheel>", self._on_mousewheel)
        item.bind("<Delete>", self._on_delete_key)
        for child in list(item.winfo_children()) + [item._thumb_label]:
            child.bind("<Button-1>", lambda e, idx=i: self._on_click(idx, e))
            child.bind("<Shift-Button-1>", lambda e, idx=i: self._on_shift_click(idx, e))
            child.bind("<Control-Button-1>", lambda e, idx=i: self._on_ctrl_click(idx, e))
            child.bind("<MouseWheel>", self._on_mousewheel)
            child.bind("<Delete>", self._on_delete_key)
        return item
    def _render(self, container_width=None, viewport_height=None):
        existing = list(self._item_frames)
        self._item_frames = []
        if not self.entries:
            for item in existing:
                item.destroy()
            # 空列表也走规范化区域：高度=视口，禁止任何滚动，第一张不会离位
            self._update_scrollregion(viewport_height)
            self.canvas.yview_moveto(0)
            return
        container_w = container_width or self.interior.winfo_width()
        if container_w <= 4:
            container_w = px(210)
        # 滚动条与 canvas 平级、位于其右侧，滚动条左边缘即 canvas 右边界，
        # 因此文字右侧只需留 RIGHT_LIST_PADDING，不再叠加滚动条宽度。
        # right_pad 已在 __init__ 换算；列宽与文字最小宽度按 DPI 换算，
        # font.measure() 用的是随 tk scaling 放大的字体，单位一致。
        right_pad = self.right_padding
        text_width = container_w - px(THUMBNAIL_COLUMN_WIDTH) - right_pad
        if text_width < px(10):
            text_width = px(10)
        import tkinter.font as tkfont
        font = tkfont.nametofont("TkDefaultFont")
        n = len(self.entries)
        try:
            for i in range(n):
                name, thumb = self.entries[i]
                if i < len(existing):
                    item = existing[i]
                    display_name = "%d. %s" % (i + 1, name)
                    wrapped = TextWrapTruncator.wrap(display_name, text_width, font, 3)
                    item._thumb_label.config(image=thumb or "", text="", bg=UNSELECTED_BG_COLOR)
                    item._name_label.config(text=wrapped, bg=UNSELECTED_BG_COLOR,
                                           fg=FILE_NAME_COLOR)
                    item._name_label.grid_configure(padx=(0, right_pad))
                    item._thumb_area.configure(bg=UNSELECTED_BG_COLOR)
                    item._thumb_area.grid_configure(
                        padx=pxp(0, THUMBNAIL_TEXT_GAP))
                    item.configure(bg=UNSELECTED_BG_COLOR)
                else:
                    item = self._create_list_item(i, name, thumb, text_width, right_pad, font)
                self._item_frames.append(item)
            for item in existing[n:]:
                item.destroy()
            log("刷新优化: 复用=%d 新建=%d 销毁=%d" % (min(n, len(existing)), max(0, n - len(existing)), max(0, len(existing) - n)))
        except Exception as e:
            log("刷新优化异常: %s\n%s" % (e, traceback.format_exc()))
            for w in self.interior.winfo_children():
                w.destroy()
            self._item_frames = []
            for i in range(n):
                name, thumb = self.entries[i]
                item = self._create_list_item(i, name, thumb, text_width, right_pad, font)
                self._item_frames.append(item)
            log("刷新优化: 全部重建=%d" % n)
        self._update_scrollregion(viewport_height)
        log("列表刷新: 条目数=%d" % n)
    def rebuild_for_dpi(self):
        """DPI 变化后按新缩放重建几何：间距、滚动条宽度、条目尺寸全部更新。
        复用分支只改文字，不会重建条目的固定宽高，因此必须销毁旧条目后
        全量 _render()，否则高 DPI 下缩略图列仍是旧尺寸。
        """
        self.right_padding = px(RIGHT_LIST_PADDING)
        try:
            self.vsb.configure(width=px(16), bd=px(1),
                               elementborderwidth=px(1))
        except Exception:
            pass
        for item in self._item_frames:
            try:
                item.destroy()
            except Exception:
                pass
        self._item_frames = []
        self._render()
        # 复用分支把条目背景统一重置为未选色；窗口 resize 也会走 _render，
        # 必须按 selected 重新着色，否则高亮会凭空消失。
        self._update_visual()
    def _on_click(self, idx, event):
        self.selected = {idx}
        self._update_visual()
        self._fire_callback()
    def _on_shift_click(self, idx, event):
        # 落在已选范围内时重置为单选（非连续收缩），与常见 Shift 连选行为不同
        if not self.selected:
            self.selected = {idx}
        else:
            lo = min(self.selected)
            hi = max(self.selected)
            if idx < lo:
                self.selected = set(range(idx, hi + 1))
            elif idx > hi:
                self.selected = set(range(lo, idx + 1))
            else:
                self.selected = {idx}
        self._update_visual()
        self._fire_callback()
    def _on_ctrl_click(self, idx, event):
        if idx in self.selected:
            self.selected.discard(idx)
        else:
            self.selected.add(idx)
        self._update_visual()
        self._fire_callback()
    def _fire_callback(self):
        if self.on_select:
            self.on_select(None)
        log("选中操作: 索引=%s" % sorted(self.selected))
    def _update_visual(self):
        for i, item in enumerate(self._item_frames):
            bg = SELECTED_BG_COLOR if i in self.selected else UNSELECTED_BG_COLOR
            item.configure(bg=bg)
            try:
                item._thumb_area.configure(bg=bg)
            except Exception:
                pass
            try:
                item._thumb_label.configure(bg=bg)
            except Exception:
                pass
            try:
                item._name_label.configure(bg=bg)
            except Exception:
                pass
    def get_selected_indices(self):
        return sorted(self.selected)
    def clear_selection(self):
        self.selected.clear()
        self._update_visual()
# ---------------------------------------------------------------------------
# 预览翻页控制器与翻页按钮
# ---------------------------------------------------------------------------
class PaginationController:
    """预览分页状态，数据源为「全部已添加照片」，与列表选中集合解耦。
    total_pages 始终 = ceil(照片总数 / 每页张数)，每页张数来自布局；
    选中/取消选中照片不会改变总页数，只会更新 current_page（跳到选中照片所在页）。
    next()/prev() 返回「翻页后是否还能继续翻」，调用方忽略该值；
    按钮可用性统一由 PaginationBar.update() 按 current_page/total_pages 决定。
    """
    def __init__(self, photo_paths, per_page):
        self.photo_paths = list(photo_paths)
        self.per_page = per_page
        self.current_page = 1
    @property
    def total_pages(self):
        # 只由「布局每页张数 + 照片总数」决定，不看选中集合
        if self.per_page <= 0 or not self.photo_paths:
            return 0
        import math
        return math.ceil(len(self.photo_paths) / self.per_page)
    def page_of_index(self, index):
        """照片在全部列表中的索引 -> 所在页码(1 起)；越界或每页张数非法时返回 1。"""
        if self.per_page <= 0 or index < 0 or index >= len(self.photo_paths):
            return 1
        return min(index // self.per_page + 1, self.total_pages or 1)
    def sync_paths(self, photo_paths):
        """照片增删后同步数据源，并把当前页码夹逼到有效范围。
        总页数仍由同步后的照片总数 + 布局推导，口径不变。
        """
        self.photo_paths = list(photo_paths)
        self.current_page = max(1, min(self.current_page, self.total_pages or 1))
    def get_current_page_photos(self):
        if self.per_page <= 0:
            return []
        start = (self.current_page - 1) * self.per_page
        return self.photo_paths[start:start + self.per_page]
    def next(self):
        if self.current_page < self.total_pages:
            self.current_page += 1
        return self.current_page < self.total_pages
    def prev(self):
        if self.current_page > 1:
            self.current_page -= 1
        return self.current_page > 1
    def reset(self):
        self.current_page = 1
    @property
    def prev_enabled(self):
        """供外部查询；当前 UI 未使用，由 PaginationBar.update() 自行判断。"""
        return self.current_page > 1
    @property
    def next_enabled(self):
        """供外部查询；当前 UI 未使用，由 PaginationBar.update() 自行判断。"""
        return self.current_page < self.total_pages
class PaginationBar:
    def __init__(self, parent, on_prev, on_next):
        self.parent = parent
        self.on_prev = on_prev
        self.on_next = on_next
        self._buttons_enabled = True
        self._cur_page = 1
        self._total_pages = 0
        self._frame = tk.Frame(parent, bg="#FFFFFF")
        self._frame.pack(side=tk.TOP, fill=tk.X, pady=pxp(2, 0))
        self._frame.pack_propagate(False)
        # 容器高度随按钮放大同步调高；pack_propagate(False) 下 pack 不会拉伸按钮，
        # 按钮高度仍等于 TButton 样式的自然高度，与工具栏按钮一致。
        # 高度按 DPI 换算，才能容纳放大后的按钮。
        self._frame.configure(height=px(PAGINATION_BAR_HEIGHT))
        # 与「添加照片」等工具栏按钮共用 TButton 样式：不传控件级 width，
        # 宽度由 style.configure("TButton", width=BUTTON_WIDTH) 统一给出。
        # 快捷键提示放回按钮文字：BUTTON_WIDTH 已放大到 16 字符，可完整容纳
        # 「上一页 (PgUp)」/「下一页 (PgDn)」；root 上的 <Prior>/<Next> 绑定不变。
        self.btn_prev = ttk.Button(self._frame, text="上一页 (PgUp)",
                                   command=self.on_prev)
        self.btn_prev.pack(side=tk.LEFT, padx=pxp(0, 6))
        self._label = tk.Label(self._frame, bg="#FFFFFF", fg="#666666",
                               text="", font=("Segoe UI", 9))
        self._label.pack(side=tk.LEFT)
        self.btn_next = ttk.Button(self._frame, text="下一页 (PgDn)",
                                   command=self.on_next)
        self.btn_next.pack(side=tk.LEFT, padx=pxp(6, 0))
    def update(self, current_page, total_pages):
        self._frame.pack(side=tk.TOP, fill=tk.X, pady=pxp(2, 0))
        # 快捷键已在按钮文字中，标签只显示页码，避免两处重复
        self._label.config(text="第%d页/共%d页" % (current_page, total_pages))
        self._cur_page = current_page
        self._total_pages = total_pages
        self._apply_state()
    def set_enabled(self, enabled):
        """整体开关；只改标志，按钮状态统一由 _apply_state() 计算。"""
        self._buttons_enabled = bool(enabled)
        self._apply_state()
    def _apply_state(self):
        """按钮可用性的唯一计算入口：整体开关 + 首/末页边界，避免两处各算一套。"""
        prev = "normal" if self._buttons_enabled and self._cur_page > 1 else "disabled"
        nxt = "normal" if self._buttons_enabled and self._cur_page < self._total_pages else "disabled"
        self.btn_prev.config(state=prev)
        self.btn_next.config(state=nxt)
# ---------------------------------------------------------------------------
# 拖拽添加控制器
# ---------------------------------------------------------------------------
class DragDropManager:
    def __init__(self, app):
        self.app = app
        self.root = app.root
        self.is_drag_active = False
        # 以下两个状态当前只写入、未被读取，保留作调试与后续功能扩展
        self._dnd_registered = False
        self._highlighted = False
    def register(self):
        try:
            from tkinterdnd2 import DND_FILES
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<DragEnter>>", self.on_drag_enter)
            self.root.dnd_bind("<<DragOver>>", self.on_drag_over)
            self.root.dnd_bind("<<DragLeave>>", self.on_drag_leave)
            self.root.dnd_bind("<<Drop>>", self.on_drop)
            self._dnd_registered = True
        except Exception:
            log("拖拽注册降级: 不可用不影响既有功能")
    def detect_target_type(self, event):
        """按坐标返回「界面」/「任务栏」；仅用于日志分类，调用方对两者一视同仁。"""
        try:
            x = getattr(event, "x_root", 0)
            y = getattr(event, "y_root", 0)
            wx = self.root.winfo_rootx()
            wy = self.root.winfo_rooty()
            ww = self.root.winfo_width()
            wh = self.root.winfo_height()
            if wx <= x <= wx + ww and wy <= y <= wy + wh:
                return "界面"
            return "任务栏"
        except Exception:
            return "界面"
    def set_drop_effect(self, event, accepted):
        """仅记录期望动作；是否真正拒绝由拖放层决定，回调未 return action。"""
        try:
            event.action = "copy" if accepted else "refuse"
        except Exception:
            pass
    def apply_accept_feedback(self):
        try:
            self.root.configure(highlightbackground="#3399FF",
                                highlightcolor="#3399FF",
                                highlightthickness=3)
            self._highlighted = True
        except Exception:
            pass
    def apply_reject_feedback(self):
        try:
            self.root.configure(highlightbackground="#CCCCCC",
                                highlightcolor="#CCCCCC",
                                highlightthickness=2)
            self._highlighted = True
        except Exception:
            pass
    def clear_feedback(self):
        try:
            self.root.configure(highlightthickness=0)
        except Exception:
            pass
        self._highlighted = False
    def on_drag_enter(self, event):
        try:
            paths = parse_dnd_file_paths(getattr(event, "data", ""))
            images = filter_image_paths(paths)
            self.set_drop_effect(event, len(images) > 0)
            # 拖放时会对非图片做过滤后添加图片，混合拖放实际是「接受」，
            # 因此只要含图片就显示接受反馈，避免视觉拒绝却实际添加。
            if len(images) > 0:
                self.apply_accept_feedback()
            else:
                self.apply_reject_feedback()
            self.is_drag_active = True
        except Exception:
            self.apply_reject_feedback()
            self.is_drag_active = True
    def on_drag_over(self, event):
        try:
            self.is_drag_active = True
        except Exception:
            pass
    def on_drag_leave(self, event):
        try:
            self.clear_feedback()
        except Exception:
            pass
        self.is_drag_active = False
    def on_drop(self, event):
        try:
            paths = parse_dnd_file_paths(getattr(event, "data", ""))
        except Exception:
            paths = []
        try:
            self.clear_feedback()
        except Exception:
            pass
        self.is_drag_active = False
        if paths:
            target = self.detect_target_type(event)
            self.app._handle_dropped_files(paths, target)
    def cancel_drag(self):
        try:
            self.clear_feedback()
        except Exception:
            pass
        self.is_drag_active = False
# ---------------------------------------------------------------------------
# 性能测量层
# ---------------------------------------------------------------------------
class PerfMeasure:
    """集中化耗时测量与日志写入，测量异常静默降级不影响业务。"""
    def measure(self, op_name, limit, func, *args, **kwargs):
        t0 = time.time()
        try:
            result = func(*args, **kwargs)
        except Exception:
            t1 = time.time()
            self._try_log(op_name, t1 - t0, limit, True)
            raise
        t1 = time.time()
        elapsed = t1 - t0
        self._try_log(op_name, elapsed, limit, elapsed > limit)
        return result
    def _try_log(self, op_name, elapsed, limit, exceeded):
        try:
            self._log(op_name, elapsed, limit, exceeded)
        except Exception:
            pass
    def _log(self, op_name, elapsed, limit, exceeded):
        flag = "超标" if exceeded else "达标"
        fmt = "%%.%df" % TIME_LOG_DIGITS
        log("%s: %s 秒 (上限 %s 秒) %s" % (op_name, fmt % elapsed, fmt % limit, flag))
# ---------------------------------------------------------------------------
# 打印机 / DEVMODE 操作
# ---------------------------------------------------------------------------
def list_printers():
    """返回可见打印机名称列表(本地 + 已连接网络打印机，默认打印机排在第一位)。"""
    import win32print
    names = []
    try:
        # 注释承诺「所有打印机」：除本地外补上已连接的网络打印机
        #（老版本 pywin32 无该常量时退化为原来的本地枚举）。
        flags = win32print.PRINTER_ENUM_NAME | win32print.PRINTER_ENUM_LOCAL
        flags |= getattr(win32print, "PRINTER_ENUM_CONNECTIONS", 0)
        for info in win32print.EnumPrinters(flags, None, 4):
            names.append(info["pPrinterName"])
    except Exception:
        pass
    if not names:
        try:
            names.append(win32print.GetDefaultPrinter())
        except Exception:
            pass
    # 去重并保持顺序，默认打印机置顶
    try:
        default = win32print.GetDefaultPrinter()
    except Exception:
        default = None
    seen = set()
    ordered = []
    if default and default in names:
        ordered.append(default)
        seen.add(default)
    for n in names:
        if n not in seen:
            ordered.append(n)
            seen.add(n)
    return ordered
def _get_devmode(printer_name):
    """打开打印机并返回 (hprinter, info, devmode)。"""
    import win32print
    hprinter = win32print.OpenPrinter(printer_name)
    info = win32print.GetPrinter(hprinter, 2)
    devmode = info.get("pDevMode")
    return hprinter, info, devmode
def apply_print_settings(printer_name, duplex_mode, paper_key):
    """将双面/纸张设置写入打印机 DEVMODE，返回可恢复的 5 个原始字段字典。
    devmode 缺失或写入异常时返回 None（调用方仍会继续打印，但之后无法恢复）；
    只保存 Duplex/Fields/PaperSize/PaperWidth/PaperLength，不是整个 DEVMODE；
    custom 纸张只写 PaperWidth/PaperLength，不设 DMPAPER_USER。
    """
    import win32print
    DM_DUPLEX = getattr(win32print, "DM_DUPLEX", 0x00001000)
    DM_PAPERSIZE = getattr(win32print, "DM_PAPERSIZE", 0x00000002)
    DM_PAPERWIDTH = getattr(win32print, "DM_PAPERWIDTH", 0x00000004)
    DM_PAPERLENGTH = getattr(win32print, "DM_PAPERLENGTH", 0x00000008)
    original = None
    hprinter = None
    try:
        hprinter, info, devmode = _get_devmode(printer_name)
        if devmode is None:
            return None
        # 保存原始 DEVMODE 字段以便打印后恢复
        original = {
            "Duplex": int(devmode.Duplex),
            "Fields": int(devmode.Fields),
            "PaperSize": int(getattr(devmode, "PaperSize", 0)),
            "PaperWidth": int(getattr(devmode, "PaperWidth", 0)),
            "PaperLength": int(getattr(devmode, "PaperLength", 0)),
        }
        # 双面
        devmode.Duplex = int(duplex_mode)
        devmode.Fields = int(devmode.Fields) | DM_DUPLEX
        # 纸张
        if paper_key:
            spec = PAPER_SIZES.get(paper_key)
            if spec is not None:
                if spec[0] == "size":
                    devmode.PaperSize = int(spec[1])
                    devmode.Fields = int(devmode.Fields) | DM_PAPERSIZE
                elif spec[0] == "custom":
                    devmode.PaperWidth = int(spec[1][0])
                    devmode.PaperLength = int(spec[1][1])
                    devmode.Fields = int(devmode.Fields) | DM_PAPERWIDTH | DM_PAPERLENGTH
        win32print.SetPrinter(hprinter, 2, info, 0)
    except Exception as e:
        log("apply_print_settings 失败: %s\n%s" % (e, traceback.format_exc()))
        original = None
    finally:
        if hprinter:
            try:
                win32print.ClosePrinter(hprinter)
            except Exception:
                pass
    return original
def restore_print_settings(printer_name, original):
    """按 apply_print_settings 保存的 5 个字段恢复 DEVMODE；original 为 None 时跳过。"""
    import win32print
    if not original:
        return
    hprinter = None
    try:
        hprinter, info, devmode = _get_devmode(printer_name)
        if devmode is None:
            return
        # 逐字段恢复原始 DEVMODE 值
        devmode.Duplex = int(original["Duplex"])
        devmode.Fields = int(original["Fields"])
        devmode.PaperSize = int(original["PaperSize"])
        devmode.PaperWidth = int(original["PaperWidth"])
        devmode.PaperLength = int(original["PaperLength"])
        win32print.SetPrinter(hprinter, 2, info, 0)
    except Exception as e:
        log("restore_print_settings 失败: %s" % e)
    finally:
        if hprinter:
            try:
                win32print.ClosePrinter(hprinter)
            except Exception:
                pass
# 模块级缓存与测量实例；页面合成图最大，单独给最小容量
_page_cache = ImageCache(max_size=IMAGE_CACHE_MAX_PAGE)
_perf = PerfMeasure()
# ---------------------------------------------------------------------------
# 图像 / 页面合成
# ---------------------------------------------------------------------------
def fit_image(img, box_w, box_h, paper_ratio=0.0):
    """将图像等比缩放放入 box_w x box_h 内(RGB)，按 paper_ratio 自动旋转。
    当前打印/预览均走 ImageCache._decode_and_fit，本函数保留供外部调用，模块内未使用。
    """
    from PIL import Image
    if img.mode != "RGB":
        img = img.convert("RGB")
    w, h = img.size
    if w <= 0 or h <= 0:
        return img
    # 自动旋转: |图片宽高比-纸张宽高比|/纸张宽高比 超过 ROTATION_THRESHOLD
    # 时单向旋转 90 度（不比较旋转前后谁更贴合，极端比例旋转后可能更偏离）
    if paper_ratio > 0:
        img_ratio = w / h
        if abs(img_ratio - paper_ratio) / paper_ratio > ROTATION_THRESHOLD:
            img = img.transpose(Image.ROTATE_90)
            w, h = img.size
    # 等比缩放，取宽高比例中较小者确保完整放入
    scale = min(box_w / w, box_h / h)
    new_w = max(1, int(round(w * scale)))
    new_h = max(1, int(round(h * scale)))
    if new_w != w or new_h != h:
        try:
            resample = Image.LANCZOS
        except AttributeError:
            resample = Image.ANTIALIAS
        img = img.resize((new_w, new_h), resample)
    return img
def compose_page(image_paths_batch, layout, printable_w, printable_h, paper_ratio=0.0):
    """把一批照片合成一页图像(白底)；旋转由缓存层按 paper_ratio 完成，本层只拼版。"""
    return _perf.measure("页面合成", PAGE_COMPOSE_TIME_LIMIT,
                         _compose_page_impl, image_paths_batch, layout, printable_w, printable_h, paper_ratio)
def _compose_page_impl(image_paths_batch, layout, printable_w, printable_h, paper_ratio=0.0):
    """页面合成实际逻辑：取缓存中已旋转/缩放的图，按布局居中粘贴到白底页上。"""
    from PIL import Image
    page = Image.new("RGB", (printable_w, printable_h), "white")
    rows, cols, _ = LAYOUTS[layout]
    gap = max(2, int(min(printable_w, printable_h) * 0.008))
    cell_w = (printable_w - gap * (cols + 1)) // cols
    cell_h = (printable_h - gap * (rows + 1)) // rows
    for idx, path in enumerate(image_paths_batch):
        fitted = _page_cache.get_fitted_image(path, (cell_w, cell_h), paper_ratio)
        if fitted is None:
            log("无法合成图像 %s" % path)
            continue
        fw, fh = fitted.size
        r = idx // cols
        c = idx % cols
        x = gap + c * (cell_w + gap) + (cell_w - fw) // 2
        y = gap + r * (cell_h + gap) + (cell_h - fh) // 2
        page.paste(fitted, (x, y))
    return page
def build_pages(image_paths, layout):
    """按布局把照片分组为每页一批。"""
    per_page = LAYOUTS[layout][2]
    pages = []
    for i in range(0, len(image_paths), per_page):
        pages.append(image_paths[i:i + per_page])
    return pages
# ---------------------------------------------------------------------------
# 打印
# ---------------------------------------------------------------------------
def do_print(image_paths, printer_name, layout, copies, duplex_mode, paper_key,
             progress_cb=None):
    """执行打印。progress_cb(current, total) 回调进度（失败页也计入进度）。
    返回 (ok, msg)：任一一页合成/绘制失败即 ok=False，并说明失败页数。
    无论成功、单页失败还是整体异常，finally 都会释放 DC 并恢复打印机设置。
    """
    from PIL import ImageWin
    import win32ui, win32con
    if not image_paths:
        return False, "没有可打印的照片"
    if not printer_name:
        return False, "未选择打印机"
    # 解析纸张宽高比，用于自动旋转判定
    paper_dims = PAPER_DIMENSIONS.get(paper_key)
    paper_ratio = (paper_dims[0] / paper_dims[1]) if paper_dims else 0.0
    # 应用双面/纸张设置并保存原始值以便恢复；返回 None 表示无法恢复，
    # 此时仍按打印机当前设置继续打印，不中断任务。
    # original / dc 先初始化，保证 finally 在任何异常路径都能安全释放。
    original = None
    dc = None
    try:
        original = apply_print_settings(printer_name, duplex_mode, paper_key)
        dc = win32ui.CreateDC()
        dc.CreatePrinterDC(printer_name)
        # 获取可打印区域像素尺寸；返回非正值时回退物理纸张尺寸(含不可打印边距)
        printable_w = dc.GetDeviceCaps(win32con.HORZRES)
        printable_h = dc.GetDeviceCaps(win32con.VERTRES)
        if printable_w <= 0 or printable_h <= 0:
            printable_w = dc.GetDeviceCaps(win32con.PHYSICALWIDTH)
            printable_h = dc.GetDeviceCaps(win32con.PHYSICALHEIGHT)
        pages = build_pages(image_paths, layout)
        total_pages = len(pages) * copies
        done = 0
        failed_pages = 0
        doc_name = "照片打印 - %d张" % len(image_paths)
        dc.StartDoc(doc_name)
        try:
            for copy in range(copies):
                for batch in pages:
                    try:
                        page_img = compose_page(batch, layout, printable_w, printable_h, paper_ratio)
                        dib = ImageWin.Dib(page_img)
                        dc.StartPage()
                        try:
                            dib.draw(dc.GetHandleOutput(),
                                     (0, 0, printable_w, printable_h))
                        finally:
                            dc.EndPage()
                    except Exception as e:
                        failed_pages += 1
                        log("打印页失败: %s" % e)
                    done += 1
                    if progress_cb:
                        progress_cb(done, total_pages)
            dc.EndDoc()
        except Exception:
            try:
                dc.AbortDoc()
            except Exception:
                pass
            raise
    except Exception as e:
        log("打印异常: %s\n%s" % (e, traceback.format_exc()))
        return False, "打印失败: %s" % e
    finally:
        # 先释放打印机 DC，再恢复 DEVMODE：DC 可能持有设置快照，顺序反了会
        # 让本次打印残留 HDC/GDI 句柄，异常时还可能卡住打印队列。
        if dc is not None:
            try:
                dc.DeleteDC()
            except Exception:
                log("释放打印机 DC 失败: %s" % traceback.format_exc())
        restore_print_settings(printer_name, original)
    if failed_pages:
        return False, "打印结束但有 %d 页失败，共 %d 页" % (failed_pages, total_pages)
    return True, "打印完成，共 %d 页" % total_pages
# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------
def make_gradient_image(width, height):
    """生成 macOS 风格蓝色按钮渐变 PhotoImage（纵向，上亮下深）+ 白色文字。
    配色取自 macOS 强调蓝 systemBlue：顶部高光 #8CC8FF -> #4AA8FF ->
    中部 #0A84FF -> 底部 #0062CC -> 暗边 #0050B4，顶部 1px 高光与底部 1px
    暗边模拟 Aqua 按钮的立体感。白色文字由调用方 tk.Button(fg="#FFFFFF") 提供。
    抽成模块级函数，让「创建按钮」与「运行时按 ttk 按钮尺寸重建」共用同一套配色，
    打印 / 删除确定 / 清空确定三个按钮因此完全一致。
    """
    w = max(1, int(width))
    h = max(1, int(height))
    img = tk.PhotoImage(width=w, height=h)
    stops = [(74, 168, 255),    # #4AA8FF 上部亮蓝
             (10, 132, 255),    # #0A84FF macOS systemBlue
             (0, 98, 204)]      # #0062CC 下部深蓝
    highlight = (140, 200, 255)  # #8CC8FF 顶部高光
    edge = (0, 80, 180)          # #0050B4 底部暗边
    def lerp(a, b, t):
        return tuple(int(round(a[i] + (b[i] - a[i]) * t)) for i in range(3))
    for y in range(h):
        t = y / (h - 1) if h > 1 else 0.5
        if t <= 0.5:
            rgb = lerp(stops[0], stops[1], t / 0.5)
        else:
            rgb = lerp(stops[1], stops[2], (t - 0.5) / 0.5)
        if h > 2 and y == 0:
            rgb = highlight
        elif h > 2 and y == h - 1:
            rgb = edge
        color = "#%02x%02x%02x" % rgb
        for x in range(w):
            img.put(color, (x, y))
    return img
class PhotoPrintApp:
    @staticmethod
    def _get_work_area():
        try:
            user32 = ctypes.windll.user32
            rect = wintypes.RECT()
            SPI_GETWORKAREA = 0x0030
            user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(rect), 0)
            return (rect.right - rect.left, rect.bottom - rect.top)
        except Exception:
            return (0, 0)
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        # 主入口 main() 已在 Tk() 前调用 enable_dpi_awareness()；这里再读一次
        # 缩放系数，保证直接构造 PhotoPrintApp(如测试/嵌入) 时同样适配。
        self._dpi_scale = read_dpi_scale(root)
        apply_tk_scaling(root)
        log("DPI 适配: 感知=%s 缩放=%.2f (%.0f%%)"
            % (_DPI_AWARENESS, self._dpi_scale, self._dpi_scale * 100))
        # 窗口与最小尺寸按 DPI 换算；工作区/屏幕读数在 DPI 感知下已是物理像素
        w, h = px(920), px(640)
        wa_w, wa_h = self._get_work_area()
        if wa_w <= 0:
            wa_w = self.root.winfo_screenwidth()
        if wa_h <= 0:
            wa_h = self.root.winfo_screenheight()
        x = (wa_w - w) // 2
        y = (wa_h - h) // 2
        self.root.geometry("%dx%d+%d+%d" % (w, h, x, y))
        self.root.minsize(px(760), px(520))
        # 初始化数据存储
        self.image_paths = []
        self.printing = False
        self._preview_ready = False
        # Tk 对 PhotoImage 只持弱引用，必须自己保存强引用；按用途分池，
        # 列表重建/预览重绘各自只清理自己的池，互不误伤。
        self._list_image_refs = []
        self._preview_image_refs = []
        self._preview_cache = ImageCache(max_size=IMAGE_CACHE_MAX_PREVIEW)
        self._thumb_cache = ImageCache(max_size=IMAGE_CACHE_MAX_THUMB)
        self._page_cache = _page_cache
        self._perf = _perf
        self._pagination = None
        # 记录上一次重绘时的选中索引：只有它变化才算「选择操作」，
        # 翻页按钮、窗口 resize、纸张/布局重绘时集合不变，不更新当前页码。
        self._last_sel_indices = None
        self._build_ui()
        # 跨显示器拖动会触发 WM_DPICHANGED，窗口随之 <Configure>；
        # 检测系数变化后重设字体/样式/列表几何并重绘预览。
        # 必须在 _build_ui 之后绑定，回调里会访问 photo_list / pagination_bar。
        self.root.bind("<Configure>", self._on_root_configure, add="+")
        self.root.after(0, self._refresh_printers)
    def _on_root_configure(self, event):
        if event.widget is not self.root:
            return
        try:
            new = read_dpi_scale(self.root)
            if abs(new - self._dpi_scale) < 0.01:
                return
            self._dpi_scale = new
            apply_tk_scaling(self.root, new)
            self._apply_win10_style()
            self.photo_list.rebuild_for_dpi()
            self.pagination_bar._frame.configure(
                height=px(PAGINATION_BAR_HEIGHT))
            # 字体变大后 ttk 按钮尺寸已变，渐变图必须按新尺寸重建，
            # 否则打印按钮会比「添加照片」小一圈。
            self._align_gradient_button(self.btn_print, label="打印按钮")
            self._draw_preview()
            log("DPI 变化: 缩放系数 -> %.2f (%.0f%%), 已重排 UI"
                % (new, new * 100))
        except Exception as e:
            log("DPI 变化处理异常: %s\n%s" % (e, traceback.format_exc()))
    # -- UI 构建 -----------------------------------------------------------
    def _apply_win10_style(self):
        """应用 Win10 风格: vista 主题 + Segoe UI 字体 + 配色。
        字号(点)固定，由 tk scaling 按 DPI 放大；像素类 padding 才需要 pxp()。
        """
        try:
            style = ttk.Style(self.root)
            style.theme_use("vista")
        except Exception:
            pass
        try:
            import tkinter.font as tkfont
            for name in ("TkDefaultFont", "TkTextFont", "TkFixedFont",
                         "TkMenuFont", "TkHeadingFont", "TkCaptionFont",
                         "TkSmallCaptionFont", "TkIconFont"):
                tkfont.nametofont(name).configure(family="Segoe UI", size=9)
        except Exception:
            pass
        # 配色: 白色窗口背景 + ttk 按钮文字 #333；
        # 打印与弹窗确定按钮为 macOS 风格蓝色纵向渐变 + 白字 tk.Button。
        try:
            style = ttk.Style(self.root)
            style.configure("TFrame", background="#FFFFFF")
            style.configure("TLabel", background="#FFFFFF", foreground="#333333")
            style.configure("TLabelframe", background="#FFFFFF")
            style.configure("TLabelframe.Label", background="#FFFFFF",
                            foreground="#333333")
            style.configure("TButton", foreground="#333333")
            # 工具栏与翻页条的所有 ttk.Button 共用此宽度；控件级不要再传 width 覆盖
            style.configure("TButton", width=BUTTON_WIDTH)
            # 垂直 padding 放大按钮高度；ttk.Button 无 height 选项，这是标准做法。
            # padding 是像素，按 DPI 换算；width=BUTTON_WIDTH 是字符单位，不换算。
            style.configure("TButton", padding=pxp(*BUTTON_PADDING))
            # 所有 ttk.Button(含删除/清空弹窗的「取消」按钮)统一使用 TButton，
            # 与「添加照片」等工具栏按钮保持相同宽高，不再单独设置紧凑样式。
            style.configure("TRadiobutton", background="#FFFFFF")
            style.configure("TCombobox", padding=pxp(*COMBOBOX_PADDING))
            self.root.configure(bg="#FFFFFF")
        except Exception:
            pass
        # 字体按 tk scaling 变化后，文件名换行宽度需按新字体重算
        try:
            self.photo_list._render()
        except Exception:
            pass
    def _build_ui(self):
        # 布局里的像素值一律经 px()/pxp() 换算；Combobox/Spinbox 的 width 是
        # 字符单位、字号是点数，二者随 tk scaling 自动放大，不再乘 DPI 系数。
        # 浅蓝提示横幅
        banner = tk.Frame(self.root, bg="#DDEBF7", height=px(30))
        banner.pack(side=tk.TOP, fill=tk.X)
        banner.pack_propagate(False)
        tk.Label(banner, bg="#DDEBF7", fg="#333333",
                 text="你希望如何打印图片？",
                 font=("Segoe UI", 10)).pack(side=tk.LEFT,
                                             padx=px(14), pady=px(5))
        # 顶部工具栏: 添加/删除/清空按钮 + 打印机选择 + 刷新按钮
        toolbar = ttk.Frame(self.root, padding=px(6))
        toolbar.pack(side=tk.TOP, fill=tk.X)
        self.btn_add = ttk.Button(toolbar, text="添加照片", command=self.add_photos)
        self.btn_add.pack(side=tk.LEFT, padx=px(2))
        self.btn_remove = ttk.Button(toolbar, text="删除选中 (Del)", command=self.remove_selected)
        self.btn_remove.pack(side=tk.LEFT, padx=px(2))
        self.btn_clear = ttk.Button(toolbar, text="清空", command=self.clear_photos)
        self.btn_clear.pack(side=tk.LEFT, padx=px(2))
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=px(6))
        ttk.Label(toolbar, text="打印机:").pack(side=tk.LEFT, padx=pxp(6, 2))
        self.cmb_printer = ttk.Combobox(toolbar, width=26, state="readonly")
        self.cmb_printer.pack(side=tk.LEFT, padx=px(8))
        self.btn_refresh = ttk.Button(toolbar, text="刷新", command=self._refresh_printers)
        self.btn_refresh.pack(side=tk.LEFT, padx=px(2))
        # 主区域: 左侧预览与设置 + 右侧照片列表
        main = ttk.Frame(self.root)
        main.pack(side=tk.TOP, fill=tk.BOTH, expand=True,
                  padx=px(8), pady=pxp(0, 8))
        # 左侧: 预览画布 + 翻页条 + 设置面板
        left = ttk.Frame(main)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        preview_frame = ttk.LabelFrame(left, text="布局预览", padding=px(4))
        preview_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True,
                           pady=pxp(0, 6))
        self.canvas = tk.Canvas(preview_frame, bg="#f0f0f0", highlightthickness=0)
        self.preview_vsb = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL,
                                         command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.preview_vsb.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.preview_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.bind("<Configure>", self._draw_preview)
        self.pagination_bar = PaginationBar(left, self._on_prev_page,
                                            self._on_next_page)
        # 设置面板
        settings = ttk.LabelFrame(left, text="打印设置", padding=px(8))
        settings.pack(side=tk.TOP, fill=tk.X)
        # 右侧: 照片列表
        right = ttk.Frame(main, width=px(220))
        right.pack_propagate(False)
        right.pack(side=tk.RIGHT, fill=tk.Y,
                   padx=pxp(LEFT_LIST_PADDING, RIGHT_LIST_EDGE_PADDING))
        ttk.Label(right, text="照片列表").pack(anchor=tk.W, pady=pxp(0, 4))
        list_frame = ttk.Frame(right)
        list_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.photo_list = PhotoListWidget(list_frame, [], on_select=self._on_list_select,
                                          on_delete=self.remove_selected,
                                          right_padding=RIGHT_LIST_PADDING)
        btn_del_frame = ttk.Frame(right)
        btn_del_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=pxp(4, 0))
        self.btn_list_delete = ttk.Button(btn_del_frame, text="删除选中 (Del)",
                                          command=self.remove_selected)
        # 高度随全局 TButton padding 放大；宽度 fill=X 铺满右侧面板，
        # 本就大于 BUTTON_WIDTH，无需再改，否则会破坏列表下方的整行布局。
        self.btn_list_delete.pack(fill=tk.X)
        # 纸张 / 布局 / 张数 行
        row1 = ttk.Frame(settings)
        row1.pack(fill=tk.X, pady=px(3))
        ttk.Label(row1, text="纸张:").pack(side=tk.LEFT)
        self.cmb_paper = ttk.Combobox(row1, values=list(PAPER_SIZES.keys()),
                                       width=12, state="readonly")
        self.cmb_paper.pack(side=tk.LEFT, padx=pxp(8, 16))
        self.cmb_paper.bind("<<ComboboxSelected>>", lambda e: self._reset_pagination_and_redraw())
        ttk.Label(row1, text="布局:").pack(side=tk.LEFT)
        self.cmb_layout = ttk.Combobox(row1, values=list(LAYOUTS.keys()),
                                        width=14, state="readonly")
        self.cmb_layout.pack(side=tk.LEFT, padx=pxp(8, 16))
        self.cmb_layout.bind("<<ComboboxSelected>>", lambda e: self._reset_pagination_and_redraw())
        ttk.Label(row1, text="张数:").pack(side=tk.LEFT)
        self.spin_copies = tk.Spinbox(row1, from_=1, to=99, width=5)
        self.spin_copies.pack(side=tk.LEFT, padx=pxp(4, 0))
        self.spin_copies.bind("<MouseWheel>", self._on_copies_mousewheel)
        # 单面 / 双面 选择行
        row2 = ttk.Frame(settings)
        row2.pack(fill=tk.X, pady=px(3))
        ttk.Label(row2, text="单双面:").pack(side=tk.LEFT)
        self.duplex_var = tk.IntVar(value=DUPLEX_NONE)
        self.rb_single = ttk.Radiobutton(row2, text="单面", variable=self.duplex_var,
                                         value=DUPLEX_NONE)
        self.rb_single.pack(side=tk.LEFT, padx=pxp(6, 10))
        self.rb_duplex = ttk.Radiobutton(row2, text="双面(长边翻页)书籍",
                                         variable=self.duplex_var, value=DUPLEX_VERTICAL)
        self.rb_duplex.pack(side=tk.LEFT)
        self.rb_duplex_short = ttk.Radiobutton(row2, text="双面(短边翻页)日历",
                                                variable=self.duplex_var, value=DUPLEX_HORIZONTAL)
        self.rb_duplex_short.pack(side=tk.LEFT)
        # 底部: 进度条 + 状态文字 + 打印/关闭按钮
        bottom = ttk.Frame(self.root, padding=px(8))
        bottom.pack(side=tk.BOTTOM, fill=tk.X)
        prog_frame = tk.Frame(bottom, height=px(25))
        prog_frame.pack_propagate(False)
        prog_frame.pack(fill=tk.X, pady=pxp(0, 6))
        self.progress = ttk.Progressbar(prog_frame, mode="determinate")
        self.progress.pack(fill=tk.BOTH, expand=True)
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(bottom, textvariable=self.status_var).pack(side=tk.LEFT)
        self.btn_print = self._create_gradient_button(bottom, "打印", self.start_print)
        self.btn_print.pack(side=tk.RIGHT, padx=pxp(6, 0))
        self.btn_cancel = ttk.Button(bottom, text="关闭", command=self.close)
        self.btn_cancel.pack(side=tk.RIGHT)
        # 绑定快捷键与窗口关闭事件
        self.root.bind("<Escape>", lambda e: self._on_escape(e))
        self.root.bind("<Delete>", lambda e: self._on_root_delete())
        self.root.bind("<Prior>", lambda e: self._on_prev_page())
        self.root.bind("<Next>", lambda e: self._on_next_page())
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.root.after(0, self._deferred_init)
        self.root.after(0, self._apply_win10_style)
        self.root.after(0, lambda: set_window_icon(self.root))
        self.root.after(0, self._sync_spinbox_height)
        # 必须在 _apply_win10_style 之后：样式生效后才能测到放大后的按钮尺寸
        self.root.after(0, self._sync_print_button_size)
        self._drag_drop = DragDropManager(self)
        self._drag_drop.register()
    def _on_escape(self, event):
        if getattr(self, "_drag_drop", None) and self._drag_drop.is_drag_active:
            self._drag_drop.cancel_drag()
            return
        self.close()
    def _on_root_delete(self):
        sel_indices = self.photo_list.get_selected_indices()
        if not sel_indices:
            return
        self.remove_selected()
        return "break"
    def _deferred_init(self):
        # 设置下拉框默认值与初始张数
        self.cmb_paper.set("A4")
        self.cmb_layout.set(list(LAYOUTS.keys())[0])
        self.spin_copies.delete(0, "end")
        self.spin_copies.insert(0, "1")
    # -- 照片列表 ----------------------------------------------------------
    def refresh_ui(self, redraw_preview=True):
        """统一的 UI 刷新入口：失效三级缓存、重建列表、(可选)重绘预览。
        添加/删除/清空等操作只调这里，避免各调用点漏掉某个缓存或某次重绘。
        """
        self._preview_cache.invalidate()
        self._thumb_cache.invalidate()
        self._page_cache.invalidate()
        self._refresh_list()
        if redraw_preview:
            self._draw_preview()
    def _add_files(self, paths):
        start_idx = len(self.image_paths)
        added = 0
        for f in paths:
            if f not in self.image_paths:
                self.image_paths.append(f)
                added += 1
        if added == 0:
            return 0
        # 不在这里重绘预览：set_selection 会触发 _on_list_select → _draw_preview，
        # 避免同一次添加产生两次完整重绘。
        self.refresh_ui(redraw_preview=False)
        self.photo_list.set_selection(set(range(start_idx, start_idx + added)))
        return added
    def add_photos(self):
        files = filedialog.askopenfilenames(
            title="选择照片",
            filetypes=[("图片文件", " ".join("*" + e for e in SUPPORTED_IMAGE_EXTENSIONS))]
        )
        if not files:
            return
        self._add_files(files)
    def _handle_dropped_files(self, paths, target_type):
        """处理拖放路径；target_type 仅用于日志，不区分「界面/任务栏」的处理方式。"""
        t0 = time.time()
        filtered = filter_image_paths(paths)
        added = self._add_files(filtered)
        # 跳过数含扩展名过滤与重复文件去重两部分，保证「文件数 = 添加 + 跳过」
        skipped = len(paths) - added
        elapsed = time.time() - t0
        # 耗时精度统一由 TIME_LOG_DIGITS 控制，勿改成 f-string 里的硬编码小数位
        log(("拖拽添加: 目标=%s, 文件数=%d, 添加=%d, 跳过=%d, 耗时=%.*f 秒")
            % (target_type, len(paths), added, skipped, TIME_LOG_DIGITS, elapsed))
        self._draw_preview()
    def remove_selected(self):
        sel_indices = self.photo_list.get_selected_indices()
        if not sel_indices:
            return
        self._confirm_delete(sel_indices)
    def _close_confirm_dialog(self, dialog):
        # 先销毁主窗口内的平滑遮罩 Label，恢复真实控件，再销毁弹窗；
        # 引用一并置空，让 PhotoImage 能被及时回收。
        label = getattr(dialog, "_scrim_label", None)
        if label is not None:
            try:
                label.destroy()
            except Exception:
                pass
        try:
            dialog._scrim_photo = None
        except Exception:
            pass
        try:
            dialog.destroy()
        except Exception:
            pass
    def _capture_dimmed_root(self):
        """截取主窗口客户区并整体调暗，返回 (PhotoImage, w, h)；失败返回 None。
        必须在创建弹窗 Toplevel 之前调用，否则弹窗会被一并截进图里，调暗后铺回
        主窗口会与真实弹窗形成重影。截图是静态快照，弹窗期间背景不实时刷新。
        """
        try:
            from PIL import ImageGrab, Image, ImageTk
            root = self.root
            if root.state() == "iconic" or not root.winfo_viewable():
                return None
            root.lift()
            root.update_idletasks()
            x = root.winfo_rootx()
            y = root.winfo_rooty()
            w = root.winfo_width()
            h = root.winfo_height()
            if w <= 0 or h <= 0:
                return None
            # 成功声明 DPI 感知时，Tk 的 winfo 坐标已是屏幕物理像素，bbox 直接用；
            # 声明失败(系统仍在位图拉伸)时，Tk 是 96 DPI 逻辑像素，必须乘缩放
            # 系数换成屏幕物理坐标，否则只会截到左上角一小块。
            s = dpi_scale()
            if is_dpi_aware():
                bbox = (x, y, x + w, y + h)
                expect = (w, h)
            else:
                bbox = (int(round(x * s)), int(round(y * s)),
                        int(round((x + w) * s)), int(round((y + h) * s)))
                expect = (max(1, int(round(w * s))),
                          max(1, int(round(h * s))))
            # Pillow >= 9.2 支持 all_screens；旧版 / 异常时退回单屏 bbox
            try:
                shot = ImageGrab.grab(bbox=bbox, all_screens=True)
            except Exception:
                shot = ImageGrab.grab(bbox=bbox)
            if shot is None or shot.width <= 0 or shot.height <= 0:
                return None
            shot = shot.convert("RGB")
            if shot.size != expect:
                log("弹窗遮罩: 截图尺寸 %s 与预期 %s 不符(感知=%s, 缩放=%.2f)"
                    % (shot.size, expect, _DPI_AWARENESS, s))
            if shot.size != (w, h):
                # 最终统一缩回 Tk 的客户区尺寸，PhotoImage 才与 Label 1:1 对齐
                shot = shot.resize((w, h))
            dark = Image.new("RGB", shot.size, (0, 0, 0))
            dimmed = Image.blend(shot, dark, DIALOG_SCRIM_DARKEN)
            photo = ImageTk.PhotoImage(dimmed)
            log("弹窗遮罩: 主窗口 %dx%d 已调暗 %.0f%%"
                % (w, h, DIALOG_SCRIM_DARKEN * 100))
            return photo, w, h
        except Exception as e:
            log("弹窗遮罩截图降级: %s\n%s" % (e, traceback.format_exc()))
            return None
    def _render_dialog_shadow(self, dialog, snapshot=None):
        """弹窗显示时，在主窗口内部铺一层「调暗后的主窗口快照」形成平滑遮罩。
        遮罩是 root 的直接子控件(单个 Label)，不是独立 Toplevel / -alpha / 落影层；
        关闭时由 _close_confirm_dialog 销毁。快照失败则不显示遮罩，弹窗照常可用。
        弹窗本身是独立 Toplevel，浮在遮罩之上，白底与变暗背景形成对比。
        """
        # 无论有没有快照，弹窗都要浮在主窗口之上
        try:
            dialog.lift()
        except Exception:
            pass
        if snapshot is None:
            log("弹窗遮罩降级: 无可用快照，本次不显示遮罩（弹窗功能不受影响）")
            return
        photo, w, h = snapshot
        try:
            label = tk.Label(self.root, image=photo, bd=0,
                             highlightthickness=0, padx=0, pady=0,
                             anchor="nw", bg=DIALOG_SCRIM_FALLBACK_BG)
            # 固定尺寸与快照 1:1 对齐；窗口被拉大时露出的边是中性灰而非亮白
            label.place(x=0, y=0, width=w, height=h)
            label.lift()
            # Tk 对 PhotoImage 只持弱引用，必须由弹窗保存强引用
            dialog._scrim_label = label
            dialog._scrim_photo = photo
        except Exception as e:
            log("弹窗遮罩渲染异常: %s\n%s" % (e, traceback.format_exc()))
    def _bind_button_tooltip(self, widget, text):
        """给按钮绑定轻量悬浮提示（快捷键说明）；异常静默降级不影响按钮功能。"""
        state = {"after": None, "win": None}
        def hide(_event=None):
            try:
                if state["after"] is not None:
                    widget.after_cancel(state["after"])
                    state["after"] = None
                if state["win"] is not None:
                    state["win"].destroy()
                    state["win"] = None
            except Exception:
                pass
        def show():
            try:
                if not widget.winfo_exists() or state["win"] is not None:
                    return
                x = widget.winfo_rootx() + widget.winfo_width() // 2
                y = widget.winfo_rooty() + widget.winfo_height() + px(4)
                tip = tk.Toplevel(widget)
                tip.overrideredirect(True)
                tk.Label(tip, text=text, bg="#333333", fg="#FFFFFF",
                         font=("Segoe UI", 8),
                         padx=px(6), pady=px(3)).pack()
                tip.geometry("+%d+%d" % (x, y))
                state["win"] = tip
            except Exception:
                pass
        def schedule(_event):
            hide()
            try:
                state["after"] = widget.after(400, show)
            except Exception:
                pass
        try:
            widget.bind("<Enter>", schedule)
            widget.bind("<Leave>", hide)
            widget.bind("<ButtonPress>", hide)
            widget.bind("<Destroy>", hide)
        except Exception:
            pass
    def _create_gradient_button(self, parent, text, command,
                                width=None, height=None, fallback_style=None):
        """创建 macOS 风格蓝色渐变 + 白色文字的 tk.Button。
        width/height 不传时用初始尺寸，创建后还需调用 _align_gradient_button()
        对齐到 ttk.Button 宽高；所有渐变按钮(打印/弹窗确定)走同一套配色。
        fallback_style 用于异常回退到 ttk.Button 时指定 style（默认全局 TButton，
        因此回退按钮同样与「添加照片」宽高一致）。
        """
        try:
            w = width or px(PRINT_BTN_GRADIENT_WIDTH)
            h = height or px(PRINT_BTN_GRADIENT_HEIGHT)
            img = make_gradient_image(w, h)
            # padx/pady 归零：tk.Button 默认内边距会让实际按钮比渐变图大一圈，
            # 归零后按钮尺寸严格等于 image 尺寸，日志里的「尺寸」才与行为一致，
            # 也才能让 _sync_print_button_size 精确对齐到 ttk 按钮的宽高。
            # fg 白字 + 渐变图蓝底 = macOS 默认按钮观感；bg 仅作无图时的兜底色
            btn = tk.Button(parent, image=img, text=text, compound="center",
                            fg="#FFFFFF", bg="#0A84FF", relief="flat", bd=0,
                            highlightthickness=0, padx=0, pady=0,
                            command=command)
            btn._gradient_img = img
            log("渐变按钮创建: text=%s 尺寸=%dx%d" % (text, w, h))
            return btn
        except Exception as e:
            log("渐变按钮创建异常: %s\n%s" % (e, traceback.format_exc()))
            kw = {"text": text, "command": command}
            if fallback_style:
                kw["style"] = fallback_style
            return ttk.Button(parent, **kw)
    def _align_gradient_button(self, btn, ref=None, label="渐变按钮"):
        """把 tk.Button 渐变按钮的外框精确对齐到 ttk.Button（默认「添加照片」）。
        tk.Button 是像素尺寸、ttk.Button 是字符尺寸，只有运行时测量才能保证一致；
        即便 bd=0/highlightthickness=0/padx=pady=0，Tk 仍会为焦点框预留每边 1px，
        因此测出差值后把渐变图缩小同等量。不同平台余量不同，故不写死。
        """
        try:
            if not hasattr(btn, "_gradient_img"):
                return  # 回退成 ttk.Button 时本身已是 TButton 宽高
            self.root.update_idletasks()
            reference = ref or self.btn_add
            w = reference.winfo_reqwidth()
            h = reference.winfo_reqheight()
            if w <= 0 or h <= 0:
                return
            img = make_gradient_image(w, h)
            btn.configure(image=img)
            # Tk 对 PhotoImage 只持弱引用，必须替换并保存新引用
            btn._gradient_img = img
            extra_w = btn.winfo_reqwidth() - w
            extra_h = btn.winfo_reqheight() - h
            if extra_w or extra_h:
                img = make_gradient_image(max(1, w - extra_w), max(1, h - extra_h))
                btn.configure(image=img)
                btn._gradient_img = img
            log("%s尺寸对齐: %dx%d (补偿 %d,%d)"
                % (label, btn.winfo_reqwidth(), btn.winfo_reqheight(),
                   extra_w, extra_h))
        except Exception as e:
            log("%s尺寸对齐失败: %s\n%s" % (label, e, traceback.format_exc()))
    def _sync_print_button_size(self):
        """打印按钮延后对齐：样式在 after(0) 的 _apply_win10_style 里才生效。"""
        self._align_gradient_button(self.btn_print, label="打印按钮")
    def _on_copies_mousewheel(self, event):
        try:
            cur = int(self.spin_copies.get())
        except Exception:
            cur = 1
        if event.delta > 0:
            cur = min(cur + 1, 99)
        else:
            cur = max(cur - 1, 1)
        self.spin_copies.delete(0, "end")
        self.spin_copies.insert(0, str(cur))
        log("张数选择框滚轮调整: %d (delta=%d)" % (cur, event.delta))
    def _sync_spinbox_height(self):
        try:
            self.root.update_idletasks()
            ref_h = self.cmb_paper.winfo_height()
            if ref_h <= 1:
                return
            cur_h = self.spin_copies.winfo_height()
            diff = ref_h - cur_h
            if diff != 0:
                self.spin_copies.pack_forget()
                self.spin_copies.pack(side=tk.LEFT, padx=(4, 0), ipady=diff / 2.0)
                self.root.update_idletasks()
            final_h = self.spin_copies.winfo_height()
            log("张数选择框高度同步: 参考=%dpx 当前=%dpx ipady=%.1f 最终=%dpx 偏差=%dpx" % (ref_h, cur_h, diff / 2.0, final_h, final_h - ref_h))
        except Exception as e:
            log("张数选择框高度同步异常: %s\n%s" % (e, traceback.format_exc()))
    def _show_confirm_dialog(self, title, message, thumb_im, placeholder_text,
                             on_confirm):
        """删除/清空共用的确认弹窗：白底 + 平滑遮罩 + 居中 + 同宽高按钮 + 快捷键。
        thumb_im 不为 None 时显示缩略图（删除弹窗），为 None 时显示多行
        placeholder_text（清空弹窗）；两者共用完全相同的布局、字体、按钮样式、
        遮罩与快捷键，保证视觉一致。
        """
        # 必须先抓主窗口快照再创建弹窗：否则 Toplevel 会被截进图里，
        # 调暗后铺回主窗口与真实弹窗形成重影。
        snapshot = self._capture_dimmed_root()
        dialog = tk.Toplevel(self.root, bg="#FFFFFF")
        dialog.title(title)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.configure(bg="#FFFFFF")
        set_window_icon(dialog)
        # 防重入：Enter 可能同时触发 dialog 绑定与获得焦点按钮的默认行为
        state = {"done": False}
        def cancel():
            if state["done"]:
                return
            state["done"] = True
            self._close_confirm_dialog(dialog)
        def confirm():
            if state["done"]:
                return
            state["done"] = True
            on_confirm(dialog)
        # 快捷键: Enter / 小键盘 Enter 确定，Esc 取消
        dialog.bind("<Escape>", lambda e: cancel())
        dialog.bind("<Return>", lambda e: confirm())
        dialog.bind("<KP_Enter>", lambda e: confirm())
        dialog.protocol("WM_DELETE_WINDOW", cancel)
        frame = tk.Frame(dialog, padx=px(16), pady=px(14), bg="#FFFFFF")
        frame.pack(fill=tk.BOTH, expand=True)
        content_frame = tk.Frame(frame, bg="#FFFFFF")
        content_frame.pack(side=tk.TOP)
        content_frame.grid_columnconfigure(0, minsize=px(CONFIRM_IMAGE_MAX_SIZE))
        thumb_area = tk.Frame(content_frame, bg="#FFFFFF",
                              width=px(CONFIRM_IMAGE_MAX_SIZE),
                              height=px(CONFIRM_IMAGE_MAX_SIZE))
        thumb_area.pack_propagate(False)
        thumb_area.grid(row=0, column=0, sticky="", pady=pxp(0, 6))
        if thumb_im is not None:
            from PIL import ImageTk
            photo = ImageTk.PhotoImage(thumb_im)
            img_label = tk.Label(thumb_area, image=photo, bg="#FFFFFF")
            # 弹窗是短生命周期窗口，用 Label 属性持有 PhotoImage 即可，
            # 不纳入 app 级引用池，避免反复打开累积。
            img_label.image = photo
            img_label.pack(expand=True)
        else:
            tk.Label(thumb_area, text=placeholder_text, bg="#FFFFFF",
                     fg="#666666", font=("Segoe UI", 12),
                     justify=tk.CENTER).pack(expand=True)
        tk.Label(content_frame, text=message, bg="#FFFFFF", fg="#333333",
                 font=("Segoe UI", 9),
                 justify=tk.CENTER).grid(row=1, column=0, sticky="")
        # 快捷键提示行：按钮文字已带 (Enter)/(Esc)，这里再给出完整说明
        tk.Label(content_frame, text="快捷键: Enter 确定  /  Esc 取消",
                 bg="#FFFFFF", fg="#999999", font=("Segoe UI", 8)).grid(
                     row=2, column=0, sticky="", pady=pxp(6, 0))
        button_frame = tk.Frame(frame, bg="#FFFFFF")
        button_frame.pack(side=tk.TOP, pady=pxp(10, 0))
        # 与「添加照片」同宽高：取消用全局 TButton；确定是 macOS 蓝渐变按钮，
        # 由 _align_gradient_button 运行时对齐到 btn_add 的宽高。
        btn_cancel = ttk.Button(button_frame, text="取消 (Esc)", command=cancel)
        btn_cancel.pack(side=tk.LEFT, padx=pxp(0, 6))
        btn_confirm = self._create_gradient_button(button_frame, "确定 (Enter)",
                                                   confirm)
        btn_confirm.pack(side=tk.LEFT)
        self._align_gradient_button(btn_confirm, ref=self.btn_add,
                                    label="确定按钮")
        self._bind_button_tooltip(btn_cancel, "按 Esc 取消，不改动照片列表")
        self._bind_button_tooltip(btn_confirm, "按 Enter 立即执行")
        try:
            btn_confirm.focus_set()
        except Exception:
            pass
        dialog.update_idletasks()
        dw = dialog.winfo_reqwidth()
        dh = dialog.winfo_reqheight()
        sx = self.root.winfo_rootx()
        sy = self.root.winfo_rooty()
        sw = self.root.winfo_width()
        sh = self.root.winfo_height()
        dx = sx + (sw - dw) // 2
        dy = sy + (sh - dh) // 2
        dialog.geometry("+%d+%d" % (dx, dy))
        # 把「弹窗创建前」抓好的调暗快照铺进主窗口内部
        self._render_dialog_shadow(dialog, snapshot)
        dialog.wait_window()
    def _confirm_delete(self, sel_indices):
        if not sel_indices:
            return
        first_idx = sel_indices[0]
        if first_idx >= len(self.image_paths):
            return
        path = self.image_paths[first_idx]
        name = os.path.basename(path)
        # 弹窗只展示第一张的缩略图，确定后会删除全部选中项，
        # 因此标题与文字都要带总数，避免用户以为只删一张。
        count = len(sel_indices)
        label_text = name if count == 1 else "%s 等 %d 张" % (name, count)
        im = self._thumb_cache.get_thumbnail(
            path, (px(CONFIRM_IMAGE_MAX_SIZE), px(CONFIRM_IMAGE_MAX_SIZE)),
            allow_upscale=False)
        self._show_confirm_dialog(
            title="删除图片 (%d 张)" % count,
            message=label_text,
            thumb_im=im,
            placeholder_text="",
            on_confirm=lambda d: self._do_delete(sel_indices, d))
    def _do_delete(self, sel_indices, dialog):
        try:
            for i in sorted(sel_indices, reverse=True):
                del self.image_paths[i]
            self.refresh_ui()
        except Exception as e:
            log("删除弹窗异常: %s\n%s" % (e, traceback.format_exc()))
        try:
            self._close_confirm_dialog(dialog)
        except Exception as e:
            log("删除弹窗关闭异常: %s\n%s" % (e, traceback.format_exc()))
    def clear_photos(self):
        if not self.image_paths:
            return
        count = len(self.image_paths)
        # 与删除弹窗共用 _show_confirm_dialog：同样的白底、平滑遮罩、字体、
        # 按钮宽高、快捷键与 macOS 蓝渐变确定按钮，只是内容改为文字提示。
        self._show_confirm_dialog(
            title="清空照片 (%d 张)" % count,
            message="确定要清空所有照片吗？",
            thumb_im=None,
            placeholder_text="共 %d 张照片\n将被移除" % count,
            on_confirm=self._do_clear)
    def _do_clear(self, dialog):
        try:
            self.image_paths = []
            self.refresh_ui()
        except Exception as e:
            log("清空弹窗异常: %s\n%s" % (e, traceback.format_exc()))
        try:
            self._close_confirm_dialog(dialog)
        except Exception as e:
            log("清空弹窗关闭异常: %s\n%s" % (e, traceback.format_exc()))
    def _refresh_list(self):
        # 真实刷新(含缩略图解码与条目渲染)必须落在 measure 内，否则日志恒为 0
        return self._perf.measure("列表刷新", LIST_REFRESH_TIME_LIMIT,
                                  self._refresh_list_impl)
    def _refresh_list_impl(self):
        # 只清理列表缩略图池；预览池由 _draw_preview_dispatch 自行管理
        self._list_image_refs.clear()
        entries = []
        for p in self.image_paths:
            name = os.path.basename(p)
            im = self._thumb_cache.get_thumbnail(
                p, (px(LIST_IMAGE_MAX_SIZE), px(LIST_IMAGE_MAX_SIZE)),
                allow_upscale=False)
            if im is not None:
                from PIL import ImageTk
                th = ImageTk.PhotoImage(im)
                self._list_image_refs.append(th)
                entries.append((name, th))
            else:
                entries.append((name, None))
        self.photo_list.refresh(entries)
    def _on_list_select(self, _event):
        # 选择变化只更新当前页码（在 _draw_preview_impl 里按选中集合判断），
        # 总页数始终由布局 + 照片总数计算，这里不做任何分页重置。
        self._draw_preview()
    # -- 预览 --------------------------------------------------------------
    def _draw_paper_frame(self):
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        # 这里的 4 是「canvas 尚未完成布局」的容差阈值，不是可见尺寸，
        # 无需按 DPI 换算；换算后反而会在高 DPI 下误判为未布局。
        if w <= 4 or h <= 4:
            return
        paper_dims = PAPER_DIMENSIONS.get(self.cmb_paper.get())
        if paper_dims:
            paper_w_ratio, paper_h_ratio = paper_dims
        else:
            paper_w_ratio, paper_h_ratio = 210, 297
        # 局部变量名用 paper_x/paper_y：不能叫 px/py，否则会遮蔽全局 px() 函数
        pad = px(16)
        avail_w = w - pad * 2
        avail_h = h - pad * 2
        if avail_w <= 0 or avail_h <= 0:
            return
        scale = min(avail_w / paper_w_ratio, avail_h / paper_h_ratio)
        pw = paper_w_ratio * scale
        ph = paper_h_ratio * scale
        paper_x = (w - pw) / 2
        paper_y = (h - ph) / 2
        self.canvas.create_rectangle(paper_x, paper_y,
                                     paper_x + pw, paper_y + ph,
                                     fill="white", outline="#888888",
                                     width=px(2))
    def _draw_preview(self, _event=None):
        # 首帧(只画纸框并调度二次重绘)也计入「预览重绘」指标；
        # 二次重绘会再产生一条该指标日志，一次用户操作共两条，属预期。
        return self._perf.measure("预览重绘", PREVIEW_REDRAW_TIME_LIMIT,
                                  self._draw_preview_dispatch)
    def _draw_preview_dispatch(self):
        self.canvas.delete("all")
        # 只清理上一批预览图；列表缩略图池不受影响
        self._preview_image_refs.clear()
        if not self._preview_ready:
            self._preview_ready = True
            self._draw_paper_frame()
            self.root.after(0, self._draw_preview)
            return
        return self._draw_preview_impl()
    def _draw_preview_impl(self):
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        if w <= 4 or h <= 4:
            return
        layout = self.cmb_layout.get()
        if layout not in LAYOUTS:
            return
        rows, cols, count = LAYOUTS[layout]
        paper_dims = PAPER_DIMENSIONS.get(self.cmb_paper.get())
        if paper_dims:
            paper_w_ratio, paper_h_ratio = paper_dims
        else:
            paper_w_ratio, paper_h_ratio = 210, 297
        paper_ratio = paper_w_ratio / paper_h_ratio
        # 同 _draw_paper_frame：局部变量避开 px/py，防止遮蔽全局 px()
        pad = px(16)
        avail_w = w - pad * 2
        avail_h = h - pad * 2
        if avail_w <= 0 or avail_h <= 0:
            return
        scale = min(avail_w / paper_w_ratio, avail_h / paper_h_ratio)
        pw = paper_w_ratio * scale
        ph = paper_h_ratio * scale
        paper_x = (w - pw) / 2
        paper_y = (h - ph) / 2
        self.canvas.create_rectangle(paper_x, paper_y,
                                     paper_x + pw, paper_y + ph,
                                     fill="white", outline="#888888",
                                     width=px(2))
        gap = px(4)
        cw = (pw - gap * (cols + 1)) / cols
        ch = (ph - gap * (rows + 1)) / rows
        sel_indices = sorted(self.photo_list.get_selected_indices())
        all_paths = list(self.image_paths)
        # 分页数据源固定为「全部已添加照片」：总页数始终由「布局每页张数 +
        # 照片总数」推导，与选中集合无关；预览内容也因此与打印范围一致。
        if self._pagination is None or self._pagination.per_page != count:
            self._pagination = PaginationController(all_paths, count)
        else:
            self._pagination.sync_paths(all_paths)
        # 只有「选中集合真的变了」才更新当前页码：跳到第一张选中照片所在页。
        # 点上一页/下一页、窗口 resize、纸张/布局重绘时集合不变，
        # current_page 保持不动；总页数在任何选择操作下都不变。
        if sel_indices != self._last_sel_indices:
            self._last_sel_indices = list(sel_indices)
            if sel_indices:
                self._pagination.current_page = self._pagination.page_of_index(sel_indices[0])
        page_paths = self._pagination.get_current_page_photos()
        for r in range(rows):
            for c in range(cols):
                x = paper_x + gap + c * (cw + gap)
                y = paper_y + gap + r * (ch + gap)
                self.canvas.create_rectangle(x, y, x + cw, y + ch,
                                             fill="#e8e8e8", outline="#bbbbbb")
                idx = r * cols + c
                if idx < len(page_paths):
                    im = self._preview_cache.get_preview_image(
                        page_paths[idx], (int(cw), int(ch)), paper_ratio)
                    if im is not None:
                        from PIL import ImageTk
                        photo = ImageTk.PhotoImage(im)
                        self._preview_image_refs.append(photo)
                        self.canvas.create_image(x + cw / 2, y + ch / 2,
                                                 image=photo)
                    else:
                        self.canvas.create_text(x + cw / 2, y + ch / 2,
                                                text=str(idx + 1), fill="#999999")
                else:
                    self.canvas.create_text(x + cw / 2, y + ch / 2,
                                            text=str(idx + 1), fill="#999999")
        self.pagination_bar.update(self._pagination.current_page,
                                   self._pagination.total_pages)
        # 照片数与总页数同口径：全部已添加照片数，不再随选择变化
        # 底部说明文字距画布底边的偏移量也按 DPI 换算
        self.canvas.create_text(w / 2, h - px(6),
                                text="布局: %s  |  照片: %d 张"
                                     % (layout, len(all_paths)),
                                fill="#666666", anchor=tk.S)
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        log("预览重绘: 当前页=%d/总页数=%d 照片数=%d 每页=%d"
            % (self._pagination.current_page, self._pagination.total_pages,
               len(all_paths), count))
    def _on_prev_page(self):
        if self._pagination is None:
            return
        self._pagination.prev()
        self._draw_preview()
    def _on_next_page(self):
        if self._pagination is None:
            return
        self._pagination.next()
        self._draw_preview()
    def _reset_pagination_and_redraw(self):
        # 纸张/布局变化：当前页码回到第 1 页；总页数仍按新布局每页张数 +
        # 照片总数重算，不受选中集合影响。
        if self._pagination is not None:
            self._pagination.reset()
        self._draw_preview()
    # -- 打印机 ------------------------------------------------------------
    def _refresh_printers(self):
        printers = list_printers()
        self.cmb_printer.config(values=printers)
        if printers:
            self.cmb_printer.set(printers[0])
    # -- 打印流程 ----------------------------------------------------------
    def start_print(self):
        if self.printing:
            return
        printer = self.cmb_printer.get()
        if not printer:
            messagebox.showwarning("提示", "请先选择打印机")
            return
        if not self.image_paths:
            messagebox.showwarning("提示", "请先添加照片")
            return
        layout = self.cmb_layout.get()
        try:
            copies = int(self.spin_copies.get())
        except Exception:
            copies = 1
        if copies < 1:
            copies = 1
        duplex = self.duplex_var.get()
        paper = self.cmb_paper.get()
        # 打印范围与预览分页同为全部已添加照片，两者口径一致；
        # 列表选择只影响当前预览页码，不影响打印内容。
        # 禁用除「关闭」外的控件并启动后台打印线程，关闭按钮保持可用以便随时退出。
        self.printing = True
        self._set_controls_state(False)
        self.progress["value"] = 0
        self.status_var.set("正在打印...")
        def worker():
            try:
                ok, msg = do_print(
                    self.image_paths, printer, layout, copies, duplex, paper,
                    progress_cb=lambda cur, tot: self._after(cur, tot),
                )
            except Exception as e:
                ok, msg = False, "打印异常: %s\n%s" % (e, traceback.format_exc())
            self._after_done(ok, msg)
        threading.Thread(target=worker, daemon=True).start()
    def _after(self, cur, tot):
        # 后台打印线程经 root.after 回主线程更新 Tk 控件，避免跨线程操作
        def _update():
            if tot > 0:
                self.progress["value"] = cur / tot * 100
            self.status_var.set("正在打印... %d/%d" % (cur, tot))
        self.root.after(0, _update)
    def _after_done(self, ok, msg):
        def _done():
            self.printing = False
            self._set_controls_state(True)
            self.progress["value"] = 100 if ok else 0
            self.status_var.set(msg)
            log(msg)
            if ok:
                messagebox.showinfo("打印完成", msg)
            else:
                messagebox.showerror("打印失败", msg)
        self.root.after(0, _done)
    def _set_controls_state(self, enabled):
        state = "normal" if enabled else "disabled"
        for w in (self.btn_add, self.btn_remove, self.btn_clear,
                  self.btn_list_delete, self.btn_print, self.btn_refresh,
                  self.cmb_layout, self.cmb_paper, self.cmb_printer,
                  self.spin_copies, self.rb_single, self.rb_duplex,
                  self.rb_duplex_short):
            try:
                w.config(state=state)
            except Exception:
                pass
        # 翻页按钮不进上面的循环：其状态由 PaginationBar._apply_state() 统一计算。
        # set_enabled 已会重算一次；恢复时再 update 一次是幂等的，确保页码同步。
        self.pagination_bar.set_enabled(enabled)
        if enabled and self._pagination is not None:
            self.pagination_bar.update(self._pagination.current_page,
                                       self._pagination.total_pages)
    def close(self):
        if self.printing:
            if not messagebox.askyesno("确定", "正在打印中，确定退出?"):
                return
        self.root.destroy()
# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------
def safe_build():
    """构建安全机制: 以 .spec 输出到临时目录 build_tmp, 体积/缩减比均达标才
    用产物 exe 替换 dist; 任一环节失败(含指标未达标)保留原 dist, 度量写日志。
    替换走同盘原子改名, copytree 仅作改名回滚失败时的兜底恢复。
    需单独调用, 未接入 main()。
    """
    import shutil
    import subprocess
    spec_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             "照片打印工具.spec")
    dist_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dist")
    build_tmp = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             "build_tmp")
    exe_name = "照片打印工具.exe"
    if os.path.exists(build_tmp):
        shutil.rmtree(build_tmp, ignore_errors=True)
    cmd = [sys.executable, "-m", "PyInstaller", "--clean", "--workpath",
           build_tmp, "--distpath", build_tmp, spec_path]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True)
    except Exception as e:
        log("构建失败: 无法启动 PyInstaller: %s" % e)
        return False
    if proc.returncode != 0:
        log("构建失败: 返回码 %d\n%s" % (proc.returncode, proc.stdout[-2000:]))
        return False
    built_exe = os.path.join(build_tmp, exe_name)
    if not os.path.exists(built_exe):
        log("构建失败: 产物未生成 %s" % built_exe)
        return False
    size = os.path.getsize(built_exe)
    ratio = (1 - size / CURRENT_EXE_SIZE_BASELINE) * 100
    size_ok = size <= EXE_SIZE_LIMIT
    ratio_ok = ratio >= VOLUME_REDUCTION_RATIO
    digits = int(VOLUME_MEASURE_PRECISION) if VOLUME_MEASURE_PRECISION > 0 else 2
    log(("exe 体积: %d 字节 (上限 %d 字节) %s; 基准 %d 字节; 缩减 %.*f%% "
         "(下限 %.*f%%) %s")
        % (size, EXE_SIZE_LIMIT, "达标" if size_ok else "超标",
           CURRENT_EXE_SIZE_BASELINE, digits, ratio, digits,
           VOLUME_REDUCTION_RATIO, "达标" if ratio_ok else "未达标"))
    if not (size_ok and ratio_ok):
        log("构建失败: 体积或缩减比未达标, 保留原 dist")
        return False
    # 先把达标产物放进独立 staging，所有可能失败的 IO 都在替换 dist 之前完成；
    # 替换时把旧 dist 改名为 backup，失败则改回，真正兑现「失败保留原 dist」。
    staging_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                               "dist_staging")
    backup_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "dist_backup")
    for d in (staging_dir, backup_dir):
        if os.path.exists(d):
            shutil.rmtree(d, ignore_errors=True)
    had_dist = os.path.exists(dist_dir)
    try:
        os.makedirs(staging_dir, exist_ok=True)
        shutil.move(built_exe, os.path.join(staging_dir, exe_name))
        if had_dist:
            os.replace(dist_dir, backup_dir)
        os.replace(staging_dir, dist_dir)
    except Exception as e:
        # 主路径用同盘 os.replace 做目录改名（原子、快，比逐文件 copytree 更抗
        # Windows 文件锁）；只有改名回滚也失败时，才用 copytree 逐文件兜底恢复。
        # dirs_exist_ok 需 Python 3.8+；更老版本会走异常分支并保留 backup 供手工恢复。
        if had_dist and os.path.exists(backup_dir) and not os.path.exists(dist_dir):
            try:
                os.replace(backup_dir, dist_dir)
            except Exception:
                try:
                    shutil.copytree(backup_dir, dist_dir, dirs_exist_ok=True)
                except Exception:
                    pass
        if not had_dist and os.path.exists(dist_dir):
            # 无原 dist 时的半成品目录不留垃圾
            shutil.rmtree(dist_dir, ignore_errors=True)
        shutil.rmtree(staging_dir, ignore_errors=True)
        log("构建失败: 替换 dist 异常, 已回滚并保留原 dist: %s" % e)
        return False
    if had_dist:
        shutil.rmtree(backup_dir, ignore_errors=True)
    shutil.rmtree(build_tmp, ignore_errors=True)
    log("构建成功: dist 已更新")
    return True
def main():
    # 必须在创建任何 Tk 窗口之前声明 DPI 感知，否则 Windows 会对整个窗口做
    # 位图拉伸，界面发糊且后续所有坐标换算都会失准。
    mode = enable_dpi_awareness()
    log("DPI 感知声明: %s" % mode)
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
        # 计时口径: 仅 UI 构造 + deiconify; Tk() 创建与 after(0) 延迟初始化不计入
        t0 = time.time()
        app = PhotoPrintApp(root)
        root.deiconify()
        t1 = time.time()
        elapsed = t1 - t0
        status = "达标" if elapsed <= STARTUP_TIME_LIMIT else "超标"
        log("启动耗时: %.*f 秒 (上限 %.1f 秒) %s"
            % (TIME_LOG_DIGITS, elapsed, STARTUP_TIME_LIMIT, status))
        root.mainloop()
    except Exception:
        log("主循环异常:\n%s" % traceback.format_exc())
        try:
            messagebox.showerror("错误", traceback.format_exc())
        except Exception:
            pass
if __name__ == "__main__":
    main()
