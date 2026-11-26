#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
照片打印工具
==========
类似 Windows 系统自带的照片打印，支持选择照片、选择打印机、纸张、布局、张数，
并新增【单面 / 双面】打印选项。
依赖: pywin32, Pillow
打包: pyinstaller --onefile --windowed
"""
import os
import sys
import threading
import time
import traceback
# === 第三方库导入 ===
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
# spec 6.1.1: 统一按钮宽度
BUTTON_WIDTH = 12
# spec 6.1.3: 启动完成时间上限(秒)
STARTUP_TIME_LIMIT = 1
# 性能指标上限(秒)
PREVIEW_REDRAW_TIME_LIMIT = 0.5
LIST_REFRESH_TIME_LIMIT = 2.0
PAGE_COMPOSE_TIME_LIMIT = 1.0
MEASURE_PRECISION = 0.001
# exe 体积指标(字节 / 百分比)
EXE_SIZE_LIMIT = 52428800
CURRENT_EXE_SIZE_BASELINE = 178330724
VOLUME_REDUCTION_RATIO = 70.0
VOLUME_MEASURE_PRECISION = 1
# spec 6.1.2: TCombobox 内边距 (水平, 垂直)
COMBOBOX_PADDING = (4, 4)
# 应用标题与日志文件路径
APP_TITLE = "照片打印工具"
LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "print_log.txt")
# ---------------------------------------------------------------------------
# 资源路径 / 窗口图标
# ---------------------------------------------------------------------------
def resource_path(rel):
    # 兼容 PyInstaller 打包后的 _MEIPASS 目录，未打包时回退到脚本所在目录
    base = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, rel)
def set_window_icon(win):
    # 打包后直接从 exe 资源取图标，小图标/大图标一次给全
    ico = sys.executable if getattr(sys, "frozen", False) else resource_path("app.ico")
    try:
        win.iconbitmap(ico)          # Windows 下 Tk 支持传 .exe 路径
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
# 图片缓存层
# ---------------------------------------------------------------------------
class ImageCache:
    """以 (路径, 尺寸) 为键缓存 PIL.Image，消除重复解码。"""
    def __init__(self):
        self._cache = {}
    def get_thumbnail(self, path, size):
        return self._get(path, size, to_rgb=True)
    def get_preview_image(self, path, size):
        return self._get(path, size, to_rgb=True)
    def get_fitted_image(self, path, box):
        return self._get(path, box, to_rgb=True)
    def invalidate(self):
        self._cache.clear()
    def _get(self, path, box, to_rgb):
        key = (path, tuple(box))
        if key in self._cache:
            return self._cache[key]
        img = self._decode_and_fit(path, box, to_rgb)
        if img is not None:
            self._cache[key] = img
        return img
    def _decode_and_fit(self, path, box, to_rgb):
        from PIL import Image
        try:
            im = Image.open(path)
            im.draft("RGB", tuple(box))
            w, h = im.size
            bw, bh = box
            if w <= 0 or h <= 0:
                return None
            scale = min(bw / w, bh / h)
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
        log("%s: %.3f 秒 (上限 %.3f 秒) %s" % (op_name, elapsed, limit, flag))
# ---------------------------------------------------------------------------
# 打印机 / DEVMODE 操作
# ---------------------------------------------------------------------------
def list_printers():
    """返回所有打印机名称列表(默认打印机排在第一位)。"""
    import win32print
    names = []
    try:
        flags = win32print.PRINTER_ENUM_NAME | win32print.PRINTER_ENUM_LOCAL
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
    """将双面/纸张设置写入打印机 DEVMODE，返回原始设置(用于恢复)。"""
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
    """恢复打印机 DEVMODE 原始设置。"""
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
# 模块级缓存与测量实例
_page_cache = ImageCache()
_perf = PerfMeasure()
# ---------------------------------------------------------------------------
# 图像 / 页面合成
# ---------------------------------------------------------------------------
def fit_image(img, box_w, box_h):
    """将图像等比缩放放入 box_w x box_h 内(RGB)。"""
    from PIL import Image
    if img.mode != "RGB":
        img = img.convert("RGB")
    w, h = img.size
    if w <= 0 or h <= 0:
        return img
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
def compose_page(image_paths_batch, layout, printable_w, printable_h):
    """把一批照片合成一页图像(白底)。"""
    return _perf.measure("页面合成", PAGE_COMPOSE_TIME_LIMIT,
                         _compose_page_impl, image_paths_batch, layout, printable_w, printable_h)
def _compose_page_impl(image_paths_batch, layout, printable_w, printable_h):
    """页面合成实际逻辑，经缓存取图。"""
    from PIL import Image
    page = Image.new("RGB", (printable_w, printable_h), "white")
    rows, cols, _ = LAYOUTS[layout]
    gap = max(2, int(min(printable_w, printable_h) * 0.008))
    cell_w = (printable_w - gap * (cols + 1)) // cols
    cell_h = (printable_h - gap * (rows + 1)) // rows
    for idx, path in enumerate(image_paths_batch):
        fitted = _page_cache.get_fitted_image(path, (cell_w, cell_h))
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
    """执行打印。progress_cb(current, total) 用于回调进度。返回 (ok, msg)。"""
    from PIL import ImageWin
    import win32ui, win32con
    if not image_paths:
        return False, "没有可打印的照片"
    if not printer_name:
        return False, "未选择打印机"
    # 应用双面/纸张设置并保存原始值以便恢复
    original = apply_print_settings(printer_name, duplex_mode, paper_key)
    try:
        dc = win32ui.CreateDC()
        dc.CreatePrinterDC(printer_name)
        # 获取打印机可打印区域尺寸(像素)
        log_px_x = dc.GetDeviceCaps(win32con.LOGPIXELSX)
        printable_w = dc.GetDeviceCaps(win32con.HORZRES)
        printable_h = dc.GetDeviceCaps(win32con.VERTRES)
        if printable_w <= 0 or printable_h <= 0:
            printable_w = dc.GetDeviceCaps(win32con.PHYSICALWIDTH)
            printable_h = dc.GetDeviceCaps(win32con.PHYSICALHEIGHT)
        pages = build_pages(image_paths, layout)
        total_pages = len(pages) * copies
        done = 0
        doc_name = "照片打印 - %d张" % len(image_paths)
        dc.StartDoc(doc_name)
        try:
            for copy in range(copies):
                for batch in pages:
                    try:
                        page_img = compose_page(batch, layout, printable_w, printable_h)
                        dib = ImageWin.Dib(page_img)
                        dc.StartPage()
                        try:
                            dib.draw(dc.GetHandleOutput(),
                                     (0, 0, printable_w, printable_h))
                        finally:
                            dc.EndPage()
                    except Exception as e:
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
        restore_print_settings(printer_name, original)
    return True, "打印完成，共 %d 页" % total_pages
# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------
class PhotoPrintApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("920x640")
        self.root.minsize(760, 520)
        set_window_icon(self.root)
        # 初始化数据存储
        self.image_paths = []
        self.thumbnails = []
        self.printing = False
        self._preview_ready = False
        self._preview_img = []
        self._preview_cache = ImageCache()
        self._thumb_cache = ImageCache()
        self._page_cache = _page_cache
        self._perf = _perf
        self._build_ui()
        self.root.after(0, self._refresh_printers)
    # -- UI 构建 -----------------------------------------------------------
    def _apply_win10_style(self):
        """应用 Win10 风格: 系统默认主题(vista) + Segoe UI 字体 + 配色。"""
        try:
            import tkinter.font as tkfont
            for name in ("TkDefaultFont", "TkTextFont", "TkFixedFont",
                         "TkMenuFont", "TkHeadingFont", "TkCaptionFont",
                         "TkSmallCaptionFont", "TkIconFont"):
                tkfont.nametofont(name).configure(family="Segoe UI", size=9)
        except Exception:
            pass
        # 配色: 白色窗口背景 + 按钮文字 #333 + 打印按钮深灰底(#333333)
        try:
            style = ttk.Style(self.root)
            style.configure("TFrame", background="#FFFFFF")
            style.configure("TLabel", background="#FFFFFF", foreground="#333333")
            style.configure("TLabelframe", background="#FFFFFF")
            style.configure("TLabelframe.Label", background="#FFFFFF",
                            foreground="#333333")
            style.configure("TButton", foreground="#333333")
            style.configure("TButton", width=BUTTON_WIDTH)
            style.configure("TRadiobutton", background="#FFFFFF")
            style.configure("TCombobox", padding=COMBOBOX_PADDING)
            self.root.configure(bg="#FFFFFF")
        except Exception:
            pass
    def _build_ui(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("vista")
        except Exception:
            pass
        self._apply_win10_style()
        # 浅蓝提示横幅
        banner = tk.Frame(self.root, bg="#DDEBF7", height=30)
        banner.pack(side=tk.TOP, fill=tk.X)
        banner.pack_propagate(False)
        tk.Label(banner, bg="#DDEBF7", fg="#333333",
                 text="你希望如何打印图片？",
                 font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=14, pady=5)
        # 顶部工具栏: 添加/删除/清空按钮 + 打印机选择
        toolbar = ttk.Frame(self.root, padding=6)
        toolbar.pack(side=tk.TOP, fill=tk.X)
        self.btn_add = ttk.Button(toolbar, text="添加照片", command=self.add_photos)
        self.btn_add.pack(side=tk.LEFT, padx=2)
        self.btn_remove = ttk.Button(toolbar, text="删除选中", command=self.remove_selected)
        self.btn_remove.pack(side=tk.LEFT, padx=2)
        self.btn_clear = ttk.Button(toolbar, text="清空", command=self.clear_photos)
        self.btn_clear.pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=6)
        ttk.Label(toolbar, text="打印机:").pack(side=tk.LEFT, padx=(6, 2))
        self.cmb_printer = ttk.Combobox(toolbar, width=26, state="readonly")
        self.cmb_printer.pack(side=tk.LEFT, padx=8)
        self.btn_refresh = ttk.Button(toolbar, text="刷新", command=self._refresh_printers)
        self.btn_refresh.pack(side=tk.LEFT, padx=2)
        # 主区域: 左侧照片列表 + 右侧预览与设置
        main = ttk.Frame(self.root)
        main.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        left = ttk.Frame(main, width=220)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 6))
        left.pack_propagate(False)
        ttk.Label(left, text="照片列表").pack(anchor=tk.W, pady=(0, 4))
        list_frame = ttk.Frame(left)
        list_frame.pack(fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(list_frame, show="tree", selectmode="extended",
                                 height=12)
        self.tree.column("#0", width=200, stretch=True)
        vsb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL,
                            command=self.tree.yview)
        self.tree.config(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<<TreeviewSelect>>", self._on_list_select)
        # 右侧: 预览画布 + 设置面板
        right = ttk.Frame(main)
        right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        preview_frame = ttk.LabelFrame(right, text="布局预览", padding=4)
        preview_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 6))
        self.canvas = tk.Canvas(preview_frame, bg="#f0f0f0", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<Configure>", self._draw_preview)
        # 设置面板
        settings = ttk.LabelFrame(right, text="打印设置", padding=8)
        settings.pack(side=tk.TOP, fill=tk.X)
        # 纸张 / 布局 / 张数 行
        row1 = ttk.Frame(settings)
        row1.pack(fill=tk.X, pady=3)
        ttk.Label(row1, text="纸张:").pack(side=tk.LEFT)
        self.cmb_paper = ttk.Combobox(row1, values=list(PAPER_SIZES.keys()),
                                       width=12, state="readonly")
        self.cmb_paper.pack(side=tk.LEFT, padx=(8, 16))
        ttk.Label(row1, text="布局:").pack(side=tk.LEFT)
        self.cmb_layout = ttk.Combobox(row1, values=list(LAYOUTS.keys()),
                                        width=14, state="readonly")
        self.cmb_layout.pack(side=tk.LEFT, padx=(8, 16))
        self.cmb_layout.bind("<<ComboboxSelected>>", lambda e: self._draw_preview())
        ttk.Label(row1, text="张数:").pack(side=tk.LEFT)
        self.spin_copies = tk.Spinbox(row1, from_=1, to=99, width=5)
        self.spin_copies.pack(side=tk.LEFT, padx=(4, 0))
        # 单面 / 双面 选择行
        row2 = ttk.Frame(settings)
        row2.pack(fill=tk.X, pady=3)
        ttk.Label(row2, text="单双面:").pack(side=tk.LEFT)
        self.duplex_var = tk.IntVar(value=DUPLEX_NONE)
        self.rb_single = ttk.Radiobutton(row2, text="单面", variable=self.duplex_var,
                                         value=DUPLEX_NONE)
        self.rb_single.pack(side=tk.LEFT, padx=(6, 10))
        self.rb_duplex = ttk.Radiobutton(row2, text="双面(长边翻页)书籍",
                                         variable=self.duplex_var, value=DUPLEX_VERTICAL)
        self.rb_duplex.pack(side=tk.LEFT)
        self.rb_duplex_short = ttk.Radiobutton(row2, text="双面(短边翻页)日历",
                                                variable=self.duplex_var, value=DUPLEX_HORIZONTAL)
        self.rb_duplex_short.pack(side=tk.LEFT)
        # 底部: 进度条 + 状态文字 + 打印/关闭按钮
        bottom = ttk.Frame(self.root, padding=8)
        bottom.pack(side=tk.BOTTOM, fill=tk.X)
        prog_frame = tk.Frame(bottom, height=25)
        prog_frame.pack_propagate(False)
        prog_frame.pack(fill=tk.X, pady=(0, 6))
        self.progress = ttk.Progressbar(prog_frame, mode="determinate")
        self.progress.pack(fill=tk.BOTH, expand=True)
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(bottom, textvariable=self.status_var).pack(side=tk.LEFT)
        self.btn_print = ttk.Button(bottom, text="打印", command=self.start_print,
                                    default="active")
        self.btn_print.pack(side=tk.RIGHT, padx=(6, 0))
        self.btn_cancel = ttk.Button(bottom, text="关闭", command=self.close)
        self.btn_cancel.pack(side=tk.RIGHT)
        # 绑定快捷键与窗口关闭事件
        self.root.bind("<Escape>", lambda e: self.close)
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.root.after(0, self._deferred_init)
    def _deferred_init(self):
        # 设置下拉框默认值与初始张数
        self.cmb_paper.set("A4")
        self.cmb_layout.set(list(LAYOUTS.keys())[0])
        self.spin_copies.delete(0, "end")
        self.spin_copies.insert(0, "1")
    # -- 照片列表 ----------------------------------------------------------
    def add_photos(self):
        files = filedialog.askopenfilenames(
            title="选择照片",
            filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp *.gif *.tiff *.tif")]
        )
        if not files:
            return
        # 去重添加，避免重复文件
        for f in files:
            if f not in self.image_paths:
                self.image_paths.append(f)
        self._preview_cache.invalidate()
        self._thumb_cache.invalidate()
        self._page_cache.invalidate()
        self._refresh_list()
        self._draw_preview()
    def remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        # 倒序删除避免索引错位
        indices = [self.tree.index(item) for item in sel]
        for i in sorted(indices, reverse=True):
            del self.image_paths[i]
        self._preview_cache.invalidate()
        self._thumb_cache.invalidate()
        self._page_cache.invalidate()
        self._refresh_list()
        self._draw_preview()
    def clear_photos(self):
        if not self.image_paths:
            return
        if messagebox.askyesno("确认", "清空所有照片?"):
            self.image_paths = []
            self._preview_cache.invalidate()
            self._thumb_cache.invalidate()
            self._page_cache.invalidate()
            self._refresh_list()
            self._draw_preview()
    def _refresh_list(self):
        self.tree.delete(*self.tree.get_children())
        self.thumbnails = []
        return self._perf.measure("列表刷新", LIST_REFRESH_TIME_LIMIT,
                                  self._refresh_list_impl)
    def _refresh_list_impl(self):
        for i, p in enumerate(self.image_paths):
            name = os.path.basename(p)
            im = self._thumb_cache.get_thumbnail(p, (48, 48))
            if im is not None:
                from PIL import ImageTk
                th = ImageTk.PhotoImage(im)
                self.thumbnails.append(th)
            else:
                th = None
                self.thumbnails.append(None)
            self.tree.insert("", "end", text="%d. %s" % (i + 1, name),
                             image=th, open=True)
    def _on_list_select(self, _event):
        pass
    # -- 预览 --------------------------------------------------------------
    def _draw_paper_frame(self):
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        if w <= 4 or h <= 4:
            return
        paper_w_ratio, paper_h_ratio = 210, 297
        pad = 16
        avail_w = w - pad * 2
        avail_h = h - pad * 2
        if avail_w <= 0 or avail_h <= 0:
            return
        scale = min(avail_w / paper_w_ratio, avail_h / paper_h_ratio)
        pw = paper_w_ratio * scale
        ph = paper_h_ratio * scale
        px = (w - pw) / 2
        py = (h - ph) / 2
        self.canvas.create_rectangle(px, py, px + pw, py + ph,
                                     fill="white", outline="#888888", width=2)
    def _draw_preview(self, _event=None):
        self.canvas.delete("all")
        self._preview_img = []
        if not self._preview_ready:
            self._preview_ready = True
            self._draw_paper_frame()
            self.root.after(0, self._draw_preview)
            return
        return self._perf.measure("预览重绘", PREVIEW_REDRAW_TIME_LIMIT,
                                  self._draw_preview_impl)
    def _draw_preview_impl(self):
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        if w <= 4 or h <= 4:
            return
        layout = self.cmb_layout.get()
        if layout not in LAYOUTS:
            return
        rows, cols, count = LAYOUTS[layout]
        paper_w_ratio, paper_h_ratio = 210, 297
        pad = 16
        avail_w = w - pad * 2
        avail_h = h - pad * 2
        if avail_w <= 0 or avail_h <= 0:
            return
        scale = min(avail_w / paper_w_ratio, avail_h / paper_h_ratio)
        pw = paper_w_ratio * scale
        ph = paper_h_ratio * scale
        px = (w - pw) / 2
        py = (h - ph) / 2
        self.canvas.create_rectangle(px, py, px + pw, py + ph,
                                     fill="white", outline="#888888", width=2)
        gap = 4
        cw = (pw - gap * (cols + 1)) / cols
        ch = (ph - gap * (rows + 1)) / rows
        for r in range(rows):
            for c in range(cols):
                x = px + gap + c * (cw + gap)
                y = py + gap + r * (ch + gap)
                self.canvas.create_rectangle(x, y, x + cw, y + ch,
                                             fill="#e8e8e8", outline="#bbbbbb")
                idx = r * cols + c
                if idx < len(self.image_paths):
                    im = self._preview_cache.get_preview_image(
                        self.image_paths[idx], (int(cw), int(ch)))
                    if im is not None:
                        from PIL import ImageTk
                        photo = ImageTk.PhotoImage(im)
                        self._preview_img.append(photo)
                        self.canvas.create_image(x + cw / 2, y + ch / 2,
                                                 image=photo)
                    else:
                        self.canvas.create_text(x + cw / 2, y + ch / 2,
                                                text=str(idx + 1), fill="#999999")
                else:
                    self.canvas.create_text(x + cw / 2, y + ch / 2,
                                            text=str(idx + 1), fill="#999999")
        self.canvas.create_text(w / 2, h - 6,
                                text="布局: %s  |  照片: %d 张"
                                     % (layout, len(self.image_paths)),
                                fill="#666666", anchor=tk.S)
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
        # 禁用控件并启动后台打印线程
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
                  self.btn_print, self.btn_refresh, self.cmb_layout,
                  self.cmb_paper, self.cmb_printer, self.spin_copies,
                  self.rb_single, self.rb_duplex, self.rb_duplex_short):
            try:
                w.config(state=state)
            except Exception:
                pass
    def close(self):
        if self.printing:
            if not messagebox.askyesno("确认", "正在打印中，确定退出?"):
                return
        self.root.destroy()
# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------
def _create_splash(root):
    """创建并显示启动界面，返回 splash 窗口。"""
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.configure(bg="#333333")
    splash.attributes("-topmost", True)
    w, h = 380, 220
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    splash.geometry("%dx%d+%d+%d" % (w, h, x, y))
    tk.Label(splash, text="照片打印工具", bg="#333333", fg="#FFFFFF",
             font=("Segoe UI", 18, "bold")).pack(expand=1)
    tk.Label(splash, text="正在启动...", bg="#333333", fg="#AAAAAA",
             font=("Segoe UI", 9)).pack(side=tk.BOTTOM, pady=(0, 16))
    splash.update_idletasks()
    return splash
def safe_build():
    """构建安全机制: 输出至临时目录, 成功替换 dist, 失败保留原产物, 度量体积写日志。"""
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
    ok = size <= EXE_SIZE_LIMIT
    log("exe 体积: %d 字节 (上限 %d 字节) %s; 基准 %d 字节; 缩减 %.2f%% (下限 %.1f%%) %s"
        % (size, EXE_SIZE_LIMIT, "达标" if ok else "超标",
           CURRENT_EXE_SIZE_BASELINE, ratio, VOLUME_REDUCTION_RATIO,
           "达标" if ratio >= VOLUME_REDUCTION_RATIO else "未达标"))
    if os.path.exists(dist_dir):
        shutil.rmtree(dist_dir, ignore_errors=True)
    shutil.move(build_tmp, dist_dir)
    log("构建成功: dist 已更新")
    return True
def main():
    try:
        t0 = time.time()
        root = tk.Tk()
        root.withdraw()
        splash = _create_splash(root)
        try:
            app = PhotoPrintApp(root)
        finally:
            splash.destroy()
        root.deiconify()
        t1 = time.time()
        elapsed = t1 - t0
        log("启动耗时: %.3f 秒 (上限 %d 秒)" % (elapsed, STARTUP_TIME_LIMIT))
        root.mainloop()
    except Exception:
        log("主循环异常:\n%s" % traceback.format_exc())
        try:
            messagebox.showerror("错误", traceback.format_exc())
        except Exception:
            pass
if __name__ == "__main__":
    main()
