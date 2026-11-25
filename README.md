# 照片打印工具

类似 Windows 系统自带的照片打印，支持选择照片、选择打印机、纸张、布局、张数，并新增 **单面 / 双面** 打印选项。启动时显示启动界面，打印按钮为默认按钮（回车触发）。

## 功能特性

- **照片选择**：多选添加照片，左侧缩略图列表（Treeview），支持删除选中、清空
  - 支持格式：`.jpg .jpeg .png .bmp .gif .tiff .tif`
- **打印机选择**：自动枚举本机打印机，默认打印机置顶，可刷新
- **纸张**：A4 / Letter / A5 / 4×6 照片纸 / 5×7 照片纸
- **布局**（每页张数）：
  - 全页（1 张/页，1×1）
  - 2 张/页（1×2）
  - 4 张/页（2×2）
  - 9 张/页（3×3）
  - 16 张/页（4×4）
- **张数**：1–99
- **单双面**（核心新增）：
  - 单面
  - 双面(长边翻页)书籍
  - 双面(短边翻页)日历
- **布局预览**：右侧画布按 A4 纵向比例（210×297）实时显示网格与照片排布
- **打印进度**：后台线程执行，进度条 + 状态文本，不阻塞界面
- **启动界面**：启动时显示深灰底启动画面，主窗口构建完成后自动切换
- **默认打印按钮**：打印按钮为窗口默认按钮，回车键直接触发打印
- **日志**：运行日志写入同目录 `print_log.txt`

## 界面布局

```
+-----------------------------------------------------------+
|  你希望如何打印图片？                                        |
+-----------------------------------------------------------+
|  [添加照片] [删除选中] [清空]  打印机:[▼] [刷新]            |
+--------------------------+--------------------------------+
| 照片列表(缩略图)          |  布局预览                       |
| (Treeview)               |  (A4 比例画布)                  |
|                          |                                 |
|                          |  +----------------------------+ |
|                          |  | 打印设置                    | |
|                          |  | 纸张:[▼] 布局:[▼] 张数:[1] | |
|                          |  | 单双面:(•单面)(○长边)(○短边)| |
|                          |  +----------------------------+ |
+--------------------------+--------------------------------+
|  [████████████] 状态: 就绪        [关闭] [打印]            |
+-----------------------------------------------------------+
```

## 运行环境

- Windows 系统
- Python 3（依赖 `pywin32`、`Pillow`）

### 依赖安装

```bash
pip install pywin32 pillow
```

## 运行方式

### 方式一：直接运行源码

```bash
python print_photo.py
```

### 方式二：运行打包后的 exe

双击 `dist\照片打印工具.exe` 即可（单文件，无需安装 Python）。

## 打包为 exe

```bash
pyinstaller --onefile --windowed ^
    --hidden-import=win32print --hidden-import=win32ui ^
    --hidden-import=win32con --hidden-import=win32api ^
    --hidden-import=PIL.ImageWin ^
    --name "照片打印工具" print_photo.py
```

产物位于 `dist\照片打印工具.exe`。

## 技术实现说明

### 启动流程与启动界面

- `main()` 创建主 `root` 后立即 `withdraw()` 隐藏，调用 `_create_splash()` 显示无边框启动界面（深灰底 #333333，居中，380×220）
- 启动界面显示期间构建 `PhotoPrintApp` 主界面，构建完成后 `splash.destroy()` + `root.deiconify()` 切换主窗口
- `root.after(0, self._deferred_init)` 延后设置纸张/布局/张数默认值，避免阻塞首屏
- 启动耗时记录到日志（上限 `STARTUP_TIME_LIMIT` 秒）

### 延迟导入优化

- `PIL`（Image/ImageTk/ImageWin）与 `win32`（win32print/win32ui/win32con）均在函数内部按需导入，启动时不加载这两个重模块，缩短启动时间
- 模块级仅保留 `os`、`sys`、`threading`、`time`、`traceback`、`tkinter` 等轻量导入
- 打包时需 `--hidden-import` 显式声明这些延迟导入模块

### 双面打印实现

双面通过修改打印机 DEVMODE 实现（`apply_print_settings` / `restore_print_settings`）：

- 写入 `devmode.Duplex` 并置位 `DM_DUPLEX` 字段标志：
  - `DUPLEX_NONE = 1` → 单面
  - `DUPLEX_VERTICAL = 2` → 双面长边翻页（书籍）
  - `DUPLEX_HORIZONTAL = 3` → 双面短边翻页（日历）
- 纸张设置：标准纸张用 `PaperSize`（DMPAPER 常量），照片纸用 `PaperWidth`/`PaperLength`（单位 0.1mm）
- **打印完成后自动还原打印机原始设置**，不污染全局配置

### 打印流程（`do_print`）

1. `apply_print_settings`：写入双面/纸张设置，保存原始值
2. `win32ui.CreateDC` + `CreatePrinterDC` 创建打印机 DC
3. `GetDeviceCaps` 获取可打印区尺寸（HORZRES/VERTRES）与 DPI
4. 按布局将照片分组为每页一批（`build_pages`）
5. 每页用 Pillow 合成白底页面（`compose_page`，等比缩放居中）
6. `ImageWin.Dib` 转为 DIB 绘制到 DC（`StartDoc` → `StartPage` → `dib.draw` → `EndPage` → `EndDoc`）
7. 异常时 `AbortDoc`
8. `restore_print_settings`：还原打印机设置

### 页面合成

- `fit_image`：等比缩放放入单元格（LANCZOS，降级 ANTIALIAS）
- `compose_page`：白底，按行列网格排布，单元格间留间隙，照片居中
- 合成分辨率取打印机可打印区实际像素，保证输出质量

### 配色与样式

- 系统主题 `vista` + Segoe UI 字体
- 窗口背景白色 #FFFFFF，标签/标签框文字 #333333
- 按钮文字 #333333，统一宽度 `BUTTON_WIDTH=12`；打印按钮深灰底 #333333 + 白字，`default="active"`
- 单选按钮背景白色 #FFFFFF
- 选择框内边距 `COMBOBOX_PADDING=(4, 4)`
- 进度条置于固定高度 25px 容器框中拉伸填充

## 代码结构

| 模块 | 说明 |
|------|------|
| `resource_path` | 解析资源路径，兼容打包（`sys._MEIPASS`）与开发模式 |
| `set_window_icon` | 设置窗口图标（打包后从 exe 资源取，开发时读 `app.ico`） |
| `log` | 日志写入 `print_log.txt` |
| `list_printers` | 枚举打印机，默认置顶 |
| `_get_devmode` | 打开打印机获取 DEVMODE |
| `apply_print_settings` / `restore_print_settings` | 双面/纸张设置写入与还原 |
| `fit_image` / `compose_page` / `build_pages` | 图像缩放与页面合成 |
| `do_print` | 执行打印，返回 `(ok, msg)` |
| `_create_splash` | 创建并显示启动界面 |
| `PhotoPrintApp` | tkinter GUI 类（含 `_build_ui`、`_draw_preview`、`start_print` 等） |
| `main` | 入口（启动界面 + 主窗口 + 主循环） |

## 快捷键

- `Enter`：触发打印（打印按钮为默认按钮）
- `Esc`：关闭程序
- 关闭时若正在打印，需确认后退出