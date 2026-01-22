# -*- coding: utf-8 -*-
"""Extensions tab for settings dialog."""

import sys
import os
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox, filedialog, simpledialog
from typing import Optional

from ...i18n import t
from ...config.defaults import RESERVED_APPS
from ...config.paths import get_app_icon_path
from ...utils.dpi import get_dpi_scale
from ...utils.system_detect import is_macos, is_windows
from ...utils.logging import log


def _get_app_id(app_info: dict) -> str:
    app_id = app_info.get("id") or ""
    return app_id.lower() if isinstance(app_id, str) else ""


class WorkflowSection:
    """单个工作流配置区域（作为 Tab 内容）"""
    
    def __init__(
        self, 
        parent: ttk.Frame, 
        workflow_key: str,
        enable_key: str,
        config: dict,
        has_keep_latex: bool = False,
        check_app_conflict = None,
    ):
        self.parent = parent
        self.workflow_key = workflow_key
        self.config = config
        self.has_keep_latex = has_keep_latex
        self.check_app_conflict = check_app_conflict
        
        # 存储应用数据 {iid: {"name": ..., "id": ..., "window_patterns": [...]}}
        self.app_data: dict[str, dict] = {}
        
        # 图标缓存
        self._icons: list = []
        
        # 使用普通 Frame 作为 Tab 内容
        self.frame = ttk.Frame(parent, padding=10)
        
        self._create_widgets(enable_key)
    
    def _create_widgets(self, enable_key: str):
        """创建 UI 组件"""
        self.frame.columnconfigure(1, weight=1)
        
        # 启用开关
        self.enabled_var = tk.BooleanVar(value=self.config.get("enabled", True))
        ttk.Checkbutton(
            self.frame, 
            text=t(enable_key),
            variable=self.enabled_var
        ).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=2)
        
        # 应用列表
        self._create_app_treeview()
        
        # 公式格式开关（仅 HTML 工作流）
        next_row = 2
        if self.has_keep_latex:
            self.keep_latex_var = tk.BooleanVar(
                value=self.config.get("keep_formula_latex", True)
            )
            ttk.Checkbutton(
                self.frame, 
                text=t("settings.extensions.keep_latex"),
                variable=self.keep_latex_var
            ).grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
            next_row = 4
        
        if self.workflow_key == "md":
            html_formatting = self.config.get("html_formatting", {})
            self.css_font_to_semantic_var = tk.BooleanVar(
                value=html_formatting.get("css_font_to_semantic", True)
            )
            self.bold_first_row_to_header_var = tk.BooleanVar(
                value=html_formatting.get("bold_first_row_to_header", True)
            )
            ttk.Checkbutton(
                self.frame,
                text=t("settings.conversion.css_font_to_semantic"),
                variable=self.css_font_to_semantic_var,
            ).grid(row=next_row, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
            ttk.Checkbutton(
                self.frame,
                text=t("settings.conversion.bold_first_row_to_header"),
                variable=self.bold_first_row_to_header_var,
            ).grid(row=next_row + 1, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
    
    def _create_app_treeview(self):
        """创建应用列表 Treeview"""
        list_frame = ttk.Frame(self.frame)
        list_frame.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=5)
        list_frame.columnconfigure(0, weight=1)
        scale = get_dpi_scale()
        column_scale = scale
        icon_scale = scale
        name_stretch = True
        pattern_stretch = False
        if is_macos():
            # macOS 字体和 DPI 偏大，收窄列宽/行高，避免挤压窗口名称列
            column_scale = min(scale, 1.25)
            icon_scale = min(scale, 1.25)
            name_stretch = False
            pattern_stretch = True
        icon_size = max(16, int(16 * icon_scale))
        base_font = tkfont.nametofont("TkDefaultFont")
        row_height = max(icon_size + 6, base_font.metrics("linespace") + 6)
        style = ttk.Style(self.frame)
        style.configure("Apps.Treeview", rowheight=row_height)
        
        # Treeview 列：图标、应用名、窗口模式
        columns = ("window_patterns",)
        self.treeview = ttk.Treeview(
            list_frame, 
            columns=columns, 
            show="tree headings",
            height=4,
            style="Apps.Treeview",
        )
        self.treeview.heading("#0", text=t("settings.extensions.app_name"))
        self.treeview.heading("window_patterns", text=t("settings.extensions.window_pattern"))
        name_col_width = int(220 * column_scale)
        name_minwidth = int(160 * column_scale)
        if is_macos():
            name_col_width = int(200 * column_scale)
            name_minwidth = int(140 * column_scale)
        self.treeview.column(
            "#0",
            width=name_col_width,
            minwidth=name_minwidth,
            stretch=name_stretch,
        )
        pattern_col_width = int(200 * column_scale)
        self.treeview.column(
            "window_patterns",
            width=pattern_col_width,
            minwidth=int(120 * column_scale),
            stretch=pattern_stretch,
        )
        
        # 双击编辑窗口模式
        self.treeview.bind("<Double-1>", self._on_double_click)
        
        # 加载已保存的应用
        for app in self.config.get("apps", []):
            if isinstance(app, dict):
                name = app.get("name", "")
                patterns = app.get("window_patterns", [])
                app_id = app.get("id", "")
                if isinstance(app_id, str):
                    app_id = app_id.lower()
            else:
                name = str(app)
                patterns = []
                app_id = ""
            path = ""
            if app_id:
                if is_macos():
                    path = self._get_macos_app_path(app_id)
                elif is_windows():
                    path = app_id

            patterns_str = ", ".join(patterns) if patterns else ""
            icon = None
            if path:
                icon = self._extract_icon(path)
            iid = self.treeview.insert("", tk.END, text=name, values=(patterns_str,), image=icon or "")
            self.app_data[iid] = {
                "name": name,
                "window_patterns": patterns,
                "id": app_id,
            }
        
        self.treeview.grid(row=0, column=0, sticky=tk.EW)
        
        # 按钮栏
        btn_frame = ttk.Frame(list_frame)
        btn_frame.grid(row=0, column=1, sticky=tk.N, padx=(5, 0))
        
        ttk.Button(btn_frame, text="+", command=self._add_app, width=3).pack(pady=2)
        ttk.Button(btn_frame, text="-", command=self._remove_app, width=3).pack(pady=2)
        ttk.Button(btn_frame, text="✎", command=self._edit_patterns, width=3).pack(pady=2)
    
    def _on_double_click(self, event):
        """双击编辑窗口模式"""
        region = self.treeview.identify("region", event.x, event.y)
        if region in ("cell", "tree"):
            row_id = self.treeview.identify_row(event.y)
            if not row_id:
                return
            self.treeview.selection_set(row_id)
            self._edit_patterns()
    
    def _edit_patterns(self):
        """编辑选中应用的窗口匹配模式"""
        selection = self.treeview.selection()
        if not selection:
            return
        
        iid = selection[0]
        app_info = self.app_data.get(iid, {})
        current_patterns = app_info.get("window_patterns", [])
        
        # 弹出编辑对话框（换行分隔）
        current_text = "\n".join(current_patterns)
        dialog = tk.Toplevel(self.frame)
        dialog.title(t("settings.extensions.edit_window_pattern"))
        dialog.transient(self.frame)
        scale = get_dpi_scale()
        dialog_width = int(420 * scale)
        dialog_height = int(260 * scale)
        dialog.geometry(f"{dialog_width}x{dialog_height}")
        dialog.minsize(dialog_width, dialog_height)
        try:
            icon_path = get_app_icon_path()
            if os.path.exists(icon_path):
                dialog.iconbitmap(icon_path)
        except Exception as e:
            log(f"Failed to set edit pattern dialog icon: {e}")
        
        ttk.Label(
            dialog, 
            text=t("settings.extensions.window_pattern_hint"),
            wraplength=380
        ).pack(pady=5, padx=10)
        
        text_widget = tk.Text(dialog, height=5, width=50)
        text_widget.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        text_widget.insert("1.0", current_text)
        
        def save():
            new_text = text_widget.get("1.0", tk.END).strip()
            new_patterns = [p.strip() for p in new_text.split("\n") if p.strip()]
            app_info["window_patterns"] = new_patterns
            patterns_str = ", ".join(new_patterns) if new_patterns else ""
            self.treeview.set(iid, "window_patterns", patterns_str)
            dialog.destroy()
        
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text=t("settings.buttons.cancel"), command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text=t("settings.buttons.save"), command=save).pack(side=tk.RIGHT, padx=5)
        
        # 延迟 grab_set 避免与热键监听冲突
        dialog.after(100, lambda: dialog.grab_set())
        dialog.wait_window()
    
    def _add_app(self):
        """添加应用（平台特定）"""
        if is_macos():
            self._add_app_macos()
        elif is_windows():
            self._add_app_windows()
        else:
            messagebox.showinfo(
                t("settings.title.info"),
                t("settings.extensions.unsupported_platform")
            )
    
    def _add_app_macos(self):
        """macOS: 使用原生对话框选择 .app"""
        try:
            from AppKit import NSOpenPanel, NSURL
            
            panel = NSOpenPanel.openPanel()
            panel.setTitle_(t("settings.extensions.select_app"))
            panel.setCanChooseFiles_(True)
            panel.setCanChooseDirectories_(False)
            panel.setAllowsMultipleSelection_(False)
            panel.setDirectoryURL_(NSURL.fileURLWithPath_("/Applications"))
            panel.setAllowedFileTypes_(["app"])
            
            result = panel.runModal()
            
            if result == 1:
                url = panel.URL()
                if url:
                    path = url.path()
                    app_name = os.path.basename(path).replace(".app", "")
                    bundle_id = self._get_macos_bundle_id(path)
                    icon = self._extract_icon(path)
                    self._add_app_to_list(app_name, bundle_id, icon, path)
        except ImportError:
            self._add_app_fallback()
        except Exception as e:
            log(f"Failed to open macOS app selector: {e}")
            self._add_app_fallback()
    
    def _add_app_fallback(self):
        """回退方案 - 手动输入应用名称"""
        app_name = simpledialog.askstring(
            t("settings.extensions.select_app"),
            t("settings.extensions.enter_app_name"),
            parent=self.frame
        )
        if app_name:
            app_path = f"/Applications/{app_name}.app"
            if not os.path.exists(app_path):
                app_path = ""
            bundle_id = self._get_macos_bundle_id(app_path) if app_path else ""
            icon = self._extract_icon(app_path) if app_path else None
            self._add_app_to_list(app_name, bundle_id, icon, app_path)
    
    def _add_app_windows(self):
        """Windows: 弹窗显示运行中的应用供选择"""
        try:
            from ...utils.win32.window import get_running_apps
        except ImportError:
            messagebox.showerror(
                t("settings.title.error"),
                t("settings.extensions.win32_import_error")
            )
            return
        
        running_apps = get_running_apps()
        apps_with_icons = []
        for app in running_apps:
            if app["name"].lower() not in RESERVED_APPS:
                app["icon"] = self._extract_icon(app.get("exe_path", ""))
                apps_with_icons.append(app)
        
        if not apps_with_icons:
            messagebox.showinfo(
                t("settings.title.info"), 
                t("settings.extensions.no_running_apps")
            )
            return
        
        selected = self._show_windows_app_selector(apps_with_icons)
        if selected:
            self._add_app_to_list(
                selected["name"], 
                selected.get("exe_path", "").lower(), 
                selected.get("icon"),
                selected.get("exe_path", ""),
            )
    
    def _show_windows_app_selector(self, apps: list) -> Optional[dict]:
        """显示应用选择对话框"""
        dialog = tk.Toplevel(self.frame)
        dialog.title(t("settings.extensions.select_app"))
        dialog.transient(self.frame)
        scale = get_dpi_scale()
        dialog_width = int(320 * scale)
        dialog_height = int(440 * scale)
        dialog.geometry(f"{dialog_width}x{dialog_height}")
        dialog.minsize(dialog_width, dialog_height)
        try:
            icon_path = get_app_icon_path()
            if os.path.exists(icon_path):
                dialog.iconbitmap(icon_path)
        except Exception as e:
            log(f"Failed to set app selector icon: {e}")
        
        # 延迟 grab_set 避免与输入法切换冲突
        dialog.after(100, lambda: dialog.grab_set())
        
        scale = get_dpi_scale()
        icon_size = max(16, int(16 * scale))
        base_font = tkfont.nametofont("TkDefaultFont")
        row_height = max(icon_size + 6, base_font.metrics("linespace") + 6)
        style = ttk.Style(dialog)
        style.configure("AppSelector.Treeview", rowheight=row_height)

        tree = ttk.Treeview(dialog, show="tree", height=15, style="AppSelector.Treeview")
        tree.heading("#0", text="")
        tree.column("#0", width=250, stretch=True)
        
        for app in apps:
            tree.insert("", tk.END, text=app["name"], image=app.get("icon") or "")
        
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        selected = [None]
        
        def on_select():
            sel = tree.selection()
            if sel:
                idx = tree.index(sel[0])
                selected[0] = apps[idx]
            dialog.destroy()
        
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text=t("settings.buttons.cancel"), command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text=t("settings.buttons.select"), command=on_select).pack(side=tk.RIGHT, padx=5)
        dialog.wait_window()
        return selected[0]
    
    def _add_app_to_list(self, app_name: str, app_id: str, icon=None, app_path: str = ""):
        """将应用添加到列表"""
        if app_name.lower() in RESERVED_APPS:
            messagebox.showerror(
                t("settings.title.error"), 
                t("settings.extensions.reserved_app_error", app=app_name)
            )
            return
        
        # 检查当前工作流中是否已存在
        new_app_info = {
            "name": app_name,
            "window_patterns": [],
            "id": app_id.lower() if isinstance(app_id, str) else "",
        }
        new_id = _get_app_id(new_app_info)
        existing = [_get_app_id(self.app_data[iid]) for iid in self.treeview.get_children()]
        if new_id in existing:
            messagebox.showinfo(
                t("settings.title.info"),
                t("settings.extensions.app_exists", app=app_name)
            )
            return
        
        # 检查跨工作流冲突（提醒并确认是否继续）
        if self.check_app_conflict:
            conflict_workflow = self.check_app_conflict(
                new_id, app_name, self.workflow_key
            )
            if conflict_workflow:
                if not messagebox.askyesno(
                    t("settings.title.warning"),
                    t("settings.extensions.app_conflict_error", app=app_name, workflow=conflict_workflow),
                ):
                    return
        
        iid = self.treeview.insert("", tk.END, text=app_name, values=("",), image=icon or "")
        self.app_data[iid] = new_app_info
    
    def _remove_app(self):
        """移除选中的应用"""
        selection = self.treeview.selection()
        if selection:
            iid = selection[0]
            if iid in self.app_data:
                del self.app_data[iid]
            self.treeview.delete(iid)
    
    def _extract_icon(self, path: str):
        """提取应用图标（平台特定）"""
        if not path:
            return None
        
        if is_macos():
            return self._extract_macos_icon(path)
        elif is_windows():
            return self._extract_windows_icon(path)
        return None

    def _get_macos_bundle_id(self, app_path: str) -> str:
        """macOS: 从 .app 获取 bundle id"""
        if not app_path:
            return ""
        try:
            from AppKit import NSBundle
            bundle = NSBundle.bundleWithPath_(app_path)
            if bundle:
                bundle_id = str(bundle.bundleIdentifier() or "")
                return bundle_id.lower() if bundle_id else ""
        except Exception as e:
            log(f"Failed to get bundle id: {e}")
        return ""

    def _get_macos_app_path(self, bundle_id: str) -> str:
        """macOS: 从 bundle id 获取 .app 路径"""
        if not bundle_id:
            return ""
        try:
            from AppKit import NSWorkspace
            ws = NSWorkspace.sharedWorkspace()
            url = ws.URLForApplicationWithBundleIdentifier_(bundle_id)
            if url:
                return str(url.path())
        except Exception as e:
            log(f"Failed to get app path from bundle id: {e}")
        return ""
    
    def _extract_macos_icon(self, app_path: str):
        """macOS: 从 .app 提取图标"""
        try:
            from AppKit import NSWorkspace, NSBitmapImageRep, NSPNGFileType, NSImage
            from PIL import Image, ImageTk
            import io
            
            ws = NSWorkspace.sharedWorkspace()
            ns_icon = ws.iconForFile_(app_path)
            
            if ns_icon:
                size = 16
                new_image = NSImage.alloc().initWithSize_((size, size))
                new_image.lockFocus()
                ns_icon.drawInRect_fromRect_operation_fraction_(
                    ((0, 0), (size, size)),
                    ((0, 0), ns_icon.size()),
                    2, 1.0
                )
                bitmap_rep = NSBitmapImageRep.alloc().initWithFocusedViewRect_(
                    ((0, 0), (size, size))
                )
                new_image.unlockFocus()
                
                if bitmap_rep:
                    png_data = bitmap_rep.representationUsingType_properties_(
                        NSPNGFileType, None
                    )
                    if png_data:
                        img = Image.open(io.BytesIO(bytes(png_data)))
                        img = img.resize((size, size), Image.Resampling.LANCZOS)
                        photo = ImageTk.PhotoImage(img)
                        self._icons.append(photo)
                        return photo
        except Exception as e:
            log(f"Failed to extract macOS icon: {e}")
        return None
    
    def _extract_windows_icon(self, exe_path: str):
        """Windows: 从 .exe 提取图标"""
        try:
            import win32gui
            import win32con
            import win32ui
            from PIL import Image, ImageTk
            
            icon_size = max(16, int(16 * get_dpi_scale()))
            ico_x = win32gui.ExtractIcon(0, exe_path, 0)
            if ico_x:
                hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
                hbmp = win32ui.CreateBitmap()
                hbmp.CreateCompatibleBitmap(hdc, icon_size, icon_size)
                
                hdc_mem = hdc.CreateCompatibleDC()
                hdc_mem.SelectObject(hbmp)
                win32gui.DrawIconEx(
                    hdc_mem.GetHandleOutput(),
                    0,
                    0,
                    ico_x,
                    icon_size,
                    icon_size,
                    0,
                    None,
                    win32con.DI_NORMAL,
                )
                
                bmpinfo = hbmp.GetInfo()
                bmpstr = hbmp.GetBitmapBits(True)
                img = Image.frombuffer(
                    'RGBA',
                    (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
                    bmpstr, 'raw', 'BGRA', 0, 1
                )
                if img.size != (icon_size, icon_size):
                    img = img.resize((icon_size, icon_size), Image.Resampling.LANCZOS)
                
                win32gui.DestroyIcon(ico_x)
                hdc_mem.DeleteDC()
                hdc.DeleteDC()
                
                photo = ImageTk.PhotoImage(img)
                self._icons.append(photo)
                return photo
        except Exception as e:
            log(f"Failed to extract Windows icon: {e}")
        return None
    
    def get_config(self) -> dict:
        """获取当前配置"""
        apps = []
        for iid in self.treeview.get_children():
            app_info = self.app_data[iid]
            apps.append(
                {
                    "name": app_info.get("name", ""),
                    "id": app_info.get("id", ""),
                    "window_patterns": app_info.get("window_patterns", []),
                }
            )
        config = {
            "enabled": self.enabled_var.get(),
            "apps": apps,
        }
        if self.has_keep_latex:
            config["keep_formula_latex"] = self.keep_latex_var.get()
        if hasattr(self, "css_font_to_semantic_var") or hasattr(self, "bold_first_row_to_header_var"):
            html_formatting = {}
            if hasattr(self, "css_font_to_semantic_var"):
                html_formatting["css_font_to_semantic"] = self.css_font_to_semantic_var.get()
            if hasattr(self, "bold_first_row_to_header_var"):
                html_formatting["bold_first_row_to_header"] = self.bold_first_row_to_header_var.get()
            config["html_formatting"] = html_formatting
        return config


class ExtensionsTab:
    """扩展设置选项卡
    
    管理可扩展工作流的设置：HTML 粘贴和 MD 粘贴。
    """
    
    def __init__(self, notebook: ttk.Notebook, config: dict):
        self.notebook = notebook
        self.config = config
        self.frame = ttk.Frame(notebook, padding=10)
        
        self._create_widgets()
        notebook.add(self.frame, text=t("settings.tab.extensions"))
    
    def _create_widgets(self):
        """创建 UI 组件"""
        ext_config = self.config.get("extensible_workflows", {})
        
        # 内部 Notebook 用于切换 HTML / MD / LaTeX 工作流
        self.inner_notebook = ttk.Notebook(self.frame)
        self.inner_notebook.pack(fill=tk.BOTH, expand=True)
        
        # HTML 工作流配置
        self.html_section = WorkflowSection(
            self.inner_notebook,
            workflow_key="html",
            enable_key="settings.extensions.html_enable",
            config=ext_config.get("html", {}),
            has_keep_latex=True,
            check_app_conflict=self._check_app_conflict,
        )
        self.inner_notebook.add(
            self.html_section.frame, 
            text=t("settings.extensions.html_title")
        )
        
        # MD 工作流配置
        self.md_section = WorkflowSection(
            self.inner_notebook,
            workflow_key="md",
            enable_key="settings.extensions.md_enable",
            config=ext_config.get("md", {}),
            has_keep_latex=False,
            check_app_conflict=self._check_app_conflict,
        )
        self.inner_notebook.add(
            self.md_section.frame, 
            text=t("settings.extensions.md_title")
        )
        
        # LaTeX 工作流配置
        self.latex_section = WorkflowSection(
            self.inner_notebook,
            workflow_key="latex",
            enable_key="settings.extensions.latex_enable",
            config=ext_config.get("latex", {}),
            has_keep_latex=False,
            check_app_conflict=self._check_app_conflict,
        )
        self.inner_notebook.add(
            self.latex_section.frame, 
            text=t("settings.extensions.latex_title")
        )

        # 文件粘贴工作流配置
        self.file_section = WorkflowSection(
            self.inner_notebook,
            workflow_key="file",
            enable_key="settings.extensions.file_enable",
            config=ext_config.get("file", {}),
            has_keep_latex=False,
            check_app_conflict=self._check_app_conflict,
        )
        self.inner_notebook.add(
            self.file_section.frame,
            text=t("settings.extensions.file_title")
        )
        
        # 说明文字
        ttk.Label(
            self.frame,
            text=t("settings.extensions.description"),
            foreground="gray",
            wraplength=400
        ).pack(pady=(10, 0), anchor=tk.W)
    
    def _check_app_conflict(self, app_id: str, app_name: str, current_workflow: str) -> Optional[str]:
        """检查应用是否在其他工作流中存在
        
        Args:
            app_id: 应用标识
            app_name: 应用名称（仅用于提示）
            current_workflow: 当前工作流键名
            
        Returns:
            如果存在冲突，返回冲突的工作流名称；否则返回 None
        """
        workflow_sections = {
            "html": (self.html_section, t("settings.extensions.html_title")),
            "md": (self.md_section, t("settings.extensions.md_title")),
            "latex": (self.latex_section, t("settings.extensions.latex_title")),
            "file": (self.file_section, t("settings.extensions.file_title")),
        }
        
        for workflow_key, (section, workflow_name) in workflow_sections.items():
            if workflow_key == current_workflow:
                continue
            
            # 检查该工作流中是否有此应用
            existing_apps = [_get_app_id(section.app_data[iid]) for iid in section.treeview.get_children()]
            if app_id in existing_apps:
                return workflow_name
        
        return None
    
    def _check_config_conflicts(self):
        """检查配置中的跨工作流应用冲突并弹窗提醒"""
        workflow_sections = {
            "html": (self.html_section, t("settings.extensions.html_title")),
            "md": (self.md_section, t("settings.extensions.md_title")),
            "latex": (self.latex_section, t("settings.extensions.latex_title")),
            "file": (self.file_section, t("settings.extensions.file_title")),
        }
        
        # 收集所有应用及其所在工作流
        app_workflows = {}  # {app_id: {"name": app_name, "workflows": [...]}}
        
        for workflow_key, (section, workflow_name) in workflow_sections.items():
            apps = [section.app_data[iid] for iid in section.treeview.get_children()]
            for app_info in apps:
                app_id = _get_app_id(app_info)
                if not app_id:
                    continue
                if app_id not in app_workflows:
                    app_workflows[app_id] = {
                        "name": app_info.get("name", ""),
                        "workflows": [],
                    }
                app_workflows[app_id]["workflows"].append(workflow_name)
        
        # 找出存在冲突的应用
        conflicts = {
            app_id: data
            for app_id, data in app_workflows.items()
            if len(data["workflows"]) > 1
        }
        
        if conflicts:
            # 构建冲突提示消息
            conflict_lines = []
            for data in conflicts.values():
                workflow_list = "、".join(data["workflows"])
                conflict_lines.append(f"• {data['name']}: {workflow_list}")
            
            conflict_msg = "\n".join(conflict_lines)
            messagebox.showwarning(
                t("settings.title.warning"),
                t("settings.extensions.config_conflict_warning", conflicts=conflict_msg)
            )
    
    def get_config(self) -> dict:
        """获取当前配置"""
        return {
            "html": self.html_section.get_config(),
            "md": self.md_section.get_config(),
            "latex": self.latex_section.get_config(),
            "file": self.file_section.get_config(),
        }
