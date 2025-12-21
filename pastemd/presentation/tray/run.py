"""Tray icon runner."""

import pystray
import sys

from ...core.state import app_state
from ...utils.system_detect import is_macos
from .icon import create_status_icon
from .menu import TrayMenuManager


class TrayRunner:
    """托盘运行器"""
    
    def __init__(self, menu_manager: TrayMenuManager):
        self.menu_manager = menu_manager
        self.icon = None
    
    def setup(self, app_name: str = "PasteMD") -> None:
        """初始化托盘图标（macOS 必须在主线程调用）"""
        # 创建初始图标
        tray_icon = create_status_icon(ok=True)
        
        # 创建托盘实例
        icon = pystray.Icon(
            app_name,
            tray_icon,
            app_name,
            self.menu_manager.build_menu()
        )
        
        # 保存图标实例
        self.icon = icon
        app_state.icon = icon
        
        # 在 macOS 上使用 setup 方法，然后在后台运行
        if is_macos():
            icon.run_detached()
    
    def run(self, app_name: str = "PasteMD") -> None:
        """启动托盘图标"""
        if self.icon is not None:
            # 已经在 setup 中初始化，直接运行
            self.icon.run()
            return
            
        # 创建初始图标
        tray_icon = create_status_icon(ok=True)
        
        # 创建托盘实例
        icon = pystray.Icon(
            app_name,
            tray_icon,
            app_name,
            self.menu_manager.build_menu()
        )
        
        # 保存图标实例到全局状态
        app_state.icon = icon
        
        # 启动托盘（阻塞运行）
        icon.run()
