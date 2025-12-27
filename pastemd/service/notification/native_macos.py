"""Native macOS notification using osascript (most compatible)."""

import subprocess
from typing import Optional

from ...utils.logging import log


class NativeMacOSNotifier:
    """原生 macOS 通知器，使用 osascript（兼容性最好）"""

    @staticmethod
    def is_available() -> bool:
        """检查是否可用（osascript 在所有 macOS 版本都可用）"""
        return True

    @staticmethod
    def notify(title: str, message: str, app_icon: Optional[str] = None,
               timeout: int = 5, group: Optional[str] = None) -> bool:
        """
        发送原生 macOS 通知

        Args:
            title: 通知标题
            message: 通知消息
            app_icon: 应用图标路径（macOS 会自动使用应用的图标）
            timeout: 超时时间（macOS 会自动管理）
            group: 分组标识符

        Returns:
            是否成功发送
        """
        try:
            # 转义引号以避免 AppleScript 语法错误
            safe_title = title.replace('"', '\\"').replace('\\', '\\\\')
            safe_message = message.replace('"', '\\"').replace('\\', '\\\\')
            
            # 使用 osascript 发送通知（兼容所有 macOS 版本）
            script = f'display notification "{safe_message}" with title "{safe_title}"'
            
            subprocess.run(
                ["osascript", "-e", script],
                check=True,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=2,
            )
            
            log(f"macOS notification sent: {title}")
            return True

        except subprocess.TimeoutExpired:
            log("macOS notification timeout")
            return False
        except subprocess.CalledProcessError as e:
            log(f"macOS notification error: {e.stderr.strip() if e.stderr else str(e)}")
            return False
        except Exception as e:
            log(f"macOS notification error: {e}")
            return False


def test_notification():
    """测试通知功能"""
    if NativeMacOSNotifier.is_available():
        print("✅ Native macOS notifier available")
        result = NativeMacOSNotifier.notify(
            "测试通知",
            "这是一条原生 macOS 通知",
            group="com.richqaq.pastemd"
        )
        print(f"Notification sent: {result}")
    else:
        print("❌ Native macOS notifier not available")


if __name__ == "__main__":
    test_notification()
