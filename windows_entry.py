"""Windows 免安装包入口：启动本地服务并打开浏览器。

PyInstaller 打包说明：
    pyinstaller --noconsole? 不要：保留控制台窗口，关闭窗口即停止服务
    --add-data "app/web;app/web"   （macOS/Linux 用 ':' 分隔）
"""

import multiprocessing
import sys
from pathlib import Path


def main():
    multiprocessing.freeze_support()
    if not getattr(sys, "frozen", False):
        sys.path.insert(0, str(Path(__file__).resolve().parent / "app"))
    import server  # noqa: WPS433  （源码/冻结两种方式都可用）
    return server.main(sys.argv[1:])


if __name__ == "__main__":
    raise SystemExit(main())
