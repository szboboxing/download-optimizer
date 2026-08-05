"""
下载流程优化工具 v3.5 — 入口文件
==================================
启动应用：python main.py

作者：下载流程优化工具 Team
许可证：MIT
"""

from app import MainApp
import tkinter as tk

if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()
