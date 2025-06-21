

import tkinter as tk
from tkinter import ttk

def build_ui(parent):
    """
    构建 '数据报表查询' 模块的界面。
    """
    frame = tk.Frame(parent)
    frame.pack(fill="both", expand=True)

    title = tk.Label(frame, text="数据报表查询模块", font=("Arial", 18))
    title.pack(pady=20)

    # 表格示例
    columns = ("编号", "姓名", "日期")
    tree = ttk.Treeview(frame, columns=columns, show="headings")
    for col in columns:
        tree.heading(col, text=col)
    tree.insert("", "end", values=("001", "示例数据", "2025-06-20"))
    tree.pack(expand=True, fill="both", padx=10, pady=10)
