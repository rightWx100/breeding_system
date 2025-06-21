import tkinter as tk
from tkinter import ttk

# 引入各模块的UI构建函数
from modules.data_entry.ui import build_ui as data_entry_ui
from modules.data_report.ui import build_ui as data_report_ui
from modules.breeding.ui import build_ui as breeding_ui

# 模块名称与UI函数映射
module_map = {
    "数据录入": data_entry_ui,
    "数据报表查询": data_report_ui,
    "种猪育种": breeding_ui
}

# 主界面创建函数
def ui_create():
    root = tk.Tk()
    root.title("育种数据管理系统")
    root.geometry("1000x600")
    root.minsize(800, 500)

    root.columnconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)

    # 顶部按钮栏（模块栏）
    module_frame = ttk.Frame(root, padding="5")
    module_frame.grid(row=0, column=0, sticky="ew")
    module_frame.columnconfigure((0, 1, 2), weight=1)

    modules = list(module_map.keys())

    # 主功能区（含导航 + 主体区域）
    main_frame = tk.Frame(root)
    main_frame.grid(row=1, column=0, sticky="nsew")
    main_frame.rowconfigure(2, weight=1)
    main_frame.columnconfigure(1, weight=1)

    # 第 1 行：功能导航提示栏
    nav_label = tk.Label(main_frame, text="", anchor= 'center', bg="#e8e8e8", font=("微软雅黑", 12), padx=10)
    nav_label.grid(row=1, column=0, columnspan=2, sticky="ew")

    # 第 2 行：功能选择 + 内容区
    sidebar = tk.Frame(main_frame, bg="#f0f0f0", width=20)
    sidebar.grid(row=2, column=0, sticky="ns")
    sidebar.grid_propagate(False)

    content_area = tk.Frame(main_frame, bg="white")
    content_area.grid(row=2, column=1, sticky="nsew")

    # 模块切换函数
    def load_module(name):
        # 清空右侧内容
        for widget in content_area.winfo_children():
            widget.destroy()
        # 加载新模块界面
        if name in module_map:
            module_map[name](content_area)

        # 清空左侧导航栏（可扩展为动态子功能）
        for widget in sidebar.winfo_children():
            widget.destroy()

        # 更新导航提示
        nav_label.config(text=f"{name} 功能导航")

    # 顶部按钮绑定
    for i, mod in enumerate(modules):
        label = tk.Label(
            module_frame, text=mod, fg="black", cursor="hand2",
            relief="ridge", padx=10, pady=5
        )
        label.grid(row=0, column=i, padx=20, sticky="ew")
        label.bind("<Double-Button-1>", lambda e, name=mod: load_module(name))

    # 默认加载第一个模块
    load_module(modules[0])

    return root

# 启动程序
if __name__ == "__main__":
    root = ui_create()
    root.mainloop()
