import tkinter as tk

def build_ui(parent):
    # 清空 parent
    for widget in parent.winfo_children():
        widget.destroy()

    # 外层容器
    content = tk.Frame(parent)
    content.pack(fill="both", expand=True)

    content.columnconfigure(1, weight=1)
    content.rowconfigure(0, weight=1)

    sidebar = tk.Frame(content, width=200, bg="#f0f0f0")
    sidebar.grid(row=0, column=0, sticky="ns")
    sidebar.grid_propagate(False)

    content_area = tk.Frame(content, bg="white")
    content_area.grid(row=0, column=1, sticky="nsew")

    def load_subform(name):
        for widget in content_area.winfo_children():
            widget.destroy()
        label = tk.Label(content_area, text=f"当前子功能：{name}", font=("Arial", 16))
        label.pack(pady=20)

    submodules = ["录入基本信息", "录入健康信息", "录入繁殖记录"]
    for sub in submodules:
        btn = tk.Button(sidebar, text=sub, anchor="w",
                        command=lambda name=sub: load_subform(name))
        btn.pack(fill="x", padx=10, pady=5)
