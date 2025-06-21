import tkinter as tk


def highlight_button(button, active_func_btn):
    """
    高亮当前按钮，并取消上一个按钮的高亮状态。

    参数：
    - button: 当前点击的按钮控件
    - active_func_btn: 一个字典，保存了当前高亮的按钮对象（用于切换样式）
    """
    # 恢复上一个高亮按钮的默认样式
    if active_func_btn["button"]:
        active_func_btn["button"].configure(bg="#f0f0f0", fg="black")

    # 设置当前按钮为高亮状态
    button.configure(bg="#c0d8ff", fg="blue")
    active_func_btn["button"] = button


def load_function_content(text, right_panel):
    """
    在右侧功能区中加载指定的功能内容。

    参数：
    - text: 当前子功能的名称
    - right_panel: 右侧显示区容器
    """
    # 清空原有内容
    for w in right_panel.winfo_children():
        w.destroy()

    # 显示新的内容
    label = tk.Label(right_panel, text=f"当前功能内容：{text}", font=("Arial", 16))
    label.pack(pady=20)


def load_subfunctions(sub_list, center_panel, right_panel):
    """
    根据主功能的子功能列表，加载中间栏的按钮并绑定功能区域内容。

    参数：
    - sub_list: 子功能名称列表
    - center_panel: 中间子功能栏容器
    - right_panel: 右侧功能显示区容器
    """
    # 清空中间子功能区
    for w in center_panel.winfo_children():
        w.destroy()

    # 创建子功能按钮
    for name in sub_list:
        btn = tk.Button(
            center_panel,
            text=name,
            anchor="center",  # 文本居中
            width=15,  # 固定宽度（单位：字符数）
            font=("Arial", 10),
            bg="#e0e0e0",
            command=lambda n=name: load_function_content(n, right_panel)
        )
        btn.pack(fill="x", padx=10, pady=5)
