import tkinter as tk
from modules.common_ui import highlight_button, load_function_content

def build_ui(parent):
    """
    构建 '种猪育种' 模块的主UI结构，分为左-中-右三栏布局：
    左侧：功能导航菜单（如“种猪管理”等）
    中侧：子功能菜单（如“种猪登记”等）
    右侧：功能实现区（显示当前操作界面）
    """

    # 清空 parent 中原有的控件内容
    for widget in parent.winfo_children():
        widget.destroy()

    # 创建外层内容容器并填满父容器
    content = tk.Frame(parent)
    content.pack(fill="both", expand=True)

    # 设置三栏中的布局规则
    content.columnconfigure(2, weight=1)  # 右侧自动拉伸
    content.rowconfigure(0, weight=1)

    # === 左侧功能区（导航栏） ===
    left_panel = tk.Frame(content, width=140, bg="#f0f0f0")
    left_panel.grid(row=0, column=0, sticky="ns")
    left_panel.grid_propagate(False)

    # === 中间子功能区 ===
    center_panel = tk.Frame(content, width=180, bg="#e0e0e0")
    center_panel.grid(row=0, column=1, sticky="ns")
    center_panel.grid_propagate(False)

    # === 右侧功能展示区 ===
    right_panel = tk.Frame(content, bg="white")
    right_panel.grid(row=0, column=2, sticky="nsew")

    # 主功能及其子功能映射表
    function_map = {
        "种猪管理": ["种猪档案登记", "核心群等级划分", "种猪状态变更"],
        "种猪选配": ["选配方案制定", "选配效果评估", "种猪配种选配报告"],
        "育种评估": ["育种值估算", "遗传进展", "繁殖性能汇总"],
        "数据查询": ["种猪个体信息查询","",""]
    }

    active_func_btn = {"button": None}

    for func_name in function_map:
        created_btn = create_btn(
            name=func_name,
            left_panel=left_panel,
            right_panel=right_panel,
            center_panel=center_panel,
            active_func_btn=active_func_btn,
            function_map=function_map
        )
        if func_name == "种猪管理":
            default_btn = created_btn

    # 初始化默认模块内容
    load_subfunctions(function_map["种猪管理"], center_panel, right_panel)
    highlight_button(default_btn, active_func_btn)


def create_btn(name, left_panel, right_panel, center_panel, active_func_btn, function_map):
    """
    创建左侧主功能按钮，绑定点击事件以加载中间子功能和高亮当前按钮。
    """
    btn = tk.Button(
        left_panel,
        text=name,
        anchor="center",
        width=16,
        bg="#f0f0f0",
        relief="flat",
        font=("Arial", 10)
    )

    btn.configure(command=lambda f=name, b=btn: (
        load_subfunctions(function_map[f], center_panel, right_panel),
        highlight_button(b, active_func_btn)
    ))

    btn.pack(pady=5)
    return btn


def load_subfunctions(sub_list, center_panel, right_panel):
    """
    加载中间子功能按钮，支持点击后在右侧显示内容并高亮子按钮。
    """
    for w in center_panel.winfo_children():
        w.destroy()

    active_sub_btn = {"button": None}

    for name in sub_list:
        btn = tk.Button(
            center_panel,
            text=name,
            anchor="center",
            width=20,
            font=("Arial", 10),
            bg="#e0e0e0",
        )

        btn.configure(command=lambda n=name, b=btn: (
            load_function_content(n, right_panel),
            highlight_button(b, active_sub_btn)
        ))

        btn.pack(pady=5)
