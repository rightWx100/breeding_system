import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd


def show_pig_registration(parent):
    label = tk.Label(parent, text="🐷 种猪档案登记界面", font=("微软雅黑", 14))
    label.pack(pady=20)

def show_core_group_grading(parent):
    label = tk.Label(parent, text="🌟 核心群等级划分界面", font=("微软雅黑", 14))
    label.pack(pady=20)

def show_status_change(parent):
    label = tk.Label(parent, text="🔁 状态变更界面", font=("微软雅黑", 14))
    label.pack(pady=20)


#种猪选配方案制定功能
def show_selection_mating(parent):
    # 清空界面
    for w in parent.winfo_children():
        w.destroy()

    # 创建顶部标签页控件
    notebook = ttk.Notebook(parent)
    notebook.pack(fill="both", expand=True)

    # 创建三个页面容器
    page_earinfo = tk.Frame(notebook)
    page_semeninfo = tk.Frame(notebook)
    page_matrix = tk.Frame(notebook)

    notebook.add(page_earinfo, text="待配耳号信息")
    notebook.add(page_semeninfo, text="公猪精液信息")
    notebook.add(page_matrix, text="选配二维表")

    # ===== 页面1：待配耳号信息 ===== #
    def setup_earinfo_page(frame):
        title = tk.Label(frame, text="请输入待配耳号信息", font=("微软雅黑", 14))
        title.pack(pady=10)

        form_frame = tk.Frame(frame)
        form_frame.pack(pady=10)

        # 母猪输入
        sow_col = tk.Frame(form_frame)
        sow_col.pack(side="left", padx=30)
        tk.Label(sow_col, text="待配母猪", font=("微软雅黑", 12)).pack(anchor="w")
        sow_text = tk.Text(sow_col, width=30, height=6, font=("微软雅黑", 11))
        sow_text.pack(pady=5)

        # 公猪输入
        boar_col = tk.Frame(form_frame)
        boar_col.pack(side="left", padx=30)
        tk.Label(boar_col, text="配种公猪", font=("微软雅黑", 12)).pack(anchor="w")
        boar_text = tk.Text(boar_col, width=30, height=6, font=("微软雅黑", 11))
        boar_text.pack(pady=5)

        # 验证耳号格式
        def validate_ear_tag(tag):
            return tag.isalnum() and len(tag) <= 15

        def validate_input():
            sow_list = [t.strip() for t in sow_text.get("1.0", "end").splitlines() if t.strip()]
            boar_list = [t.strip() for t in boar_text.get("1.0", "end").splitlines() if t.strip()]
            errors = []
            for s in sow_list:
                if not validate_ear_tag(s):
                    errors.append(f"母猪耳号格式错误：{s}")
            for b in boar_list:
                if not validate_ear_tag(b):
                    errors.append(f"公猪耳号格式错误：{b}")
            if errors:
                messagebox.showwarning("耳号格式错误", "\n".join(errors))
            else:
                messagebox.showinfo("验证通过", f"母猪 {len(sow_list)} 头，公猪 {len(boar_list)} 头")
                # 可以保存 sow_list/boar_list 为全局或外部变量以供后续页使用

        tk.Button(frame, text="验证耳号", font=("微软雅黑", 11), command=validate_input).pack(pady=10)

    # ===== 页面2：公猪精液信息 ===== #
    def setup_semeninfo_page(frame):
        tk.Label(frame, text="公猪精液信息录入", font=("微软雅黑", 14)).pack(pady=10)

        # ========== 顶部 Tab 结构 ==========
        tab_control = ttk.Notebook(frame)
        tab_control.pack(fill="both", expand=True, padx=10, pady=10)

        # --- 1. 文件导入页面 ---
        file_tab = tk.Frame(tab_control)
        tab_control.add(file_tab, text="导入Excel文件")

        file_status = tk.StringVar(value="未选择文件")

        def import_file():
            file_path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx *.xls")])
            if not file_path:
                return
            try:
                df = pd.read_excel(file_path)  # 读取 Excel 文件
                if "个体号" not in df.columns or "可用份数" not in df.columns:
                    messagebox.showerror("格式错误", "Excel 中必须包含列：‘个体号’ 和 ‘可用份数’")
                    return

                data_preview.delete("1.0", "end")
                for _, row in df.iterrows():
                    ear = str(row["个体号"]).strip()
                    dose = str(row["可用份数"]).strip()
                    line = f"{ear} - 可用份数：{dose}"
                    data_preview.insert("end", line + "\n")

                file_status.set(f"已导入文件：{file_path}")
            except Exception as e:
                messagebox.showerror("导入失败", f"读取文件出错：\n{e}")

        tk.Button(file_tab, text="选择Excel文件", command=import_file, font=("微软雅黑", 11)).pack(pady=5)
        tk.Label(file_tab, textvariable=file_status, font=("微软雅黑", 10), fg="gray").pack()

        data_preview = tk.Text(file_tab, width=60, height=10, font=("Consolas", 11))
        data_preview.pack(pady=10)

        # --- 2. 手动录入页面 ---
        manual_tab = tk.Frame(tab_control)
        tab_control.add(manual_tab, text="手动录入信息")

        form = tk.Frame(manual_tab)
        form.pack(pady=10)

        tk.Label(form, text="个体号：", font=("微软雅黑", 11)).grid(row=0, column=0, sticky="e", padx=5, pady=3)
        entry_id = tk.Entry(form, font=("微软雅黑", 11))
        entry_id.grid(row=0, column=1, padx=5)

        tk.Label(form, text="精液可用份数：", font=("微软雅黑", 11)).grid(row=1, column=0, sticky="e", padx=5, pady=3)
        entry_dose = tk.Entry(form, font=("微软雅黑", 11))
        entry_dose.grid(row=1, column=1, padx=5)

        manual_result = tk.Text(manual_tab, width=50, height=8, font=("Consolas", 11))
        manual_result.pack(pady=10)

        semen_list = []

        def add_manual_record():
            ear = entry_id.get().strip()
            dose = entry_dose.get().strip()
            if not ear or not dose.isdigit():
                messagebox.showwarning("输入错误", "请输入有效的个体号和数字份数")
                return
            semen_list.append((ear, int(dose)))
            manual_result.insert("end", f"{ear} - 可用份数：{dose}\n")
            entry_id.delete(0, "end")
            entry_dose.delete(0, "end")

        tk.Button(manual_tab, text="添加记录", command=add_manual_record, font=("微软雅黑", 11)).pack()

    # ===== 页面3：选配二维表 ===== #
    def setup_matrix_page(frame):
        tk.Label(frame, text="种猪选配结果表", font=("微软雅黑", 14)).pack(pady=10)
        matrix = tk.Text(frame, width=80, height=15, font=("Consolas", 11))
        matrix.pack(pady=10)

        def mock_fill():
            # 示例表格生成逻辑
            matrix.delete("1.0", "end")
            matrix.insert("end", "母猪耳号\t公猪耳号\n")
            matrix.insert("end", "-" * 30 + "\n")
            for i in range(5):
                matrix.insert("end", f"SOW{i+1:03d}\tBOAR{i+1:03d}\n")

        def export():
            try:
                with open("选配结果表.txt", "w", encoding="utf-8") as f:
                    f.write(matrix.get("1.0", "end"))
                messagebox.showinfo("导出成功", "结果已保存为：选配结果表.txt")
            except Exception as e:
                messagebox.showerror("导出失败", str(e))

        btn_frame = tk.Frame(frame)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="生成示例表", font=("微软雅黑", 11), command=mock_fill).pack(side="left", padx=5)
        tk.Button(btn_frame, text="导出结果", font=("微软雅黑", 11), command=export).pack(side="left", padx=5)

    # 初始化各页内容
    setup_earinfo_page(page_earinfo)
    setup_semeninfo_page(page_semeninfo)
    setup_matrix_page(page_matrix)

# 所有功能统一注册
function_handlers = {
    "种猪档案登记": show_pig_registration,
    "核心群等级划分": show_core_group_grading,
    "种猪状态变更": show_status_change,
    "选配方案制定":show_selection_mating,
    # 可继续添加...
}