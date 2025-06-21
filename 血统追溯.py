import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# 全局变量用于存储文件路径
pedigree_path = None
query_path = None

def select_pedigree_file():
    global pedigree_path
    pedigree_path = filedialog.askopenfilename(title="选择三代系谱文件", filetypes=[("Excel 文件", "*.xlsx")])
    pedigree_label.config(text=f"三代系谱文件：{pedigree_path if pedigree_path else '未选择'}")

def select_query_file():
    global query_path
    query_path = filedialog.askopenfilename(title="选择查询个体号文件", filetypes=[("Excel 文件", "*.xlsx")])
    query_label.config(text=f"个体号文件：{query_path if query_path else '未选择'}")

def find_paternal_ancestors_batch(individual_ids, sire_dict):
    result = {}
    for ind_id in individual_ids:
        ancestors = []
        current = ind_id
        while True:
            sire = sire_dict.get(current)
            if pd.isna(sire) or sire == '' or sire not in sire_dict:
                break
            ancestors.append(sire)
            current = sire
        result[ind_id] = ancestors
    return result

def run_analysis():
    if not pedigree_path or not query_path:
        messagebox.showwarning("警告", "请先选择两个输入文件！")
        return

    try:
        # 读取数据
        df = pd.read_excel(pedigree_path)
        df1 = pd.read_excel(query_path)

        if 'id' not in df.columns or 'sire' not in df.columns:
            messagebox.showerror("错误", "三代系谱文件中必须包含列：id, sire")
            return
        if '个体号' not in df1.columns:
            messagebox.showerror("错误", "个体号文件中必须包含列：个体号")
            return

        sire_dict = dict(zip(df['id'], df['sire']))
        individual_list = df1['个体号']
        paternal_ancestors = find_paternal_ancestors_batch(individual_list, sire_dict)

        # 转为 DataFrame
        df_out = pd.DataFrame.from_dict(paternal_ancestors, orient='index')
        df_out.columns = [f'父系第{i+1}代' for i in range(df_out.shape[1])]
        df_out.index.name = '个体号'

        df_out['最终代次数'] = df_out.notna().sum(axis=1)
        df_out['最高代次祖先'] = df_out.apply(lambda row: row.dropna().iloc[-2] if row.dropna().shape[0] > 0 else None, axis=1)

        # 保存结果
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", title="保存结果为", filetypes=[("Excel 文件", "*.xlsx")])
        if save_path:
            df_out.to_excel(save_path)
            messagebox.showinfo("成功", f"文件已保存到：\n{save_path}")
        else:
            messagebox.showinfo("取消", "未选择保存位置，操作取消。")

    except Exception as e:
        messagebox.showerror("错误", f"处理过程中出现异常：\n{str(e)}")


# GUI 设置
root = tk.Tk()
root.title("父系祖先追溯工具")
root.geometry("500x300")

tk.Label(root, text="请选择输入文件：", font=("Arial", 12)).pack(pady=10)

# 选择三代系谱文件
tk.Button(root, text="选择三代系谱文件", command=select_pedigree_file, width=30).pack()
pedigree_label = tk.Label(root, text="三代系谱文件：未选择", wraplength=450, anchor='w', justify='left')
pedigree_label.pack(pady=5)

# 选择个体号文件
tk.Button(root, text="选择查询个体号文件", command=select_query_file, width=30).pack()
query_label = tk.Label(root, text="个体号文件：未选择", wraplength=450, anchor='w', justify='left')
query_label.pack(pady=5)

# 开始计算按钮
tk.Button(root, text="确认并开始计算", command=run_analysis, bg='lightgreen', width=30, height=2).pack(pady=20)

root.mainloop()
