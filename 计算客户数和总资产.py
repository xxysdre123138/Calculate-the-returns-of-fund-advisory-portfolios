import pandas as pd
import tkinter as tk
from tkinter import filedialog

def process_excel_summary():
    # === 打开文件选择窗口 ===
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="请选择Excel文件", filetypes=[("Excel Files", "*.xlsx")])
    if not file_path:
        print("❌ 未选择文件")
        return

    # === 读取文件 ===
    df = pd.read_excel(file_path)

    # === 清洗列名，去除前后空格 ===
    df.columns = df.columns.str.strip()

    # === 在过滤前读取“客户小计”行的“客户数”并打印 ===
    # NEW: 尝试找到包含“客户数”的列名（如：客户数、客户数(户)…）
    cust_col = None
    for c in df.columns:
        if "客户数" in c:
            cust_col = c
            break

    if cust_col is not None:
        # 在任一列出现“客户小计”的行中取第一条
        mask_subtotal = df.apply(lambda s: s.astype(str).str.contains("客户小计", na=False), axis=0).any(axis=1)
        subtotal_rows = df[mask_subtotal]
        if not subtotal_rows.empty:
            val = str(subtotal_rows.iloc[0][cust_col]).replace(",", "")
            try:
                val_num = float(val)
            except ValueError:
                val_num = pd.to_numeric(val, errors="coerce")
            print(f"📌 客户小计（列：{cust_col}）：{val_num}")
        else:
            print("ℹ️ 未找到包含“客户小计”的行，跳过打印。")
    else:
        print("ℹ️ 未找到包含“客户数”的列，跳过打印。")

    # === 去除“客户数”中的逗号，转为数字 ===
    df["客户数"] = pd.to_numeric(df["客户数"].astype(str).str.replace(",", ""), errors='coerce')

    # === 去除“总资产(元)”中的逗号，转为数字 ===
    df["总资产(元)"] = pd.to_numeric(df["总资产(元)"].astype(str).str.replace(",", ""), errors='coerce')

    # === 删除组合名称为空的行（如“汇总”行） ===
    df = df[df["组合名称"].notna()]

    # === 分组汇总（包含总资产为0的） ===
    grouped = df.groupby("组合名称", as_index=False).agg({
        "客户数": "sum",
        "总资产(元)": "sum"
    })

    # === 新增一列“总资产(万元)” ===
    grouped["总资产(万元)"] = grouped["总资产(元)"] / 10000

    import os
    import datetime

    # === 输出结果到当前项目目录 ===
    # file_name = os.path.basename(file_path).replace(".xlsx", "_组合汇总结果.xlsx")
    yesterday_str = (datetime.date.today() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    file_name = f"客户数和总资产计算结果_{yesterday_str}.xlsx"
    output_path = os.path.join(os.getcwd(), file_name)
    grouped.to_excel(output_path, index=False)

    print(f"✅ 汇总完成，已保存为：{output_path}")

# === 运行主程序 ===
if __name__ == "__main__":
    process_excel_summary()

