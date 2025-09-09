import pandas as pd
import tkinter as tk
from tkinter import filedialog

def summarize_contract_flow():
    # === 打开文件选择窗口 ===
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="请选择Excel文件", filetypes=[("Excel 文件", "*.xlsx")])
    if not file_path:
        print("❌ 未选择文件")
        return

    # === 读取Excel文件 ===
    df = pd.read_excel(file_path)

    # === 清洗列名，去除前后空格 ===
    df.columns = df.columns.str.strip()

    # === 提取“客户去重”行的【签约客户数(户)、解约客户数(户)】两个单元格的值 ===
    sign_col = "签约客户数(户)"
    cancel_col = "解约客户数(户)"

    if sign_col in df.columns and cancel_col in df.columns:
        # 精确匹配组合名称 == 客户去重
        row_mask = df["组合名称"].astype(str).str.strip().eq("客户去重")
        row = df[row_mask]

        # 如果没找到，则兜底匹配（任意列包含“客户去重”）
        if row.empty:
            any_mask = df.astype(str).apply(lambda s: s.str.contains("客户去重", na=False))
            row = df[any_mask.any(axis=1)]

        if not row.empty:
            ridx = row.index[0]

            def _to_num(x):
                return pd.to_numeric(str(x).replace(",", "").replace("，", ""), errors="coerce")

            val_sign = _to_num(df.at[ridx, sign_col])
            val_cancel = _to_num(df.at[ridx, cancel_col])

            # Excel 坐标转换
            def col_idx_to_letter(idx0: int) -> str:
                n = idx0 + 1
                s = ""
                while n > 0:
                    n, r = divmod(n - 1, 26)
                    s = chr(65 + r) + s
                return s

            excel_row = ridx + 2  # 默认第1行为表头
            sign_addr = f"{col_idx_to_letter(df.columns.get_loc(sign_col))}{excel_row}"
            cancel_addr = f"{col_idx_to_letter(df.columns.get_loc(cancel_col))}{excel_row}"

            print(f"📌 单元格（客户去重，{sign_col}） -> {sign_addr} = {val_sign}")
            print(f"📌 单元格（客户去重，{cancel_col}） -> {cancel_addr} = {val_cancel}")
        else:
            print("ℹ️ 未定位到“客户去重”行，跳过提取。")
    else:
        print("ℹ️ 缺少【签约客户数(户)】或【解约客户数(户)】列，无法提取。")

    # === 要提取和处理的列 ===
    required_columns = [
        "组合名称",
        "签约客户数(户)",
        "转入资金(元)",
        "解约客户数(户)",
        "转出资金(元)"
    ]

    # === 检查是否所有列都存在 ===
    for col in required_columns:
        if col not in df.columns:
            print(f"❌ 缺少列：{col}")
            return

    # === 去除“组合名称”为空的行 ===
    df = df[df["组合名称"].notna()]

    # === 去除千位分隔符，转为数值 ===
    for col in required_columns[1:]:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", ""), errors='coerce')

    # === 分类汇总 ===
    grouped = df.groupby("组合名称", as_index=False).agg({
        "签约客户数(户)": "sum",
        "转入资金(元)": "sum",
        "解约客户数(户)": "sum",
        "转出资金(元)": "sum"
    })

    # === 新增计算列 ===
    grouped["新增金额（万元）"] = grouped["转入资金(元)"] / 10000
    grouped["减少金额（万元）"] = grouped["转出资金(元)"] / 10000

    # === 导出到当前目录的汇总表格 ===
    import os
    import datetime
    # === 自动生成文件名（昨天的日期） ===
    yesterday_str = (datetime.date.today() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    output_file_name = f"签解约客户数和资金增减计算结果_{yesterday_str}.xlsx"

    # === 保存到项目目录 ===
    output_path = os.path.join(os.getcwd(), output_file_name)
    grouped.to_excel(output_path, index=False)

    print(f"✅ 汇总完成，结果已保存为：{output_path}")

# === 运行主程序 ===
if __name__ == "__main__":
    summarize_contract_flow()