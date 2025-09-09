import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os


def merge_combination_data():
    # === 选择三张表格 ===
    root = tk.Tk()
    root.withdraw()

    print("📄 请选择第一张表格（投资收益数据）")
    file1 = filedialog.askopenfilename(title="选择表格1", filetypes=[("Excel files", "*.xlsx")])
    print("📄 请选择第二张表格（资金与客户数变化）")
    file2 = filedialog.askopenfilename(title="选择表格2", filetypes=[("Excel files", "*.xlsx")])
    print("📄 请选择第三张表格（总资产与客户数）")
    file3 = filedialog.askopenfilename(title="选择表格3", filetypes=[("Excel files", "*.xlsx")])

    if not file1 or not file2 or not file3:
        print("❌ 有文件未选择，程序终止。")
        return

    # === 读取三张表格 ===
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)
    df3 = pd.read_excel(file3)

    # === 清洗列名，去除空格 ===
    df1.columns = df1.columns.str.strip()
    df2.columns = df2.columns.str.strip()
    df3.columns = df3.columns.str.strip()

    # === 合并表2：签约、解约、转入转出金额 ===
    df_merge2 = df2[["组合名称", "签约客户数(户)", "解约客户数(户)", "转入资金(元)", "转出资金(元)"]].copy()
    df_merge2["新增金额（万元）"] = df_merge2["转入资金(元)"] / 10000
    df_merge2["减少金额（万元）"] = df_merge2["转出资金(元)"] / 10000
    df_merge2.drop(["转入资金(元)", "转出资金(元)"], axis=1, inplace=True)

    # === 合并表3：客户数与总资产 ===
    df_merge3 = df3[["组合名称", "客户数", "总资产(万元)"]].copy()

    # === 合并所有数据到表格1 ===
    df_final = df1.copy()

    # 注意 merge：左连接，以表格1为主，避免因组合顺序不同导致遗漏
    df_final = pd.merge(df_final, df_merge2, how="left", on="组合名称")
    df_final = pd.merge(df_final, df_merge3, how="left", on="组合名称")

    # === 处理“货币增强”组合：仅保留第一次匹配结果 ===
    if (df_final["组合名称"] == "货币增强").sum() > 1:
        idx = df_final[df_final["组合名称"] == "货币增强"].index[0]
        dup_idx = df_final[df_final["组合名称"] == "货币增强"].index[1:]
        df_final.loc[dup_idx, ["签约客户数(户)", "解约客户数(户)", "新增金额（万元）", "减少金额（万元）", "客户数",
                               "总资产(万元)"]] = None

    df_final["总份额（万份）"] = df_final["总资产(万元)"] / df_final["组合净值"]

    df_final["净值日期"] = pd.to_datetime(df_final["净值日期"]).dt.strftime("%Y/%m/%d")
    df_final["起始日期"] = pd.to_datetime(df_final["起始日期"]).dt.strftime("%Y/%m/%d")

    # === 组合排序 ===
    # 固定顺序（注意：不含最后一行“货币增强”）
    custom_order = [
        "股债平衡", "债券稳健", "货币增强", "股票精选", "量化睿选", "同业存单",
        "固收增强", "债券臻选", "指增严选", "偏股智选", "消费严选", "短债优选",
        "红利优选", "先进制造"
    ]

    # 提取最后一行“货币增强”
    last_row = df_final.iloc[[-1]] if df_final.iloc[-1]["组合名称"] == "货币增强" else pd.DataFrame()

    # 排除最后一行并按指定顺序排序
    df_except_last = df_final.iloc[:-1] if not last_row.empty else df_final.copy()
    df_except_last = df_except_last.set_index("组合名称").loc[custom_order].reset_index()

    # 合并排序结果 + 原末尾“货币增强”
    df_final = pd.concat([df_except_last, last_row], ignore_index=True)

    # 统一把“组合名称”显示为“……组合”
    df_final["组合名称"] = (
            df_final["组合名称"].astype(str).str.strip()
            .str.replace(r"组合$", "", regex=True)  # 先去掉已有的“组合”后缀（防止重复）
            + "组合"
    )

    # 删除不需要的列
    df_final.drop(columns=["策略名称", "组合代码"], inplace=True)

    # 设置导出列顺序
    columns_order = [
        "组合名称",  # 自定义加入
        "净值日期",
        "组合净值",
        "基准净值",
        "组合累计收益",
        "基准累计收益",
        "超额收益",
        "组合年化收益",
        "基准年化收益",
        "签约客户数(户)",
        "解约客户数(户)",
        "客户数",
        "新增金额（万元）",
        "减少金额（万元）",
        "总资产(万元)",
        "总份额（万份）",
        "运行天数",
        "起始日期"
    ]

    # 重新排序
    df_final = df_final[columns_order]

    # 2) 统一改成“图1”里的标准列名
    rename_to_standard = {
        "净值日期": "日期",
        "组合累计收益": "组合累计收益*",
        "签约客户数(户)": "签约客户数",
        "解约客户数(户)": "解约客户数",
        "客户数": "累计客户数",
        "总资产(万元)": "总资产（万元）",  # 全角括号
        # “总份额（万份）”本身已是目标名，这里不必改
    }
    df_final.rename(columns=rename_to_standard, inplace=True)

    # 3) 可选：再按“图1”最终顺序确保一次（不写也行，顺序已保持）
    final_cols = [
        "组合名称", "日期", "组合净值", "基准净值", "组合累计收益*",
        "基准累计收益", "超额收益", "组合年化收益", "基准年化收益",
        "签约客户数", "解约客户数", "累计客户数",
        "新增金额（万元）", "减少金额（万元）",
        "总资产（万元）", "总份额（万份）", "运行天数", "起始日期",
    ]
    df_final = df_final[final_cols]

    #    # === 输出文件 ===
    # output_path = os.path.join(os.getcwd(), "合并结果_基金组合数据.xlsx")

    # === 输出文件（自动带昨天日期，保存在脚本所在目录） ===
    import datetime

    script_dir = os.path.dirname(os.path.abspath(__file__))  # 脚本所在目录
    yesterday_str = (datetime.date.today() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    output_path = os.path.join(script_dir, f"基金数据统计结果_{yesterday_str}.xlsx")

    df_final.to_excel(output_path, index=False)
    print(f"✅ 合并完成，已保存到：{output_path}")


# === 执行主程序 ===
if __name__ == "__main__":
    merge_combination_data()
