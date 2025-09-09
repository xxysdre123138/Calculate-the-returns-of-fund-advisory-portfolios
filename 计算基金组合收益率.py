import pandas as pd
from tkinter import Tk, filedialog
import datetime
import os
from tkinter import Tk, filedialog

def process_net_value_file():
    # 弹出文件选择窗口
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="请选择净值数据Excel文件",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not file_path:
        print("❌ 未选择文件，程序终止。")
        return

    # 读取Excel
    df = pd.read_excel(file_path)

    # 转换净值日期为 datetime.date 类型
    df["净值日期"] = pd.to_datetime(df["净值日期"]).dt.date

    # 获取最新两天的日期
    unique_dates = sorted(df["净值日期"].unique(), reverse=True)
    if len(unique_dates) < 2:
        print("❌ 数据中不足两个日期，无法执行。")
        return

    t_date = unique_dates[0]
    t_1_date = unique_dates[1]

    # 提取最新日期的全部数据
    df_t = df[df["净值日期"] == t_date]

    # 提取前一天的“活钱管理”策略数据
    df_t1_currency = df[
        (df["净值日期"] == t_1_date) &
        (df["策略名称"] == "活钱管理")
    ]

    # 合并结果
    df_result = pd.concat([df_t, df_t1_currency], ignore_index=True)

    # 构造输出文件名
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_name = f"{base_name}_净值提取结果_{t_date}.xlsx"
    output_path = os.path.join(os.path.dirname(file_path), output_name)

    # 保留前导0
    df_result["组合代码"] = df_result["组合代码"].astype(str).str.zfill(4)

    from datetime import datetime

    # 设置起始日期（例如2025年5月10日）
    start_date = pd.to_datetime("2025-05-10")

    # 确保净值日期为 datetime 类型
    df_result["净值日期"] = pd.to_datetime(df_result["净值日期"])

    # === 弹窗让用户选择“起始日期表格” ===
    root = Tk()
    root.withdraw()
    file_start_date = filedialog.askopenfilename(title="请选择包含起始日期的Excel文件",
                                                 filetypes=[("Excel Files", "*.xlsx *.xls")])

    # === 读取起始日期数据 ===
    df_start = pd.read_excel(file_start_date)

    # 确保列名统一（根据你截图）
    df_start = df_start.rename(columns={"组合名称": "组合名称", "起始日期": "起始日期"})

    # === 合并到 df_result 中 ===
    df_result = df_result.merge(df_start, on="组合名称", how="left")

    # 添加运行天数列
    df_result["运行天数"] = (df_result["净值日期"] - df_result["起始日期"]).dt.days

    df_result["净值日期"] = df_result["净值日期"].dt.date
    df_result["起始日期"] = df_result["起始日期"].dt.date

    # 增加累计收益列
    df_result["组合累计收益"] = df_result["组合净值"] - 1
    df_result["基准累计收益"] = df_result["基准净值"] - 1

    # 增加超额收益列
    df_result["超额收益"] = df_result["组合累计收益"] - df_result["基准累计收益"]

    # 年化收益列
    df_result["组合年化收益"] = df_result["组合累计收益"] / df_result["运行天数"] * 365
    df_result["基准年化收益"] = df_result["基准累计收益"] / df_result["运行天数"] * 365

    # # 保存为新文件
    # with pd.ExcelWriter("组合净值结果.xlsx", engine='xlsxwriter', date_format='yyyy/m/d') as writer:
    #     df_result.to_excel(writer, index=False)
    # print(f"✅ 提取完成，保存为：{output_path}")
    # print(f"📌 包含 {t_date} 全部数据 + {t_1_date} 的活钱管理策略，共 {len(df_result)} 条记录")

    # 获取当前脚本所在目录（project目录）
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # output_file = os.path.join(script_dir, "组合净值结果.xlsx")

    import datetime
    yesterday_str = (datetime.date.today() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
    output_file = os.path.join(script_dir, f"组合收益率计算结果_{yesterday_str}.xlsx")

    # 保存为 Excel 文件
    with pd.ExcelWriter(output_file, engine='xlsxwriter', date_format='yyyy/m/d') as writer:
        df_result.to_excel(writer, index=False)

    print(f"✅ 提取完成，文件已保存为：{output_file}")
    print(f"📌 包含 {t_date} 全部数据 + {t_1_date} 的活钱管理策略，共 {len(df_result)} 条记录")

# 运行主程序
if __name__ == "__main__":
    process_net_value_file()



