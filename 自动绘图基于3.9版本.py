import os, re, math
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
from tkinter import Tk, filedialog
from pathlib import Path

# ========== 全局外观（中文/负号）==========
mpl.rcParams['font.family'] = ['Microsoft YaHei', 'SimHei', 'Arial Unicode MS']
mpl.rcParams['axes.unicode_minus'] = False

def choose_uniform_ticks(n_points: int, target_labels: int = 12):
    """
    在 0..n_points-1 上挑出 ~target_labels 个等距刻度。
    - 保证首尾都有
    - 小样本不会报错
    - 去重、排序
    """
    target_labels = max(3, int(target_labels))
    if n_points <= target_labels:
        return np.arange(n_points)

    step = int(np.ceil(n_points / target_labels))
    pos = np.arange(0, n_points, step)

    # 保证首尾
    if pos[0] != 0:
        pos = np.r_[0, pos]
    if pos[-1] != n_points - 1:
        pos = np.r_[pos, n_points - 1]

    return np.unique(pos)

# ========== 小工具 ==========
def pick_excel(title):
    Tk().withdraw()
    p = filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel 文件", "*.xlsx;*.xls")]
    )
    return p

def _norm_col(s: str) -> str:
    """列名规范化：去空格/全角空格、小写，并做常见别名归一"""
    s0 = str(s).strip().replace("\u3000", "").lower()
    s0 = s0.replace("−", "-")
    alias = {
        "组合名称": "组合", "组合名": "组合", "名称": "组合", "sheet": "组合",
        "leftmin": "left_min", "left_max": "left_max",
        "leftste": "left_step", "left_ste": "left_step",
        "rightmin": "right_min", "rightmax": "right_max",
        "rightstep": "right_step",
    }
    return alias.get(s0, s0)

def _num_clean(series: pd.Series) -> pd.Series:
    """数字清洗：全角负号→半角；去千分位逗号；转数值"""
    return pd.to_numeric(
        series.astype(str).str.replace("−", "-", regex=False).str.replace(",", "", regex=False),
        errors="coerce"
    )

def _safe_name(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", str(name))

def match_sheet_name(cfg_name: str, sheets: list) -> str or None:
    """把‘组合名称’匹配到数据sheet：完全相等；加/去‘组合’后缀；包含关系"""
    cfg_name = str(cfg_name).strip()
    for s in sheets:
        if s == cfg_name:
            return s
    for s in sheets:
        if s == cfg_name + "组合":
            return s
    for s in sheets:
        if s.replace("组合", "") == cfg_name:
            return s
    cand = [s for s in sheets if (cfg_name in s) or (s.replace("组合","") in cfg_name)]
    if cand:
        return sorted(cand, key=len, reverse=True)[0]
    return None

REQUIRED_COLS = {"组合", "left_min", "left_max", "left_step", "right_min", "right_max", "right_step"}

def read_axes_config_from_workbook(xls_path: str) -> pd.DataFrame or None:
    """扫描该工作簿所有 sheet，找出包含必需列的那张轴参数表"""
    try:
        xls = pd.ExcelFile(xls_path)
    except Exception as e:
        print(f"❌ 无法打开参数工作簿：{e}")
        return None
    for sh in xls.sheet_names:
        try:
            raw = pd.read_excel(xls_path, sheet_name=sh)
        except Exception:
            continue
        if raw is None or raw.empty:
            continue
        tmp = raw.copy()
        tmp.columns = [_norm_col(c) for c in tmp.columns]
        if REQUIRED_COLS.issubset(set(tmp.columns)):
            df = tmp[list(REQUIRED_COLS)].copy()
            for col in ["left_min","left_max","left_step","right_min","right_max","right_step"]:
                df[col] = _num_clean(df[col])
            df["组合"] = df["组合"].astype(str).str.strip()
            df = df.dropna(subset=["组合"]).reset_index(drop=True)
            print(f"✅ 在『{Path(xls_path).name}』的 sheet「{sh}」识别到轴参数表")
            return df
    return None

# ========== 1) 选择数据工作簿 ==========
data_wb = pick_excel("请选择【数据工作簿】（包含14个组合各自的sheet）")
if not data_wb:
    raise SystemExit("未选择数据工作簿，程序退出。")
save_dir = Path(data_wb).parent
xls_data = pd.ExcelFile(data_wb)
ALL_SHEETS = xls_data.sheet_names

# ========== 2) 选择轴参数工作簿（可选；取消则在数据工作簿中查找）==========
cfg_wb = pick_excel("请选择【轴参数工作簿】（可与数据同一文件；若取消将自动在数据工作簿中查找）")
if cfg_wb:
    cfg_df = read_axes_config_from_workbook(cfg_wb)
else:
    print("未选择单独的轴参数工作簿，将在数据工作簿中尝试查找。")
    cfg_df = read_axes_config_from_workbook(data_wb)

if cfg_df is None:
    raise SystemExit("⚠️ 未找到轴参数表（需包含：组合/left_min/left_max/left_step/right_min/right_max/right_step）。")

# ========== 绘图参数（与你单张图风格一致）=========
FIGSIZE = (12, 8)
XTICK_STEP = 61
BAR_WIDTH = 0.8
BOTTOM_SPACE = 0.18

# ========== （可选）选择导出目录 ==========
# root = Tk(); root.withdraw()
# folder_pick = filedialog.askdirectory(title="选择导出图片的文件夹（可取消将保存到Excel同目录）")
# if folder_pick: save_dir = Path(folder_pick)

# ========== 批量绘图 ==========
for _, row in cfg_df.iterrows():
    combo_name = str(row["组合"]).strip()
    sheet_name = match_sheet_name(combo_name, ALL_SHEETS)
    if not sheet_name:
        print(f"⚠️ 未找到与『{combo_name}』匹配的数据sheet，跳过。")
        continue

    # 从第41行开始读取数据
    try:
        df = pd.read_excel(data_wb, sheet_name=sheet_name, skiprows=40)
    except Exception as e:
        print(f"❌ 读取『{sheet_name}』失败：{e}")
        continue

    # 定位必要列
    if "日期" not in df.columns:
        print(f"⚠️ 『{sheet_name}』缺少【日期】列，跳过。")
        continue
    col_combo = next((c for c in df.columns if str(c).strip().startswith("组合累计收益")), None)
    col_bench = next((c for c in df.columns if str(c).strip().startswith("基准累计收益")), None)
    col_excess= next((c for c in df.columns if str(c).strip().startswith("超额收益")), None)
    col_assets= next((c for c in df.columns if "总资产" in str(c)), None)
    col_shares= next((c for c in df.columns if "总份额" in str(c)), None)
    needed = [col_combo, col_bench, col_excess, col_assets, col_shares]
    if any(v is None for v in needed):
        print(f"⚠️ 『{sheet_name}』缺列：{[('组合累计收益',col_combo),('基准累计收益',col_bench),('超额收益',col_excess),('总资产',col_assets),('总份额',col_shares)]}，跳过。")
        continue

    # 清洗/排序
    df = df.copy()
    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    for c in [col_combo, col_bench, col_excess, col_assets, col_shares]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df.dropna(subset=["日期"]).sort_values("日期").reset_index(drop=True)
    N = len(df)
    if N == 0:
        print(f"⚠️ 『{sheet_name}』无有效数据，跳过。")
        continue

    x = np.arange(N)

    # —— 开始绘图 —— #
    plt.close('all')
    title = combo_name

    fig, ax1 = plt.subplots(figsize=FIGSIZE, num=title)
    fig.subplots_adjust(bottom=BOTTOM_SPACE)

    # 左轴（收益）范围与刻度（来自轴参数；缺失则用数据兜底）
    lmin, lmax, lstep = row["left_min"], row["left_max"], row["left_step"]
    if pd.isna(lmin) or pd.isna(lmax) or pd.isna(lstep):
        vmin = np.nanmin([df[col_combo].min(), df[col_bench].min(), df[col_excess].min()])
        vmax = np.nanmax([df[col_combo].max(), df[col_bench].max(), df[col_excess].max()])
        pad = 0.01
        ax1.set_ylim(vmin - pad, vmax + pad)
        ax1.yaxis.set_major_formatter(lambda v, pos: f"{v*100:.0f}%")
    else:
        lmin, lmax, lstep = float(lmin), float(lmax), float(lstep)
        ax1.set_ylim(lmin, lmax)
        ax1.yaxis.set_major_formatter(lambda v, pos: f"{v*100:.0f}%")
        ax1.set_yticks(np.arange(lmin, lmax + 1e-12, lstep))

    # 右轴
    ax2 = ax1.twinx()
    ax1.set_zorder(3); ax2.set_zorder(2); ax1.patch.set_alpha(0)

    ax2.bar(x, df[col_assets], color="lightgray", label="总资产（万元）",
            width=BAR_WIDTH, align="center", zorder=1)
    ax2.plot(x, df[col_shares], color="gold", linewidth=1.2, alpha=0.8,
             label="总份额（万份）", zorder=2)

    rmin = 0.0 if pd.isna(row["right_min"]) else float(row["right_min"])
    rstep= 1000.0 if pd.isna(row["right_step"]) else float(row["right_step"])
    if pd.isna(row["right_max"]):
        rmax_data = float(np.nanmax([df[col_assets].max(), df[col_shares].max(), rstep]))
        rmax = math.ceil(rmax_data / rstep) * rstep
    else:
        rmax = float(row["right_max"])
    ax2.set_ylim(rmin, rmax)
    ax2.set_yticks(np.arange(rmin, rmax + 1e-9, rstep))

    # … ax2.set_ylim(...); ax2.set_yticks(...)

    # 右轴单位（放在坐标轴内部右上角，避免被 tight_layout 裁切）
    ax2.text(0.995, 1.01, "万元",
             transform=ax2.transAxes,
             ha="right", va="bottom",
             fontsize=10, color="#666")

    # 左轴三条线
    ax1.plot(x, df[col_combo],  label="组合累计收益", color="#1f77b4", linewidth=1.6, zorder=5)
    ax1.plot(x, df[col_bench],  label="基准累计收益", color="#ff7f0e", linewidth=1.6, zorder=6)
    ax1.plot(x, df[col_excess], label="超额收益",   color="#d62728", linewidth=1.6, zorder=5)

    # X 轴刻度（每 61 个数据点）
    # tick_pos = np.arange(0, N, XTICK_STEP)
    # if len(tick_pos) >= 2 and tick_pos[-1] >= N - 1:
    #     tick_pos = tick_pos[:-1]
    # ax1.set_xticks(tick_pos)
    # ax1.set_xticklabels(df["日期"].dt.strftime("%Y/%m/%d").iloc[tick_pos], rotation=0, ha='center')
    # ax1.set_xlim(-0.5, N - 0.5)

    # X 轴刻度：等量化抽样，自动适配长短
    tick_pos = choose_uniform_ticks(N, target_labels=10)  # 想更稀/更密改这个数
    ax1.set_xticks(tick_pos)
    ax1.set_xticklabels(
        df["日期"].dt.strftime("%Y/%m/%d").iloc[tick_pos],
        rotation=0, ha='center'
    )
    ax1.set_xlim(-0.5, N - 0.5)

    # 字号
    ax1.tick_params(axis='x', labelsize=8)
    ax1.tick_params(axis='y', labelsize=8)
    ax2.tick_params(axis='y', labelsize=8)

    # 图例（底部）
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    handles = lines1 + lines2
    labels  = labels1 + labels2
    ax1.legend(handles, labels,
               loc='upper center', bbox_to_anchor=(0.5, -0.08),
               ncol=5, frameon=False, prop={'size':10})

    # 标题/布局/窗口名
    ax1.set_title(title)
    fig.tight_layout()
    try:
        fig.canvas.manager.set_window_title(title)
    except Exception:
        pass

    # 保存
    out_png = save_dir / f"{_safe_name(title)}.png"
    fig.savefig(out_png, dpi=150)
    plt.close(fig)
    print(f"✅ 已保存：{out_png}")

print("🎉 全部完成。")
