#!/usr/bin/env python3
"""
週復盤處理工具
從時間軌跡 CSV 生成週度統計報告、圖表和 Markdown 報告
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
import matplotlib.patches as mpatches
from datetime import datetime, timedelta
from scipy.stats import variation
import os
import shutil
import glob
import io
import sys
from pathlib import Path

# 設定中文字體
matplotlib.rcParams['font.family'] = ['Heiti TC', 'Arial Unicode MS', 'sans-serif']
matplotlib.rcParams['axes.unicode_minus'] = False

# 任務顏色對應表
TASK_COLOR_MAP = {
    "學科學習": "#1f568d",
    "課後學習": "#1d7fc4",
    "休閒閱讀": "#2d9ccc",
    "社團活動": "#a05eb7",
    "自主學習": "#b0145a",
    "創業項目": "#f18825",
    "競賽相關": "#f2b525",
    "課外籌備": "#39b5a4",
    "專題論文": "#eb4d39",
    "講座活動": "#14876b",
    "健康活動": "#19a44d",
    "有薪工作": "#9a5325",
    "媒體經營": "#9d5bb7",
    "個人專案": "#e73829",
    "應考考試": "#e04d37"
}

# 輸入文件
INPUT_FILE = Path("時間軌跡.csv")


def load_and_preprocess_data(filepath: Path) -> pd.DataFrame:
    """
    讀取 CSV 文件並進行預處理
    """
    if not filepath.exists():
        raise FileNotFoundError(f"找不到檔案：{filepath}")
    
    df = pd.read_csv(filepath, encoding='utf-8-sig')
    
    # 欄位映射
    columns_map = {
        '項目名稱': 'project_name',
        '任務名稱': 'task_name',
        '開始日期': 'start_date',
        '開始時間': 'start_time',
        '結束日期': 'end_date',
        '結束時間': 'end_time',
        '持續時間（分鐘）': 'duration_min',
        '持續時間': 'duration',
        '備註': 'note'
    }
    
    # 只保留需要的欄位
    available_columns = [col for col in columns_map.keys() if col in df.columns]
    df = df[available_columns]
    df = df.rename(columns=columns_map)
    
    # 切分 project_code 和 project_contents
    df[['project_code', 'project_contents']] = df['note'].str.split('_', n=1, expand=True)
    
    # 轉換日期和數值
    df['start_date'] = pd.to_datetime(df['start_date'], errors='coerce')
    df['end_date'] = pd.to_datetime(df['end_date'], errors='coerce')
    df['duration_min'] = pd.to_numeric(df['duration_min'], errors='coerce')
    df['duration_hours'] = (df['duration_min'] / 60).round(2)
    
    return df


def filter_by_week(df: pd.DataFrame, selected_monday: datetime) -> pd.DataFrame:
    """
    篩選指定週的數據（週一到週日）
    """
    if selected_monday.weekday() != 0:
        raise ValueError("指定的日期必須是週一（Monday）")
    if selected_monday > datetime.now():
        raise ValueError("指定的日期不能是未來日期")
    
    week_end = selected_monday + timedelta(days=7)
    filtered_df = df[(df['start_date'] >= selected_monday) & (df['start_date'] < week_end)].copy()
    
    return filtered_df


def save_plot(fig, filename: str, base_date: datetime):
    """
    儲存圖表，檔名使用日期前綴
    """
    date_prefix = base_date.strftime('%m-%d')
    safe_filename = f"{date_prefix}_{filename}"
    fig.savefig(safe_filename, dpi=800, bbox_inches='tight')
    print(f"Saved: {safe_filename}")


def plot_time_proportion(df: pd.DataFrame, base_date: datetime):
    """
    繪製時間比例圖（專案和任務的圓餅圖）
    """
    fig, axs = plt.subplots(1, 2, figsize=(12, 5))
    
    # 專案時間比例
    project_time = df.groupby('project_name')['duration_min'].sum()
    axs[0].pie(project_time, labels=project_time.index, autopct='%1.1f%%')
    axs[0].set_title('各項目時間比例')
    
    # 任務時間比例
    task_time = df.groupby('task_name')['duration_min'].sum()
    colors = [TASK_COLOR_MAP.get(t, "#cccccc") for t in task_time.index]
    axs[1].pie(task_time, labels=task_time.index, colors=colors, autopct='%1.1f%%')
    axs[1].set_title('各任務時間比例')
    
    save_plot(fig, "time_proportion.png", base_date)
    plt.close(fig)


def plot_time_aggregation(df: pd.DataFrame, base_date: datetime):
    """
    繪製累積工作時間折線圖
    """
    daily = df.groupby('start_date')['duration_hours'].sum().sort_index()
    cumulative = daily.cumsum()
    
    fig = plt.figure(figsize=(10, 4))
    threshold = 40
    plt.plot(cumulative.index, cumulative.values, marker='o')
    plt.xticks(rotation=45)
    plt.title('累積工作時間')
    plt.ylabel('累積小時')
    
    if (cumulative > threshold).any():
        plt.ylim(0, cumulative.max() + 3)
    else:
        plt.ylim(0, threshold)
    
    plt.tight_layout()
    save_plot(fig, "time_aggregation_cumulative.png", base_date)
    plt.close(fig)


def plot_project_details(df: pd.DataFrame, base_date: datetime):
    """
    繪製專案代碼時間消耗長條圖，並列出各專案的 project_contents
    """
    # 只保留同時有 project_code 與 project_contents（且非空白）的列
    cond_code = df['project_code'].notna() & df['project_code'].astype(str).str.strip().ne('')
    cond_contents = df['project_contents'].notna() & df['project_contents'].astype(str).str.strip().ne('')
    df_valid = df.loc[cond_code & cond_contents].copy()
    
    if df_valid.empty:
        print("No rows with both project_code and project_contents — nothing to plot.")
        return
    
    # 確保數值型態
    df_valid['duration_min'] = pd.to_numeric(df_valid['duration_min'], errors='coerce').fillna(0.0)
    
    # 以 project_code 彙總總時長
    agg = df_valid.groupby('project_code')['duration_min'].sum().sort_values(ascending=False)
    
    # 找出每個 project_code 下，哪個 task_name 貢獻最多時間（作為顏色對應依據）
    code_task = df_valid.groupby(['project_code', 'task_name'])['duration_min'].sum().reset_index()
    idx = code_task.groupby('project_code')['duration_min'].idxmax()
    dominant_map = code_task.loc[idx].set_index('project_code')['task_name'].to_dict()
    
    labels = agg.index.astype(str)
    
    # 以 dominant task_name 去對應 task_color_map（找不到則灰色）
    colors = [TASK_COLOR_MAP.get(dominant_map.get(code, ''), "#cccccc") for code in labels]
    
    fig = plt.figure(figsize=(8, 4))
    bars = plt.bar(labels, agg.values, color=colors)
    plt.title('專案代碼時間消耗')
    plt.ylabel('分鐘')
    plt.xlabel('專案代碼')
    plt.xticks(rotation=45, ha='right')
    
    # 長條上方註記數值
    y_max = agg.values.max() if len(agg.values) else 0
    offset = max(1.0, y_max * 0.01)
    for bar, val in zip(bars, agg.values):
        plt.text(bar.get_x() + bar.get_width() / 2, val + offset, f"{val:.0f}", 
                ha='center', va='bottom', fontsize=8)
    
    # 圖例：顯示本圖中出現的 task_name 與顏色
    seen = {}
    patches = []
    for code in labels:
        task = dominant_map.get(code, '')
        color = TASK_COLOR_MAP.get(task, "#cccccc")
        key = task if task else '其他'
        if key not in seen:
            seen[key] = color
            patches.append(mpatches.Patch(color=color, label=key))
    
    if patches:
        plt.legend(handles=patches, title="對應任務", bbox_to_anchor=(1.02, 1), 
                  loc='upper left', borderaxespad=0.)
    
    plt.tight_layout()
    save_plot(fig, "project_details.png", base_date)
    plt.close(fig)
    
    # 列出每個 project_code 對應的 project_contents（去重）
    contents = df_valid[['project_code', 'project_contents']].copy()
    grouped = contents.groupby('project_code')['project_contents'].unique()
    for code, items in grouped.items():
        if len(items):
            print(f"Project {code}:")
            for item in items:
                print(f"  - {item}")


def _safe_series_sum(s):
    """安全計算 Series 總和"""
    return float(np.nansum(s.values)) if len(s) else 0.0


def _hhi(weights: pd.Series) -> float:
    """計算 HHI (Herfindahl-Hirschman Index)"""
    w = weights.astype(float)
    tot = _safe_series_sum(w)
    if tot <= 0:
        return np.nan
    p = (w / tot).values
    return float(np.sum(p**2))


def _shannon_entropy(weights: pd.Series, base=np.e) -> float:
    """計算 Shannon Entropy"""
    w = weights.astype(float)
    tot = _safe_series_sum(w)
    if tot <= 0:
        return np.nan
    p = (w / tot).values
    p = p[p > 0]
    if len(p) == 0:
        return 0.0
    return float(-(p * np.log(p) / np.log(base)).sum())


def _gini(array_like) -> float:
    """計算 Gini 係數"""
    x = np.array(array_like, dtype=float)
    x = x[~np.isnan(x)]
    if x.size == 0:
        return np.nan
    if np.all(x == 0):
        return 0.0
    x_sorted = np.sort(x)
    n = x_sorted.size
    cumx = np.cumsum(x_sorted)
    g = 1 + (2.0 / (n * x_sorted.sum())) * np.sum((np.arange(1, n + 1)) * x_sorted) - (n + 1) / n
    return float(g)


def statistical_report(df: pd.DataFrame,
                       long_threshold_min: int = 360,   # 6 小時
                       short_threshold_min: int = 1,    # 1 分鐘
                       top_k: int = 5) -> str:
    """
    產生週期性時間使用統計報告（文字輸出）
    返回報告文字內容
    """
    # 基本防呆
    required = ['project_name', 'task_name', 'start_date', 'duration_min']
    for col in required:
        if col not in df.columns:
            raise KeyError(f"缺少必要欄位：{col}")
    
    if df.empty or df['duration_min'].fillna(0).sum() <= 0:
        return "---- 統計報告 ----\n本期無有效資料。\n"
    
    # 安全數值化
    df = df.copy()
    df['duration_min'] = pd.to_numeric(df['duration_min'], errors='coerce').fillna(0)
    
    # I. 摘要所需核心彙總
    total_min = _safe_series_sum(df['duration_min'])
    total_hours = total_min / 60.0
    
    # 日別彙總
    daily = df.groupby('start_date', dropna=False)['duration_min'].sum().sort_index()
    mean_daily = float(daily.mean()) if len(daily) else np.nan
    std_daily = float(daily.std(ddof=1)) if len(daily) > 1 else 0.0
    cv_daily = float(variation(daily)) if len(daily) > 1 and daily.mean() != 0 else np.nan
    
    # 任務級彙總
    num_tasks = int(len(df))
    avg_task = float(df['duration_min'].mean()) if num_tasks else np.nan
    longest = float(df['duration_min'].max()) if num_tasks else np.nan
    shortest = float(df['duration_min'].min()) if num_tasks else np.nan
    
    # 專案/任務分布
    proj_agg = df.groupby('project_name', dropna=False)['duration_min'].sum().sort_values(ascending=False)
    task_agg = df.groupby('task_name', dropna=False)['duration_min'].sum().sort_values(ascending=False)
    
    hhi_proj = _hhi(proj_agg)
    hhi_task = _hhi(task_agg)
    ent_proj = _shannon_entropy(proj_agg, base=2)
    ent_task = _shannon_entropy(task_agg, base=2)
    gini_proj = _gini(proj_agg.values)
    gini_task = _gini(task_agg.values)
    
    top_proj_name = str(proj_agg.index[0]) if len(proj_agg) else ""
    top_task_name = str(task_agg.index[0]) if len(task_agg) else ""
    top_proj_share = (proj_agg.iloc[0] / total_min) if len(proj_agg) and total_min > 0 else np.nan
    top_task_share = (task_agg.iloc[0] / total_min) if len(task_agg) and total_min > 0 else np.nan
    
    # 高峰/低谷日
    if len(daily):
        peak_day, peak_val = daily.idxmax(), float(daily.max())
        low_day, low_val = daily.idxmin(), float(daily.min())
    else:
        peak_day = low_day = None
        peak_val = low_val = np.nan
    
    # 個人標準
    lack_productivity_days = int((daily < 240).sum())
    
    # 資料完整性
    missing_contents = int(df['project_contents'].isna().sum()) if 'project_contents' in df.columns else None
    
    # 異常
    long_sessions_cnt = int((df['duration_min'] >= long_threshold_min).sum())
    short_sessions_cnt = int((df['duration_min'] <= short_threshold_min).sum())
    
    # 選用：weekday / week_label（若存在則加一段簡述）
    weekday_blurb = None
    if 'weekday' in df.columns:
        wd = df.groupby('weekday')['duration_min'].sum().sort_index()
        if len(wd):
            peak_wd = int(wd.idxmax())
            weekday_blurb = f"按週幾彙總顯示，第 {peak_wd + 1} 天工時最高：{wd.max():.1f} 分。"
    
    weeklabel_blurb = None
    if 'week_label' in df.columns:
        wk = df.groupby('week_label')['duration_min'].sum().sort_index()
        if len(wk) > 1:
            weeklabel_blurb = "跨週資料分布：" + ", ".join([f"{k}: {v:.1f} 分" for k, v in wk.items()])
    
    # 報告輸出
    lines = []
    lines.append("---- 統計報告（回測） ----")
    lines.append("")
    
    # I. 摘要
    lines.append("[摘要]")
    lines.append(f"- 總工時：{total_min:.1f} 分（{total_hours:.2f} 小時）")
    lines.append(f"- 每日平均：{mean_daily:.1f} 分；標準差：{std_daily:.1f} 分；CV：{(cv_daily if not np.isnan(cv_daily) else float('nan')):.3f}")
    if len(proj_agg):
        lines.append(f"- 主要專案：{top_proj_name}（佔比 {top_proj_share:.2%}）；HHI(專案)={hhi_proj:.3f}，Entropy(專案)={ent_proj:.3f}，Gini(專案)={gini_proj:.3f}")
    if len(task_agg):
        lines.append(f"- 主要任務：{top_task_name}（佔比 {top_task_share:.2%}）；HHI(任務)={hhi_task:.3f}，Entropy(任務)={ent_task:.3f}，Gini(任務)={gini_task:.3f}")
    
    # II. 工時總覽
    lines.append("")
    lines.append("[工時總覽]")
    lines.append(f"- 任務數：{num_tasks}")
    lines.append(f"- 平均單任務時長：{avg_task:.1f} 分")
    lines.append(f"- 最長任務：{longest:.1f} 分；最短任務：{shortest:.1f} 分")
    lines.append(f"- 低於最低生產力門檻天數：{lack_productivity_days} 天")
    
    # III. 日別模式
    lines.append("")
    lines.append("[日別模式]")
    if peak_day is not None:
        lines.append(f"- 高峰日：{peak_day}（{peak_val:.1f} 分）")
        lines.append(f"- 低谷日：{low_day}（{low_val:.1f} 分）")
    else:
        lines.append("- 本期無日別資料")
    
    # IV. 專案分析（Top-K）
    lines.append("")
    lines.append("[專案分析]")
    if len(proj_agg):
        top_proj_display = proj_agg.head(top_k)
        for i, (name, mins) in enumerate(top_proj_display.items(), start=1):
            lines.append(f"  {i}. {name}: {mins:.1f} 分（{mins/total_min:.2%}）")
        lines.append(f"- HHI(專案)：{hhi_proj:.3f}；Entropy(專案)：{ent_proj:.3f}；Gini(專案)：{gini_proj:.3f}")
    else:
        lines.append("- 無專案彙總")
    
    # V. 任務分析（Top-K）
    lines.append("")
    lines.append("[任務分析]")
    if len(task_agg):
        top_task_display = task_agg.head(top_k)
        for i, (name, mins) in enumerate(top_task_display.items(), start=1):
            lines.append(f"  {i}. {name}: {mins:.1f} 分（{mins/total_min:.2%}）")
        lines.append(f"- HHI(任務)：{hhi_task:.3f}；Entropy(任務)：{ent_task:.3f}；Gini(任務)：{gini_task:.3f}")
    else:
        lines.append("- 無任務彙總")
    
    # VIII. 補充（若附加欄位存在）
    if weekday_blurb:
        lines.append("")
        lines.append("[補充：週幾分布]")
        lines.append(f"- {weekday_blurb}")
    if weeklabel_blurb:
        lines.append("")
        lines.append("[補充：跨週分布]")
        lines.append(f"- {weeklabel_blurb}")
    
    # IX. 規則式觀察
    lines.append("")
    lines.append("[觀察]")
    if not np.isnan(top_proj_share) and top_proj_share > 0.5:
        lines.append("- 時間高度集中於單一專案。")
    if len(daily) > 1 and std_daily > max(1e-9, mean_daily) * 0.5:
        lines.append("- 日別工時波動偏大。")
    if not np.isnan(avg_task) and avg_task < 30:
        lines.append("- 任務偏碎片化。")
    elif not np.isnan(avg_task) and avg_task > 120:
        lines.append("- 任務偏向長段深度工作。")
    if not np.isnan(hhi_proj) and hhi_proj > 0.30:
        lines.append("- 專案集中度偏高。")
    elif not np.isnan(hhi_proj):
        lines.append("- 專案分散度較高。")
    if lack_productivity_days > 2:
        lines.append("- 生產力不足天數偏多。")
    
    return "\n".join(lines)


def save_statistical_report_md(df: pd.DataFrame, filename: str = "weekly_report.md"):
    """
    生成統計報告並存成 Markdown 文件
    """
    report_text = statistical_report(df)
    
    with open(filename, "w", encoding="utf-8") as f:
        f.write("# 統計報告 (回測)\n\n")
        f.write("```\n")
        f.write(report_text)
        f.write("\n```\n")
    
    print(f"Saved report to {filename}")


def integrate_all_move(base_date: datetime,
                       output_root: str = 'integration',
                       include_process: bool = False,
                       move_csv: bool = False):
    """
    移動本次產出到 ./integration/MM-DD 下
    """
    if isinstance(base_date, datetime):
        base_date = base_date.date()
    
    date_prefix = base_date.strftime('%m-%d')
    display_name = base_date.strftime('%m/%d')
    target_dir = os.path.join(output_root, date_prefix)
    os.makedirs(target_dir, exist_ok=True)
    
    moved = []
    missing = []
    
    # 單一檔案清單
    files_to_move = [
        'weekly_report.md',
        'statistical_report.md'
    ]
    if include_process:
        files_to_move.insert(0, 'process.py')
    if move_csv:
        files_to_move.insert(0, '時間軌跡.csv')
    
    for fname in files_to_move:
        if os.path.exists(fname):
            dest = os.path.join(target_dir, os.path.basename(fname))
            shutil.move(fname, dest)
            moved.append(fname)
        else:
            missing.append(fname)
    
    # 移動 PNG 圖檔
    for p in glob.glob(f"{date_prefix}_*.png"):
        dest = os.path.join(target_dir, os.path.basename(p))
        shutil.move(p, dest)
        moved.append(p)
    
    # 移動其他目錄（如果存在）
    for d in ['src', 'data', 'reports']:
        if os.path.isdir(d):
            dest_dir = os.path.join(target_dir, d)
            if os.path.exists(dest_dir):
                for item in os.listdir(d):
                    src_item = os.path.join(d, item)
                    dest_item = os.path.join(dest_dir, item)
                    shutil.move(src_item, dest_item)
                try:
                    os.rmdir(d)
                except OSError:
                    pass
                moved.append(f"{d}/ (merged)")
            else:
                shutil.move(d, dest_dir)
                moved.append(d + '/')
    
    print(f"Moved into: {target_dir}  (display name: {display_name})")
    if moved:
        print("Moved:")
        for m in moved:
            print("  -", m)
    if missing:
        print("Missing (skipped):")
        for s in missing:
            print("  -", s)
    
    return target_dir


def main():
    """
    主程式流程
    """
    print("=" * 50)
    print("週復盤處理工具")
    print("=" * 50)
    
    # 1. 讀取數據
    print(f"\n正在讀取數據：{INPUT_FILE}")
    df = load_and_preprocess_data(INPUT_FILE)
    print(f"讀取完成，共 {len(df)} 筆記錄")
    
    # 2. 選擇週期
    print("\n請輸入週一的日期（格式：YYYY-MM-DD）")
    date_input = input("Assigned date (YYYY-MM-DD): ").strip()
    selected_monday = datetime.strptime(date_input, '%Y-%m-%d')
    
    # 3. 篩選數據
    df_week = filter_by_week(df, selected_monday)
    print(f"\n篩選後數據：{len(df_week)} 筆記錄")
    print(f"週期：{selected_monday.strftime('%Y-%m-%d')} ~ {(selected_monday + timedelta(days=6)).strftime('%Y-%m-%d')}")
    
    if df_week.empty:
        print("警告：該週期沒有數據！")
        return
    
    # 4. 生成圖表
    print("\n正在生成圖表...")
    plot_time_proportion(df_week, selected_monday)
    plot_time_aggregation(df_week, selected_monday)
    plot_project_details(df_week, selected_monday)
    
    # 5. 生成統計報告
    print("\n正在生成統計報告...")
    report_text = statistical_report(df_week)
    print(report_text)
    
    # 6. 儲存報告
    save_statistical_report_md(df_week, "weekly_report.md")
    
    # 7. 整合文件
    print("\n正在整合文件...")
    print(f"輸出區間：{selected_monday} ~ {selected_monday + timedelta(days=6)}")
    print(f"執行日期：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    integrate_all_move(selected_monday)
    
    print("\n處理完成！")


if __name__ == "__main__":
    main()

