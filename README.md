# 週復盤模組

一個用於時間追蹤與週復盤分析的 Python 工具集，可從時間軌跡數據生成週報、統計圖表和 Excel 報告。

## 📋 專案簡介

本專案提供了一套完整的時間追蹤分析工具，能夠：

- 從 CSV 時間軌跡數據生成週度統計報告
- 按專案和任務進行時間聚合分析
- 生成可視化圖表（專案詳情、時間比例、累積時間）
- 提供互動式專案查詢功能
- 輸出 Excel 格式的總覽報告

## 📁 專案結構

```
週復盤模組/
├── README.md                 # 本文件
├── 時間軌跡.csv              # 原始時間追蹤數據（輸入）
├── total_review.xlsx         # 週度統計 Excel 報告（輸出）
├── weekly_split.py           # 週度數據分割與 Excel 生成工具
├── final_review.py           # 專案查詢與統計工具
├── process.py             # 數據處理與可視化筆記本
└── integration/              # 週度報告輸出目錄
    ├── 09-08/                # 2025-09-08 週期報告
    │   ├── 09-08_project_details.png
    │   ├── 09-08_time_aggregation_cumulative.png
    │   ├── 09-08_time_proportion.png
    │   └── weekly_report.md
    ├── 09-15/                # 2025-09-15 週期報告
    └── ...                   # 其他週期報告
```

## 🚀 快速開始

### 環境需求

- Python 3.7+
- 依賴套件：
  - `pandas`（用於數據處理）

### 安裝依賴

```bash
pip install pandas
```

### 數據格式

輸入文件 `時間軌跡.csv` 應包含以下欄位：

- `開始日期`：格式為 `YYYY/MM/DD`
- `持續時間（分鐘）`：數值（分鐘）
- `項目名稱`：專案名稱
- `任務名稱`：任務名稱
- `備註`：可選，用於 `final_review.py` 的專案代碼切分

## 📖 使用說明

### 1. 生成週度 Excel 報告

執行 `weekly_split.py` 從 `時間軌跡.csv` 生成 `total_review.xlsx`：

```bash
python weekly_split.py
```

**輸出說明：**

- `projects_long`：長格式專案統計（週期、專案名稱、總分鐘、總小時）
- `tasks_long`：長格式任務統計（週期、任務名稱、總分鐘、總小時）
- `projects_wide`：寬格式專案統計（週期為行，專案為列）
- `tasks_wide`：寬格式任務統計（週期為行，任務為列）

**週期定義：** 每週從週一（Monday）開始，到週日（Sunday）結束。

### 2. 專案查詢工具

執行 `final_review.py` 進行互動式專案查詢：

```bash
python final_review.py
```

**功能說明：**

- 讀取並處理 `時間軌跡.csv`
- 自動將 `備註` 欄位切分為 `project_code` 和 `project_contents`（以 `_` 分隔）
- 支援根據 Project Code 查詢專案統計
- 顯示總投入時間、紀錄筆數、內容分佈和最近紀錄

**使用範例：**

```
>> 請輸入代碼 (q 離開): AP
📊 專案報告: AP
總投入時間 : 10 小時 30 分鐘
總紀錄筆數 : 25 筆
【內容分佈 Top 5】
 - 內容A : 5hr 20min
 - 內容B : 3hr 10min
...
```

### 3. 數據處理與可視化

使用 `process.ipynb` Jupyter Notebook 進行：

- 數據清洗與轉換
- 生成週度報告（Markdown）
- 生成可視化圖表：
  - 專案詳情圖
  - 時間比例圖
  - 時間聚合累積圖

## 📊 輸出範例

### 週報格式（`integration/*/weekly_report.md`）

週報包含以下統計資訊：

- **摘要**：總工時、每日平均、主要專案/任務
- **工時總覽**：任務數、平均時長、最長/最短任務
- **日別模式**：高峰日、低谷日
- **專案分析**：各專案時間分佈與集中度指標（HHI、Entropy、Gini）
- **任務分析**：各任務時間分佈與集中度指標
- **資料完整性**：缺失數據統計
- **異常情況**：超過 6 小時任務、極短任務等

### Excel 報告格式

`total_review.xlsx` 包含四個工作表，提供不同視角的時間統計：

- **長格式**：適合詳細查看每個專案/任務的週度數據
- **寬格式**：適合橫向比較不同週期的專案/任務時間變化

## 🔧 技術細節

### weekly_split.py

- **純 Python 實現**：不依賴第三方 Excel 庫（如 openpyxl），直接生成 Excel XML 格式
- **週期計算**：使用 `week_bounds()` 函數計算每週的週一到週日範圍
- **數據聚合**：使用 `defaultdict` 進行高效的時間累加

### final_review.py

- **數據預處理**：自動處理 UTF-8 BOM、欄位映射、缺失值
- **專案代碼切分**：從 `備註` 欄位提取 `project_code` 和 `project_contents`
- **互動式查詢**：支援大小寫不敏感的專案代碼查詢

## 📝 注意事項

1. **日期格式**：確保 `時間軌跡.csv` 中的日期格式為 `YYYY/MM/DD`
2. **編碼**：CSV 文件應使用 UTF-8 編碼（支援 UTF-8 BOM）
3. **週期定義**：週期以週一為起始日，週日為結束日
4. **專案代碼切分**：`final_review.py` 使用 `_` 作為分隔符，格式為 `CODE_CONTENTS`

## 🛠️ 開發說明

### 擴展功能

- 修改 `weekly_split.py` 可調整週期計算邏輯或輸出格式
- 修改 `final_review.py` 可調整專案代碼切分規則
- 使用 `process.ipynb` 可自訂可視化圖表和報告格式

### 貢獻

歡迎提交 Issue 或 Pull Request 來改進此專案。

## 📄 授權

本專案為個人時間追蹤工具，供學習與個人使用。
# Personal-Time-Management
