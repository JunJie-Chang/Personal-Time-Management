import pandas as pd
import os
import sys

# ================= 設定區 =================
# 原始 CSV 檔案路徑
RAW_FILE_PATH = '時間軌跡.csv'

# 輸出處理後的檔案路徑 (可供檢查用)
PROCESSED_FILE_PATH = 'processed_records.csv'
# =========================================

def preprocess_data(filepath):
    """
    讀取原始 CSV，重新命名欄位，並切分 Project Code 與 Contents。
    """
    if not os.path.exists(filepath):
        print(f"錯誤: 找不到檔案 '{filepath}'")
        sys.exit(1)
        
    try:
        # 讀取 CSV (處理 utf-8-sig 以去除 BOM)
        df = pd.read_csv(filepath, encoding='utf-8-sig')
        
        # 1. 欄位映射 (中文 -> 英文)
        # 根據 process.ipynb 的結構進行對應
        column_mapping = {
            '備註': 'note',
            '持續時間（分鐘）': 'duration_min',
            '任務名稱': 'task_name',
            '開始日期': 'start_date',
            '開始時間': 'start_time',
            '項目名稱': 'project_name'
        }
        df.rename(columns=column_mapping, inplace=True)
        
        # 確保必要欄位存在
        if 'note' not in df.columns or 'duration_min' not in df.columns:
            print("錯誤: CSV 格式不符，找不到 '備註' 或 '持續時間（分鐘）' 欄位。")
            sys.exit(1)

        # 2. 處理 NaN 值 (將無備註的填為空字串，避免錯誤)
        df['note'] = df['note'].fillna('')

        # 3. 切分 Logic
        # 邏輯: 嘗試以第一個 "_" 切分。
        # 若有 "_": 左邊是 Code, 右邊是 Content
        # 若無 "_": 整串是 Code, Content 為 None (參照你的 process.ipynb 邏輯)
        def split_project_info(note):
            if not note:
                return None, None
            
            note_str = str(note).strip()
            if '_' in note_str:
                parts = note_str.split('_', 1)
                return parts[0].strip(), parts[1].strip()
            else:
                # 若無底線，假設整串為 Code (如: "游庭皓" -> Code="游庭皓")
                # 如果你想改變邏輯 (例如 "財務管理 Chap.13" 要切分)，請修改此處
                return note_str, None

        # 應用切分邏輯
        df[['project_code', 'project_contents']] = df['note'].apply(
            lambda x: pd.Series(split_project_info(x))
        )
        
        # 過濾掉沒有 project_code 的列 (原本備註為空或是純雜訊)
        clean_df = df.dropna(subset=['project_code'])
        
        print(f"資料處理完成。原始筆數: {len(df)}, 有效專案筆數: {len(clean_df)}")
        return clean_df
        
    except Exception as e:
        print(f"讀取或處理檔案時發生錯誤: {e}")
        sys.exit(1)

def generate_report(df, code):
    """篩選並計算總時間"""
    # 統一轉為字串並去除空白後比對
    mask = df['project_code'].astype(str).str.strip().str.lower() == code.strip().lower()
    project_data = df[mask]
    
    if project_data.empty:
        print(f"\n[結果] 找不到 Project Code 為 '{code}' 的紀錄。")
        return

    # 計算統計數據
    total_minutes = project_data['duration_min'].sum()
    record_count = len(project_data)
    
    hours = int(total_minutes // 60)
    minutes = int(total_minutes % 60)
    
    print("\n" + "="*40)
    print(f"📊 專案報告: {project_data.iloc[0]['project_code']}") # 顯示原始大小寫
    print("="*40)
    print(f"總投入時間 : {hours} 小時 {minutes} 分鐘 ({total_minutes} min, {total_minutes/60:.2f} hr)")
    print(f"總紀錄筆數 : {record_count} 筆")
    print("-" * 40)
    
    # 顯示詳細內容分佈 (Group by Contents)
    if 'project_contents' in project_data.columns:
        print("【內容分佈 Top 5】")
        content_stats = project_data.groupby('project_contents')['duration_min'].sum().sort_values(ascending=False).head(5)
        for content, mins in content_stats.items():
            if content: # 排除 None
                h = int(mins // 60)
                m = int(mins % 60)
                print(f" - {content:<20} : {h}hr {m}min")
    
    print("\n【最近 3 筆紀錄】")
    cols_to_show = ['start_date', 'task_name', 'project_contents', 'duration_min']
    # 只顯示存在的欄位
    valid_cols = [c for c in cols_to_show if c in project_data.columns]
    print(project_data.tail(3)[valid_cols].to_string(index=False))
    print("="*40 + "\n")

def main():
    print("--- 專案時間統計與處理工具 ---")
    print(f"正在讀取並處理: {RAW_FILE_PATH} ...")
    
    # 1. 執行前處理
    df = preprocess_data(RAW_FILE_PATH)
    
    # (選項) 儲存處理後的檔案
    # df.to_csv(PROCESSED_FILE_PATH, index=False, encoding='utf-8-sig')
    
    # 2. 進入查詢迴圈
    print("\n系統就緒。請輸入 Project Code 查詢 (例如: AP, LBM)。")
    
    while True:
        user_input = input(">> 請輸入代碼 (q 離開): ").strip()
        
        if user_input.lower() == 'q':
            print("程式結束。")
            break
            
        if not user_input:
            continue
            
        generate_report(df, user_input)

if __name__ == "__main__":
    main()
