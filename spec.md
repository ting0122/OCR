專案規格書：掃描檔 PDF 定點數據擷取與 Excel 自動化 (OCR-to-Excel Automation)
1. 專案目標 (Objective)
建立一個 Python 自動化程式，用於處理掃描成圖片格式的 PDF 文件（多頁）。程式需根據使用者預先設定的「像素座標 (Pixel Coordinates)」，將圖片中特定區域裁切並進行 OCR（光學字元辨識），提取數值後整理為陣列，最終寫入 Excel 指定欄位。

2. 系統流程 (Workflow)
讀取配置 (Config Load)：程式讀取使用者設定檔 (JSON/YAML)，獲取需要擷取的欄位名稱及對應座標。

PDF 轉換 (Convert)：將多頁 PDF 轉換為高解析度圖片序列。

區域裁切 (Crop)：針對每一頁圖片，根據設定檔中的座標裁切出特定區塊 (ROI, Region of Interest)。

影像前處理 (Pre-process)：對裁切後的圖片進行灰階、二值化處理，以提升 OCR 準確率。

文字辨識 (OCR)：使用 Tesseract OCR 讀取圖片中的數字/文字。

數據結構化 (Data Formatting)：將每一頁的數據整理為一組 Array/List。

輸出 (Export)：將整理好的數據寫入 Excel 檔。

3. 技術堆疊 (Tech Stack)
語言: Python 3

核心套件:

pdf2image: 將 PDF 轉為圖片 (需安裝 Poppler)。

pytesseract: 封裝 Google Tesseract-OCR 引擎。

Pillow (PIL): 圖片處理與裁切。

pandas & openpyxl: Excel 讀寫與數據處理。

json: 管理座標設定檔。

4. 功能需求規格 (Functional Requirements)
4.1 使用者配置模組 (User Configuration)
為了讓使用者能輕易增加欄位，不需修改程式碼，採用 JSON 格式 作為設定檔。

檔案名稱: config.json

座標定義: 使用 (x, y, width, height) 或 (left, top, right, bottom) 格式。

範例結構:

JSON
{
  "input_pdf": "data/scan_log_2024.pdf",
  "output_excel": "output/result.xlsx",
  "fields": [
    {
      "name": "溫度_T1",
      "box": [100, 200, 300, 250], 
      "page_index": "all" 
    },
    {
      "name": "壓力_P1",
      "box": [100, 350, 300, 400],
      "page_index": "all"
    }
  ]
}
註: box 代表擷取區域的像素位置：[x座標, y座標, 右下角x座標, 右下角y座標]。使用者只需量測一次即可填入。

4.2 PDF 處理與影像轉換 (Image Conversion)
程式需支援多頁 PDF 讀取。

轉換時建議設定 DPI 為 300 以上，以確保 OCR 辨識度。

輸入: .pdf 檔案路徑。

輸出: List[PIL.Image] (圖片物件列表)。

4.3 區域裁切與 OCR 核心 (Extraction Core)
程式需遍歷 config.json 中的 fields 清單。

針對每個欄位，在對應的頁面上進行裁切 (Crop)。

影像增強 (關鍵步驟):

將裁切圖片轉為灰階 (Grayscale)。

應用閾值 (Thresholding) 進行二值化，去除雜訊，讓數字變黑、背景變白。

OCR 設定:

設定 Tesseract 參數 --psm 7 (假設擷取目標為單行文字)。

若只需辨識數字，可設定白名單參數 (whitelist digits)。

4.4 數據輸出 (Data Output)
記憶體內結構: 程式運算過程中，產生如下的 Array 結構（List of Dictionaries）：

Python
[
  {"Page": 1, "溫度_T1": 85.5, "壓力_P1": 101.3},
  {"Page": 2, "溫度_T1": 86.0, "壓力_P1": 101.2},
  ...
]
Excel 儲存:

將上述結構轉為 Pandas DataFrame。

儲存為 .xlsx 檔案。

5. 程式碼邏輯大綱 (Pseudo Code)
Python
import pytesseract
from pdf2image import convert_from_path
import pandas as pd
import json

def load_config(path):
    # 讀取 JSON 設定檔
    pass

def preprocess_image(image):
    # 轉灰階 -> 二值化 -> 去雜訊
    return processed_image

def extract_data_from_pdf(config):
    # 1. 將 PDF 轉為圖片列表
    images = convert_from_path(config['input_pdf'], dpi=300)
    
    extracted_data = []

    # 2. 遍歷每一頁
    for page_num, img in enumerate(images, start=1):
        page_record = {"Page": page_num}
        
        # 3. 遍歷使用者設定的每個欄位
        for field in config['fields']:
            # 取得座標 (left, top, right, bottom)
            coords = field['box']
            
            # 裁切圖片
            cropped_img = img.crop(coords)
            
            # 影像前處理
            clean_img = preprocess_image(cropped_img)
            
            # 執行 OCR (配置為只辨識數字與小數點)
            text = pytesseract.image_to_string(
                clean_img, 
                config='--psm 7 -c tessedit_char_whitelist=0123456789.'
            )
            
            # 存入暫存字典
            page_record[field['name']] = text.strip()
        
        # 將此頁數據加入總表
        extracted_data.append(page_record)
        
    return extracted_data

def save_to_excel(data, output_path):
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False)
    print(f"數據已儲存至 {output_path}")

# Main Execution Flow
if __name__ == "__main__":
    conf = load_config('config.json')
    result_array = extract_data_from_pdf(conf)
    save_to_excel(result_array, conf['output_excel'])
6. 使用者操作流程 (User Manual Draft)
準備工作:

掃描文件並存為 PDF。

使用小畫家或截圖軟體打開其中一頁，將游標移到要擷取的數字左上角與右下角，記下座標 (x1, y1, x2, y2)。

設定欄位:

打開 config.json。

複製貼上新的欄位區塊，填入剛剛記下的座標與欄位名稱。

執行程式:

執行 Python script。

查看結果:

打開生成的 Excel 檔確認數據。