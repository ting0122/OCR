# 掃描檔 PDF 定點數據擷取與 Excel 自動化 (OCR-to-Excel)

依 `config.json` 設定的像素座標，從多頁掃描 PDF 中裁切指定區域並進行 OCR，將結果寫入 Excel。

## 環境需求

- **Python 3.9+**
- **Poppler**：供 `pdf2image` 轉 PDF 為圖片  
  - Ubuntu/Debian: `sudo apt install poppler-utils`  
  - macOS: `brew install poppler`
- **Tesseract OCR**  
  - Ubuntu/Debian: `sudo apt install tesseract-ocr`  
  - macOS: `brew install tesseract`

## 安裝

```bash
pip install -r requirements.txt
```

## 設定檔 `config.json`

| 鍵 | 說明 |
|----|------|
| `input_pdf` | 輸入 PDF 路徑（例：`data/scan_log_2024.pdf`） |
| `output_excel` | 輸出 Excel 路徑（例：`output/result.xlsx`） |
| `fields` | 欄位陣列，每個欄位含 `name`、`box`、`page_index` |

**box**：`[left, top, right, bottom]` 像素座標，對應擷取區域的左上與右下角。  
**page_index**：`"all"` 表示每一頁都擷取；也可填數字指定單一頁。  
**cell**（選填）：Excel 儲存格參照（如 `"B2"`），該欄位會寫入此欄、從此列開始，每頁一列；標題會寫在上一列。  
**ocr_whitelist**（選填）：OCR 允許的字元，例如 `"0123456789."` 只辨識數字與小數點；設為 `""` 則不限制（辨識全部字元，適合血色素等可能含英文或特殊符號的欄位）。  
**lang**（選填）：Tesseract 語系，預設 `"eng"`；若需中文可設 `"chi_tra+eng"`。

頂層可設 **omit_page_column**（選填）：設為 `true` 時不輸出 Page 欄，僅輸出各欄位，直向列出（每欄一個欄位、每頁一列）。  
頂層可設 **page_cell**（選填）：有輸出頁碼時，頁碼欄的起始儲存格，預設 `"A2"`（標題在 A1，頁碼從 A2 開始）。

### 取得座標

1. 用小畫家或截圖軟體開啟 PDF 其中一頁的畫面（或匯出的一張圖）。
2. 將游標移到要擷取數字的**左上角**與**右下角**，記下 (x1, y1) 與 (x2, y2)。
3. 在 `config.json` 的 `box` 填入 `[x1, y1, x2, y2]`。

## 執行

```bash
python ocr_to_excel.py
```

或指定設定檔路徑：

```bash
OCR_CONFIG=my_config.json python ocr_to_excel.py
```

## 輸出格式

程式會產生 List of dict，再寫入 Excel，例如：

| Page | 溫度_T1 | 壓力_P1 |
|------|---------|---------|
| 1    | 85.5    | 101.3   |
| 2    | 86.0    | 101.2   |

## 目錄結構建議

```
OCR/
├── config.json       # 設定檔
├── ocr_to_excel.py   # 主程式
├── requirements.txt
├── data/             # 放置輸入 PDF
└── output/           # 輸出 Excel（會自動建立）
```

## 規格說明

詳見專案內 `spec.md`。
