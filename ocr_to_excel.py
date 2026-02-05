#!/usr/bin/env python3
"""
掃描檔 PDF 定點數據擷取 (OCR)
根據 config.json 設定的像素座標裁切區域並進行 OCR，將結果以 JSON Array 格式輸出至標準輸出。
"""

import json
import os
from pathlib import Path

import pytesseract
from pdf2image import convert_from_path
from PIL import Image, ImageOps


def load_config(path: str = "config.json") -> dict:
    """讀取 JSON 設定檔，取得輸入 PDF 與欄位座標。"""
    config_path = Path(path)
    if not config_path.exists():
        raise FileNotFoundError(f"設定檔不存在: {path}")
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def preprocess_image(image: Image.Image) -> Image.Image:
    """
    影像前處理：轉灰階 -> 二值化 -> 去雜訊。
    讓數字變黑、背景變白，提升 OCR 準確率。
    """
    # 轉灰階
    gray = image.convert("L")
    # 二值化：Otsu 或固定閾值，此處用 0/255 二值化，確保黑字白底
    threshold = 150
    # 使用 Image.eval 進行像素級二值化，確保相容性
    binary = Image.eval(gray, lambda p: 255 if p > threshold else 0)
    # Image.eval 輸出已經是 L 模式 (0-255)，所以不需要再次 convert("L")
    return binary


def extract_data_from_pdf(config: dict) -> list[dict]:
    """
    依設定將 PDF 轉為圖片，對每頁依欄位座標裁切、前處理、OCR，回傳 List[dict]。
    """
    input_pdf = config.get("input_pdf")
    if not input_pdf or not Path(input_pdf).exists():
        raise FileNotFoundError(f"找不到 PDF 檔案: {input_pdf}")

    # 1. PDF 轉為高解析度圖片列表 (DPI >= 300)
    images = convert_from_path(input_pdf, dpi=300)
    extracted_data = []

    # 2. 遍歷每一頁
    for page_num, img in enumerate(images, start=1):
        page_record: dict = {"Page": page_num}

        # 3. 遍歷使用者設定的每個欄位
        for field in config.get("fields", []):
            # 若欄位指定特定頁，可依 page_index 過濾（此處支援 "all"）
            page_index = field.get("page_index", "all")
            if page_index != "all" and page_num != page_index:
                continue

            coords = field["box"]
            # box: [left, top, right, bottom] 像素
            left, top, right, bottom = coords[0], coords[1], coords[2], coords[3]
            cropped_img = img.crop((left, top, right, bottom))

            clean_img = preprocess_image(cropped_img)

            # OCR：單行；可依欄位設定 whitelist（空字串 = 不限制，辨識全部字元，利於血色素等）
            whitelist = field.get("ocr_whitelist", "0123456789.")
            lang = field.get("lang", "eng")
            tesseract_config = "--psm 7"
            if whitelist:
                tesseract_config += f" -c tessedit_char_whitelist={whitelist}"
            text = pytesseract.image_to_string(
                clean_img,
                config=tesseract_config,
                lang=lang,
            )
            value = text.strip()
            page_record[field["name"]] = value

        extracted_data.append(page_record)

    return extracted_data


def main() -> None:
    config_path = os.environ.get("OCR_CONFIG", "config.json")
    conf = load_config(config_path)
    result_array = extract_data_from_pdf(conf)

    # 將結果以 JSON 格式輸出到標準輸出
    # 使用 print 和 json.dumps 確保直接輸出純粹的 JSON 陣列
    print(json.dumps(result_array, ensure_ascii=False))


if __name__ == "__main__":
    main()
