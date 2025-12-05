from pdf2image import convert_from_path
import pytesseract
from PIL import Image
import pandas as pd

# 将 PDF 转为图像
def pdf_to_images(pdf_path):
    images = convert_from_path(pdf_path)
    print(f"成功转换 {len(images)} 页图像")  # 输出转换的图像页数
    return images

# 使用 OCR 提取文本
def ocr_from_images(images):
    extracted_text = []
    for i, img in enumerate(images):
        text = pytesseract.image_to_string(img, lang='chi_sim')  # 中文识别
        print(f"第 {i+1} 页提取的文本：")
        print(text)  # 输出每一页提取的文本
        extracted_text.append(text)
    return extracted_text

# 假设 OCR 提取的文本是按行存储的，可以进一步处理
def process_extracted_text(extracted_text):
    all_data = []
    for page_text in extracted_text:
        rows = page_text.split("\n")  # 按行分割
        for row in rows:
            columns = row.split()  # 按空格分列
            if columns:  # 确保行不为空
                all_data.append(columns)
    return all_data

# 转换文本为 Excel 文件
def save_to_excel(all_data, excel_path):
    if all_data:  # 如果有数据
        df = pd.DataFrame(all_data)
        df.to_excel(excel_path, index=False)
        print(f"文件已成功保存为 {excel_path}")
    else:
        print("没有数据可以保存到 Excel 文件。")

# 使用示例
pdf_file = 'source/123.pdf'  # 替换为你的 PDF 文件路径
excel_file = 'output'  # 输出的 Excel 文件路径（完整文件名）

images = pdf_to_images(pdf_file)
extracted_text = ocr_from_images(images)
all_data = process_extracted_text(extracted_text)
save_to_excel(all_data, excel_file)
