import pytesseract
from PIL import Image
import pandas as pd
import os


# 设置 Tesseract 路径（如果需要的话）
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# 读取图像文件夹中的所有图像并使用 OCR 提取文本
def ocr_from_images(image_folder):
    extracted_text = []
    # 获取文件夹中所有图片文件（可以是 .png、.jpg 等）
    image_files = [f for f in os.listdir(image_folder) if f.endswith('.png') or f.endswith('.jpg')]
    print(f"读取到 {len(image_files)} 张图片")  # 打印读取到的图片数量

    for i, image_file in enumerate(image_files):
        image_path = os.path.join(image_folder, image_file)
        img = Image.open(image_path)
        text = pytesseract.image_to_string(img, lang='chi_sim')  # 中文识别
        print(f"第 {i + 1} 张图片提取的文本：")
        print(text)  # 输出每一张图片提取的文本
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


# 保存每张图像的 OCR 结果为单独的 Excel 文件
def save_to_excel(extracted_text, output_folder):
    # 确保输出文件夹存在
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 为每张图片生成一个单独的 Excel 文件
    for i, text in enumerate(extracted_text):
        all_data = process_extracted_text([text])
        excel_file = os.path.join(output_folder, f"page_{i + 1}.xlsx")
        if all_data:  # 如果有数据
            df = pd.DataFrame(all_data)
            df.to_excel(excel_file, index=False)
            print(f"第 {i + 1} 张图片的 OCR 结果已保存为 {excel_file}")
        else:
            print(f"第 {i + 1} 张图片没有提取到文本。")


# 使用示例
image_folder = 'png'  # 图像文件夹路径，包含所有需要提取文本的图片
output_folder = 'output_excel'  # 输出的 Excel 文件夹路径

# 从图像中提取文本
extracted_text = ocr_from_images(image_folder)

# 保存每张图片的 OCR 结果为单独的 Excel 文件
save_to_excel(extracted_text, output_folder)
