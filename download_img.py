import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
import requests
from io import BytesIO
import os
from urllib.parse import urlparse

# --- 用户配置区 ---
SOURCE_EXCEL_FILE = r'C:\Users\Lenovo\Desktop\BCG\女装-消费品\图片处理-polo.xlsx'  # 替换成你的Excel文件名
OUTPUT_EXCEL_FILE = r'C:\Users\Lenovo\Desktop\BCG\女装-消费品\图片处理结果-拉夫劳伦.xlsx' # 输出的新文件名

URL_COLUMN_NAME = '商品主图'             # 包含图片链接的列名
IMAGE_INSERT_COLUMN_NAME = '图片预览'     # 新增的用于存放图片的列名

# 设置图片在单元格中的尺寸 (可以根据需要调整)
MAX_IMAGE_WIDTH = 150  # 图片最大宽度 (像素)
MAX_IMAGE_HEIGHT = 150 # 图片最大高度 (像素)
# --- 配置区结束 ---

def download_and_insert_images():
    """
    主函数：读取Excel，下载图片并插入，最后保存新文件。
    """
    # 检查源文件是否存在
    if not os.path.exists(SOURCE_EXCEL_FILE):
        print(f"错误：源文件 '{SOURCE_EXCEL_FILE}' 不存在。请检查文件名和路径。")
        return

    # 1. 使用pandas读取Excel数据
    try:
        df = pd.read_excel(SOURCE_EXCEL_FILE)
        print("成功读取Excel文件。")
    except Exception as e:
        print(f"读取Excel文件时出错: {e}")
        return

    # 2. 检查图片链接列是否存在
    if URL_COLUMN_NAME not in df.columns:
        print(f"错误：在Excel中找不到名为 '{URL_COLUMN_NAME}' 的列。")
        print(f"可用的列有: {list(df.columns)}")
        return
        
    # 3. 在DataFrame中找到图片链接列的位置，并在其后插入新列名
    url_col_index = df.columns.get_loc(URL_COLUMN_NAME)
    df.insert(url_col_index + 1, IMAGE_INSERT_COLUMN_NAME, '') # 插入一个空列作为占位符
    
    # 4. 将更新后的DataFrame写入到新的Excel文件，为后续操作图片做准备
    df.to_excel(OUTPUT_EXCEL_FILE, index=False)
    print(f"已创建新文件 '{OUTPUT_EXCEL_FILE}' 并写入基础数据。")

    # 5. 使用openpyxl打开新创建的Excel文件，以进行图片插入操作
    wb = load_workbook(OUTPUT_EXCEL_FILE)
    ws = wb.active

    # 6. 找到新列的列号和列字母
    # openpyxl列号从1开始，所以索引+1
    image_col_idx = url_col_index + 2 
    image_col_letter = get_column_letter(image_col_idx)

    # 7. 调整新列的宽度和行的高度
    ws.column_dimensions[image_col_letter].width = MAX_IMAGE_WIDTH / 7  # 粗略转换像素到宽度单位
    for i in range(2, len(df) + 2): # 从第2行开始，因为第1行是标题
        ws.row_dimensions[i].height = MAX_IMAGE_HEIGHT * 0.75 # 粗略转换像素到高度单位
    print("已调整图片列的单元格大小。")

    # 8. 遍历每一行，下载并插入图片
    print("\n开始下载并插入图片...")
    # iter_rows(min_row=2) 从第二行开始遍历
    for index, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=False)):
        # pandas的索引是0开始，所以用index
        url = df.loc[index, URL_COLUMN_NAME]
        
        # 目标单元格
        target_cell = ws.cell(row=index + 2, column=image_col_idx)

        # 检查URL是否有效
        if pd.isna(url) or not str(url).startswith('http'):
            print(f"第 {index + 2} 行：跳过无效或空的URL。")
            continue

        try:
            # 下载图片
            response = requests.get(url, stream=True, timeout=10)
            response.raise_for_status()  # 如果下载失败，会抛出异常
            
            # 将图片数据读入内存
            image_data = BytesIO(response.content)
            img = Image(image_data)
            
            # 计算缩放比例，保持纵横比
            ratio = min(MAX_IMAGE_WIDTH / img.width, MAX_IMAGE_HEIGHT / img.height)
            img.width = img.width * ratio
            img.height = img.height * ratio

            # 将图片添加到工作表
            ws.add_image(img, target_cell.coordinate)
            
            print(f"第 {index + 2} 行：成功插入图片。")

        except requests.exceptions.RequestException as e:
            print(f"第 {index + 2} 行：下载图片失败 - {e}")
        except Exception as e:
            print(f"第 {index + 2} 行：处理图片时出错 - {e}")
            
    # 9. 保存最终的Excel文件
    try:
        wb.save(OUTPUT_EXCEL_FILE)
        print(f"\n所有任务完成！结果已保存到 '{OUTPUT_EXCEL_FILE}'。")
    except Exception as e:
        print(f"保存最终文件时出错: {e}")

# --- 运行主函数 ---
if __name__ == "__main__":
    download_and_insert_images()