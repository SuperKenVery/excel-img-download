from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
import httpx
from httpx_retries import RetryTransport, Retry
import asyncio
from io import BytesIO
import os
from pathlib import Path
from tqdm.asyncio import tqdm_asyncio
from tqdm import tqdm

# --- 用户配置区 ---
SOURCE_EXCEL_FILES = [
    r'/Users/ken/Codes/hsy/excel-img-download/test/图片处理-polo.xlsx'
]  # 要处理的文件名列表

URL_COLUMN_NAMES = ['商品主图', '商品图片']             # 包含图片链接的列名
IMAGE_INSERT_COLUMN_NAME = '图片预览'     # 新增的用于存放图片的列名

# 设置图片在单元格中的尺寸 (可以根据需要调整)
MAX_IMAGE_WIDTH = 150  # 图片最大宽度 (像素)
MAX_IMAGE_HEIGHT = 150 # 图片最大高度 (像素)
# --- 配置区结束 ---

async def download_image(url):
    """
    下载图片并返回处理后的Image对象，使用指数退避重试机制

    Args:
        url (str): 图片URL

    Returns:
        Image: 处理后的Image对象，如果下载失败返回None
    """
    retry = Retry(total=5)

    try:
        async with httpx.AsyncClient(transport=RetryTransport(retry=retry)) as client:
            response = await client.get(url, timeout=5.0)
            response.raise_for_status()  # 如果下载失败，会抛出异常

            # 将图片数据读入内存
            image_data = BytesIO(response.content)
            img = Image(image_data)

            # 计算缩放比例，保持纵横比
            ratio = min(MAX_IMAGE_WIDTH / img.width, MAX_IMAGE_HEIGHT / img.height)
            img.width = int(img.width * ratio)
            img.height = int(img.height * ratio)

            return img

    except httpx.RequestError as e:
        tqdm.write(f"下载图片失败: {e}")
        return None
    except Exception as e:
        tqdm.write(f"处理图片时出错: {e}")
        return None

async def process_one_worksheet(ws):
    """
    处理单个工作表中的图片下载和插入

    Args:
        ws: openpyxl的工作表对象
    """
    print(f"\n开始处理工作表: '{ws.title}'")

    # 1. 查找URL列的位置
    url_col_idx = None
    for cell in ws[1]:  # 第一行是标题行
        for keyword in URL_COLUMN_NAMES:
            if type(cell.value)==str and keyword in cell.value:
                url_col_idx = cell.column
                break

    # 检查URL列是否存在
    if url_col_idx is None:
        print(f"工作表 '{ws.title}'：找不到名为 '{URL_COLUMN_NAMES}' 的列，跳过此工作表。")
        return

    print(f"工作表 '{ws.title}'：找到URL列 '{URL_COLUMN_NAMES}' 在第 {url_col_idx} 列。")

    # 2. 在URL列后插入新列
    image_col_idx = url_col_idx + 1
    ws.insert_cols(image_col_idx)

    # 3. 设置新列的标题
    ws.cell(row=1, column=image_col_idx).value = IMAGE_INSERT_COLUMN_NAME

    # 4. 调整新列的宽度和行的高度
    image_col_letter = get_column_letter(image_col_idx)
    ws.column_dimensions[image_col_letter].width = MAX_IMAGE_WIDTH / 7  # 粗略转换像素到宽度单位
    for i in range(2, ws.max_row + 1):  # 从第2行开始，因为第1行是标题
        ws.row_dimensions[i].height = MAX_IMAGE_HEIGHT * 0.75  # 粗略转换像素到高度单位

    # 5. 遍历每一行，下载并插入图片
    print(f"工作表 '{ws.title}'：开始下载并插入图片...")

    download_tasks = [
        process_line(ws, line_idx, url_col_idx,image_col_idx )
        for line_idx in range(2, ws.max_row+1)
    ]

    await tqdm_asyncio.gather(*download_tasks, desc=f"工作表{ws.title}")

async def process_line(worksheet, line_idx, src_col_idx, dst_col_idx):
    url_cell = worksheet.cell(row=line_idx, column=src_col_idx)
    url = url_cell.value
    if url is None or url=="":
        print(f"工作表 '{worksheet.title}'：第 {line_idx} 行：跳过空的URL。")

    img = await download_image(url)
    if img is None:
        print(f"工作表 '{worksheet.title}'：第 {line_idx} 行：跳过无效或下载失败的URL。")
        return

    target_cell = worksheet.cell(row=line_idx, column=dst_col_idx)
    worksheet.add_image(img, target_cell.coordinate)

async def process_one_excel_file(source: str, dest: str):
    """
    主函数：读取Excel，对每个工作表下载图片并插入，最后保存新文件。
    """
    # 1. 使用openpyxl读取Excel数据
    wb = load_workbook(source)
    print(f"成功读取Excel文件，包含 {len(wb.sheetnames)} 个工作表: {wb.sheetnames}")

    # 2. 对每个工作表执行处理
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        await process_one_worksheet(ws)

    # 3. 保存最终的Excel文件
    try:
        wb.save(dest)
        print(f"\n所有任务完成！结果已保存到 '{dest}'。")
    except Exception as e:
        print(f"保存最终文件时出错: {e}")

async def process_excel_files():
    for src in SOURCE_EXCEL_FILES:
        filename = Path(src)
        dest = filename.parent / f"{filename.stem}-图片已下载{filename.suffix}"
        print(f"正在处理{src}")
        await process_one_excel_file(src, str(dest))

# --- 运行主函数 ---
if __name__ == "__main__":
    asyncio.run(process_excel_files())
