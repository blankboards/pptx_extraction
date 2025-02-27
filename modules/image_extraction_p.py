from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import os
import logging
from paddleocr import PaddleOCR
from modules.config import OUTPUT_DIR

# 配置日志
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

# 初始化 PaddleOCR
ocr = PaddleOCR(use_angle_cls=True, lang='ch', use_gpu=False)  # 可根据需要启用 GPU

def ensure_dir(directory):
    """确保目录存在，若不存在则创建"""
    if not os.path.exists(directory):
        os.makedirs(directory)
        logger.info(f"Created directory: {directory}")

def extract_images_from_ppt_paddleocr(file_path, output_dir=OUTPUT_DIR, output_format="text"):
    """
    从 PPT 文件中提取图片，并使用 PaddleOCR 识别图片中的文本。

    Args:
        file_path (str): PPT 文件路径。
        output_dir (str): 输出目录路径，默认为 OUTPUT_DIR。
        output_format (str): 输出格式，可选 "text"（默认）或 "json"。

    Returns:
        list: 包含每张图片识别文本的列表。
    """
    try:
        image_texts = []
        presentation = Presentation(file_path)
        logger.info(f"Processing PPT file: {file_path}")

        for slide_number, slide in enumerate(presentation.slides, start=1):
            slide_folder = os.path.join(output_dir, f"slide_{slide_number}", "image")
            ensure_dir(slide_folder)
            image_index = 1

            for shape in slide.shapes:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    image = shape.image
                    image_bytes = image.blob
                    image_name = f"image_{image_index}.jpg"
                    image_path = os.path.join(slide_folder, image_name)

                    # 保存图片
                    with open(image_path, "wb") as f:
                        f.write(image_bytes)
                    logger.debug(f"Saved image: {image_path}")

                    # 检查并提取文本
                    if contains_text(image_path):
                        text = process_image_for_ocr(image_path)
                        if text:
                            text_entry = {
                                "slide": slide_number,
                                "image": image_index,
                                "text": text
                            }
                            if output_format == "json":
                                image_texts.append(text_entry)
                            else:
                                image_texts.append(f"\nSlide {slide_number}, Image {image_index} Text:\n{text}")

                            # 保存到文件
                            text_file_path = os.path.join(slide_folder, f"slide_{slide_number}_image_{image_index}_text.txt")
                            with open(text_file_path, "w", encoding="utf-8") as f:
                                f.write(text)
                            logger.info(f"Extracted text from {image_name}: {text[:50]}...")
                        else:
                            logger.debug(f"No readable text in {image_name}")
                    else:
                        logger.debug(f"No text detected in {image_name}")
                    image_index += 1

        logger.info(f"Completed processing {file_path}. Extracted {len(image_texts)} image texts.")
        return image_texts

    except Exception as e:
        logger.error(f"Failed to process PPT file {file_path}: {e}")
        return []

def process_image_for_ocr(image_path):
    """
    使用 PaddleOCR 处理图片并提取文本。

    Args:
        image_path (str): 图片文件路径。

    Returns:
        str: 提取的文本内容。
    """
    try:
        result = ocr.ocr(image_path, cls=True)
        if not result or not result[0]:
            return ""

        texts = []
        for line in result[0]:
            if line and len(line[1][0]) >= 2 and line[1][1] > 0.5:  # 置信度 > 0.5，文本长度 >= 2
                texts.append(line[1][0])
        full_text = '\n'.join(texts).strip()
        return full_text if full_text else ""

    except Exception as e:
        logger.error(f"OCR processing failed for {image_path}: {e}")
        return ""

def contains_text(image_path):
    """
    使用 PaddleOCR 判断图片是否含有文本。

    Args:
        image_path (str): 图片文件路径。

    Returns:
        bool: 如果图片中含有文本，返回 True；否则返回 False。
    """
    try:
        result = ocr.ocr(image_path, cls=True)
        if result and result[0]:
            valid_texts = [
                line for line in result[0]
                if line and len(line[1][0]) >= 2 and line[1][1] > 0.5
            ]
            return len(valid_texts) > 0
        return False

    except Exception as e:
        logger.error(f"Text detection failed for {image_path}: {e}")
        return False

# 测试代码
if __name__ == "__main__":
    sample_ppt = "path/to/your/sample.pptx"
    result = extract_images_from_ppt_paddleocr(sample_ppt, output_format="text")
    for text in result:
        print(text)