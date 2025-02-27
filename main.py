from modules.ppt_text_extraction import extract_text_from_ppt, extract_metadata
from modules.image_extraction_t import extract_images_from_ppt_tesseract
from modules.image_extraction_p import extract_images_from_ppt_paddleocr
from modules.ai_optimizer import optimize_text_with_ai
from modules.utils import setup_logger, validate_file_type
from modules.config import PPTX_FILE, OUTPUT_DIR, PPTX_FILE_2, OUTPUT_DIR_2
import warnings
import re

logger = setup_logger()

def process_ppt_file(file_path):
    warnings.filterwarnings("ignore", category=UserWarning, module="PIL.Image")
    if not validate_file_type(file_path, ['.ppt', '.pptx', '.pot', '.potx', '.pps', '.ppsx', 'pptm', 'pdf']):
        logger.error(f"Invalid file type: {file_path}. Skipping.")
        return

    logger.info(f"Processing file: {file_path}")
    
    try:
        # 提取元数据
        metadata = extract_metadata(file_path)
        metadata_output = "\n".join([f"{key}: {value}" for key, value in metadata.items()])
        logger.info("Metadata extracted successfully.")
        
        # 提取幻灯片文本
        text_output = extract_text_from_ppt(file_path)
        if not text_output:
            logger.warning("No text extracted from PPT slides.")
        
        # 提取图片文本（使用 PaddleOCR）
        # image_output = extract_images_from_ppt_tesseract(file_path, OUTPUT_DIR_2)
        image_output = extract_images_from_ppt_paddleocr(file_path, OUTPUT_DIR_2)
        if not image_output:
            logger.warning("No image text extracted.")
        
        # 清理无关内容（如重复的水印）
        cleaned_text_output = clean_text_output(text_output + image_output)
        combined_output = "\n".join([metadata_output] + cleaned_text_output)
        
        # optimized_text = combined_output
        # AI 优化文本
        optimized_text = optimize_text_with_ai(combined_output)
        if not optimized_text:
            logger.error("Text optimization failed.")
            return
        
        # 保存结果
        output_file = f"{OUTPUT_DIR_2}/optimized_output.txt"
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(optimized_text)
        logger.info(f"File processed successfully: {output_file}")
    
    except Exception as e:
        logger.error(f"Error processing file {file_path}: {str(e)}")

def clean_text_output(text_list):
    """清理提取的文本，去除冗余信息（如水印）并优化结构"""
    cleaned_output = []
    watermark_pattern = re.compile(r'stablediffusionweb\.com')  # 假设这是常见的水印
    for slide in text_list:
        # 跳过空幻灯片
        if not slide.strip():
            continue
        # 去除水印
        cleaned_slide = watermark_pattern.sub('', slide)
        # 清理多余换行和空格
        cleaned_slide = re.sub(r'\n\s*\n', '\n', cleaned_slide.strip())
        cleaned_output.append(cleaned_slide)
    return cleaned_output

def main():
    # process_ppt_file(PPTX_FILE)
    process_ppt_file(PPTX_FILE_2)

if __name__ == "__main__":
    main()