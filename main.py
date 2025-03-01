# main.py
from modules.ppt_text_extraction import extract_text_from_ppt, extract_metadata, extract_text_from_ppt_legacy, extract_metadata_from_ppt_legacy
from modules.image_extraction_t import extract_images_from_ppt_tesseract
from modules.image_extraction_p import extract_images_from_ppt_paddleocr, extract_images_from_ppt_legacy
from modules.ai_optimizer import optimize_text_with_ai
from modules.utils import setup_logger, validate_file_type
from modules.config import PPTX_FILE, OUTPUT_DIR, PPTX_FILE_2, OUTPUT_DIR_2
import warnings
import re
import os
from concurrent.futures import ThreadPoolExecutor
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

logger = setup_logger()

# GPU option from environment variable (default to False if not set)
USE_GPU = os.getenv("USE_GPU", "False").lower() in ("true", "1", "yes")
logger.info(f"GPU enabled: {USE_GPU}")

# Thread pool for async processing
executor = ThreadPoolExecutor(max_workers=2)

def process_ppt_file(file_path):
    warnings.filterwarnings("ignore", category=UserWarning, module="PIL.Image")
    if not validate_file_type(file_path, ['.ppt', '.pptx', '.pot', '.potx', '.pps', '.ppsx', '.pptm', '.pdf']):
        logger.error(f"Invalid file type: {file_path}. Skipping.")
        return

    logger.info(f"Processing file: {file_path}")
    
    try:
        # Asynchronous processing function
        def process_file(file_path):
            ext = os.path.splitext(file_path.lower())[1]
            is_pptx = ext == '.pptx'

            # Extract metadata
            metadata = extract_metadata(file_path) if is_pptx else extract_metadata_from_ppt_legacy(file_path)
            if "Error" in metadata:
                logger.error(f"Metadata extraction failed: {metadata['Error']}")
                raise Exception(f"Metadata extraction failed: {metadata['Error']}")
            metadata_output = "\n".join([f"{key}: {value}" for key, value in metadata.items()])
            logger.info("Metadata extracted successfully.")

            # Extract slide text
            text_output = extract_text_from_ppt(file_path) if is_pptx else extract_text_from_ppt_legacy(file_path)
            if not text_output:
                logger.warning("No text extracted from PPT slides.")
                text_output = []

            # Extract image text (using PaddleOCR with GPU option)
            image_output = extract_images_from_ppt_paddleocr(file_path, OUTPUT_DIR_2, use_gpu=USE_GPU) if is_pptx else extract_images_from_ppt_legacy(file_path, OUTPUT_DIR_2, use_gpu=USE_GPU)
            if not image_output:
                logger.warning("No image text extracted.")
                image_output = []

            # Clean and combine text
            cleaned_text_output = clean_text_output(text_output + image_output)
            combined_output = "\n".join([metadata_output] + cleaned_text_output)
            if not combined_output.strip():
                logger.warning("No combined output generated")
                combined_output = "No content extracted"

            # Optimize text with AI
            optimized_text = optimize_text_with_ai(combined_output)
            if not optimized_text:
                logger.warning("Text optimization returned empty, using combined output as fallback")
                optimized_text = combined_output

            # Save result
            output_file = os.path.abspath(os.path.join(OUTPUT_DIR_2, "optimized_output.txt"))
            os.makedirs(OUTPUT_DIR_2, exist_ok=True)
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(optimized_text)
            logger.info(f"File processed successfully: {output_file}")

            return optimized_text

        # Run processing in a thread with timeout
        future = executor.submit(process_file, file_path)
        optimized_text = future.result(timeout=120)  # 2 分钟超时

    except TimeoutError:
        logger.error(f"Processing timed out for file: {file_path}")
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