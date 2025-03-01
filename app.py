# app.py
# Author: BLAIR
# Backend interface for PPT content extraction and optimization

import os
import logging
import time
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from modules.ppt_text_extraction import extract_text_from_ppt, extract_metadata, extract_text_from_ppt_legacy, extract_metadata_from_ppt_legacy
from modules.image_extraction_p import extract_images_from_ppt_paddleocr, extract_images_from_ppt_legacy
from modules.ai_optimizer import optimize_text_with_ai
from modules.utils import setup_logger, validate_file_type
from modules.config import OUTPUT_DIR_2
import re
import warnings
import socket
from dotenv import load_dotenv
from concurrent.futures import ThreadPoolExecutor
import tempfile

load_dotenv()

LOG_FILE = os.getenv("LOG_FILE", "ppt_processor.log")
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler(LOG_FILE), logging.StreamHandler()]
)
logger = setup_logger()

USE_GPU = os.getenv("USE_GPU", "False").lower() in ("true", "1", "yes")
logger.info(f"GPU enabled: {USE_GPU}")

app = Flask(__name__, static_folder='static', static_url_path='')
CORS(app, resources={r"/api/*": {"origins": "*"}})
executor = ThreadPoolExecutor(max_workers=1)

PROJECT_ROOT = os.path.abspath(os.path.dirname(__file__))
OUTPUT_DIR = os.path.abspath(os.getenv("OUTPUT_DIR", OUTPUT_DIR_2))
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    logger.info(f"Created output directory: {OUTPUT_DIR}")

SUPPORTED_FORMATS = ['.ppt', '.pptx', '.pot', '.potx', '.pps', '.ppsx', '.pptm', '.pdf']

def clean_text_output(text_list):
    try:
        cleaned_output = []
        watermark_pattern = re.compile(r'stablediffusionweb\.com')
        for slide in text_list:
            if not slide or not slide.strip():
                continue
            cleaned_slide = watermark_pattern.sub('', slide)
            cleaned_slide = re.sub(r'\n\s*\n', '\n', cleaned_slide.strip())
            cleaned_output.append(cleaned_slide)
        return cleaned_output
    except Exception as e:
        logger.error(f"Error cleaning text output: {str(e)}", exc_info=True)
        return []

@app.route('/')
def serve_index():
    try:
        return send_from_directory(app.static_folder, 'index.html')
    except Exception as e:
        logger.error(f"Error serving index.html: {str(e)}", exc_info=True)
        return jsonify({"error": "Failed to load index page"}), 500

@app.route('/<path:path>')
def serve_static(path):
    if os.path.exists(os.path.join(app.static_folder, path)):
        return send_from_directory(app.static_folder, path)
    return jsonify({"error": "File not found"}), 404

@app.route('/api/process_ppt', methods=['POST'])
def process_ppt():
    logger.info("Received POST request to /api/process_ppt")
    file_path = None
    try:
        if 'file' not in request.files:
            logger.error("No file part in the request")
            return jsonify({"error": "No file uploaded"}), 400
        
        file = request.files['file']
        if file.filename == '':
            logger.error("No file selected")
            return jsonify({"error": "No file selected"}), 400

        # 使用临时文件避免覆盖和权限问题
        temp_dir = tempfile.gettempdir()
        file_path = os.path.join(temp_dir, f"temp_{os.urandom(8).hex()}_{file.filename}")
        logger.info(f"Saving file to temporary path: {file_path}")
        file.save(file_path)

        if not validate_file_type(file_path, SUPPORTED_FORMATS):
            os.remove(file_path)
            logger.error(f"Invalid file type: {file.filename}. Supported formats: {SUPPORTED_FORMATS}")
            return jsonify({"error": f"Invalid file type: {file.filename}. Supported formats: {', '.join(SUPPORTED_FORMATS)}"}), 400

        if not os.path.exists(file_path):
            logger.error(f"File does not exist after saving: {file_path}")
            return jsonify({"error": "File save failed"}), 500

        logger.info(f"Starting processing for file: {file_path}")
        warnings.filterwarnings("ignore", category=UserWarning, module="PIL.Image")

        def process_file(file_path):
            ext = os.path.splitext(file_path.lower())[1]
            is_pptx = ext == '.pptx'

            metadata = extract_metadata(file_path) if is_pptx else extract_metadata_from_ppt_legacy(file_path)
            if "Error" in metadata:
                logger.error(f"Metadata extraction failed: {metadata['Error']}")
                metadata_output = "Metadata extraction failed\n"
            else:
                metadata_output = "\n".join([f"{key}: {value}" for key, value in metadata.items()])
            logger.info("Metadata processed")

            text_output = extract_text_from_ppt(file_path) if is_pptx else extract_text_from_ppt_legacy(file_path)
            if not text_output:
                logger.warning("No text extracted from PPT slides")
                text_output = []

            image_output = extract_images_from_ppt_paddleocr(file_path, OUTPUT_DIR, use_gpu=USE_GPU) if is_pptx else extract_images_from_ppt_legacy(file_path, OUTPUT_DIR, use_gpu=USE_GPU)
            if not image_output:
                logger.warning("No image text extracted")
                image_output = []

            cleaned_text_output = clean_text_output(text_output + image_output)
            combined_output = "\n".join([metadata_output] + cleaned_text_output)
            if not combined_output.strip():
                logger.warning("No combined output generated")
                combined_output = "No content extracted"

            optimized_text = optimize_text_with_ai(combined_output) or combined_output
            output_file = os.path.abspath(os.path.join(OUTPUT_DIR, f"optimized_output_{file.filename}.txt"))
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(optimized_text)
            logger.info(f"Optimized text saved to {output_file}")

            return optimized_text, output_file

        future = executor.submit(process_file, file_path)
        optimized_text, output_file = future.result(timeout=120)

        # 重试删除文件
        for _ in range(3):  # 尝试 3 次
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
                    logger.info(f"Temporary file removed: {file_path}")
                break
            except PermissionError:
                logger.warning(f"Retrying file removal: {file_path}")
                time.sleep(1)  # 等待 1 秒重试

        return jsonify({
            "message": "File processed successfully",
            "output_file": output_file,
            "optimized_text": optimized_text
        }), 200

    except PermissionError as e:
        logger.error(f"Permission denied: {str(e)}", exc_info=True)
        if file_path and os.path.exists(file_path):
            for _ in range(3):
                try:
                    os.remove(file_path)
                    logger.info(f"Temporary file removed after permission error: {file_path}")
                    break
                except PermissionError:
                    logger.warning(f"Retrying file removal after permission error: {file_path}")
                    time.sleep(1)
        return jsonify({"error": "Permission denied"}), 403
    except TimeoutError:
        logger.error(f"Processing timed out for file: {file_path}")
        if file_path and os.path.exists(file_path):
            os.remove(file_path)
            logger.info(f"Temporary file removed after timeout: {file_path}")
        return jsonify({"error": "Processing timed out"}), 504
    except Exception as e:
        logger.error(f"Error processing file: {str(e)}", exc_info=True)
        if file_path and os.path.exists(file_path):
            for _ in range(3):
                try:
                    os.remove(file_path)
                    logger.info(f"Temporary file removed after error: {file_path}")
                    break
                except PermissionError:
                    logger.warning(f"Retrying file removal after error: {file_path}")
                    time.sleep(1)
        return jsonify({"error": "Processing failed", "details": str(e)}), 500

@app.route('/health', methods=['GET'])
def health_check():
    logger.info("Health check requested")
    return jsonify({"status": "healthy", "message": "PPT Processor server is running"}), 200

def check_port(host, port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(1)
        try:
            s.bind((host, port))
            return True
        except OSError:
            logger.error(f"Port {port} is already in use")
            return False

if __name__ == "__main__":
    HOST = os.getenv("HOST", "0.0.0.0")
    PORT = int(os.getenv("PORT", 5000))
    MAX_PORT_ATTEMPTS = int(os.getenv("MAX_PORT_ATTEMPTS", 5))

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    static_dir = os.path.join(PROJECT_ROOT, 'static')
    if not os.path.exists(static_dir):
        os.makedirs(static_dir, exist_ok=True)
        logger.info(f"Created static directory: {static_dir}")

    if not os.path.exists(os.path.join(static_dir, 'index.html')):
        logger.warning("index.html not found in static folder. Please place it there.")

    try:
        for attempt in range(MAX_PORT_ATTEMPTS):
            if check_port(HOST, PORT):
                logger.info(f"Starting Flask server on {HOST}:{PORT}")
                app.run(host=HOST, port=PORT, debug=False, threaded=False)
                break
            else:
                PORT += 1
                logger.info(f"Trying next port: {PORT}")
        else:
            logger.error(f"Failed to find an available port after {MAX_PORT_ATTEMPTS} attempts")
            raise SystemExit(f"Startup failed: No available port found after {MAX_PORT_ATTEMPTS} attempts")
    except Exception as e:
        logger.error(f"Server startup failed: {str(e)}")
        raise