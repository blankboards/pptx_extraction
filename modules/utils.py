import logging
import os
from typing import List, Optional

# 默认日志格式
DEFAULT_LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"


def validate_file_type(file_path: str, valid_extensions: List[str], raise_exception: bool = False) -> bool:
    """
    验证文件路径是否有效且扩展名在允许的列表中。

    Args:
        file_path (str): 文件路径。
        valid_extensions (List[str]): 允许的扩展名列表（如 ['.pptx', '.pdf']）。
        raise_exception (bool): 是否在验证失败时抛出异常，默认为 False。

    Returns:
        bool: 如果文件有效且扩展名匹配，返回 True；否则返回 False。

    Raises:
        FileNotFoundError: 如果 raise_exception=True 且文件不存在。
        ValueError: 如果 raise_exception=True 且扩展名无效。
    """
    logger = logging.getLogger(__name__)

    # 检查文件是否存在
    if not os.path.isfile(file_path):
        error_msg = f"The file {file_path} does not exist."
        logger.error(error_msg)
        if raise_exception:
            raise FileNotFoundError(error_msg)
        return False

    # 获取文件扩展名
    _, ext = os.path.splitext(file_path)
    ext = ext.lower()

    # 规范化扩展名列表（去除多余的点号）
    valid_extensions = [e.lower().lstrip('.') for e in valid_extensions]

    # 检查扩展名是否匹配
    is_valid = ext[1:] in valid_extensions  # 去掉点号比较
    if not is_valid:
        error_msg = f"Invalid file type: {ext}. Expected one of {valid_extensions}."
        logger.warning(error_msg)
        if raise_exception:
            raise ValueError(error_msg)
    
    logger.debug(f"Validated file: {file_path} with extension {ext}")
    return is_valid


def setup_logger(
    name: str = __name__,
    level: int = logging.INFO,
    log_file: Optional[str] = None,
    log_format: str = DEFAULT_LOG_FORMAT
) -> logging.Logger:
    """
    配置并返回一个日志器，支持控制台和文件输出。

    Args:
        name (str): 日志器名称，默认为当前模块名。
        level (int): 日志级别，默认为 logging.INFO。
        log_file (Optional[str]): 日志文件路径，若提供则记录到文件。
        log_format (str): 日志输出格式，默认为时间-级别-消息。

    Returns:
        logging.Logger: 配置好的日志器实例。
    """
    # 获取或创建日志器
    logger = logging.getLogger(name)
    
    # 避免重复配置
    if logger.handlers:
        logger.debug("Logger already configured, skipping setup.")
        return logger

    # 设置日志级别
    logger.setLevel(level)

    # 创建格式器
    formatter = logging.Formatter(log_format)

    # 添加控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # 如果指定了日志文件，添加文件处理器
    if log_file:
        try:
            os.makedirs(os.path.dirname(log_file), exist_ok=True)
            file_handler = logging.FileHandler(log_file, encoding="utf-8")
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
            logger.info(f"Logging to file: {log_file}")
        except Exception as e:
            logger.error(f"Failed to set up file logging: {e}")

    logger.info(f"Logger initialized with level {logging.getLevelName(level)}")
    return logger


# 测试代码
if __name__ == "__main__":
    # 设置日志
    logger = setup_logger(level=logging.DEBUG, log_file="logs/test.log")
    
    # 测试文件验证
    valid_extensions = ['ppt', 'pptx', 'pdf']
    test_files = [
        "sample.pptx",
        "sample.doc",
        "nonexistent.pptx"
    ]
    
    for file in test_files:
        result = validate_file_type(file, valid_extensions, raise_exception=False)
        logger.info(f"File {file} is valid: {result}")