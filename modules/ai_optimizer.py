# import google.generativeai as genai
# import os
# from dotenv import load_dotenv
# import logging

# # 配置日志
# logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
# logger = logging.getLogger(__name__)

# # 载入环境变量
# load_dotenv()

# # 设置 Gemini API Key
# genai.configure(api_key=os.getenv("GOOGLE_GEMINI_API_KEY"))

# def optimize_text_with_ai(text):
#     try:
#         return text
#         # 选择 Gemini 2 Flash 模型
#         model = genai.GenerativeModel("gemini-1.5-flash")

#         # 发送请求
#         response = model.generate_content(f"请优化以下文本，使其更加通顺，去除乱码，并在每句话末尾加上问号：\n\n{text}")

#         # 返回优化后的文本
#         return response.text if response and response.text else text
#     except Exception as e:
#         logger.error(f"Error calling Google AI API: {e}")
#         return text

###############################################

# import re
# import logging
# import random

# # 配置日志
# logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
# logger = logging.getLogger(__name__)

# def optimize_text_with_ai(text):
#     """
#     本地优化 PPT 提取的文本，使其更生动、适合有声课件。
#     改进图片文本处理，跳过空内容，增强语流。
#     """
#     try:
#         # 分割文本为行
#         lines = text.split('\n')
#         optimized_lines = []
#         in_list = False
#         slide_count = 0
        
#         # 多样化的引导语和过渡词
#         title_openers = ["本节我们来聊聊", "接下来我们看看", "现在我们进入", "这一部分是关于"]
#         list_openers = ["首先来看", "我们先说说", "这里有几点", "让我们了解一下"]
#         list_connectors = ["接着是", "还有", "另外", "再比如"]
#         formula_openers = ["这里有个公式", "我们来看一个公式", "简单来说，这有个计算过程"]
#         image_openers = ["这里有一张图片", "我们来看一张图", "这部分有张示例图"]
        
#         for line in lines:
#             # 跳过空行
#             if not line.strip():
#                 continue
            
#             # 处理元数据（只保留关键信息）
#             if ':' in line and slide_count == 0:
#                 if not any(k in line for k in ["Title", "Author", "Created", "Modified"]):
#                     continue
#                 optimized_lines.append(line)
#                 continue
            
#             # 检测幻灯片分隔符
#             if line.startswith('@@@Slide_'):
#                 slide_count += 1
#                 if slide_count > 1:
#                     optimized_lines.append("\n好了，刚才的内容就先讲到这里。")
#                 optimized_lines.append(f"\n{random.choice(title_openers)}幻灯片 {slide_count} 的内容：")
#                 in_list = False
#                 continue
            
#             # 清理乱码和水印
#             cleaned_line = re.sub(r'stablediffusionweb\.com|WHATEV-VEER\.', '', line).strip()
#             if not cleaned_line or cleaned_line in ['.', ',']:  # 跳过空行或仅剩标点
#                 continue
            
#             # 检测图片文本
#             if 'Image' in line:
#                 if cleaned_line == line.strip():  # 无清理，说明有实际内容
#                     optimized_lines.append(f"{random.choice(image_openers)}，内容是：{cleaned_line}。")
#                 else:  # 被清理为空，添加默认说明
#                     optimized_lines.append(f"{random.choice(image_openers)}，具体内容请参考幻灯片。")
#                 continue
            
#             # 检测标题（非列表、非公式）
#             if not (cleaned_line.startswith(('▪', '-', '•')) or '=' in cleaned_line or '∑' in cleaned_line or 'σ' in cleaned_line):
#                 if in_list:
#                     optimized_lines.append("以上就是这部分的要点，挺有意思吧？")
#                     in_list = False
#                 optimized_lines.append(f"{random.choice(title_openers)} {cleaned_line}。")
#                 continue
            
#             # 检测列表项
#             if cleaned_line.startswith(('▪', '-', '•')):
#                 if not in_list:
#                     optimized_lines.append(f"{random.choice(list_openers)}：")
#                     in_list = True
#                     connector = ""
#                 else:
#                     connector = random.choice(list_connectors)
#                 content = cleaned_line[1:].strip()
#                 optimized_lines.append(f"{connector} {content}，这部分其实挺重要的。")
#                 continue
            
#             # 检测公式
#             if '=' in cleaned_line or '∑' in cleaned_line or 'σ' in cleaned_line:
#                 if in_list:
#                     optimized_lines.append("这些要点就先讲到这儿。")
#                     in_list = False
#                 optimized_lines.append(f"{random.choice(formula_openers)}：{cleaned_line}，它其实是在帮助模型更好地工作。")
#                 continue
        
#         # 添加结束语
#         if in_list:
#             optimized_lines.append("以上就是这部分的要点啦。")
#         optimized_lines.append("\n好了，这就是今天的全部内容，希望大家有所收获，咱们下次再见！")
        
#         # 合并并返回
#         optimized_text = "\n".join(optimized_lines)
#         logger.info("Text optimization completed successfully.")
#         return optimized_text
    
#     except Exception as e:
#         logger.error(f"Error optimizing text: {e}")
#         return text

# # 测试代码
# if __name__ == "__main__":
#     sample_text = """Title: N/A
#     @@@Slide_4@@@
#     卷积运算可视化
#     Slide 4, Image 1 Text:
#     WHATEV-VEER.
#     stablediffusionweb.com
#     @@@Slide_5@@@
#     典型应用场景
#     Slide 5, Image 1 Text:
#     stablediffusionweb.com
#     Slide 5, Image 2 Text:
#     stablediffusionweb.com
#     Slide 5, Image 3 Text:
#     stablediffusionweb.com
#     """
#     result = optimize_text_with_ai(sample_text)
#     print(result)

#################################

import re
import logging
import random

# 配置日志
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

def optimize_text_with_ai(text):
    """
    本地优化 PPT 提取的文本，使其更自然、适合有声课件讲解。
    """
    try:
        lines = text.split('\n')
        optimized_lines = []
        in_list = False
        slide_count = 0
        prev_title = ""

        # 多样化的引导语和连接词
        slide_openers = ["现在我们来看看", "接下来聊聊", "这一节我们讲讲", "让我们进入"]
        title_openers = ["首先是", "那么我们说说", "这里有个重点", "值得一提的是"]
        list_openers = ["先来看看", "我们聊聊", "这里包括", "比如说"]
        list_connectors = ["然后是", "接着是", "还有呢", "另外一点"]
        formula_openers = ["这里有个公式", "看看这个计算", "简单解释一下"]
        image_openers = ["这里有张图", "我们看个示例", "这有个图片说明"]

        for line in lines:
            if not line.strip():
                continue

            # 处理元数据
            if ':' in line and slide_count == 0:
                optimized_lines.append(line)
                continue

            # 检测幻灯片分隔符
            if line.startswith('@@@Slide_'):
                slide_count += 1
                if slide_count > 1:
                    optimized_lines.append("\n那我们就先讲到这里，接下来看看新的内容。")
                optimized_lines.append(f"\n{random.choice(slide_openers)}第 {slide_count} 张幻灯片：")
                in_list = False
                continue

            # 清理水印和乱码
            cleaned_line = re.sub(r'stablediffusionweb\.com|WHATEV-VEER\.', '', line).strip()
            if not cleaned_line:
                continue

            # 处理图片文本
            if 'Image' in cleaned_line:
                if cleaned_line.endswith('Text:'):
                    optimized_lines.append(f"{random.choice(image_openers)}，具体内容可以参考幻灯片。")
                else:
                    content = cleaned_line.split('Text:')[-1].strip()
                    if content:
                        optimized_lines.append(f"{random.choice(image_openers)}，上面写着：{content}。")
                continue

            # 检测标题（非列表、非公式）
            if not (cleaned_line.startswith(('▪', '-', '•')) or '=' in cleaned_line or '∑' in cleaned_line or 'σ' in cleaned_line):
                if in_list:
                    optimized_lines.append("这些要点就先讲到这里。")
                    in_list = False
                prev_title = cleaned_line
                optimized_lines.append(f"{random.choice(title_openers)} {cleaned_line}。")
                continue

            # 检测列表项
            if cleaned_line.startswith(('▪', '-', '•')):
                if not in_list:
                    optimized_lines.append(f"{random.choice(list_openers)}：")
                    in_list = True
                    connector = ""
                else:
                    connector = random.choice(list_connectors)
                content = cleaned_line[1:].strip()
                optimized_lines.append(f"{connector} {content}，挺关键的吧？")
                continue

            # 检测公式
            if '=' in cleaned_line or '∑' in cleaned_line or 'σ' in cleaned_line:
                if in_list:
                    optimized_lines.append("这些要点就先讲到这里。")
                    in_list = False
                formula_desc = "它描述了模型的计算过程" if "∑" in cleaned_line else "它让模型更有效"
                optimized_lines.append(f"{random.choice(formula_openers)}：{cleaned_line}，{formula_desc}。")
                continue

        # 添加结束语
        if in_list:
            optimized_lines.append("这些要点就先讲到这里。")
        optimized_lines.append("\n好了，今天的内容就到这儿，希望大家收获满满，下次再聊！")

        optimized_text = "\n".join(optimized_lines)
        logger.info("Text optimization completed successfully.")
        return optimized_text

    except Exception as e:
        logger.error(f"Error optimizing text: {e}")
        return text

# 测试代码
if __name__ == "__main__":
    sample_text = """Title: N/A
    @@@Slide_1@@@
    深度学习基础
    @@@Slide_2@@@
    课程内容概览
    ▪ 神经网络基本原理
    Slide 5, Image 1 Text:
    stablediffusionweb.com
    """
    result = optimize_text_with_ai(sample_text)
    print(result)