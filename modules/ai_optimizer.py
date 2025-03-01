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

# # 虚假AI（缺少API）
# import re
# import logging
# import random

# # 配置日志
# logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
# logger = logging.getLogger(__name__)

# def optimize_text_with_ai(text):
#     """
#     本地优化 PPT 提取的文本，使其更自然、适合有声课件讲解。
#     """
#     try:
#         lines = text.split('\n')
#         optimized_lines = []
#         in_list = False
#         slide_count = 0
#         prev_title = ""

#         # 多样化的引导语和连接词
#         slide_openers = ["现在我们来看看", "接下来聊聊", "这一节我们讲讲", "让我们进入"]
#         title_openers = ["首先是", "那么我们说说", "这里有个重点", "值得一提的是"]
#         list_openers = ["先来看看", "我们聊聊", "这里包括", "比如说"]
#         list_connectors = ["然后是", "接着是", "还有呢", "另外一点"]
#         formula_openers = ["这里有个公式", "看看这个计算", "简单解释一下"]
#         image_openers = ["这里有张图", "我们看个示例", "这有个图片说明"]

#         for line in lines:
#             if not line.strip():
#                 continue

#             # 处理元数据
#             if ':' in line and slide_count == 0:
#                 optimized_lines.append(line)
#                 continue

#             # 检测幻灯片分隔符
#             if line.startswith('@@@Slide_'):
#                 slide_count += 1
#                 if slide_count > 1:
#                     optimized_lines.append("\n那我们就先讲到这里，接下来看看新的内容。")
#                 optimized_lines.append(f"\n{random.choice(slide_openers)}第 {slide_count} 张幻灯片：")
#                 in_list = False
#                 continue

#             # 清理水印和乱码
#             cleaned_line = re.sub(r'stablediffusionweb\.com|WHATEV-VEER\.', '', line).strip()
#             if not cleaned_line:
#                 continue

#             # 处理图片文本
#             if 'Image' in cleaned_line:
#                 if cleaned_line.endswith('Text:'):
#                     optimized_lines.append(f"{random.choice(image_openers)}，具体内容可以参考幻灯片。")
#                 else:
#                     content = cleaned_line.split('Text:')[-1].strip()
#                     if content:
#                         optimized_lines.append(f"{random.choice(image_openers)}，上面写着：{content}。")
#                 continue

#             # 检测标题（非列表、非公式）
#             if not (cleaned_line.startswith(('▪', '-', '•')) or '=' in cleaned_line or '∑' in cleaned_line or 'σ' in cleaned_line):
#                 if in_list:
#                     optimized_lines.append("这些要点就先讲到这里。")
#                     in_list = False
#                 prev_title = cleaned_line
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
#                 optimized_lines.append(f"{connector} {content}，挺关键的吧？")
#                 continue

#             # 检测公式
#             if '=' in cleaned_line or '∑' in cleaned_line or 'σ' in cleaned_line:
#                 if in_list:
#                     optimized_lines.append("这些要点就先讲到这里。")
#                     in_list = False
#                 formula_desc = "它描述了模型的计算过程" if "∑" in cleaned_line else "它让模型更有效"
#                 optimized_lines.append(f"{random.choice(formula_openers)}：{cleaned_line}，{formula_desc}。")
#                 continue

#         # 添加结束语
#         if in_list:
#             optimized_lines.append("这些要点就先讲到这里。")
#         optimized_lines.append("\n好了，今天的内容就到这儿，希望大家收获满满，下次再聊！")

#         optimized_text = "\n".join(optimized_lines)
#         logger.info("Text optimization completed successfully.")
#         return optimized_text

#     except Exception as e:
#         logger.error(f"Error optimizing text: {e}")
#         return text

# # 测试代码
# if __name__ == "__main__":
#     sample_text = """Title: N/A
#     @@@Slide_1@@@
#     深度学习基础
#     @@@Slide_2@@@
#     课程内容概览
#     ▪ 神经网络基本原理
#     Slide 5, Image 1 Text:
#     stablediffusionweb.com
#     """
#     result = optimize_text_with_ai(sample_text)
#     print(result)

##########################################

import re
import logging
import random
from transformers import pipeline
import spacy

# 加载 Spacy 模型
nlp = spacy.load("zh_core_web_sm")

# 加载 Transformers 模型（文本生成）
generator = pipeline("text-generation", model="uer/gpt2-chinese-cluecorpussmall", max_length=50)

# 配置日志
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

def optimize_text_with_ai(text):
    """
    使用 Transformers 和 Spacy 优化 PPT 文本，生成自然、复杂的讲解内容。
    """
    try:
        # 预处理文本
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        optimized_lines = []
        slide_count = 0
        current_section = "metadata"
        slide_buffer = []
        prev_context = None

        # 多样化表达库
        metadata_ends = ["课程资料齐全，我们马上开讲！", "文本就位，接下来直入主题！"]
        slide_intros = ["我们先从这里入手，聊聊", "接下来带大家走进", "现在一起探讨", "那就让我们开始"]
        section_ends = ["这部分先告一段落，接下来有新亮点", "聊到这里，我们转向新内容", "先到这儿，后续更精彩"]
        final_closings = ["今天的课程到此结束，大家收获如何？下次见！", 
                          "好了，这节内容就先画个句号，咱们下次再续！"]

        for i, line in enumerate(lines):
            # 元数据处理
            if current_section == "metadata" and ':' in line:
                key, value = line.split(':', 1)
                optimized_lines.append(f"{key.strip()}: {value.strip()}")
                if "Revision" in key:
                    optimized_lines.append(f"\n{random.choice(metadata_ends)}")
                    current_section = "slides"
                continue

            # 幻灯片分隔符
            slide_match = re.match(r'(?:接下来聊聊|让我们进入|现在我们来看)第 (\d+) 张幻灯片：', line)
            if slide_match or re.match(r'@@@Slide_\d+@@@', line):
                slide_count += 1
                if slide_buffer:
                    optimized_lines.append(_process_slide(slide_buffer, prev_context, slide_count - 1))
                    slide_buffer = []
                if slide_count > 1:
                    optimized_lines.append(f"\n{random.choice(section_ends)}")
                optimized_lines.append(f"\n{random.choice(slide_intros)}第 {slide_count} 张幻灯片：")
                continue

            # 清理无用内容
            cleaned_line = re.sub(r'[▪•\-\t]|stablediffusionweb\.com|WHATEV-VEER\.|\s{2,}', ' ', line).strip()
            if not cleaned_line:
                continue

            # 添加到幻灯片缓冲区
            slide_buffer.append(cleaned_line)

        # 处理最后一个幻灯片
        if slide_buffer:
            final_text = _process_final_slide(slide_buffer, prev_context, slide_count)
            optimized_lines.append(final_text)
        optimized_lines.append(f"\n{random.choice(final_closings)}")

        optimized_text = "\n".join(optimized_lines)
        logger.info("Text optimization completed successfully.")
        return optimized_text

    except Exception as e:
        logger.error(f"Error optimizing text: {e}", exc_info=True)
        return text

def _process_slide(slide_lines, prev_context, slide_num):
    """处理单个幻灯片的文本"""
    doc = nlp(" ".join(slide_lines))
    sentences = [sent.text.strip() for sent in doc.sents]
    narrative = []
    list_items = []

    for i, sent in enumerate(sentences):
        # 处理公式
        if any(symbol in sent for symbol in ['=', '∑', 'σ']):
            purpose = "揭示模型计算的关键" if '∑' in sent else "提升模型的效率"
            narrative.append(f"这里有个关键点，我们来看公式：{sent}，它{purpose}。")
            continue
        
        # 处理图片
        if 'Image' in sent:
            content = sent.split('Text:')[-1].strip() if 'Text:' in sent else ""
            narrative.append(f"幻灯片上有个直观的展示{'，内容是：' + content if content else '，具体请看幻灯片'}。")
            continue

        # 处理列表或正文
        if len(sent.split()) < 10 and (':' in sent or sent.endswith('，')):
            list_items.append(sent.strip('：，'))
        else:
            if list_items:
                narrative.append(_format_list(list_items, prev_context))
                list_items = []
            transition = _generate_transition(prev_context, sent)
            narrative.append(f"{transition}{sent}。")
            prev_context = sent

    if list_items:
        narrative.append(_format_list(list_items, prev_context))

    return " ".join(narrative)

def _process_final_slide(slide_lines, prev_context, slide_num):
    """将最后一个幻灯片处理为完整段落"""
    doc = nlp(" ".join(slide_lines))
    sentences = [sent.text.strip() for sent in doc.sents]
    narrative = [f"\n在第 {slide_num} 张幻灯片中，我们深入剖析了知识点的细节。"]

    for i, sent in enumerate(sentences):
        if any(symbol in sent for symbol in ['=', '∑', 'σ']):
            purpose = "揭示了计算的核心逻辑" if '∑' in sent else "让模型运行更高效"
            narrative.append(f"其中一个关键点是公式：{sent}，它{purpose}，")
        elif 'Image' in sent:
            content = sent.split('Text:')[-1].strip() if 'Text:' in sent else ""
            narrative.append(f"幻灯片上通过图示{'展示了' + content if content else '清晰呈现了相关内容'}，")
        else:
            if i == 0:
                narrative.append(f"首先是{sent}，这为我们理解整体框架奠定了基础，")
            elif i == len(sentences) - 1:
                narrative.append(f"最后谈到{sent}，它不仅总结了前面的内容，也为后续学习提供了启发。")
            else:
                narrative.append(f"接着是{sent}，进一步丰富了我们的视角，")

    return " ".join(narrative).rstrip('，') + "。"

def _generate_transition(prev_context, current_line):
    """使用 Transformers 生成自然过渡语"""
    if not prev_context:
        return "我们先来看看，"
    prompt = f"{prev_context}，接下来是{current_line}"
    try:
        generated = generator(prompt, num_return_sequences=1, max_new_tokens=10)[0]['generated_text']
        transition = generated.split('，')[-2] + "，" if len(generated.split('，')) > 1 else "接着是，"
    except Exception:
        transition = random.choice(["顺着这个思路，", "再来看看，", "基于此，"])
    return transition

def _format_list(items, prev_context):
    """格式化列表为自然叙述"""
    if not items:
        return ""
    intro = "具体来说，"
    connectors = ["接着是", "然后聊到", "再看看", "另外还有"]
    sentences = [intro]
    for i, item in enumerate(items):
        connector = "" if i == 0 else random.choice(connectors)
        sentences.append(f"{connector}{item}{'，这点很关键' if i % 2 == 0 else '，也很重要'}")
    return " ".join(sentences) + "。"
