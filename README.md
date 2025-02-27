# 课件语音助手 (CoursewareSpeechMaker)

## 简介
课件语音助手是一个基于 Python 的智能工具，可以快速将 PPT 课件转换为有声课件。通过先进的文本提取和语音合成技术，轻松实现课件语音化。

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Build Status](https://img.shields.io/badge/build-passing-brightgreen.svg)

## 功能特点
- 📄 支持 PPT 文本自动提取
- 🔊 基于 PaddleSpeech 的智能语音合成
- 🎙️ 多种语音模式选择
- 📦 支持多种音频格式导出
- 🚀 简单易用的 API 接口

## 技术栈
- Python 3.8+
- python-pptx
- PaddleSpeech
- pydub

## 安装依赖
```bash
# 克隆项目
git clone https://github.com/yourusername/courseware-speech-maker.git
cd courseware-speech-maker

# 创建虚拟环境
python -m venv venv
source venv/bin/activate  # Unix/macOS
# venv\Scripts\activate   # Windows

# 安装依赖
pip install -r requirements.txt
```

## 快速开始

### 基本使用
```python
from courseware_speech_maker import CoursewareSpeechMaker

# 创建语音课件
ppt_path = 'your_courseware.pptx'
maker = CoursewareSpeechMaker(ppt_path)

# 生成语音课件
audio_file = maker.generate_audiobook()
print(f"语音课件已生成：{audio_file}")
```

### 高级配置
```python
# 自定义语音参数
maker = CoursewareSpeechMaker(
    ppt_path, 
    output_dir='output',
    voice_type='female',
    speech_speed=1.0
)

# 导出多种格式
maker.export_multiple_formats(['wav', 'mp3'])
```

## 配置参数
| 参数 | 类型 | 默认值 | 描述 |
|------|------|--------|------|
| ppt_path | str | 必填 | PPT 文件路径 |
| output_dir | str | 'output' | 输出目录 |
| voice_type | str | 'female' | 语音类型 |
| speech_speed | float | 1.0 | 语音播放速度 |

## 项目结构
```
courseware-speech-maker/
│
├── src/
│   ├── __init__.py
│   ├── speech_maker.py
│   └── text_extractor.py
│
├── tests/
│   └── test_speech_maker.py
│
├── examples/
│   └── demo.py
│
├── requirements.txt
├── README.md
└── setup.py
```

## 常见问题
1. **支持哪些 PPT 格式？**
   - 支持 .pptx 格式
   - 不支持早期 .ppt 格式

2. **语音生成耗时？**
   - 取决于 PPT 页数和文本长度
   - 约 1-2 秒

## 许可证
MIT License

## 贡献
欢迎提交 Pull Request 和 Issue

## 联系
- 邮箱：your_email@example.com
- 项目地址：https://github.com/yourusername/courseware-speech-maker