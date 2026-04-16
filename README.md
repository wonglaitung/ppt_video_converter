# PPT 视频生成器

基于 Python 的智能演示文稿和视频生成工具，能够从文本内容自动生成专业的 PPT 演示文稿，并进一步转换为带有语音讲解的视频文件。

## 功能特性

- **PPT 生成**：基于大模型（Qwen）分析文本，生成结构化 PPT
- **视频生成**：从文本或 PPT 生成带语音讲解的视频
- **PDF 转 Word**：将 PDF 每页转为图片插入 Word 文档
- **多种风格**：商务、简约、科技、创意四种专业风格
- **语音合成**：支持多种中文语音类型
- **样式保留**：使用 unoconv 保留 PPT 完整样式

## 快速开始

### 安装依赖

```bash
# Python 包
pip install -r requirements.txt

# 系统依赖（Linux）
sudo apt-get install poppler-utils ffmpeg unoconv libreoffice
```

### 基本使用

```bash
# 从文本生成 PPT
python3 utils/ppt_generator.py "文本内容" --style business --output output.pptx

# 从文件生成 PPT
python3 utils/ppt_generator.py --file input.txt --style minimal

# 从文本生成视频
python3 utils/video_generator.py "文本内容" --style business --voice female --output video.mp4

# 从 PPT 文件生成视频（保留完整样式）
python3 utils/video_generator.py --pptx presentation.pptx --voice female --output video.mp4

# PDF 转 Word（每页作为图片插入）
python3 utils/pdf_to_word.py input.pdf

# PDF 转 Word（指定参数）
python3 utils/pdf_to_word.py input.pdf --output result.docx --dpi 120 --quality 80
```

## 风格支持

| 风格 | 设计参考 |
|------|---------|
| `business` | McKinsey/BCG 咨询报告 |
| `minimal` | Apple Keynote |
| `tech` | NVIDIA GTC/Apple WWDC |
| `creative` | Behance/Dribbble 热门设计 |

## 语音选项

| 语音代码 | 说明 |
|---------|------|
| `female` | 女声 - 晓晓 |
| `male` | 男声 - 云希 |
| `female_news` | 女声新闻 - 晓伊 |
| `male_news` | 男声新闻 - 云健 |
| `female_gentle` | 女声温柔 - 晓辰 |
| `male_cheerful` | 男声活泼 - 云扬 |

## PDF 转 Word 参数

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--dpi` | 图片 DPI，数值越大越清晰 | 100 |
| `--quality` | JPEG 质量 (1-100) | 70 |
| `--output` | 输出文件路径 | 自动生成 |

**示例**：

```bash
# 默认参数（压缩率高，适合阅读）
python3 utils/pdf_to_word.py input.pdf

# 更高清晰度（文件稍大）
python3 utils/pdf_to_word.py input.pdf --dpi 120 --quality 80

# 指定输出路径
python3 utils/pdf_to_word.py input.pdf --output result.docx
```

## 项目结构

```
ppt_video_converter/
├── utils/                    # 核心工具模块
│   ├── __init__.py
│   ├── ppt_generator.py      # PPT 生成器
│   ├── video_generator.py    # 视频生成器
│   └── pdf_to_word.py        # PDF 转 Word 工具
├── llm_services/             # 大模型服务
│   └── qwen_engine.py        # Qwen 接口
├── data/                     # 数据目录
│   └── raw/                  # 原始输入文件
├── output/                   # 输出文件目录
│   ├── presentations/        # PPT 文件
│   ├── videos/               # 视频文件
│   └── word/                 # Word 文件
├── requirements.txt          # Python 依赖
├── set_key.sh                # API 密钥配置
└── README.md                 # 项目文档
```

## 配置说明

### API 密钥配置

编辑 `set_key.sh` 文件设置 Qwen API 密钥：

```bash
export QWEN_API_KEY=your_api_key_here
```

使用前运行：
```bash
source set_key.sh
```

### 输出目录

- PPT 文件：`output/presentations/`
- 视频文件：`output/videos/`
- Word 文件：`output/word/`

## 技术栈

- **语言**: Python 3.10+
- **PPT 生成**: python-pptx
- **大模型**: Qwen
- **语音合成**: edge-tts
- **PDF 处理**: PyMuPDF (fitz) + reportlab + poppler-utils
- **Word 生成**: python-docx
- **图像处理**: Pillow
- **视频合成**: FFmpeg
- **PPT 转换**: unoconv + LibreOffice

## 常见问题

### Q1: PPT 生成失败 - `Qwen大模型接口不可用`

确保 `llm_services/qwen_engine.py` 存在且 API 密钥已正确配置。

### Q2: PDF 中转方案不可用

```bash
sudo apt-get install poppler-utils
pip install reportlab
```

### Q3: FFmpeg 不可用

```bash
sudo apt-get install ffmpeg
```

### Q4: 视频样式与 PPT 不一致

安装 unoconv 以保留 PPT 完整样式：

```bash
sudo apt-get install unoconv libreoffice
```

## 开发规范

详见 [AGENTS.md](AGENTS.md)。

## 许可证

MIT License