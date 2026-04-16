# PPT 视频生成器项目

> 基于 Python 的智能演示文稿和视频生成工具

## 📋 目录

- [快速开始](#快速开始)
- [项目概述](#项目概述)
- [安装依赖](#安装依赖)
- [使用指南](#使用指南)
- [核心模块](#核心模块)
- [开发规范](#开发规范)
- [配置说明](#配置说明)
- [常见问题](#常见问题)
- [扩展开发](#扩展开发)

---

## 快速开始

```bash
# 安装依赖
pip install -r requirements.txt
sudo apt-get install poppler-utils ffmpeg

# 从文本生成 PPT
python3 utils/ppt_generator.py "项目介绍：这是一个创新的AI项目，致力于解决..."

# 从文本生成视频
python3 utils/video_generator.py "项目介绍：这是一个创新的AI项目，致力于解决..." --output video.mp4
```

---

## 项目概述

### 核心功能

| 功能 | 描述 | 文件 |
|------|------|------|
| **PPT 生成** | 基于大模型分析文本，生成专业 PPT | `utils/ppt_generator.py` |
| **视频生成** | 从文本/PPT 生成带语音讲解的视频 | `utils/video_generator.py` |

### 技术栈

- **语言**: Python 3.10+
- **PPT 生成**: python-pptx
- **大模型**: Qwen（`llm_services.qwen_engine`）
- **语音合成**: edge-tts
- **PDF 处理**: reportlab + poppler-utils（可选）
- **图像处理**: Pillow（备选方案）
- **视频合成**: FFmpeg

### 风格支持

| 风格 | 设计参考 |
|------|---------|
| `business` | McKinsey/BCG 咨询报告 |
| `minimal` | Apple Keynote |
| `tech` | NVIDIA GTC/Apple WWDC |
| `creative` | Behance/Dribbble 热门设计 |

---

## 安装依赖

### Python 包

```bash
pip install -r requirements.txt
```

**依赖列表**:
- `python-pptx>=0.6.21` - PPT 文件生成
- `edge-tts>=6.1.0` - 语音合成
- `Pillow>=10.0.0` - 图像处理
- `reportlab>=4.0.0` - PDF 生成（可选但推荐）

### 系统依赖（Linux）

```bash
sudo apt-get install poppler-utils ffmpeg
```

| 依赖 | 用途 |
|------|------|
| `poppler-utils` | PDF 转图片（PDF 中转方案） |
| `ffmpeg` | 视频合成 |

### 大模型依赖

确保 `llm_services/qwen_engine.py` 模块可用：

```python
from llm_services.qwen_engine import chat_with_llm
```

---

## 使用指南

### PPT 生成器

```bash
# 基本用法
python3 utils/ppt_generator.py "文本内容" --style business --output output.pptx

# 从文件读取
python3 utils/ppt_generator.py --file input.txt --style minimal

# 启用推理模式（更深入分析）
python3 utils/ppt_generator.py "文本内容" --thinking

# 可用风格
--style {business,minimal,tech,creative}
```

### 视频生成器

```bash
# 从文本生成视频
python3 utils/video_generator.py "文本内容" --style business --voice female --output video.mp4

# 从文件生成视频
python3 utils/video_generator.py --file input.txt --style minimal --voice male

# 从 PPT 文件生成视频
python3 utils/video_generator.py --pptx presentation.pptx --voice female --output video.mp4

# 强制使用 PDF 中转方案
python3 utils/video_generator.py "文本内容" --use-pdf

# 强制使用 Pillow 方案
python3 utils/video_generator.py "文本内容" --no-pdf
```

### 语音选项

| 语音代码 | 说明 |
|---------|------|
| `female` | 女声 - 晓晓 |
| `male` | 男声 - 云希 |
| `female_news` | 女声新闻 - 晓伊 |
| `male_news` | 男声新闻 - 云健 |
| `female_gentle` | 女声温柔 - 晓辰 |
| `male_cheerful` | 男声活泼 - 云扬 |

---

## 核心模块

### 项目结构

```
/data/ppt_video_converter/
├── utils/                    # 核心工具模块
│   ├── __init__.py
│   ├── ppt_generator.py      # PPT 生成器
│   └── video_generator.py    # 视频生成器
├── llm_services/             # 大模型服务
│   └── qwen_engine.py        # Qwen 接口
├── .iflow/                   # iFlow 配置
│   ├── commands/             # 自定义命令
│   └── hooks/                # Git 钩子
├── requirements.txt          # Python 依赖
└── AGENTS.md                 # 项目文档
```

### PPT 生成器 (`utils/ppt_generator.py`)

**导出函数**:
```python
from utils.ppt_generator import generate_ppt, create_presentation, analyze_text_with_llm

# 从文本生成 PPT
result = generate_ppt(text="...", style='business', output_path='output.pptx')

# 使用大模型分析文本
outline = analyze_text_with_llm(text="...")

# 根据大纲创建 PPT
prs = create_presentation(outline, style='business', output_path='output.pptx')
```

**关键函数**:
- `get_style_config(style)` - 获取样式配置
- `create_title_slide()` - 创建封面页
- `create_toc_slide()` - 创建目录页
- `create_content_slide()` - 创建内容页
- `create_conclusion_slide()` - 创建结尾页
- `set_background()` - 设置背景（渐变/纯色）
- `add_decorations()` - 添加装饰元素

### 视频生成器 (`utils/video_generator.py`)

**导出函数**:
```python
from utils.video_generator import generate_video_from_text, generate_video_from_pptx

# 从文本生成视频
result = generate_video_from_text(text="...", style='business', voice='female', output_path='video.mp4')

# 从 PPT 文件生成视频
result = generate_video_from_pptx(pptx_path='presentation.pptx', voice='female', output_path='video.mp4')
```

**关键函数**:
- `generate_pdf_from_slides()` - 生成 PDF（PDF 中转方案）
- `convert_pdf_to_images()` - PDF 转图片序列
- `create_slide_image()` - 创建内容页图片（Pillow 方案）
- `create_title_slide_image()` - 创建封面页图片
- `create_conclusion_slide_image()` - 创建结尾页图片
- `generate_audio()` - 生成语音文件
- `get_audio_duration()` - 获取音频时长

---

## 开发规范

### 代码风格

- **编码**: UTF-8
- **规范**: PEP 8
- **注释**: 中文
- **命名**: 英文（函数、变量）

### 编程规范

本项目遵循 [.iflow/commands/programmer_skill.md](.iflow/commands/programmer_skill.md) 定义的规范。

#### 核心原则

**修改完即测试（最高优先级）**
```bash
# 每次修改后执行
python3 -m py_compile utils/ppt_generator.py
python3 -m py_compile utils/video_generator.py
```

#### 开发流程

1. 需求分析 → 2. 整体设计 → 3. 公共代码提取 → 4. 零重复代码 → 5. 修改完即测试

#### 代码质量要求

- 单一职责原则
- DRY 原则（不重复自己）
- 可读性优先
- 与现有代码风格一致
- 测试友好
- 完善的错误处理
- 避免硬编码路径
- 持续验证

#### 测试检查清单

每次修改后必须确认：
- [ ] 语法检查通过
- [ ] 新增功能测试通过
- [ ] 现有功能测试通过
- [ ] 异常处理正确
- [ ] 代码风格符合规范

**只有所有测试通过后，才能继续下一步。**

### 字体配置

项目自动加载中文字体，优先顺序：
1. `/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc` (文泉驿正黑)
2. `/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc` (Noto Sans CJK)
3. `/System/Library/Fonts/PingFang.ttc` (macOS)
4. `C:/Windows/Fonts/msyh.ttc` (Windows)

### 大模型提示词设计

**核心原则**:
1. 保留关键实体名称（人名、公司名、产品名、项目名、地点等）
2. 保留关键数据（数字、百分比、金额、日期、指标等）
3. 保持信息完整

**格式要求**:
- 主标题：5-15 字
- 副标题：10-20 字
- 3-8 个章节
- 每章节 2-6 个要点
- 每要点控制在 100 字以内

---

## 配置说明

### 视频配置 (`VIDEO_CONFIG`)

```python
VIDEO_CONFIG = {
    'width': 1920,           # 视频宽度
    'height': 1080,          # 视频高度
    'fps': 24,               # 帧率
    'default_duration': 5,   # 默认每页时长（秒）
    'use_pdf_method': True,  # 优先使用 PDF 中转方案
}
```

### 样式配置结构

```python
{
    'name': '风格名称',
    'description': '风格描述',
    'gradient': {
        'type': 'linear' or 'solid',
        'angle': 135,
        'stops': [...]
    },
    'title_color': RGBColor(...),
    'subtitle_color': RGBColor(...),
    'text_color': RGBColor(...),
    'accent_color': RGBColor(...),
    'highlight_color': RGBColor(...),
    'decorations': {...},
    'fonts': {
        'title': {'name': '...', 'size': 44, 'bold': True},
        'subtitle': {...},
        'heading': {...},
        'body': {...},
        'caption': {...},
    }
}
```

---

## 常见问题

### Q1: PPT 生成失败 - `RuntimeError: Qwen大模型接口不可用`

**解决方案**: 确保 `llm_services/qwen_engine.py` 存在且可导入。

### Q2: PDF 中转方案不可用 - `RuntimeError: poppler-utils未安装`

**解决方案**:
```bash
sudo apt-get install poppler-utils
pip install reportlab
```

或使用 `--no-pdf` 参数强制使用 Pillow 方案。

### Q3: FFmpeg 不可用 - `RuntimeError: ffmpeg未安装`

**解决方案**:
```bash
sudo apt-get install ffmpeg
```

### Q4: 中文字体显示为方框

**解决方案**: 确保系统中安装了中文字体，项目会自动检测常用字体路径。

---

## 扩展开发

### 添加新风格

在 `get_style_config()` 函数中添加：

```python
styles['custom'] = {
    'name': '自定义风格',
    'description': '风格描述',
    'gradient': {...},
    'title_color': RGBColor(...),
    # ... 其他配置
}
```

### 添加新语音

在 `TTS_VOICES` 字典中添加：

```python
TTS_VOICES['custom_voice'] = 'zh-CN-CustomVoiceNeural'
```

### 自定义装饰元素

在 `add_decorations()` 函数中添加新的装饰元素处理逻辑。

---

## iFlow 集成

项目配置了自定义 iFlow 命令（skills）：

| 命令 | 说明 |
|------|------|
| `ppt_generator_skill.md` | PPT 生成器技能 |
| `video_generator_skill.md` | 视频生成器技能 |
| `programmer_skill.md` | 程序员技能 |

这些命令可以通过 iFlow CLI 的 `/` 前缀调用。

---

## 版本信息

- **Python 版本**: 3.10+
- **FFmpeg**: 已安装 (`/usr/bin/ffmpeg`)
- **创建日期**: 2026-02-21