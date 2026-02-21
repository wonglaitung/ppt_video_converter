#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
工具模块
"""

from .ppt_generator import generate_ppt, create_presentation, analyze_text_with_llm

__all__ = [
    'generate_ppt',
    'create_presentation',
    'analyze_text_with_llm'
]
