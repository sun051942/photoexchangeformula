#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
LaTeX转Word OMML格式转换器
使用latex2mathml将LaTeX转换为MathML，然后转换为Word的OMML格式
"""

import os
from pathlib import Path
from latex2mathml.converter import convert as latex2mathml
from lxml import etree
from docx.oxml import parse_xml
from latex_cleaner import extract_main_formula

def latex_to_word_math(latex_str):
    """
    将LaTeX转换为Word可以识别的数学公式格式

    Args:
        latex_str: LaTeX格式的公式字符串

    Returns:
        Word OMML格式的XML元素
    """
    try:
        # 首先清理和简化LaTeX
        print(f"原始LaTeX: {latex_str}")
        cleaned_latex = extract_main_formula(latex_str)
        print(f"清理后LaTeX: {cleaned_latex}")

        # 检查清理后的LaTeX是否为空
        if not cleaned_latex or not cleaned_latex.strip():
            print("警告: 清理后的LaTeX为空")
            return None

        # 转换为MathML，添加超时保护
        try:
            mathml = latex2mathml(cleaned_latex)
        except Exception as e:
            print(f"MathML转换失败: {e}")
            return None

        if not mathml:
            print("警告: MathML转换结果为空")
            return None

        # 使用XSLT将MathML转换为OMML
        # 这是微软提供的标准转换
        try:
            omml_xml = mathml_to_omml_simple(mathml)
            return omml_xml
        except Exception as e:
            print(f"OMML转换失败: {e}")
            return None

    except Exception as e:
        print(f"LaTeX转换失败: {e}")
        import traceback
        traceback.print_exc()
        return None


def mathml_to_omml_simple(mathml_str):
    """Convert MathML with Office's MML2OMML stylesheet, preserving structure."""
    candidates = []
    for env_name in ('ProgramFiles', 'ProgramFiles(x86)'):
        base = os.environ.get(env_name)
        if base:
            for version in ('Office16', 'Office15', 'Office14'):
                candidates.extend((
                    Path(base) / 'Microsoft Office' / 'root' / version / 'MML2OMML.XSL',
                    Path(base) / 'Microsoft Office' / version / 'MML2OMML.XSL',
                ))
    stylesheet = next((path for path in candidates if path.is_file()), None)
    if stylesheet is None:
        raise FileNotFoundError('未找到 Microsoft Office 的 MML2OMML.XSL')
    transform = etree.XSLT(etree.parse(str(stylesheet)))
    mathml = etree.fromstring(mathml_str.encode('utf-8'))
    result = transform(mathml)
    if result.getroot() is None:
        raise ValueError('MathML 转 OMML 得到空结果')
    return etree.tostring(result.getroot(), encoding='unicode')


def insert_latex_to_word(doc, latex_str):
    """
    在Word文档中插入LaTeX公式

    Args:
        doc: python-docx Document对象
        latex_str: LaTeX格式的公式
    """
    try:
        # 转换LaTeX为OMML
        omml_xml = latex_to_word_math(latex_str)

        if omml_xml:
            try:
                # 创建段落
                paragraph = doc.add_paragraph()

                # 解析OMML XML并插入
                omml_element = parse_xml(omml_xml)
                paragraph._element.append(omml_element)

                return True
            except Exception as e:
                print(f"插入OMML失败: {e}，降级为纯文本")
                # 如果XML解析失败，插入纯文本
                doc.add_paragraph(latex_str)
                return False
        else:
            # 如果转换失败，插入纯文本
            print("OMML转换失败，插入纯文本LaTeX")
            doc.add_paragraph(latex_str)
            return False

    except Exception as e:
        print(f"插入公式失败: {e}")
        import traceback
        traceback.print_exc()
        # 降级方案：插入纯文本
        try:
            doc.add_paragraph(latex_str)
        except:
            pass
        return False
