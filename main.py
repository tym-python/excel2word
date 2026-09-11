# !/usr/bin/env python3
# -*-coding:utf-8 -*-

"""
# File       : main.py
# Time       ：2026-06-04 9:38
# Author     ：t
# Func     ：
"""
import os, sys,io
import re

from openpyxl import load_workbook
from datetime import datetime
from docxtpl import DocxTemplate,InlineImage
from docx.shared import Mm,Inches
from jinja2 import Environment
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
def money_filter(value, fmt="#,##0.00"):
    """
    Jinja2 filter：把数字格式化成指定格式的字符串。
    用法：{{ 金额 | money }}                → 1,233,454.00
          {{ 金额 | money('#,##0') }}       → 1,233,454
          {{ 金额 | money('¥#,##0.00') }}   → ¥1,233,454.00
          {{ 金额 | money('0.00%') }}       → 12.34%
    """
    if value is None or value == "":
        return ""
    if isinstance(value, str):
        return value          # 字符串不动
    if isinstance(value, bool):
        return value
    try:
        return f"{value:{fmt}}"
    except Exception:
        return str(value)

class excel2word():
    def __init__(self, DATE_FORMAT='%Y年%m月%d日', input_DIR_NAME='', OUTPUT_DIR_NAME='output', Image_DIR_NAME='img'):
        '''
        :param DATE_FORMAT:
        :param OUTPUT_DIR_NAME:     输出文件夹名称
        :param Image_DIR_NAME:  插图文件夹
        '''
        self.DATE_FORMAT = DATE_FORMAT
        # 1. 获取程序所在根目录
        self.base_dir = self.get_base_dir()
        self.input_dir = os.path.join(self.base_dir, input_DIR_NAME)
        self.output_dir = os.path.join(self.base_dir, OUTPUT_DIR_NAME)
        self.img_dir = os.path.join(self.base_dir, Image_DIR_NAME)

    def get_base_dir(self):
        """兼容性路径获取"""
        if os.environ.get('PYCHARM_HOSTED') == '1':
            return os.path.dirname(os.path.abspath(__file__))
        else:
            # 直接运行脚本，返回脚本文件所在的目录
            return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # 上一级目录（根目录）

    def find_key_containing(self, data, keyword):
        """从字典的键中查找包含 keyword 的键（不区分大小写），返回第一个匹配的键，找不到返回 None"""
        if isinstance(data, dict):
            keyword_lower = keyword.lower()
            for key in data.keys():
                if keyword_lower in key.lower():
                    return key
        elif isinstance(data, list):
            keyword_lower = keyword.lower()
            for key in data:
                if keyword_lower in key.lower():
                    return key
        return None

    def format_cell_value(self, cell, doc=None, date_format=None):
        '''
        处理数据格式，文本，日期，图片
        :param cell:
        :param date_format:
        :return:
        '''
        if not date_format:
            date_format = self.DATE_FORMAT
        value = cell.value
        if value is None:
            return ''
        if isinstance(value, datetime):
            result = date_format
            result = result.replace('%Y', str(value.year))
            result = result.replace('%m', f"{value.month}")
            result = result.replace('%d', f"{value.day}")
            return result
        elif re.findall('png$|jpg$|jpeg$', str(value).lower()):
            # 判断是否为图片路径
            img_path = self.get_imagePath(self.img_dir, str(value))
            if img_path and doc:
                img = InlineImage(
                    doc,  # 关联的文档对象
                    image_descriptor=img_path,  # 图片文件路径（支持本地文件、BytesIO等）
                    width=Mm(50),  # 设置图片宽度为 50 毫米（高度自动按比例）
                    # 也可以使用英寸：width=Inches(1.5)
                )
                return img
            else:
                return str(value)
        elif isinstance(value, (int, float)):
            return value
        return value

    def get_outline(self, doc):
        """
        从 Document 对象中提取标题大纲
        返回列表，每个元素为 (level, text)
        """
        outline = []
        for para in doc.paragraphs:
            style_name = para.style.name if para.style else ''
            if style_name.startswith('Heading'):
                try:
                    level = int(style_name.split()[-1])
                except:
                    level = 1
                text = para.text.strip()
                if text:
                    outline.append((level, text))
        if outline:
            highest_level = min(level for level, _ in outline)
            first_highest = next(text for level, text in outline if level == highest_level)
            safe_title = "".join(c for c in first_highest if c not in r'\/:*?"<>|')
            return safe_title
        else:
            return ''

    def get_imagePath(self, img_dir, imgName):
        img_path = os.path.join(img_dir, imgName)
        if not os.path.exists(img_path):
            return None
        else:
            return img_path

    def main(self, excelName='data.xlsx', wordName='template.docx'):
        excel_path = os.path.join(self.input_dir, excelName)
        template_path = os.path.join(self.input_dir, wordName)

        # 检查文件是否存在
        if not os.path.exists(excel_path):
            print(f"[ERROR]：找不到Excel文件 {excel_path}")
            input("按Enter键退出...")
            return
        if not os.path.exists(template_path):
            print(f"[ERROR]：找不到Word模板 {template_path}")
            input("按Enter键退出...")
            return

        # 创建输出文件夹
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.img_dir, exist_ok=True)

        # 读取Excel（data_only=True 获取公式计算后的值）
        wb = load_workbook(excel_path, data_only=True)
        ws = wb.active

        # 获取表头（第一行）
        headers = [cell.value for cell in ws[1] if cell.value is not None]
        if not headers:
            print("[ERROR]：Excel第一行没有列标题")
            input("按Enter键退出...")
            return
        print(f"[OK] 检测到列标题：{headers}")

        # 获取数据行（从第二行开始）
        rows = list(ws.iter_rows(min_row=2, values_only=False))
        # 在 for 循环之前，只建一次 env
        jinja_env = Environment()
        jinja_env.filters['money'] = money_filter

        count = 0
        for idx, row in enumerate(rows, start=2):
            # 跳过全空行
            if all(cell.value is None or str(cell.value).strip() == '' for cell in row):
                continue

            # 加载Word模板
            doc = DocxTemplate(template_path)
            # 构建数据字典,构建图片需要doc
            data = {}
            for i, header in enumerate(headers):
                if i < len(row):
                    cell = row[i]
                    data[header] = self.format_cell_value(cell, doc)

            doc.render(data, jinja_env)

            custom_title = data.get(self.find_key_containing(data, '文件标题'), '').strip()
            if custom_title:
                safe_title = "".join(c for c in custom_title if c not in r'\/:*?"<>|')
            else:
                # 获取大纲，根据大纲输出title
                safe_title = self.get_outline(doc)

            if not safe_title:
            # 使用单位+年份构造文件名
                if '单位名称' in data.keys():
                    unit = data.get('单位名称', '').strip()
                else:
                    unit_col = self.find_key_containing(data, '单位')
                    unit = data.get(unit_col, '').strip()
                yearName = str(data.get('年份', '')).strip()
                if unit:
                    # 去除Windows文件名非法字符
                    safe_unit = "".join(c for c in unit+yearName if c not in r'\/:*?"<>|')
                    output_name = f"{safe_unit}.docx"
                else:
                    output_name = f"doc_{idx}.docx"
            else:
                output_name = f"{safe_title}.docx"

            output_path = os.path.join(self.output_dir, output_name)

            # 避免文件名重复（如果已存在则加序号）
            counter = 1
            original_path = output_path
            while os.path.exists(output_path):
                name, ext = os.path.splitext(original_path)
                output_path = f"{name}_{counter}{ext}"
                counter += 1

            doc.save(output_path)
            print(f"[OK] 已生成：{output_path}")
            count += 1

        print(f"\n[DONE] 全部完成！共生成 {count} 个文档。")
        print(f"输出目录：{self.output_dir}")
        wb.close()
        print(f'----【{datetime.now()}】运行结束。----')

if __name__ == '__main__':
    print(f'----【{datetime.now()}】开始运行。----')
    c = excel2word(input_DIR_NAME=r'D:\AppData\xwechat_files\wxid_vcnil19ggaqm22_0bf8\msg\file\2026-09')
    # c.main(excelName='数据源.xlsx',wordName='模板.docx')

    # c = excel2word(input_DIR_NAME='')
    c.main()
