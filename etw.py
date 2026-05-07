#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ETW (Excel To Word) - Excel转Word命令行工具
版本: 1.1
"""

import argparse
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl
from docx import Document
from docx.table import Table
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


# 预定义行名排序（Word表格第一列）
PREDEFINED_ROW_NAMES = [
    "试验标识", "用例编号", "用例标题", "问题等级", "问题标题", "问题描述",
    "原因分析", "整改措施", "整改方", "回归步骤", "回归人员", "问题单编号",
    "问题单类型", "整改方联系人", "回归单编号", "回归单类型"
]

VERSION = "1.1"


def parse_args():
    """解析命令行参数"""
    parser = argparse.ArgumentParser(
        description="ETW - Excel转Word命令行工具",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument("-v", "--version", action="version", version=f"ETW v{VERSION}")
    parser.add_argument("-i", required=True, help="输入Excel文件路径")
    parser.add_argument("-o", required=True, help="输出Word文件路径")
    parser.add_argument("-F", "--force", action="store_true", help="启用覆盖模式")
    
    # CK参数：根字段匹配（1-4个）
    for i in range(1, 5):
        parser.add_argument(f"--CK{i}", dest=f"CK{i}", help=f"根字段{i}")
    
    # HM参数：覆盖字段（1-20个）
    for i in range(1, 21):
        parser.add_argument(f"--HM{i}", dest=f"HM{i}", help=f"覆盖字段{i}")
    
    return parser.parse_args()


def read_excel(file_path: str) -> Tuple[List[str], List[Dict[str, str]]]:
    """读取Excel文件，返回：(列名列表, 数据行列表)"""
    try:
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        
        headers = []
        for cell in ws[1]:
            headers.append(str(cell.value).strip() if cell.value else "")
        
        data_rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            row_dict = {}
            for i, header in enumerate(headers):
                if header:
                    row_dict[header] = str(row[i]).strip() if row[i] else ""
            if row_dict:
                data_rows.append(row_dict)
        
        wb.close()
        return headers, data_rows
    
    except FileNotFoundError:
        print(f"错误: 找不到Excel文件 '{file_path}'", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"错误: 读取Excel文件失败 - {e}", file=sys.stderr)
        sys.exit(1)


def get_word_table_row_value(table: Table, row_name: str) -> Optional[str]:
    """从Word表格中获取指定行名对应的值（第二列）"""
    for row in table.rows:
        if row.cells[0].text.strip() == row_name:
            return row.cells[1].text.strip()
    return None


def set_word_table_row_value(table: Table, row_name: str, value: str):
    """设置Word表格中指定行名的值（第二列）"""
    for row in table.rows:
        if row.cells[0].text.strip() == row_name:
            row.cells[1].text = value
            return True
    return False


def match_table_by_ck(table: Table, excel_row: Dict[str, str], ck_fields: List[Tuple[str, str]]) -> bool:
    """根据CK字段检查表格是否匹配"""
    if not ck_fields:
        return False  # 无CK时，不匹配任何表格
    
    for ck_col, ck_row_name in ck_fields:
        excel_val = excel_row.get(ck_col, "")
        word_val = get_word_table_row_value(table, ck_row_name)
        if excel_val != word_val:
            return False
    return True


def add_heading_4(doc: Document, text: str):
    """添加4级标题"""
    heading = doc.add_heading(text, level=4)
    heading.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT


def create_table_from_row(doc: Document, excel_row: Dict[str, str], excel_headers: List[str]):
    """从Excel行数据创建表格"""
    # 添加4级标题：用例编号:用例标题
    case_id = excel_row.get("用例编号", "")
    case_title = excel_row.get("用例标题", "")
    title_text = f"{case_id}:{case_title}" if case_id or case_title else "未命名"
    add_heading_4(doc, title_text)
    
    # 创建2列表格（字段名 | 值）
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    
    # 设置表头
    header_row = table.rows[0]
    header_row.cells[0].text = "字段名"
    header_row.cells[1].text = "值"
    # 加粗表头
    for cell in header_row.cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True
    
    # 按预定义顺序添加数据行
    for row_name in PREDEFINED_ROW_NAMES:
        if row_name in excel_row and excel_row[row_name]:
            row = table.add_row()
            row.cells[0].text = row_name
            row.cells[1].text = excel_row[row_name]
    
    # 添加Excel中有的其他字段（不在预定义列表中）
    for col_name in excel_headers:
        if col_name not in PREDEFINED_ROW_NAMES and excel_row.get(col_name):
            exists = False
            for row in table.rows[1:]:  # 跳过表头
                if row.cells[0].text.strip() == col_name:
                    exists = True
                    break
            if not exists:
                row = table.add_row()
                row.cells[0].text = col_name
                row.cells[1].text = excel_row[col_name]
    
    return table


def process_word(input_excel: str, output_word: str, ck_fields: List[Tuple[str, str]], hm_fields: List[str], force_mode: bool):
    """处理Excel转Word
    
    逻辑规则：
    1. 输出文件不存在 → 新增模式（创建新文档，每个Excel行一个带标题的表格）
    2. 输出文件存在 + 无CK参数 → 追加模式（在原文档末尾追加新的表格）
    3. 输出文件存在 + 有CK参数 → 覆盖模式（根据CK匹配更新已有表格的值）
    """
    # 读取Excel
    excel_headers, excel_data = read_excel(input_excel)
    
    if not excel_data:
        print("警告: Excel文件中没有数据", file=sys.stderr)
        return
    
    # 检查输出文件是否存在
    output_exists = Path(output_word).exists()
    
    if not output_exists:
        # 场景1: 输出文件不存在 → 新增模式
        print(f"新增模式: 创建新文档")
        doc = Document()
        doc.add_heading('Excel数据转换结果', level=1)
        
        for excel_idx, excel_row in enumerate(excel_data):
            create_table_from_row(doc, excel_row, excel_headers)
            doc.add_paragraph()
        
        print(f"创建表格数: {len(excel_data)}")
    
    elif ck_fields:
        # 场景2: 输出文件存在 + 有CK参数 → 覆盖模式
        print(f"覆盖模式: 加载现有文档 '{output_word}'")
        doc = Document(output_word)
        
        num_tables = len(doc.tables)
        print(f"Word文档表格数: {num_tables}, Excel数据行数: {len(excel_data)}")
        
        updated_count = 0
        for excel_idx, excel_row in enumerate(excel_data):
            # 遍历Word中所有表格，找匹配的
            matched = False
            for table_idx, table in enumerate(doc.tables):
                if match_table_by_ck(table, excel_row, ck_fields):
                    # 找到匹配的表格，更新值
                    for col_name, value in excel_row.items():
                        if hm_fields and col_name not in hm_fields:
                            continue
                        if set_word_table_row_value(table, col_name, value):
                            updated_count += 1
                    matched = True
                    print(f"  Excel第{excel_idx+1}行匹配Word表格{table_idx+1}")
                    break  # 匹配一个表格后跳出
            
            if not matched and not force_mode:
                # 未匹配且非强制模式：追加新表格
                print(f"  Excel第{excel_idx+1}行未匹配，新增表格")
                create_table_from_row(doc, excel_row, excel_headers)
                doc.add_paragraph()
        
        print(f"更新字段数: {updated_count}")
    
    else:
        # 场景3: 输出文件存在 + 无CK参数 → 追加模式
        print(f"追加模式: 加载现有文档 '{output_word}'，在末尾追加")
        doc = Document(output_word)
        
        doc.add_paragraph()  # 添加分隔
        doc.add_heading('追加数据', level=2)
        
        for excel_idx, excel_row in enumerate(excel_data):
            create_table_from_row(doc, excel_row, excel_headers)
            doc.add_paragraph()
        
        print(f"追加表格数: {len(excel_data)}")
    
    # 保存文档
    try:
        doc.save(output_word)
        print(f"成功: 已保存Word文档 '{output_word}'")
    except Exception as e:
        print(f"错误: 保存Word文档失败 - {e}", file=sys.stderr)
        sys.exit(1)


def main():
    """主函数"""
    args = parse_args()
    
    # 收集CK字段
    ck_fields = []
    for i in range(1, 5):
        value = getattr(args, f"CK{i}", None)
        if value:
            if ":" in value:
                parts = value.split(":", 1)
                ck_fields.append((parts[0], parts[1]))
            else:
                ck_fields.append((value, value))
    
    # 收集HM字段
    hm_fields = []
    for i in range(1, 21):
        value = getattr(args, f"HM{i}", None)
        if value:
            hm_fields.append(value)
    
    # 检查覆盖模式参数
    if args.force:
        if not ck_fields:
            print("错误: 覆盖模式(-F)必须配合CK参数使用", file=sys.stderr)
            sys.exit(1)
    
    # 执行转换
    process_word(args.i, args.o, ck_fields, hm_fields, args.force)


if __name__ == "__main__":
    main()