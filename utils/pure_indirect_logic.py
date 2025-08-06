# -*- coding: utf-8 -*-
"""
Pure INDIRECT Logic - 只提取你程式碼中的核心INDIRECT處理邏輯
不包含GUI，只有純邏輯
"""

import re
import os
import openpyxl
from urllib.parse import unquote

def resolve_indirect_pure(indirect_content, workbook_path, sheet_name, current_cell=None):
    """
    純INDIRECT解析邏輯 - 提取自你的unified_indirect_resolver
    
    Args:
        indirect_content: INDIRECT函數內容 (例如: D32&"!"&"A8")
        workbook_path: Excel文件路徑
        sheet_name: 工作表名稱
        current_cell: 當前儲存格 (例如: B32)
        
    Returns:
        str: 解析後的引用 (例如: 工作表2!A8)
    """
    try:
        print(f"Pure INDIRECT logic starting...")
        print(f"Content: {indirect_content}")
        print(f"Current cell: {current_cell}")
        
        # 載入工作簿
        workbook = openpyxl.load_workbook(workbook_path, data_only=False)
        worksheet = workbook[sheet_name]
        
        # 獲取外部連結映射
        external_links_map = get_external_links_map(workbook, workbook_path)
        
        # 修復外部引用
        fixed_content = fix_external_references(indirect_content, external_links_map)
        print(f"After external fix: {fixed_content}")
        
        # 解析字串連接
        if '&' in fixed_content:
            result = resolve_concatenation(fixed_content, worksheet, current_cell)
            print(f"Concatenation result: {result}")
            return result
        else:
            # 簡單引用，移除引號
            if fixed_content.startswith('"') and fixed_content.endswith('"'):
                fixed_content = fixed_content[1:-1]
            print(f"Simple reference result: {fixed_content}")
            return fixed_content
            
    except Exception as e:
        print(f"Error in pure INDIRECT logic: {e}")
        return None

def get_external_links_map(workbook, workbook_path):
    """獲取外部連結映射 - 提取自你的邏輯"""
    external_links_map = {}
    
    try:
        if hasattr(workbook, '_external_links'):
            external_links = workbook._external_links
            if external_links:
                for i, link in enumerate(external_links, 1):
                    if hasattr(link, 'file_link') and link.file_link:
                        file_path = link.file_link.Target
                        if file_path:
                            decoded_path = unquote(file_path)
                            if decoded_path.startswith('file:///'):
                                decoded_path = decoded_path[8:]
                            elif decoded_path.startswith('file://'):
                                decoded_path = decoded_path[7:]
                            
                            external_links_map[str(i)] = decoded_path
        
        # 如果沒有找到，推斷常見的外部連結
        if not external_links_map:
            base_dir = os.path.dirname(workbook_path)
            common_files = [
                "Link1.xlsx", "Link2.xlsx", "Link3.xlsx",
                "File1.xlsx", "File2.xlsx", "File3.xlsx",
                "Data.xlsx", "GDP.xlsx", "Test.xlsx"
            ]
            
            index = 1
            for filename in common_files:
                full_path = os.path.join(base_dir, filename)
                if os.path.exists(full_path):
                    external_links_map[str(index)] = full_path
                    index += 1
                    
    except Exception as e:
        print(f"Error getting external links: {e}")
    
    return external_links_map

def fix_external_references(content, external_links_map):
    """修復外部引用 - 提取自你的邏輯"""
    try:
        def replace_ref(match):
            ref_num = match.group(1)
            if ref_num in external_links_map:
                full_path = external_links_map[ref_num]
                decoded_path = unquote(full_path) if isinstance(full_path, str) else full_path
                if decoded_path.startswith('file:///'):
                    decoded_path = decoded_path[8:]
                
                filename = os.path.basename(decoded_path)
                directory = os.path.dirname(decoded_path)
                return f"'[{directory}\\{filename}]'"
            return f"[Unknown_{ref_num}]"
        
        pattern = r'\[(\d+)\]'
        return re.sub(pattern, replace_ref, content)
    except:
        return content

def resolve_concatenation(content, worksheet, current_cell=None):
    """解析字串連接 - 提取自你的邏輯"""
    try:
        # 按 & 分割（智能處理引號內的&）
        parts = smart_split_by_ampersand(content)
        print(f"Split parts: {parts}")
        
        result_parts = []
        for part in parts:
            part = part.strip()
            print(f"Processing part: {part}")
            
            # 字串常數
            if (part.startswith('"') and part.endswith('"')) or \
               (part.startswith("'") and part.endswith("'")):
                value = part[1:-1]
                result_parts.append(value)
                print(f"  String constant: {value}")
            
            # 儲存格引用
            elif re.match(r'^\$?[A-Z]+\$?\d+$', part):
                try:
                    cell_value = worksheet[part].value
                    result_parts.append(str(cell_value) if cell_value is not None else "")
                    print(f"  Cell {part}: {cell_value}")
                except:
                    result_parts.append("")
                    print(f"  Cell {part}: Error reading")
            
            # ROW()函數
            elif 'ROW()' in part.upper() and current_cell:
                try:
                    row_num = int(re.search(r'\d+', current_cell).group())
                    if '+' in part:
                        match = re.search(r'ROW\(\)\s*\+\s*(\d+)', part, re.IGNORECASE)
                        if match:
                            add_num = int(match.group(1))
                            result_parts.append(str(row_num + add_num))
                        else:
                            result_parts.append(str(row_num))
                    else:
                        result_parts.append(str(row_num))
                    print(f"  ROW(): {result_parts[-1]}")
                except:
                    result_parts.append("ROW()")
                    print(f"  ROW(): Error")
            
            # COLUMN()函數
            elif 'COLUMN()' in part.upper() and current_cell:
                try:
                    col_letters = re.search(r'[A-Z]+', current_cell).group()
                    col_num = 0
                    for char in col_letters:
                        col_num = col_num * 26 + (ord(char) - ord('A') + 1)
                    result_parts.append(str(col_num))
                    print(f"  COLUMN(): {col_num}")
                except:
                    result_parts.append("COLUMN()")
                    print(f"  COLUMN(): Error")
            
            else:
                # 其他，保持原樣
                result_parts.append(part)
                print(f"  Other: {part}")
        
        final_result = ''.join(result_parts)
        print(f"Final concatenation result: {final_result}")
        return final_result
        
    except Exception as e:
        print(f"Error in concatenation: {e}")
        return content

def smart_split_by_ampersand(content):
    """按 & 分割，但不會分割引號內的 & - 提取自你的邏輯"""
    try:
        parts = []
        current_part = ""
        in_quotes = False
        quote_char = None
        
        i = 0
        while i < len(content):
            char = content[i]
            
            # 處理引號
            if char in ['"', "'"] and not in_quotes:
                in_quotes = True
                quote_char = char
                current_part += char
            elif char == quote_char and in_quotes:
                in_quotes = False
                quote_char = None
                current_part += char
            elif char == '&' and not in_quotes:
                # 分割點
                if current_part.strip():
                    parts.append(current_part.strip())
                current_part = ""
            else:
                current_part += char
            
            i += 1
        
        # 加最後一部分
        if current_part.strip():
            parts.append(current_part.strip())
        
        return parts
    except Exception as e:
        print(f"Error in smart split: {e}")
        return [content]

def process_formula_with_pure_indirect(formula, workbook_path, sheet_name, current_cell=None):
    """
    使用純邏輯處理包含INDIRECT的公式
    
    Args:
        formula: 公式字串 (例如: =INDIRECT(D32&"!"&"A8"))
        workbook_path: Excel文件路徑
        sheet_name: 工作表名稱
        current_cell: 當前儲存格地址
        
    Returns:
        dict: {
            'has_indirect': bool,
            'original_formula': str,
            'resolved_formula': str,
            'success': bool,
            'error': str or None
        }
    """
    try:
        # 檢查是否包含INDIRECT
        if not formula or 'INDIRECT' not in formula.upper():
            return {
                'has_indirect': False,
                'original_formula': formula,
                'resolved_formula': formula,
                'success': True,
                'error': None
            }
        
        print(f"Processing formula with pure INDIRECT logic: {formula}")
        
        # 找到INDIRECT函數內容
        indirect_match = re.search(r'INDIRECT\(([^)]+)\)', formula, re.IGNORECASE)
        if not indirect_match:
            return {
                'has_indirect': True,
                'original_formula': formula,
                'resolved_formula': formula,
                'success': False,
                'error': 'Could not extract INDIRECT content'
            }
        
        indirect_content = indirect_match.group(1)
        print(f"INDIRECT content: {indirect_content}")
        
        # 使用純邏輯解析
        resolved_ref = resolve_indirect_pure(indirect_content, workbook_path, sheet_name, current_cell)
        
        if resolved_ref:
            # 替換INDIRECT函數為解析後的引用
            resolved_formula = formula.replace(indirect_match.group(0), resolved_ref)
            print(f"Resolved formula: {resolved_formula}")
            
            return {
                'has_indirect': True,
                'original_formula': formula,
                'resolved_formula': resolved_formula,
                'success': True,
                'error': None
            }
        else:
            return {
                'has_indirect': True,
                'original_formula': formula,
                'resolved_formula': formula,
                'success': False,
                'error': 'INDIRECT resolution failed'
            }
        
    except Exception as e:
        print(f"Error processing formula: {e}")
        return {
            'has_indirect': True,
            'original_formula': formula,
            'resolved_formula': formula,
            'success': False,
            'error': str(e)
        }


# 測試函數
if __name__ == "__main__":
    # 測試用例
    test_formula = '=INDIRECT(D32&"!"&"A8")'
    test_workbook = r'C:\Users\user\Excel_tools_develop\Excel_tools_develop_v70\File5_v4.xlsx'
    test_sheet = "工作表1"
    test_cell = "B32"
    
    print("=== 測試純INDIRECT邏輯 ===")
    
    try:
        result = process_formula_with_pure_indirect(test_formula, test_workbook, test_sheet, test_cell)
        
        print(f"測試結果:")
        print(f"  Has INDIRECT: {result['has_indirect']}")
        print(f"  Success: {result['success']}")
        print(f"  Original: {result['original_formula']}")
        print(f"  Resolved: {result['resolved_formula']}")
        print(f"  Error: {result['error']}")
        
        if result['success'] and '!A8' in result['resolved_formula']:
            print("🎉 純INDIRECT邏輯工作正常！")
        else:
            print("❌ 純INDIRECT邏輯需要調整")
            
    except Exception as e:
        print(f"測試失敗: {e}")
    
    input("按Enter退出...")