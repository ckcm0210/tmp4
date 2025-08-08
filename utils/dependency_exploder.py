# -*- coding: utf-8 -*-
"""
Dependency Exploder - 公式依賴鏈遞歸分析器
"""

import re
import os
from urllib.parse import unquote
from utils.openpyxl_resolver import read_cell_with_resolved_references

class DependencyExploder:
    """公式依賴鏈爆炸分析器"""
    
    def __init__(self, max_depth=10, range_expand_threshold=5):
        self.max_depth = max_depth
        self.range_expand_threshold = range_expand_threshold  # Range展開閾值
        self.visited_cells = set()
        self.circular_refs = []
    
    def explode_dependencies(self, workbook_path, sheet_name, cell_address, current_depth=0, root_workbook_path=None):
        """
        遞歸展開公式依賴鏈
        
        Args:
            workbook_path: Excel 檔案路徑
            sheet_name: 工作表名稱
            cell_address: 儲存格地址 (如 A1)
            current_depth: 當前遞歸深度
            
        Returns:
            dict: 依賴樹結構
        """
        # 創建唯一標識符
        cell_id = f"{workbook_path}|{sheet_name}|{cell_address}"
        
        # 檢查遞歸深度限制
        if current_depth >= self.max_depth:
            # 決定顯示格式
            current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
            if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
                filename = os.path.basename(workbook_path)
                if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                    filename = filename.rsplit('.', 1)[0]
                display_address = f"[{filename}]{sheet_name}!{cell_address}"
            else:
                display_address = f"{sheet_name}!{cell_address}"
            
            return {
                'address': display_address,
                'workbook_path': workbook_path,
                'sheet_name': sheet_name,
                'cell_address': cell_address,
                'value': 'Max depth reached',
                'formula': None,
                'type': 'limit_reached',
                'children': [],
                'depth': current_depth,
                'error': 'Maximum recursion depth reached'
            }
        
        # 檢查循環引用
        if cell_id in self.visited_cells:
            self.circular_refs.append(cell_id)
            # 決定顯示格式
            current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
            if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
                filename = os.path.basename(workbook_path)
                if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                    filename = filename.rsplit('.', 1)[0]
                display_address = f"[{filename}]{sheet_name}!{cell_address}"
            else:
                display_address = f"{sheet_name}!{cell_address}"
            
            return {
                'address': display_address,
                'workbook_path': workbook_path,
                'sheet_name': sheet_name,
                'cell_address': cell_address,
                'value': 'Circular reference',
                'formula': None,
                'type': 'circular_ref',
                'children': [],
                'depth': current_depth,
                'error': 'Circular reference detected'
            }
        
        # 標記為已訪問
        self.visited_cells.add(cell_id)
        
        try:
            # 讀取儲存格內容
            cell_info = read_cell_with_resolved_references(workbook_path, sheet_name, cell_address)
            
            if 'error' in cell_info:
                # 決定顯示格式
                current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
                if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
                    filename = os.path.basename(workbook_path)
                    if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                        filename = filename.rsplit('.', 1)[0]
                    display_address = f"[{filename}]{sheet_name}!{cell_address}"
                else:
                    display_address = f"{sheet_name}!{cell_address}"
                
                return {
                    'address': display_address,
                    'workbook_path': workbook_path,
                    'sheet_name': sheet_name,
                    'cell_address': cell_address,
                    'value': 'Error',
                    'formula': None,
                    'type': 'error',
                    'children': [],
                    'depth': current_depth,
                    'error': cell_info['error']
                }
            
            # 基本節點信息
            original_formula = cell_info.get('formula')
            # 增強的公式清理：處理雙反斜線、URL 編碼和雙引號
            fixed_formula = None
            if original_formula:
                # 步驟1: 處理雙反斜線
                fixed_formula = original_formula.replace('\\\\', '\\')
                # 步驟2: 解碼 URL 編碼字符（如 %20 -> 空格）
                from urllib.parse import unquote
                fixed_formula = unquote(fixed_formula)
                # 步驟3: 處理雙引號問題 - 將 ''path'' 改為 'path'
                import re
                # 匹配 ''...'' 模式並替換為 '...'
                fixed_formula = re.sub(r"''([^']*?)''", r"'\1'", fixed_formula)

            # === 修復：所有節點都顯示完整檔案信息 ===
            # 提取檔案名
            filename = os.path.basename(workbook_path)
            dir_path = os.path.dirname(workbook_path)
            
            # 使用 root_workbook_path 來判斷是否為外部引用
            current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
            
            if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
                # 外部引用：準備 short 和 full 兩種格式
                # Short format: [filename.xlsx]sheet!cell
                short_display_address = f"[{filename}]{sheet_name}!{cell_address}"
                # Full format: 'C:\path\[filename.xlsx]sheet'!cell
                full_display_address = f"'{dir_path}\\[{filename}]{sheet_name}'!{cell_address}"
                # 預設使用 short format
                display_address = short_display_address
            else:
                # === 修復：本地引用也顯示檔案名 ===
                # Short format: [filename.xlsx]sheet!cell (與外部引用格式一致)
                short_display_address = f"[{filename}]{sheet_name}!{cell_address}"
                # Full format: 'C:\path\[filename.xlsx]sheet'!cell
                full_display_address = f"'{dir_path}\\[{filename}]{sheet_name}'!{cell_address}"
                # 預設使用 short format
                display_address = short_display_address

            node = {
                'address': display_address,
                'short_address': short_display_address,
                'full_address': full_display_address,
                'workbook_path': workbook_path,
                'sheet_name': sheet_name,
                'cell_address': cell_address,
                'value': cell_info.get('display_value', 'N/A'),
                'calculated_value': cell_info.get('calculated_value', 'N/A'),
                'formula': fixed_formula,
                'type': cell_info.get('cell_type', 'unknown'),
                'children': [],
                'depth': current_depth,
                'error': None
            }
            
            # 如果是公式，解析依賴關係
            if cell_info.get('cell_type') == 'formula' and cell_info.get('formula'):
                references = self.parse_formula_references(cell_info['formula'], workbook_path, sheet_name)
                
                # 遞歸展開每個引用
                for ref in references:
                    try:
                        child_node = self.explode_dependencies(
                            ref['workbook_path'],
                            ref['sheet_name'],
                            ref['cell_address'],
                            current_depth + 1,
                            root_workbook_path or workbook_path
                        )
                        node['children'].append(child_node)
                    except Exception as e:
                        # 添加錯誤節點
                        # 決定顯示格式
                        current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
                        if os.path.normpath(current_workbook_path) != os.path.normpath(ref['workbook_path']):
                            filename = os.path.basename(ref['workbook_path'])
                            if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                                filename = filename.rsplit('.', 1)[0]
                            error_display_address = f"[{filename}]{ref['sheet_name']}!{ref['cell_address']}"
                        else:
                            error_display_address = f"{ref['sheet_name']}!{ref['cell_address']}"
                        
                        error_node = {
                            'address': error_display_address,
                            'workbook_path': ref['workbook_path'],
                            'sheet_name': ref['sheet_name'],
                            'cell_address': ref['cell_address'],
                            'value': 'Error',
                            'formula': None,
                            'type': 'error',
                            'children': [],
                            'depth': current_depth + 1,
                            'error': str(e)
                        }
                        node['children'].append(error_node)
            
            # 移除已訪問標記（允許在不同分支中重複訪問）
            self.visited_cells.discard(cell_id)
            
            return node
            
        except Exception as e:
            # 移除已訪問標記
            self.visited_cells.discard(cell_id)
            
            # 決定顯示格式
            current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
            if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
                filename = os.path.basename(workbook_path)
                if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                    filename = filename.rsplit('.', 1)[0]
                display_address = f"[{filename}]{sheet_name}!{cell_address}"
            else:
                display_address = f"{sheet_name}!{cell_address}"
            
            return {
                'address': display_address,
                'workbook_path': workbook_path,
                'sheet_name': sheet_name,
                'cell_address': cell_address,
                'value': 'Error',
                'formula': None,
                'type': 'error',
                'children': [],
                'depth': current_depth,
                'error': str(e)
            }
    
    def parse_formula_references(self, formula, current_workbook_path, current_sheet_name):
        """
        使用修復後的公式引用解析器（基於dependency_exploder_old.py的proven方法）
        同時支援範圍引用的智能處理
        """
        if not formula or not formula.startswith('='):
            return []

        references = []
        processed_spans = set()

        # 正規化反斜線處理
        normalized_formula = formula.replace('\\\\', '\\')

        # 正則表達式匹配絕對引用（本地和外部）
        # 修復：使用proven pattern from dependency_exploder_old.py
        abs_pattern = r"((?:''[^']*''|'[^']+'|[^'!,=+\-*/^&()<> ]+)!)\$?([A-Z]{1,3})\$?([0-9]{1,7})"
        
        for match in re.finditer(abs_pattern, normalized_formula):
            sheet_part_raw = match.group(1)  # e.g., "'C:\\path\\[file.xlsx]Sheet1'!" or "Sheet1!"
            col = match.group(2)
            row = match.group(3)
            cell_address = f"{col}{row}"
            
            # 標記此部分字符串為已處理
            processed_spans.add(match.span())

            # 檢查是否為外部引用
            if '[' in sheet_part_raw and ']' in sheet_part_raw:
                # 增強的外部路徑清理（使用proven方法）
                decoded_ref = unquote(sheet_part_raw)
                cleaned_ref = decoded_ref.strip("\' ! \"").strip()
                cleaned_ref = cleaned_ref.replace('\\\\', '\\')
                
                # 特殊處理雙引號模式: ''path''!
                if cleaned_ref.startswith("'") and cleaned_ref.endswith("'"):
                    cleaned_ref = cleaned_ref[1:-1]
                    cleaned_ref = cleaned_ref.strip()
                
                try:
                    workbook_part, sheet_name = cleaned_ref.rsplit(']', 1)
                    workbook_part += ']'
                    dir_path, file_name = workbook_part.rsplit('[', 1)
                    file_name = file_name.rstrip(']')
                    
                    workbook_path = os.path.normpath(os.path.join(dir_path, file_name))
                    
                    references.append({
                        'workbook_path': workbook_path,
                        'sheet_name': sheet_name,
                        'cell_address': cell_address,
                        'type': 'external'
                    })
                except ValueError:
                    continue
            else:
                # 本地絕對引用
                sheet_name = sheet_part_raw.strip("\'!")
                references.append({
                    'workbook_path': current_workbook_path,
                    'sheet_name': sheet_name,
                    'cell_address': cell_address,
                    'type': 'local_absolute'
                })

        # 處理相對引用（使用proven方法避免路徑組件的誤匹配）
        path_exclusions = set()
        
        # 找出所有引號區段
        quoted_sections = re.finditer(r"'[^']*'", normalized_formula)
        for quoted_match in quoted_sections:
            path_exclusions.add(quoted_match.span())
        
        # 找出所有方括號區段
        bracketed_sections = re.finditer(r"\[[^\]]*\]", normalized_formula)
        for bracket_match in bracketed_sections:
            path_exclusions.add(bracket_match.span())
        
        # 更精確的相對引用模式
        rel_pattern = r"(?<![A-Za-z0-9_\\/])([A-Z]{1,3})([0-9]{1,7})(?![A-Za-z0-9_\\/])"
        
        for match in re.finditer(rel_pattern, normalized_formula):
            is_processed = False
            match_start, match_end = match.span()
            
            # 檢查是否在已處理的絕對引用範圍內
            for span_start, span_end in processed_spans:
                if span_start <= match_start and match_end <= span_end:
                    is_processed = True
                    break
            
            # 檢查是否在路徑排除區域內
            if not is_processed:
                for excl_start, excl_end in path_exclusions:
                    if excl_start <= match_start and match_end <= excl_end:
                        is_processed = True
                        break
            
            # 額外的上下文檢查
            if not is_processed:
                context_start = max(0, match_start - 5)
                context_end = min(len(normalized_formula), match_end + 5)
                context = normalized_formula[context_start:context_end]
                
                path_indicators = ['\\', '/', '[', ']', '.xl']
                if any(indicator in context for indicator in path_indicators):
                    is_processed = True
            
            if not is_processed:
                col = match.group(1)
                row = match.group(2)
                
                try:
                    col_len = len(col)
                    row_num = int(row)
                    
                    if 1 <= col_len <= 3 and 1 <= row_num <= 1048576:
                        cell_ref = f"{col}{row}"
                        if cell_ref.lower() not in ['md9', 'ab12', 'cd34', 'xy99']:
                            references.append({
                                'workbook_path': current_workbook_path,
                                'sheet_name': current_sheet_name,
                                'cell_address': cell_ref,
                                'type': 'relative'
                            })
                except ValueError:
                    continue

        # === 新功能：處理範圍引用 ===
        # 在所有單個儲存格引用處理完後，檢查是否有範圍引用需要處理
        # 使用簡單的範圍模式匹配並應用range處理邏輯
        range_pattern = r"([A-Z]{1,3}\d{1,7}:[A-Z]{1,3}\d{1,7})"
        
        for range_match in re.finditer(range_pattern, normalized_formula):
            range_ref = range_match.group(1)
            range_start, range_end = range_match.span()
            
            # 檢查這個範圍是否已經被處理過
            is_range_processed = False
            for span_start, span_end in processed_spans:
                if span_start <= range_start and range_end <= span_end:
                    is_range_processed = True
                    break
            
            # 檢查是否在路徑排除區域內
            if not is_range_processed:
                for excl_start, excl_end in path_exclusions:
                    if excl_start <= range_start and range_end <= excl_end:
                        is_range_processed = True
                        break
            
            if not is_range_processed:
                # 處理範圍引用
                range_refs = self._process_range_reference(
                    range_ref, current_workbook_path, current_sheet_name, 'current'
                )
                references.extend(range_refs)

        return references
    
    def _process_range_reference(self, range_ref, workbook_path, sheet_name, ref_type):
        """
        處理range引用，根據大小決定展開或摘要
        
        Args:
            range_ref: Range引用 (如 A1:B5)
            workbook_path: 工作簿路徑
            sheet_name: 工作表名稱
            ref_type: 引用類型
            
        Returns:
            list: 處理後的引用列表
        """
        try:
            # 計算range大小
            cell_count = self._calculate_range_size(range_ref)
            
            if cell_count <= self.range_expand_threshold:
                # 小範圍：展開為個別儲存格
                return self._expand_range_to_cells(range_ref, workbook_path, sheet_name, ref_type)
            else:
                # 大範圍：創建摘要節點
                return self._create_range_summary(range_ref, workbook_path, sheet_name, ref_type, cell_count)
                
        except Exception as e:
            print(f"Warning: Could not process range {range_ref}: {e}")
            # 發生錯誤時，創建單個摘要節點
            return [{
                'workbook_path': workbook_path,
                'sheet_name': sheet_name,
                'cell_address': range_ref,
                'type': f'{ref_type}_range_error',
                'is_range_summary': True,
                'range_info': f'Error processing range: {e}'
            }]
    
    def _calculate_range_size(self, range_ref):
        """計算range包含的儲存格數量"""
        try:
            # 移除$符號並分割range
            clean_range = range_ref.replace('$', '').strip()
            start_cell, end_cell = clean_range.split(':')
            
            # 解析起始儲存格
            start_col, start_row = self._parse_cell_address(start_cell.strip())
            # 解析結束儲存格  
            end_col, end_row = self._parse_cell_address(end_cell.strip())
            
            # 計算行列數量
            row_count = abs(end_row - start_row) + 1
            col_count = abs(end_col - start_col) + 1
            
            return row_count * col_count
            
        except Exception as e:
            print(f"Warning: Could not calculate range size for {range_ref}: {e}")
            return 999  # 返回大數值，強制使用摘要模式
    
    def _parse_cell_address(self, cell_address):
        """解析儲存格地址為列號和行號"""
        import re
        match = re.match(r'([A-Z]+)(\d+)', cell_address.upper())
        if not match:
            raise ValueError(f"Invalid cell address: {cell_address}")
        
        col_letters = match.group(1)
        row_num = int(match.group(2))
        
        # 轉換列字母為數字 (A=1, B=2, ...)
        col_num = 0
        for char in col_letters:
            col_num = col_num * 26 + (ord(char) - ord('A') + 1)
        
        return col_num, row_num
    
    def _expand_range_to_cells(self, range_ref, workbook_path, sheet_name, ref_type):
        """將range展開為個別儲存格引用"""
        try:
            clean_range = range_ref.replace('$', '').strip()
            start_cell, end_cell = clean_range.split(':')
            
            start_col, start_row = self._parse_cell_address(start_cell.strip())
            end_col, end_row = self._parse_cell_address(end_cell.strip())
            
            # 確保起始位置小於結束位置
            min_col, max_col = min(start_col, end_col), max(start_col, end_col)
            min_row, max_row = min(start_row, end_row), max(start_row, end_row)
            
            references = []
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    # 轉換列號回字母
                    col_letters = self._col_num_to_letters(col)
                    cell_address = f"{col_letters}{row}"
                    
                    references.append({
                        'workbook_path': workbook_path,
                        'sheet_name': sheet_name,
                        'cell_address': cell_address,
                        'type': f'{ref_type}_from_range',
                        'original_range': range_ref
                    })
            
            return references
            
        except Exception as e:
            print(f"Warning: Could not expand range {range_ref}: {e}")
            return []
    
    def _col_num_to_letters(self, col_num):
        """將列號轉換為字母 (1=A, 2=B, ...)"""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(ord('A') + (col_num % 26)) + result
            col_num //= 26
        return result
    
    def _create_range_summary(self, range_ref, workbook_path, sheet_name, ref_type, cell_count):
        """創建range摘要節點"""
        # 生成range的hash值用於顯示
        import hashlib
        range_hash = hashlib.md5(f"{workbook_path}|{sheet_name}|{range_ref}".encode()).hexdigest()[:8]
        
        # 計算維度信息
        try:
            clean_range = range_ref.replace('$', '').strip()
            start_cell, end_cell = clean_range.split(':')
            start_col, start_row = self._parse_cell_address(start_cell.strip())
            end_col, end_row = self._parse_cell_address(end_cell.strip())
            
            row_count = abs(end_row - start_row) + 1
            col_count = abs(end_col - start_col) + 1
            dimension_info = f"{row_count}行×{col_count}列"
        except:
            dimension_info = f"{cell_count}個儲存格"
        
        return [{
            'workbook_path': workbook_path,
            'sheet_name': sheet_name,
            'cell_address': range_ref,
            'type': f'{ref_type}_range_summary',
            'is_range_summary': True,
            'range_info': f'Range摘要 (Hash: {range_hash}, {dimension_info}, 共{cell_count}個儲存格)'
        }]
    
    def _normalize_formula_paths(self, formula):
        """
        標準化公式中的路徑，將雙反斜線轉為單反斜線
        
        Args:
            formula: 原始公式字符串
            
        Returns:
            str: 標準化後的公式字符串
        """
        if not formula:
            return formula
        
        # 使用正則表達式找到所有外部引用路徑並標準化
        def normalize_path_match(match):
            full_match = match.group(0)
            path_part = match.group(1)
            
            # 標準化路徑部分
            normalized_path = os.path.normpath(path_part)
            
            # 重建完整的引用
            return full_match.replace(path_part, normalized_path)
        
        # 匹配外部引用中的路徑部分
        external_ref_pattern = r"'([^']*\[[^\]]+\][^']*)'!"
        normalized_formula = re.sub(external_ref_pattern, normalize_path_match, formula)
        
        return normalized_formula
    
    def get_explosion_summary(self, root_node):
        """
        獲取爆炸分析摘要
        
        Args:
            root_node: 根節點
            
        Returns:
            dict: 摘要信息
        """
        def count_nodes(node):
            count = 1
            for child in node.get('children', []):
                count += count_nodes(child)
            return count
        
        def get_max_depth(node):
            if not node.get('children'):
                return node.get('depth', 0)
            return max(get_max_depth(child) for child in node['children'])
        
        def count_by_type(node, type_counts=None):
            if type_counts is None:
                type_counts = {}
            
            node_type = node.get('type', 'unknown')
            type_counts[node_type] = type_counts.get(node_type, 0) + 1
            
            for child in node.get('children', []):
                count_by_type(child, type_counts)
            
            return type_counts
        
        return {
            'total_nodes': count_nodes(root_node),
            'max_depth': get_max_depth(root_node),
            'type_distribution': count_by_type(root_node),
            'circular_references': len(self.circular_refs),
            'circular_ref_list': self.circular_refs
        }


def explode_cell_dependencies(workbook_path, sheet_name, cell_address, max_depth=10, range_expand_threshold=5):
    """
    便捷函數：爆炸分析指定儲存格的依賴關係
    
    Args:
        workbook_path: Excel 檔案路徑
        sheet_name: 工作表名稱
        cell_address: 儲存格地址
        max_depth: 最大遞歸深度
        range_expand_threshold: Range展開閾值（小於等於此數量的range會展開為個別儲存格）
        
    Returns:
        tuple: (依賴樹, 摘要信息)
    """
    exploder = DependencyExploder(max_depth=max_depth, range_expand_threshold=range_expand_threshold)
    dependency_tree = exploder.explode_dependencies(workbook_path, sheet_name, cell_address)
    summary = exploder.get_explosion_summary(dependency_tree)
    
    return dependency_tree, summary


# 測試函數
if __name__ == "__main__":
    # 測試用例
    test_workbook = r"C:\Users\user\Desktop\pytest\test.xlsx"
    test_sheet = "Sheet1"
    test_cell = "A1"
    
    try:
        tree, summary = explode_cell_dependencies(test_workbook, test_sheet, test_cell)
        print("Dependency Tree:")
        print(tree)
        print("\nSummary:")
        print(summary)
    except Exception as e:
        print(f"Test failed: {e}")
