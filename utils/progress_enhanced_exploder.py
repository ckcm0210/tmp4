# -*- coding: utf-8 -*-
"""
Enhanced Dependency Exploder with Progress Display
增強版依賴關係爆炸分析器 - 包含進度顯示和日誌累積
"""

import re
import os
from urllib.parse import unquote
from utils.openpyxl_resolver import read_cell_with_resolved_references
from utils.range_processor import range_processor, process_formula_ranges
from utils.pure_indirect_logic import process_formula_with_pure_indirect

class ProgressCallback:
    """進度回調接口 - 支持實時訊息和累積日誌"""
    def __init__(self, progress_var=None, popup_window=None, log_text_widget=None):
        self.progress_var = progress_var
        self.popup_window = popup_window
        self.log_text_widget = log_text_widget
        self.current_step = 0
        self.total_steps = 0
        
    def update_progress(self, message, step=None):
        """更新進度訊息 - 同時更新實時顯示和累積日誌"""
        if step is not None:
            self.current_step = step
            
        # 格式化訊息
        if self.total_steps > 0:
            progress_text = f"[{self.current_step}/{self.total_steps}] {message}"
        else:
            progress_text = message
            
        # 更新實時進度標籤
        if self.progress_var:
            self.progress_var.set(progress_text)
            
        # 累積到日誌區域
        if self.log_text_widget:
            try:
                import datetime
                timestamp = datetime.datetime.now().strftime("%H:%M:%S")
                log_entry = f"[{timestamp}] {progress_text}\n"
                
                self.log_text_widget.config(state='normal')
                self.log_text_widget.insert('end', log_entry)
                self.log_text_widget.see('end')  # 自動滾動到最新
                self.log_text_widget.config(state='disabled')
            except Exception as e:
                print(f"Log update error: {e}")
                
        # 更新視窗
        if self.popup_window:
            try:
                self.popup_window.update()
            except:
                pass  # 視窗可能已關閉
                
        # 控制台輸出
        print(f"[Explode Progress] {progress_text}")
        
    def set_total_steps(self, total):
        """設置總步驟數"""
        self.total_steps = total
        self.current_step = 0

class EnhancedDependencyExploder:
    """增強版公式依賴鏈爆炸分析器 - 包含進度顯示"""
    
    def __init__(self, max_depth=10, range_expand_threshold=5, progress_callback=None):
        self.max_depth = max_depth
        self.range_expand_threshold = range_expand_threshold
        self.visited_cells = set()
        self.circular_refs = []
        self.progress_callback = progress_callback or ProgressCallback()
        self.processed_count = 0
        
    def explode_dependencies(self, workbook_path, sheet_name, cell_address, current_depth=0, root_workbook_path=None):
        """
        遞歸展開公式依賴鏈 - 增強版包含進度顯示
        """
        # 更新進度
        if current_depth == 0:
            self.progress_callback.update_progress("正在初始化依賴關係分析...")
            self.processed_count = 0
        
        self.processed_count += 1
        
        # 創建唯一標識符
        cell_id = f"{workbook_path}|{sheet_name}|{cell_address}"
        
        # 顯示當前處理的儲存格
        filename = os.path.basename(workbook_path)
        current_ref = f"{filename}!{sheet_name}!{cell_address}"
        self.progress_callback.update_progress(
            f"正在分析 {current_ref} (深度: {current_depth}/{self.max_depth}, 已處理: {self.processed_count})"
        )
        
        # 檢查遞歸深度限制
        if current_depth >= self.max_depth:
            self.progress_callback.update_progress(f"警告：達到最大遞歸深度限制 ({self.max_depth})")
            return self._create_limit_node(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path)
        
        # 檢查循環引用
        if cell_id in self.visited_cells:
            self.circular_refs.append(cell_id)
            self.progress_callback.update_progress(f"警告：檢測到循環引用 {current_ref}")
            return self._create_circular_node(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path)
        
        # 標記為已訪問
        self.visited_cells.add(cell_id)
        
        try:
            # 顯示讀取進度
            self.progress_callback.update_progress(f"正在讀取儲存格內容: {current_ref}")
            
            # 讀取儲存格內容
            cell_info = read_cell_with_resolved_references(workbook_path, sheet_name, cell_address)
            
            if 'error' in cell_info:
                self.progress_callback.update_progress(f"錯誤：無法讀取 {current_ref} - {cell_info['error']}")
                return self._create_error_node(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path, cell_info['error'])
            
            # 處理公式清理和INDIRECT解析
            original_formula = cell_info.get('formula')
            fixed_formula = None
            resolved_formula = None
            indirect_info = None
            
            if original_formula:
                self.progress_callback.update_progress(f"正在處理公式: {current_ref}")
                fixed_formula = self._clean_formula(original_formula)
                
                # === 新增：INDIRECT函數處理 ===
                if 'INDIRECT' in fixed_formula.upper():
                    self.progress_callback.update_progress(f"正在解析INDIRECT函數: {current_ref}")
                    indirect_info = process_formula_with_pure_indirect(
                        fixed_formula, workbook_path, sheet_name, cell_address
                    )
                    if indirect_info['has_indirect'] and indirect_info['success']:
                        resolved_formula = indirect_info['resolved_formula']
                        self.progress_callback.update_progress(f"INDIRECT解析完成，resolved: {resolved_formula}")
                    elif indirect_info['has_indirect']:
                        self.progress_callback.update_progress(f"INDIRECT解析失敗: {indirect_info['error']}")

            # 創建節點
            node = self._create_node(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path, cell_info, fixed_formula, resolved_formula, indirect_info)
            
            # 如果是公式，解析依賴關係
            if cell_info.get('cell_type') == 'formula' and cell_info.get('formula'):
                self.progress_callback.update_progress(f"正在解析公式依賴關係: {current_ref}")
                
                # === 新增：處理範圍地址 ===
                # 使用resolved_formula如果有INDIRECT，否則使用原始公式
                formula_for_ranges = resolved_formula if resolved_formula else cell_info['formula']
                ranges = process_formula_ranges(formula_for_ranges, workbook_path, sheet_name)
                if ranges:
                    self.progress_callback.update_progress(f"找到 {len(ranges)} 個範圍，正在處理...")
                    
                    for i, range_info in enumerate(ranges, 1):
                        try:
                            range_display = f"{os.path.basename(range_info['workbook_path'])}!{range_info['sheet_name']}!{range_info['address']}"
                            self.progress_callback.update_progress(f"正在處理範圍 {i}/{len(ranges)}: {range_display}")
                            
                            # 創建範圍節點
                            range_node = self._create_range_node(range_info, current_depth + 1, root_workbook_path)
                            node['children'].append(range_node)
                            
                        except Exception as e:
                            self.progress_callback.update_progress(f"錯誤：處理範圍失敗 {range_display} - {str(e)}")
                            error_node = self._create_error_node(
                                range_info['workbook_path'], range_info['sheet_name'], range_info['address'], 
                                current_depth + 1, root_workbook_path, str(e)
                            )
                            node['children'].append(error_node)
                
                # 處理單個儲存格引用 - 使用resolved_formula如果有INDIRECT
                formula_to_parse = resolved_formula if resolved_formula else cell_info['formula']
                # 使用dependency_exploder的新解析邏輯
                from utils.dependency_exploder import DependencyExploder
                temp_exploder = DependencyExploder(range_expand_threshold=self.range_expand_threshold)
                references = temp_exploder.parse_formula_references(formula_to_parse, workbook_path, sheet_name)
                
                if references:
                    self.progress_callback.update_progress(f"找到 {len(references)} 個儲存格引用，正在遞歸分析...")
                    
                    # 遞歸展開每個引用
                    for i, ref in enumerate(references, 1):
                        try:
                            ref_display = f"{os.path.basename(ref['workbook_path'])}!{ref['sheet_name']}!{ref['cell_address']}"
                            self.progress_callback.update_progress(f"正在處理引用 {i}/{len(references)}: {ref_display}")
                            
                            child_node = self.explode_dependencies(
                                ref['workbook_path'],
                                ref['sheet_name'],
                                ref['cell_address'],
                                current_depth + 1,
                                root_workbook_path or workbook_path
                            )
                            node['children'].append(child_node)
                        except Exception as e:
                            self.progress_callback.update_progress(f"錯誤：處理引用失敗 {ref_display} - {str(e)}")
                            error_node = self._create_error_node(
                                ref['workbook_path'], ref['sheet_name'], ref['cell_address'], 
                                current_depth + 1, root_workbook_path, str(e)
                            )
                            node['children'].append(error_node)
                
                if not ranges and not references:
                    self.progress_callback.update_progress(f"公式中未找到可解析的引用或範圍")
            
            # 移除已訪問標記（允許在不同分支中重複訪問）
            self.visited_cells.discard(cell_id)
            
            # 根節點完成時顯示總結
            if current_depth == 0:
                total_nodes = self._count_nodes(node)
                max_depth = self._get_max_depth(node)
                self.progress_callback.update_progress(
                    f"分析完成！共處理 {self.processed_count} 次，生成 {total_nodes} 個節點，最大深度: {max_depth}"
                )
            
            return node
            
        except Exception as e:
            # 移除已訪問標記
            self.visited_cells.discard(cell_id)
            
            self.progress_callback.update_progress(f"錯誤：處理 {current_ref} 時發生異常 - {str(e)}")
            return self._create_error_node(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path, str(e))
    
    def _clean_formula(self, formula):
        """清理公式中的路徑問題"""
        if not formula:
            return formula
            
        # 步驟1: 處理雙反斜線
        fixed_formula = formula.replace('\\\\', '\\')
        # 步驟2: 解碼 URL 編碼字符（如 %20 -> 空格）
        fixed_formula = unquote(fixed_formula)
        # 步驟3: 處理雙引號問題 - 將 ''path'' 改為 'path'
        fixed_formula = re.sub(r"''([^']*?)''", r"'\1'", fixed_formula)
        
        return fixed_formula
    
    def _create_node(self, workbook_path, sheet_name, cell_address, current_depth, root_workbook_path, cell_info, fixed_formula, resolved_formula=None, indirect_info=None):
        """創建標準節點 - 支持INDIRECT信息"""
        # 決定顯示格式
        current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
        if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path) or current_depth == 0:
            filename = os.path.basename(workbook_path)
            dir_path = os.path.dirname(workbook_path)
            short_display_address = f"[{filename}]{sheet_name}!{cell_address}"
            full_display_address = f"'{dir_path}\\[{filename}]{sheet_name}'!{cell_address}"
            display_address = short_display_address
        else:
            display_address = f"{sheet_name}!{cell_address}"
            short_display_address = display_address
            full_display_address = display_address

        # 基本節點信息
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
        
        # === 新增：INDIRECT信息 ===
        if indirect_info and indirect_info['has_indirect'] and indirect_info['success']:
            node['has_indirect'] = True
            node['resolved_formula'] = resolved_formula
        else:
            node['has_indirect'] = False
            
        return node
    
    def _create_limit_node(self, workbook_path, sheet_name, cell_address, current_depth, root_workbook_path):
        """創建深度限制節點"""
        display_address = self._get_display_address(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path)
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
    
    def _create_circular_node(self, workbook_path, sheet_name, cell_address, current_depth, root_workbook_path):
        """創建循環引用節點"""
        display_address = self._get_display_address(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path)
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
    
    def _create_error_node(self, workbook_path, sheet_name, cell_address, current_depth, root_workbook_path, error_msg):
        """創建錯誤節點"""
        display_address = self._get_display_address(workbook_path, sheet_name, cell_address, current_depth, root_workbook_path)
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
            'error': error_msg
        }
    
    def _create_range_node(self, range_info, current_depth, root_workbook_path):
        """創建範圍節點"""
        workbook_path = range_info['workbook_path']
        sheet_name = range_info['sheet_name']
        range_address = range_info['address']
        
        # 決定顯示格式
        current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
        if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
            filename = os.path.basename(workbook_path)
            if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                filename = filename.rsplit('.', 1)[0]
            display_address = f"[{filename}]{sheet_name}!{range_address}"
            short_display_address = display_address
            full_display_address = f"'{os.path.dirname(workbook_path)}\\[{os.path.basename(workbook_path)}]{sheet_name}'!{range_address}"
        else:
            display_address = f"{sheet_name}!{range_address}"
            short_display_address = display_address
            full_display_address = display_address
        
        # 構建範圍值顯示 - 新格式：9Rx1C | Hash: abc123def456
        rows = range_info.get('rows', 0)
        columns = range_info.get('columns', 0)
        hash_short = range_info.get('hash_short', 'N/A')
        
        # 組合顯示值：使用新格式
        range_value = f"{rows}Rx{columns}C | Hash: {hash_short}"
        
        return {
            'address': display_address,
            'short_address': short_display_address,
            'full_address': full_display_address,
            'workbook_path': workbook_path,
            'sheet_name': sheet_name,
            'cell_address': range_address,
            'value': range_value,
            'calculated_value': range_value,
            'formula': None,
            'type': 'range',
            'children': [],
            'depth': current_depth,
            'error': range_info.get('error'),
            # 範圍特有字段
            'range_info': {
                'dimensions': {
                    'rows': rows,
                    'columns': columns,
                    'total_cells': range_info.get('total_cells', 0),
                    'dimension_summary': f"{rows}行 x {columns}列"
                },
                'hash': {
                    'full_hash': range_info.get('hash', 'N/A'),
                    'short_hash': hash_short,
                    'content_summary': range_info.get('content_summary', '無內容摘要')
                }
            }
        }
    
    def _get_display_address(self, workbook_path, sheet_name, cell_address, current_depth, root_workbook_path):
        """獲取顯示地址"""
        current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
        if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
            filename = os.path.basename(workbook_path)
            if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                filename = filename.rsplit('.', 1)[0]
            return f"[{filename}]{sheet_name}!{cell_address}"
        else:
            return f"{sheet_name}!{cell_address}"
    
    def _count_nodes(self, node):
        """計算節點總數"""
        count = 1
        for child in node.get('children', []):
            count += self._count_nodes(child)
        return count
    
    def _get_max_depth(self, node):
        """獲取最大深度"""
        if not node.get('children'):
            return node.get('depth', 0)
        return max(self._get_max_depth(child) for child in node['children'])
    
    # 移除舊的parse_formula_references方法，現在使用dependency_exploder.py中的新方法
    
    def get_explosion_summary(self, root_node):
        """獲取爆炸分析摘要"""
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


def explode_cell_dependencies_with_progress(workbook_path, sheet_name, cell_address, max_depth=10, range_expand_threshold=5, progress_callback=None):
    """
    便捷函數：爆炸分析指定儲存格的依賴關係 - 包含進度顯示
    
    Args:
        workbook_path: Excel 檔案路徑
        sheet_name: 工作表名稱
        cell_address: 儲存格地址
        max_depth: 最大遞歸深度
        range_expand_threshold: Range展開閾值（小於等於此數量的range會展開為個別儲存格）
        progress_callback: 進度回調對象
        
    Returns:
        tuple: (依賴樹, 摘要信息)
    """
    exploder = EnhancedDependencyExploder(max_depth=max_depth, range_expand_threshold=range_expand_threshold, progress_callback=progress_callback)
    dependency_tree = exploder.explode_dependencies(workbook_path, sheet_name, cell_address)
    summary = exploder.get_explosion_summary(dependency_tree)
    
    return dependency_tree, summary