# `INDEX(MATCH)` 解析功能 - 技術設計方案 (v2.0)

## 1. 目標 (Objective)

本方案旨在擴充現有的 `progress_enhanced_exploder.py` 模組，賦予其解析 `INDEX` 及 `INDEX(MATCH)` 這類「計算型引用」公式的能力。最終目標是將這類公式還原為一個靜態的、單一的儲存格地址（例如 `'[data.xlsx]Sheet1'!I5`），以便我們的依賴關係分析引擎可以繼續追蹤其後續的依賴鏈路。

## 2. 核心原理 (Core Principle)

我們將採用一個更穩健的、不受公式位置影響的解析策略。對於一個包含 `INDEX` 的複雜公式（如 `=A1+INDEX(...)`），我們的處理流程如下：

1.  **偵測與提取**: 在整個公式字串中搜尋 `INDEX(`，並透過精確的括號配對演算法，提取出 `INDEX` 函式的完整內容，包括其三個核心參數：`array`, `row_num`, `column_num`。
2.  **遞迴解析參數**: 對 `row_num` 和 `column_num` 參數進行解析。如果它們本身是另一個需要計算的函式（如 `MATCH` 或 `OFFSET`），我們將重用專案中已有的「安全 COM 計算引擎」來獲得它們的數值結果。
3.  **計算最終地址**: 根據 `array` 參數的起始儲存格，以及計算出的行列偏移量，精確定位到最終指向的單一儲存格地址。
4.  **靜態替換**: 將原始公式中的整個 `INDEX(...)` 部分，替換為我們計算出的靜態地址，以便進行後續的依賴分析。

## 3. 應用情境與示範 (Scenarios & Demonstrations)

### 情境一：簡單 `INDEX` (純數字偏移)

*   **公式範例**: `=INDEX(C3:Z100, 3, 7)`
*   **解析結果**: `I5`

### 情境二：`INDEX` 配合簡單 `MATCH`

*   **公式範例**: `=INDEX(A1:A100, MATCH("Apple", B1:B100, 0))`
*   **解析思路**: 首先安全計算 `MATCH(...)` 的值（假設為10），然後計算出 `INDEX` 的結果。
*   **解析結果**: `A10`

### 情境三：複雜 `INDEX(MATCH, MATCH)` 內嵌於 VLOOKUP (更真實的情境)

*   **公式範例**: `=VLOOKUP(A1, 'C:\Users\user\Documents\Financial Reports\2025\[Q3_Forecast_Internal.xlsx]Assumptions'!$A:$D, INDEX({1,2,3,4}, 0, MATCH(B1, 'C:\Users\user\Documents\Financial Reports\2025\[Q3_Forecast_Internal.xlsx]Assumptions'!$A$1:$D$1, 0)), FALSE)`
*   **解析思路**:
    1.  我們的解析器會首先偵測到 `INDEX` 函式。
    2.  **解析 `array`**: 識別出 `array` 是一個常數陣列 `{1,2,3,4}`。
    3.  **解析 `row_num`**: 識別出 `row_num` 是 `0`。
    4.  **解析 `column_num`**: 識別出 `column_num` 是一個 `MATCH` 函式。
    5.  **安全計算 `MATCH`**: 讀取當前檔案 `B1` 儲存格的值（假設為 "Q2"），然後呼叫安全 COM 計算引擎，計算 `MATCH("Q2", 'C:\...\Assumptions'!$A$1:$D$1, 0)`，假設得到結果為 `3`。
    6.  **計算 `INDEX` 結果**: 計算 `INDEX({1,2,3,4}, 0, 3)`，得到結果 `3`。
    7.  **靜態替換**: 將原始公式中的 `INDEX(...)` 部分替換為計算結果 `3`。
    8.  **最終公式 (用於後續分析)**: `=VLOOKUP(A1, 'C:\...\Assumptions'!$A:$D, 3, FALSE)`。這樣，我們就成功消除了一個動態引用。

---

## 4. Python 程式碼實現思路 (示範性)

根據您的建議，我們將此邏輯封裝在一個獨立的新檔案 `utils/index_resolver.py` 中。

```python
# 檔案: utils/index_resolver.py

import re
# ... 其他必要的 import ...

class IndexResolver:
    """一個專門用來解析 INDEX 公式的類別。"""

    def __init__(self, exploder_context):
        """接收一個包含所有上下文的 exploder 物件。"""
        self.exploder = exploder_context

    def resolve_formula(self, formula):
        """在公式中尋找並取代所有 INDEX 函式。"""
        # 使用一個 while 迴圈來處理可能巢狀的 INDEX
        while "INDEX(" in formula.upper():
            # 尋找最內層的 INDEX 函式進行解析
            match = self._find_innermost_index(formula)
            if not match:
                break # 找不到更多可解析的 INDEX

            index_expression = match.group(0) # 例如 "INDEX(C3:Z100, 3, 7)"
            
            # 核心解析邏輯
            static_ref = self._resolve_single_index(index_expression)

            if static_ref:
                # 如果成功解析，就用靜態引用取代掉公式中的 INDEX(...) 部分
                formula = formula.replace(index_expression, static_ref)
            else:
                # 如果解析失敗，為了避免無限迴圈，將這個 INDEX 標記為已處理失敗
                formula = formula.replace(index_expression, f"UNRESOLVED_INDEX({index_expression})")
        
        return formula

    def _find_innermost_index(self, formula):
        """使用正規表示式和括號計數來找到最內層的 INDEX 函式。"""
        # ... 此處為複雜的正規表示式和括號匹配邏輯 ...
        pass

    def _resolve_single_index(self, index_expression):
        """解析單個 INDEX(...) 表達式。"""
        # 步驟 A: 從表達式中提取三個核心參數
        params = self._parse_index_parameters(index_expression)
        if not params: return None
        array_str, row_num_str, col_num_str = params

        # 步驟 B: 解析每一個參數
        array_info = self._analyze_array_range(array_str)
        if not array_info: return None

        row_num = self._resolve_parameter_value(row_num_str)
        if row_num is None: return None

        col_num = 1
        if col_num_str:
            col_num = self._resolve_parameter_value(col_num_str)
            if col_num is None: return None

        # 步驟 C: 計算最終的靜態地址
        return self._calculate_final_address(array_info, row_num, col_num)

    def _resolve_parameter_value(self, param_str):
        """解析一個參數，無論它是數字、儲存格引用還是函式。"""
        param_str = param_str.strip()
        if param_str.isdigit():
            return int(param_str)
        
        # 如果是 MATCH 或其他函式，呼叫「安全 COM 計算引擎」
        # 這裡的 self.exploder.safe_calculate(...) 就是我們重用的核心功能
        # 它會啟動隔離的 Excel 來計算 param_str 的值
        result = self.exploder.safe_calculate(param_str)
        
        if isinstance(result, (int, float)):
            return int(result)
        return None

    def _analyze_array_range(self, array_str):
        """解析 array 參數，返回其範圍和上下文資訊。"""
        # ... 此處為解析範圍字串的邏輯 ...
        pass

    def _calculate_final_address(self, array_info, row_num, col_num):
        """根據範圍資訊和偏移量，計算最終地址。"""
        # ... 此處為我們之前討論過的座標計算邏輯 ...
        pass

```

---

## 5. 整合方案 (Integration Plan)

1.  建立新檔案 `utils/index_resolver.py` 並將上述 `IndexResolver` 類別的邏輯放入其中。
2.  在 `utils/progress_enhanced_exploder.py` 的頂部，新增 `from .index_resolver import IndexResolver`。
3.  在 `EnhancedDependencyExploder` 的 `__init__` 方法中，建立一個 `self.index_resolver = IndexResolver(self)` 的實例。
4.  在 `explode_dependencies` 的主迴圈中，增加一個邏輯分支：在處理 `INDIRECT` 之前，先檢查公式是否包含 `INDEX`。如果包含，則呼叫 `self.index_resolver.resolve_formula(formula)`，並用其返回的、被簡化過的公式繼續後續的分析。

這個方案將新功能的複雜性完美地封裝在一個獨立的模組中，保持了主分析引擎的整潔，是一個非常清晰且可維護的架構。
