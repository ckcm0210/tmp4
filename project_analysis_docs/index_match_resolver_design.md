# `INDEX(MATCH)` 解析功能 - 技術設計方案

## 1. 目標 (Objective)

本方案旨在擴充現有的 `progress_enhanced_exploder.py` 模組，賦予其解析 `INDEX` 及 `INDEX(MATCH)` 這類「計算型引用」公式的能力。最終目標是將這類公式還原為一個靜態的、單一的儲存格地址（例如 `Sheet1!I5`），以便我們的依賴關係分析引擎可以繼續追蹤其後續的依賴鏈路。

---

## 2. 核心原理 (Core Principle)

我們將遵循您提出的、極具洞察力的三步走策略來解決這個問題。`INDEX` 函式的結構為 `INDEX(array, row_num, [column_num])`，我們的解析將圍繞這三個參數展開：

1.  **解析 `array` (範圍)**: 準確地識別出 `INDEX` 函式的第一個參數所代表的數據範圍，並確定其左上角的起始儲存格。
2.  **解析 `row_num` & `column_num` (行列偏移量)**: 計算出第二和第三個參數的最終**數值**。這是最關鍵的一步，因為這兩個參數可能是數字、另一個儲存格的引用，或是一個需要計算的函式（如 `MATCH`）。
3.  **計算最終地址**: 根據 `array` 的起始點和計算出的行列偏移量，精確定位到最終指向的單一儲存格地址。

---

## 3. 應用情境與示範 (Scenarios & Demonstrations)

為了驗證此方案的可行性，我們構想了幾個由簡到繁的典型情境。

### 情境一：簡單 `INDEX` (純數字偏移)

*   **公式範例**: `=INDEX(C3:Z100, 3, 7)`
*   **解析思路**:
    1.  **解析 `array`**: 識別出範圍是 `C3:Z100`，其起始點為 `C3`。
    2.  **解析偏移量**: `row_num` 是數字 `3`，`column_num` 是數字 `7`。
    3.  **計算最終地址**: 
        *   目標行 = `C3` 的行號 `3` + `row_num` `3` - 1 = `5`。
        *   目標欄 = `C3` 的欄號 `C` (第3欄) + `column_num` `7` - 1 = `I` (第9欄)。
        *   **最終結果**: `I5`。

### 情境二：`INDEX` 配合簡單 `MATCH`

*   **公式範例**: `=INDEX(A1:A100, MATCH("Apple", B1:B100, 0))`
*   **解析思路**:
    1.  **解析 `array`**: 識別出範圍是 `A1:A100`，起始點為 `A1`。
    2.  **解析偏移量**: 
        *   `row_num` 是一個函式：`MATCH("Apple", B1:B100, 0)`。
        *   **執行安全計算**: 我們將 `MATCH(...)` 這部分公式，放入我們隔離的 Excel 實例中進行計算。假設計算結果為 `10`。
        *   `column_num` 未提供，預設為 `1`。
    3.  **計算最終地址**: 
        *   目標行 = `A1` 的行號 `1` + `row_num` `10` - 1 = `10`。
        *   目標欄 = `A1` 的欄號 `A` (第1欄) + `column_num` `1` - 1 = `A` (第1欄)。
        *   **最終結果**: `A10`。

### 情境三：複雜 `INDEX(MATCH, MATCH)` 配合外部引用

*   **公式範例**: `=INDEX('[data.xlsx]Prices'!A1:D100, MATCH(A1, '[data.xlsx]Prices'!A1:A100, 0), MATCH(B1, '[data.xlsx]Prices'!A1:D1, 0))`
*   **解析思路**:
    1.  **解析 `array`**: 識別出範圍是外部檔案 `[data.xlsx]` 的工作表 `Prices` 上的 `A1:D100`。起始點為 `A1`。
    2.  **解析偏移量**: 
        *   `row_num` 是 `MATCH(A1, '[data.xlsx]Prices'!A1:A100, 0)`。
        *   `column_num` 是 `MATCH(B1, '[data.xlsx]Prices'!A1:D1, 0)`。
        *   **執行安全計算**: 我們需要執行兩次安全計算。首先，讀取當前檔案中 `A1` 的值（假設為 "CPU"）和 `B1` 的值（假設為 "Q2"）。然後：
            a. 計算 `MATCH("CPU", '[data.xlsx]Prices'!A1:A100, 0)`，假設得到結果 `5`。
            b. 計算 `MATCH("Q2", '[data.xlsx]Prices'!A1:D1, 0)`，假設得到結果 `3`。
    3.  **計算最終地址**: 
        *   目標行 = `A1` 的行號 `1` + `row_num` `5` - 1 = `5`。
        *   目標欄 = `A1` 的欄號 `A` (第1欄) + `column_num` `3` - 1 = `C` (第3欄)。
        *   **最終結果**: `'[data.xlsx]Prices'!C5`。

---

## 4. Python 程式碼實現思路 (示範性)

以下是實現上述邏輯的示範性 Python 程式碼，包含了詳細的註解以解釋每一步。**請注意：這不是最終的、可直接執行的程式碼，而是一個清晰的、用於溝通和確認的實現思路。**

```python
# 這是一個新的、假想的函式，將會被整合到 EnhancedDependencyExploder 類別中
def resolve_index_formula(self, formula, context_cell):
    """主函式，負責解析一個完整的 INDEX 公式。"""
    print(f"[INDEX-RESOLVER] 開始解析 INDEX 公式: {formula}")

    # 步驟 A: 從公式中提取三個核心參數
    # 我們需要一個強大的正規表示式來處理可能的巢狀結構
    params = self._parse_index_parameters(formula)
    if not params:
        print("[INDEX-RESOLVER] 無法解析參數，終止。")
        return None # 無法解析，返回 None

    array_str, row_num_str, col_num_str = params
    print(f"[INDEX-RESOLVER] 成功提取參數: ARRAY=[{array_str}], ROW=[{row_num_str}], COL=[{col_num_str}]")

    # 步驟 B: 解析每一個參數
    # B.1 - 解析範圍 (array)
    # `_analyze_array_range` 是一個需要我們實現的輔助函式
    # 它需要返回範圍的起始儲存格地址 (如 C3) 和其所屬的工作簿/工作表路徑
    array_info = self._analyze_array_range(array_str)
    if not array_info:
        print(f"[INDEX-RESOLVER] 無法解析 ARRAY 範圍: {array_str}")
        return None
    print(f"[INDEX-RESOLVER] ARRAY 範圍資訊: {array_info}")

    # B.2 - 解析行號 (row_num)
    # `_resolve_parameter_value` 是另一個核心輔助函式
    row_num = self._resolve_parameter_value(row_num_str, context_cell)
    if row_num is None or not isinstance(row_num, int):
        print(f"[INDEX-RESOLVER] 無法將 ROW 參數解析為一個整數: {row_num_str}")
        return None
    print(f"[INDEX-RESOLVER] 解析後的 ROW 偏移量: {row_num}")

    # B.3 - 解析欄號 (col_num)
    col_num = 1 # 如果 col_num 為空，預設為 1
    if col_num_str:
        col_num = self._resolve_parameter_value(col_num_str, context_cell)
        if col_num is None or not isinstance(col_num, int):
            print(f"[INDEX-RESOLVER] 無法將 COL 參數解析為一個整數: {col_num_str}")
            return None
    print(f"[INDEX-RESOLVER] 解析後的 COL 偏移量: {col_num}")

    # 步驟 C: 計算最終的靜態地址
    final_address = self._calculate_final_address(array_info, row_num, col_num)
    print(f"[INDEX-RESOLVER] 計算出的最終靜態地址: {final_address}")

    return final_address

def _resolve_parameter_value(self, param_str, context_cell):
    """解析一個參數，無論它是數字、儲存格引用還是函式。"""
    param_str = param_str.strip()

    # 情況一：參數是純數字
    if param_str.isdigit():
        return int(param_str)

    # 情況二：參數是另一個函式 (MATCH, OFFSET, etc.)
    # 這是您提出的絕妙想法：重用我們為 INDIRECT 建立的安全計算引擎
    if "MATCH(" in param_str.upper() or "OFFSET(" in param_str.upper():
        print(f"    [PARAM-RESOLVER] 檢測到巢狀函式: {param_str}。正在使用安全 COM 計算...")
        # `_calculate_indirect_safely` 雖然名字叫 indirect, 但它本質上是一個通用的
        # 「安全公式計算器」，我們完全可以重用它來計算 MATCH 等函式的結果。
        # 我們需要傳遞完整的上下文，包括當前工作簿、工作表和儲存格地址。
        result_dict = self._calculate_indirect_safely(
            param_str, 
            self.current_workbook_path, # 假設我們有這些上下文
            self.current_sheet_name,
            context_cell
        )
        
        # 檢查計算是否成功，並且結果是否為數字
        if result_dict and result_dict['success'] and isinstance(result_dict['static_reference'], (int, float)):
            return int(result_dict['static_reference'])
        else:
            return None # 計算失敗

    # 情況三：參數是單一儲存格引用
    # 我們需要使用 openpyxl 從已載入的工作簿中讀取其值
    # ... 此處省略讀取儲存格值的程式碼 ...

    return None # 未知或無法處理的參數類型


def _calculate_final_address(self, array_info, row_num, col_num):
    """根據範圍資訊和偏移量，計算最終地址。"""
    from openpyxl.utils import get_column_letter, column_index_from_string

    start_cell_str = array_info['start_cell'] # 例如 "C3"
    
    # 從起始儲存格解析出欄和列的數字
    match = re.match(r"([A-Z]+)([0-9]+)", start_cell_str)
    start_col_str, start_row_str = match.groups()
    start_col_idx = column_index_from_string(start_col_str)
    start_row_idx = int(start_row_str)

    # 計算最終的欄和列索引（偏移量從 1 開始）
    final_col_idx = start_col_idx + col_num - 1
    final_row_idx = start_row_idx + row_num - 1

    # 將最終的欄索引轉換回字母
    final_col_str = get_column_letter(final_col_idx)

    # 組合最終的儲存格地址
    final_cell_address = f"{final_col_str}{final_row_idx}"

    # 加上工作簿和工作表的前綴
    return f"{array_info['prefix']}{final_cell_address}"

```

---

## 5. 整合至 `progress_enhanced_exploder` 的初步構想

這個全新的 `resolve_index_formula` 函式將會被整合到 `EnhancedDependencyExploder` 類別的 `explode_dependencies` 主迴圈中。目前的迴圈邏輯是：

1.  讀取一個儲存格。
2.  檢查是否包含 `INDIRECT`，如果是，則呼叫 `INDIRECT` 解析器。
3.  如果不是，則使用正規表示式解析靜態引用。

我們將在其中加入一個新的步驟：

1.  讀取一個儲存格。
2.  檢查是否以 `=INDEX(...` 開頭，如果是，則呼叫我們新的 `resolve_index_formula` 函式。如果成功解析出靜態地址，則將其作為新的依賴節點加入分析隊列。
3.  如果不是 `INDEX`，再檢查是否包含 `INDIRECT`...
4.  以此類推。

透過這種方式，我們可以逐步增強分析引擎的能力，使其能夠處理越來越多樣、越來越複雜的 Excel 公式。
