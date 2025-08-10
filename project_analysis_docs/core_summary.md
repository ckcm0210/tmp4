# `core` 資料夾模組深度分析報告 (最終完整版)

## 引言

本文檔旨在對 `core` 資料夾內所有「實際運作中的」Python 模組，進行深入且詳盡的功能、設計與互動分析。目標是為未來的開發者或維護人員提供一份清晰、易於理解的技術說明，使其能快速掌握本專案的核心邏輯與架構。

---

## `data_processor.py`

### 1. 總體功用 (Overall Purpose)

此模組扮演著一個「UI 數據後處理器」的角色。它本身不負責從 Excel 讀取原始數據，而是專門處理那些已經被掃描並載入到使用者介面 (UI) 表格中的數據。其主要職責是根據 UI 上當前的狀態（例如使用者是否過濾了結果），從中提取數據用於特定功能，例如「摘要」或「報告」。

### 2. 詳細組件分析 (Detailed Component Analysis)

**Functions**

*   `_get_summary_data(controller)`:
    *   **用途與情境**: 這是一個內部輔助函式，當使用者點擊「摘要」相關功能時，此函式被呼叫以獲取當前 UI 上可見的數據。它解決了「只處理使用者看的見的數據」這一問題。
    *   **參數詳解**: 
        *   `controller`: 接收主 `WorksheetController` 物件。它需要從 `controller.view.result_tree` 來讀取 UI 表格中的所有項目，並從 `controller.all_formulas` 來獲取原始的、未經篩選的完整數據列表，以便進行比較。
    *   **返回詳解**: 返回一個元組 `(list, bool)`。`list` 是從 UI 表格中直接抓取的所有數據行的列表；`bool` 是一個標誌，如果列表中的項目數量與原始數據總量不符，則為 `True`，表示當前使用者正在篩選結果。

*   `get_unique_external_links(formulas_to_summarize, tree_columns)`:
    *   **用途與情境**: 此函式用於實現「摘要外部連結」功能。它會遍歷一個給定的公式列表，並從中找出所有指向外部 Excel 檔案的連結路徑（例如 `'C:\path\[file.xlsx]Sheet1'!A1`）。
    *   **參數詳解**:
        *   `formulas_to_summarize`: 一個列表，其中每個元素都是一條公式的數據（通常來自 `_get_summary_data`）。
        *   `tree_columns`: 一個欄位名稱列表，用於動態地找到 "formula" 欄位的索引，增加了程式的靈活性，避免了硬編碼。
    *   **返回詳解**: 返回一個經過排序且去除了重複項的字串列表，其中每個字串都是一個獨一無二的外部連結路徑。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: 
    *   `re`: 用於執行正規表示式，以精準匹配公式中的外部連結格式。
    *   `utils.range_optimizer`: 從中匯入了 `smart_range_display` 函式，用於將一系列儲存格地址轉換為更易讀的格式（例如 `A1:A5, B10`）。
*   **被匯入 (Imported By)**: 
    *   `core.worksheet_summary`: 呼叫此模組的 `_get_summary_data` 函式來獲取數據，為彈出摘要視窗做準備。

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **潛在效能問題**: 在 `get_unique_external_links` 函式中，`re.compile()` 被放在函式內部。這意味著每次呼叫該函式時，都會重新進行一次正規表示式的「編譯」操作。雖然單次操作很快，但如果此功能被頻繁觸發，將編譯操作移至模組的頂層作為一個全域常數（`EXTERNAL_LINK_PATTERN = re.compile(...)`），可以確保只編譯一次，從而提升效能。
*   **語法或結構建議**: 此檔案曾匯入 `smart_range_display` 函式但未使用，這是一個小瑕疵，在先前的步驟中已被修正。維持匯入語句的乾淨對程式碼健康至關重要。

---

## `dual_pane_controller.py`

### 1. 總體功用 (Overall Purpose)

此模組為「檢查模式 (Inspect Mode)」提供了核心的雙窗格 (Dual-Pane) 佈局框架。它會建立兩個可以獨立運作、並排顯示的分析面板，每個面板都能夠連接到 Excel、掃描並分析一個指定的儲存格。

### 2. 詳細組件分析 (Detailed Component Analysis)

**Class: `DualPaneController`**
*   **描述**: 這是管理整個雙窗格視圖的頂層控制器。它負責建立左右兩個窗格，並管理它們各自的控制器 (`InspectPaneController`)。
*   **方法**:
    *   `__init__(self, parent_frame, root_app)`: 建構函式。接收一個父級 Tkinter 框架和主應用程式視窗作為參數，並呼叫 `setup_dual_pane_layout` 來初始化介面。
    *   `setup_dual_pane_layout(self)`: 使用 `ttk.PanedWindow` 建立一個可由使用者拖動調整大小的雙窗格佈局。接著，為左右兩個窗格分別實例化一個 `InspectPaneController`。
    *   `get_left_controller(self)` / `get_right_controller(self)`: 提供外部獲取左或右窗格控制器實例的接口。
    *   `reset_both_panes(self)`: 提供一個方便的「一鍵重置」功能，同時清理左右兩個面板的狀態。

**Class: `InspectPaneController`**
*   **描述**: 這是單一分析面板的控制器，是這個模組的核心。它封裝了單一面板的所有 UI 元素和業務邏輯，從連接 Excel 到分析儲存格並顯示結果。
*   **方法**:
    *   `__init__(self, ...)`: 初始化單一面板，設定好所有需要的變數。
    *   `setup_pane_ui(self)`: 建立此面板中的所有 Tkinter UI 元件。
    *   `connect_to_excel(self)`: 處理連接到當前運行的 Excel 應用程式的邏輯。
    *   `scan_current_cell(self)`: 從輸入框獲取儲存格地址，讀取該儲存格的屬性，然後呈現結果。
    *   `display_cell_analysis(self, cell_data)`: 將掃描到的儲存格數據格式化後，顯示在結果文字區域中。
    *   `reset_pane(self)`: 將此面板的 UI 和內部數據重置到初始狀態。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: 
    *   `tkinter`
    *   `ui.worksheet.controller.WorksheetController`
*   **被匯入 (Imported By)**: 
    *   沒有檔案匯入此模組 (可能被動態載入)。

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **語法或結構建議:**
    *   匯入了 `WorksheetController`，但在程式碼中並未使用到它，應予以移除。
    *   `InspectPaneController` 類別的職責較多，可考慮將其拆分為獨立的 View 和 Controller。
    *   錯誤捕獲過於寬泛，建議改為捕獲更具體的錯誤類型。
*   **可增強功能:** `display_cell_analysis` 方法中有一個 `TODO` 註解，是明確的增強方向。

### 5. 待處理的 `import` 語句
*   **已修正:** `import win32com.client` 已被移至檔案頂部。

---

## `excel_connector.py`

### 1. 總體功用 (Overall Purpose)
此模組包含專門用於管理與 Excel 應用程式連接的函式。它處理重新連接、啟動 Excel 視窗以及尋找其他已開啟工作簿路徑等操作。

### 2. 詳細組件分析 (Detailed Component Analysis)
*   `_perform_excel_reconnection(...)`: 內部函式，執行實際的重新連接邏輯。
*   `reconnect_to_excel(controller)`: 公開函式，由 UI 呼叫，處理重新連接並更新 UI。
*   `activate_excel_window(controller)`: 將已連接的 Excel 視窗帶到前景。
*   `find_external_workbook_path(controller, file_name)`: 尋找外部工作簿的完整路徑。

### 3. 模組互動 (Interactions)
*   **匯入 (Imports):** `os`, `win32com.client`, `tkinter.messagebox`, `win32gui`, `win32con`
*   **被匯入 (Imported By):** `utils.excel_helpers`, `ui.worksheet_ui`, `core.worksheet_tree`, `core.worksheet_export`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)
*   **結構建議:** 錯誤處理中的 `except Exception as e:` 捕獲過於寬泛。

### 5. 待處理的 `import` 語句
*   此檔案中沒有發現在程式碼中間的 `import` 語句。

---

## `excel_scanner.py`

### 1. 總體功用 (Overall Purpose)
這是執行核心「掃描」功能的關鍵模組。它連接到 Excel，識別目標範圍，遍歷儲存格以尋找所有公式，並在掃描期間管理 UI 回饋（如進度條）。

### 2. 詳細組件分析 (Detailed Component Analysis)
*   `_get_formulas_from_excel(...)`: 內部函式，使用高效的 `SpecialCells` 方法來獲取所有包含公式的儲存格並處理它們。
*   `refresh_data(controller, ...)`: 由 UI 呼叫的主要公開函式，負責協調整個掃描流程，包括連接、確定範圍、定義進度回呼、呼叫掃描函式及更新最終結果。

### 3. 模組互動 (Interactions)
*   **匯入 (Imports):** `os`, `time`, `win32com.client`, `tkinter.messagebox`, `psutil`, `win32gui`, `win32process`, `win32con`, `core.formula_classifier`, `core.worksheet_tree`
*   **被匯入 (Imported By):** `core.worksheet_export`, `core.formula_comparator`, `ui.modes.inspect_mode`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)
*   **結構建議:** `refresh_data` 函式非常長且有多層巢狀的 `try...except`，是重構的主要候選者。應將其分解為更小的函式以提高可讀性。

### 5. 待處理的 `import` 語句
*   **已修正:** `import traceback` 已被移至檔案頂部。

---

## `formula_classifier.py`

### 1. 總體功用 (Overall Purpose)
一個專一的輔助模組，其唯一功能是將給定的公式字串分類為「external link」、「local link」或「formula」。

### 2. 詳細組件分析 (Detailed Component Analysis)
*   `classify_formula_type(formula)`: 模組中唯一的函式，根據一系列規則（正規表示式、'!' 字元等）來確定公式類型。

### 3. 模組互動 (Interactions)
*   **匯入 (Imports):** `.link_analyzer`
*   **被匯入 (Imported By):** `core.excel_scanner`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)
*   **語法建議:** 發現並移除了未使用的 `import re` 語句。

### 5. 待處理的 `import` 語句
*   此檔案中沒有發現在程式碼中間的 `import` 語句。

---

## `formula_comparator.py`

### 1. 總體功用 (Overall Purpose)
此模組提供了「Excel 公式比較器」分頁的主要 UI 和邏輯。它設定了一個雙面板視圖，使用者可以在其中載入來自兩個工作表的數據，然後將公式從一邊同步到另一邊。

### 2. 詳細組件分析 (Detailed Component Analysis)
*   **Class: `ExcelFormulaComparator`**: 封裝了比較器分頁的全部功能，包括 UI 設定、掃描觸發和同步邏輯。
    *   `setup_ui()`: 建立所有 UI 元件。
    *   `scan_worksheet_full()` / `scan_worksheet_selected()`: 掃描按鈕的事件處理常式。
    *   `sync_formulas(...)`: 執行同步的核心邏輯。
    *   `_get_active_controller(...)`: 獲取當前活動的控制器（左或右面板）。

### 3. 模組互動 (Interactions)
*   **匯入 (Imports):** `tkinter`, `win32com.client`, `re`, `time`, `ui.worksheet.controller`, `ui.worksheet.view`, `core.excel_scanner`, `core.worksheet_tree`
*   **被匯入 (Imported By):** `main`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)
*   **結構建議:** 類別很大，可以考慮將 UI 設定和後端邏輯進一步分離。

### 5. 待處理的 `import` 語句
*   此檔案的 `import` 語句在先前的操作中已全部修正。

---

## `graph_generator.py`

### 1. 總體功用 (Overall Purpose)
此模組負責產生一個自包含的、可互動的 HTML 檔案來視覺化依賴關係圖。它將所有必要的 HTML、CSS 和一個自訂的 JavaScript 函式庫嵌入到單一檔案中。

### 2. 詳細組件分析 (Detailed Component Analysis)
*   **Class: `GraphGenerator`**: 處理整個圖表產生過程。
    *   `generate_graph()`: 協調整個流程：計算節點位置、產生 HTML 內容、寫入檔案並在瀏覽器中打開。
    *   `_generate_standalone_html()`: 核心方法，將包含大量 JavaScript 程式碼的 HTML 範本與節點數據結合。

### 3. 模組互動 (Interactions)
*   **匯入 (Imports):** `os`, `webbrowser`, `json`
*   **被匯入 (Imported By):** `core.worksheet_tree`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)
*   **結構建議:** 將巨大的 JavaScript 字串嵌入 Python 程式碼中使得維護極其困難。強烈建議將 JavaScript 和 CSS 分離到獨立的 `.js` 和 `.css` 檔案中，Python 只負責讀取範本並注入數據。
*   **可增強功能:** 可考慮使用成熟的第三方圖表函式庫（如官方的 vis.js、D3.js）來取代自訂的 JavaScript 實現，以獲得更強大的功能和更好的效能。

### 5. 待處理的 `import` 語句
*   此檔案中沒有發現在程式碼中間的 `import` 語句。

---

## `link_analyzer.py`

### 1. 總體功用 (Overall Purpose)
一個專門的 Excel 公式字串解析器。其主要目標是從一個公式中找出所有儲存格引用（外部、內部等），提取它們，並檢索它們的值。

### 2. 詳細組件分析 (Detailed Component Analysis)
*   `is_external_link_regex_match(formula_str)`: 快速檢查公式是否包含外部連結。
*   `get_referenced_cell_values(...)`: 模組中最核心、最複雜的函式，使用一系列正規表示式來尋找並處理公式中的每一個儲存格引用。
*   `parse_external_path_and_sheet(...)`: 從引用字串中解析出檔案路徑和工作表名稱。

### 3. 模組互動 (Interactions)
*   **匯入 (Imports):** `re`, `os`
*   **被匯入 (Imported By):** `core.worksheet_tree`, `core.formula_classifier`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)
*   **潛在效能問題:** `get_referenced_cell_values` 函式在每次被呼叫時都會重新編譯五個正規表示式。應將這些模式移到模組級別作為常數，以提高效能。
*   **結構建議:** `get_referenced_cell_values` 函式非常長且複雜，應將其分解為更小的輔助函式。

### 5. 待處理的 `import` 語句
*   此檔案中沒有發現在程式碼中間的 `import` 語句。

---

## `mode_manager.py`

### 1. 總體功用 (Overall Purpose)
此模組充當應用程式 UI 的狀態機。它定義了不同的應用程式模式（「正常」和「檢查」）並管理它們之間的轉換，負責應用特定於模式的設定。

### 2. 詳細組件分析 (Detailed Component Analysis)
*   **Enum: `AppMode`**: 為不同的應用程式模式提供清晰的常數。
*   **Class: `ModeManager`**: 持有應用程式的當前狀態並處理所有與模式相關的邏輯。
    *   `register_mode_switch_callback(...)`: 實現觀察者模式，允許其他模組在模式變更時接收通知。
    *   `toggle_mode()`: 處理狀態轉換邏輯。

### 3. 模組互動 (Interactions)
*   **匯入 (Imports):** `tkinter`, `enum`
*   **被匯入 (Imported By):** `main`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)
*   **結構建議:** 程式碼結構良好，使用了設定字典和回呼系統，乾淨且易於擴展。

### 5. 待處理的 `import` 語句
*   此檔案中沒有發現在程式碼中間的 `import` 語句。

---

## `worksheet_export.py`

### 1. 總體功用 (Overall Purpose)
提供將數據匯出到 Excel 檔案以及從 Excel 檔案匯入公式以更新當前工作表的功能。

### 2. 詳細組件分析 (Detailed Component Analysis)
*   `export_formulas_to_excel(controller)`: 處理匯出流程，使用 `openpyxl` 建立新的 Excel 檔案。
*   `import_and_update_formulas(controller)`: 處理匯入和更新的複雜流程，包括讀取檔案、禁用 Excel 自動計算、請求使用者確認以及將公式寫回當前工作表。

### 3. 模組互動 (Interactions)
*   **匯入 (Imports):** `openpyxl`, `os`, `tkinter`, `re`, `core.excel_connector`, `core.excel_scanner`
*   **被匯入 (Imported By):** `ui.worksheet_ui`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)
*   **結構建議:** `import_and_update_formulas` 函式非常長，應分解為更小的輔助函式。
*   **效能:** 在批量更新前禁用了 Excel 的自動計算和事件，這是非常好的效能優化實踐。

### 5. 待處理的 `import` 語句
*   **已修正:** `import time` 已被移至檔案頂部。

---

## `worksheet_summary.py`

### 1. 總體功用 (Overall Purpose)
提供「摘要外部連結」功能的觸發器。它充當主控制器和摘要 UI 視窗之間的中間人。

### 2. 詳細組件分析 (Detailed Component Analysis)
*   `summarize_external_links(controller)`: 當使用者需要產生摘要時呼叫的主要函式。它從 UI 獲取數據，然後建立並開啟一個新的 `SummaryWindow` 來顯示它。

### 3. 模組互動 (Interactions)
*   **匯入 (Imports):** `tkinter`, `ui.summary_window`, `core.data_processor`
*   **被匯入 (Imported By):** `ui.worksheet_ui`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)
*   **結構建議:** 程式碼簡單、乾淨且結構良好。

### 5. 待處理的 `import` 語句
*   此檔案中沒有發現在程式碼中間的 `import` 語句。

---

## `worksheet_tree.py`

### 1. 總體功用 (Overall Purpose)
一個功能非常龐大且複雜的模組，是「UI 互動邏輯中心」。它以一組獨立函式的形式，負責處理主視窗中公式列表 (Treeview) 的各種互動操作，如篩選、排序、點擊、雙擊、右鍵選單、前往引用和彈出「依賴關係爆炸」視窗等。

### 2. 詳細組件分析 (Detailed Component Analysis)
*   `apply_filter(...)`: 根據輸入框內容篩選 Treeview。
*   `sort_column(...)`: 處理點擊欄位標題排序的邏輯。
*   `go_to_reference(...)`: 處理導航到 Excel 中特定儲存格的複雜邏輯。
*   `on_select(...)` / `on_double_click(...)`: 處理在 Treeview 中選擇或雙擊項目的事件。
*   `explode_dependencies_popup(...)`: 一個巨大的函式，從頭開始建立整個「依賴關係爆炸」彈出視窗及其所有互動邏輯。

### 3. 模組互動 (Interactions)
*   **匯入 (Imports):** `tkinter`, `os`, `re`, `win32com.client`, `json`, `datetime`, `traceback` 以及多個 `core` 和 `utils` 套件中的模組。
*   **被匯入 (Imported By):** `ui.worksheet_ui`, `core.formula_comparator`, `core.excel_scanner`, `ui.modes.inspect_mode`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)
*   **結構建議:** **這是整個專案中最需要重構的檔案。** 它違反了單一職責原則。強烈建議將其拆分，例如將 `explode_dependencies_popup` 相關邏輯獨立成新模組，將導航相關函式組合進一個 `Navigation` 類別，並將事件處理常式移回它們所服務的控制器中。
*   **語法建議:** 檔案中存在 `from core.worksheet_tree import go_to_reference` 這樣的自我匯入，這是不好的實踐，應予以修正。

### 5. 待處理的 `import` 語句
*   此檔案的 `import` 語句在先前的操作中已全部修正。