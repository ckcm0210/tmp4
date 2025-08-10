# `utils` 資料夾模組深度分析報告 (最終完整版)

## 引言

本文檔旨在對 `utils` (Utilities) 資料夾內所有「實際運作中的」Python 模組，進行深入且詳盡的功能、設計與互動分析。`utils` 資料夾是本專案的「工具箱」，它包含了專案中最核心、最複雜的底層演算法和輔助工具，是實現所有上層功能的基礎。對此資料夾的理解，是掌握本專案精髓的關鍵。

---

## `dependency_converter.py`

### 1. 總體功用 (Overall Purpose)

此模組扮演著「數據化妝師」的角色。它的核心職責是接收由 `progress_enhanced_exploder` 產生的、原始的、樹狀結構的依賴關係數據，並將其精心轉換和美化成一種可以直接被圖表視覺化函式庫（如 `graph_generator.py`）所使用的「節點 (Node)」和「邊 (Edge)」列表。它處理了所有與「顯示」相關的細節，例如為節點產生標籤、為不同檔案分配顏色等。

### 2. 詳細組件分析 (Detailed Component Analysis)

*   `convert_tree_to_graph_data(dependency_tree_data)`:
    *   **用途與情境**: 這是此模組最主要的公開函式。當「依賴爆炸」分析完成後，`core.worksheet_tree` 會呼叫此函式，將分析結果傳入，以獲取用於生成視覺化圖表的標準化數據。
    *   **執行流程**: 
        1.  呼叫內部遞迴函式 `collect_filenames` 遍歷整個依賴樹，收集所有涉及到的外部檔名。
        2.  呼叫 `_generate_unique_colors_for_files` 為每一個獨特的檔名分配一個專屬顏色。
        3.  再次呼叫內部遞迴函式 `traverse_tree`，第二次遍歷依賴樹。這一次，它會為樹中的每一個節點建立一個對應的「圖表節點」字典，並使用各種輔助函式 (`_create_short_address`, `_create_short_formula` 等) 來產生不同詳盡程度的標籤和提示文字。
        4.  最終返回 `nodes_data` 和 `edges_data` 兩個列表。

*   `_generate_unique_colors_for_files(filenames)`:
    *   **用途與情境**: 為圖表中的節點提供視覺區分。它接收一個檔名列表，為每個檔名分配一個獨特的顏色。為了美觀，它會優先使用一個預定義的調色盤，如果檔名過多，它還能透過 `colorsys` 函式庫動態生成新的、區分度高的顏色。

*   `_create_...` / `_format_...` 系列輔助函式:
    *   **用途與情境**: 這些函式是數據美化的核心。它們負責將原始的、可能很長的數據（如公式、儲存格地址、計算結果）轉換成多種格式，以適應不同的顯示需求（例如，節點上的簡短顯示、滑鼠懸停時的完整提示等）。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: `os`, `re`, `colorsys`, `urllib.parse`
*   **被匯入 (Imported By)**: `core.worksheet_tree`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **潛在效能問題**: `convert_tree_to_graph_data` 函式會遍歷整個依賴樹兩次。對於非常深或闊的依賴樹，這可能會帶來不必要的效能開銷。可以將其重構為單次遍歷：在第一次（也是唯一一次）遍歷樹時，動態地收集遇到的新檔名，並即時為其分配顏色，從而避免第二次完整的遍歷。

### 5. 待處理的 `import` 語句

*   **已修正**: `import colorsys`, `import urllib.parse` 等已在先前的操作中移至檔案頂部。

---

## `excel_helpers.py`

### 1. 總體功用 (Overall Purpose)

此模組是 `SummaryWindow`（摘要視窗）的專屬「執行官」。它包含了兩個非常高階的函式，這兩個函式封裝了使用者在摘要視窗中可以執行的最複雜的操作：在 Excel 中高亮顯示受影響的儲存格，以及對外部連結執行批量取代。

### 2. 詳細組件分析 (Detailed Component Analysis)

*   `select_ranges_in_excel(summary_window, ...)`:
    *   **用途與情境**: 當使用者在摘要視窗中點擊「Go to Excel and Select Affected Ranges」按鈕時，此函式被呼叫。
    *   **執行流程**: 它首先從 UI 中獲取使用者當前選擇的連結，然後從快取中找出所有包含此連結的儲存格地址。在進行一系列檢查（如 Excel 是否連接）後，它會呼叫 `_perform_excel_selection` 來執行實際的選取操作。
    *   **設計模式**: 這是「UI 邏輯」與「核心邏輯」分離的良好示範。此函式負責所有與使用者互動的部分（如彈出警告框），而將與 Excel 的直接通訊交給了 `_perform_excel_selection`。

*   `replace_links_in_excel(summary_window, ...)`:
    *   **用途與情境**: 這是「Perform Replacement in Excel」按鈕背後的巨人。它是一個極其複雜但非常重要的函式，負責安全地執行連結取代操作。
    *   **執行流程**: 
        1.  **大量驗證**: 在執行任何操作前，它會進行多達十幾項的嚴格驗證，包括：使用者是否選擇了舊連結？新連結是否為空？新連結的格式是否正確？新舊連結指向的檔案是否存在？新檔案中是否包含了所有必要的 worksheet？等等。這使得此功能非常穩健。
        2.  **使用者確認**: 所有驗證通過後，它會彈出一個詳細的確認對話框，二次確認使用者的意圖。
        3.  **效能優化**: 在執行取代前，它會透過 COM 介面暫時關閉 Excel 的自動計算和事件更新，以極大地提升批量寫入的效能。
        4.  **批量執行**: 將需要更新的儲存格分批次（batching），逐批寫入 Excel，並在 UI 上更新進度條。
        5.  **狀態恢復**: 使用 `finally` 區塊確保無論成功或失敗，Excel 的原始設定（如自動計算）都會被恢復。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: `tkinter`, `os`, `re`, `openpyxl`, `core.excel_connector`
*   **被匯入 (Imported By)**: `ui.summary_window`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **結構建議**: `replace_links_in_excel` 函式是整個專案中最需要被重構的函式。它接收了 18 個參數，是典型的「程式碼壞味道」。應將其重構，改為只接收 `summary_window` 一個物件，並將其巨大的邏輯拆分為多個更小的、職責單一的私有函式。

### 5. 待處理的 `import` 語句

*   無。

---

## `excel_io.py`

### 1. 總體功用 (Overall Purpose)

此模組提供與 Excel 檔案 I/O 相關的輔助函式。它的特別之處在於能夠在不依賴 Excel 主程式執行的情況下，直接讀取磁碟上的檔案內容，並且為了兼容性，能同時處理新版 (`.xlsx`) 和舊版 (`.xls`) 的檔案格式。

### 2. 詳細組件分析 (Detailed Component Analysis)

*   `read_external_cell_value(...)`: 
    *   **用途與情境**: 這是此模組的核心功能，用於從一個外部（未在主程式中開啟）的 Excel 檔案裡讀取特定單一儲存格的值。
    *   **執行流程**: 它會先判斷檔案的副檔名。如果是 `.xlsx` 或 `.xlsm`，它會使用 `openpyxl` 函式庫來讀取。如果是舊版的 `.xls`，它會轉而使用 `xlrd` 函式庫來讀取。這種對不同格式的處理能力，大大增強了工具的適用範圍。
*   `find_matching_sheet(...)`: 一個輔助函式，用於在一個透過 COM 連接的 Excel 工作簿物件中，根據名稱尋找特定的工作表。
*   `get_sheet_by_name(...)`: 功能同上，但操作的對象是 `openpyxl` 的工作簿物件。
*   `calculate_similarity(...)`: 一個通用的字串相似度計算函式，使用「編輯距離 (Levenshtein distance)」演算法。它計算將一個字串轉換成另一個所需的最少編輯次數，從而得出一個 0.0 到 1.0 之間的相似度分數。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: `os`, `re`, `openpyxl`, `xlrd`
*   **被匯入 (Imported By)**: `core.worksheet_tree`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **結構建議**: `calculate_similarity` 函式的功能與「Excel I/O」這個模組的主題無關。為了讓模組的職責更單一，建議將這個純演算法函式移動到一個更通用的輔助模組中，例如一個新建的 `utils/string_utils.py` 或之前被廢棄的 `utils/helpers.py`。
*   **依賴管理**: 此模組動態匯入了 `xlrd`。這意味著如果使用者需要分析舊版 `.xls` 檔案，就必須手動安裝這個第三方函式庫，否則程式會在執行時出錯。應在專案的說明文件中明確指出這個「可選依賴」。

### 5. 待處理的 `import` 語句

*   **已修正**: `import xlrd` 已被移至檔案頂部。

---

## `openpyxl_resolver.py`

### 1. 總體功用 (Overall Purpose)

此模組是專案中一個設計非常精良的部分，它實現了一個「增強版」的 `openpyxl` 讀取器。`openpyxl` 在讀取含有外部連結的公式時，會將其顯示為 `=[1]Sheet1!A1` 這樣的內部索引，而不是人類可讀的檔案路徑。此模組的核心功能就是解決這個問題，它能夠自動將這些索引解析為完整的、帶有磁碟路徑的公式字串。

### 2. 詳細組件分析 (Detailed Component Analysis)

*   **設計模式**: 它巧妙地運用了「**裝飾器模式 (Decorator Pattern)**」。它沒有修改 `openpyxl` 的原始碼，而是定義了三個 `View` 類別（`ResolvedWorkbookView`, `ResolvedSheetView`, `ResolvedCellView`），分別「包裝」了 `openpyxl` 對應的 `Workbook`, `Worksheet`, 和 `Cell` 物件。
*   **`ResolvedCellView`**: 這是實現解析的關鍵。它攔截了所有對儲存格 `.value` 屬性的存取請求。當外部程式碼試圖讀取一個公式儲存格的值時，它不會立即返回值，而是先呼叫內部的 `_resolve_formula_string` 函式，將 `[1]` 這樣的索引替換成真實的檔案路徑後，再返回這個「已解析」的公式字串。
*   **`load_resolved_workbook(...)`**: 這是外部模組與本模組互動的主要入口。它接收一個檔案路徑，使用 `openpyxl`（或 `safe_cache`）載入工作簿，然後立刻用 `ResolvedWorkbookView` 將其包裝起來，返回給呼叫者一個功能被增強過的「解析版」工作簿物件。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: `openpyxl`, `os`, `re`, `.safe_cache`, `traceback`
*   **被匯入 (Imported By)**: `utils.progress_enhanced_exploder`, `core.worksheet_tree`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **優點**: 這是專案中軟體設計的最佳實踐之一。它在不侵入第三方函式庫的前提下，優雅地擴展了其功能，使得上層程式碼在處理外部連結時變得異常簡單。同時，它還整合了 `safe_cache`，兼顧了效能。

### 5. 待處理的 `import` 語句

*   **已修正**: `from .safe_cache import ...`, `import traceback` 已被移至檔案頂部。

---

## `progress_enhanced_exploder.py`

### 1. 總體功用 (Overall Purpose)

這是整個專案的**核心分析引擎**，也是所有 `INDIRECT` 解析方案的最終、權威版本。它實現了一個「超安全」的依賴關係分析器，能夠遞迴地「引爆」一個儲存格的所有公式引用，建立一個完整的依賴樹，同時為長時間的分析提供即時的進度回報。

### 2. 詳細組件分析 (Detailed Component Analysis)

*   **Class: `ProgressCallback`**: 一個獨立的進度回報類別，可以與 UI 的進度條和日誌文字框掛鉤，將後端分析的進度即時傳遞給前端。
*   **Class: `EnhancedDependencyExploder`**: 核心分析引擎類別。
    *   `_open_workbook_for_calculation(...)`: 這是此模組安全性的基石。它**從不**使用使用者當前開啟的 Excel，而是透過 `win32com.client.DispatchEx` 來建立一個全新的、完全獨立的、隱藏的 Excel 程序。它會記錄下這個新程序的 PID。
    *   `_calculate_indirect_safely(...)`: 使用上述建立的「沙箱化」Excel 實例來安全地計算 `INDIRECT` 函數的內容。
    *   `_ultra_safe_cleanup()`: 這是專案中最值得稱讚的安全設計。在分析結束後，它會執行一個多階段的清理流程：1. 正常關閉所有它自己建立的 COM 物件。 2. 強制 Python 進行記憶體回收。 3. 根據之前記錄的 PID，使用 `psutil` 函式庫來檢查這些 Excel 程序是否還在執行。 4. 如果還在，則**強制終止**這些由它自己建立的「殭屍程序」。這從根本上杜絕了因程式錯誤而導致的 Excel 程序殘留和檔案鎖定問題。
    *   `explode_dependencies(...)`: 主要的遞迴分析函式，在每個關鍵步驟都插入了 `progress_callback` 的呼叫，以實現進度匯報。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: `win32com.client`, `pythoncom`, `psutil`, `utils.openpyxl_resolver`, `utils.range_processor` 等。
*   **被匯入 (Imported By)**: `core.worksheet_tree`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **優點**: 這是專業級的 COM 自動化程式碼。對程序隔離、資源釋放和錯誤處理的考慮非常周全，是整個專案中最穩定、最可靠的部分。

### 5. 待處理的 `import` 語句

*   **已修正**: 所有在函式內部的 `import` 語句都已被移至檔案頂部。

---

## `range_optimizer.py`

### 1. 總體功用 (Overall Purpose)

此模組是一個演算法密集型的工具，其核心功能是「範圍優化」。它接收一個可能包含大量、離散的儲存格地址列表，並透過演算法將其合併、轉換為人類可讀的、最簡潔的範圍表示方式（例如，將 `A1, A2, A3, B1, B2, B3` 優化為 `A1:B3`）。

### 2. 詳細組件分析 (Detailed Component Analysis)

*   `parse_excel_address(addr)`: 一個強大的地址解析器，能夠理解並標準化多種使用者輸入格式，包括 `A1:B2`（範圍）、`A:C`（整列）、`1:5`（整行）等。
*   `optimize_ranges(parsed_addresses)`: 核心優化演算法。它不僅僅是尋找連續的儲存格，其內部還包含一個 `detect_rectangles` 函式，試圖在所有離散的點中找出最大的矩形區域，以實現最高效的壓縮。
*   `smart_range_display(addresses)`: 對外提供的主要函式，它整合了所有內部邏輯，並在最終結果過於零散時，用 `...` 來進行智慧截斷，以提升顯示效果。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: `re`, `collections`, `openpyxl.utils`
*   **被匯入 (Imported By)**: `ui.visualizer`, `ui.summary_window`, `core.worksheet_tree`, `core.data_processor`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **演算法**: `detect_rectangles` 函式中使用的演算法非常巧妙，但在處理極端情況（例如數萬個完全不相鄰的儲存格）時，其多層迴圈可能會有效能瓶頸。不過對於絕大多數真實場景，其效能是足夠的。

### 5. 待處理的 `import` 語句

*   無。

---

## `range_processor.py`

### 1. 總體功用 (Overall Purpose)

一個「範圍處理器」，專門處理公式中的範圍引用（如 `SUM(A1:B100)`）。它的核心職責是計算範圍的維度（大小），並為範圍內的**所有數據**計算一個「內容雜湊值 (Content Hash)」。這個雜湊值是一個獨一無二的指紋，可以用來在不逐一比對每個儲存格的情況下，快速判斷兩個不同位置的大範圍內容是否完全一致。

### 2. 詳細組件分析 (Detailed Component Analysis)

*   **Class: `RangeProcessor`**: 封裝了所有與範圍處理相關的邏輯。
*   `calculate_range_content_hash(...)`: 核心功能，透過讀取範圍內所有儲存格的內容來計算 SHA256 雜湊值。它還會對範圍內的內容進行統計，生成如「100數值, 50文字」這樣的摘要。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: `re`, `hashlib`, `openpyxl`, `os`
*   **被匯入 (Imported By)**: `utils.progress_enhanced_exploder`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **效能**: `calculate_range_content_hash` 函式每次處理一個新範圍時，都會重新從磁碟讀取整個 Excel 檔案。可以考慮改進 `RangeProcessor` 類別，讓它能夠快取已開啟的 `openpyxl` 工作簿物件，避免重複的檔案讀取。

### 5. 待處理的 `import` 語句

*   無。

---

## `safe_cache.py`

### 1. 總體功用 (Overall Purpose)

此模組提供了一個執行緒安全、高效能的記憶體快取系統，專門用於儲存和複用由 `openpyxl` 開啟的工作簿物件。它的存在是為了避免在程式的不同部分重複地從磁碟讀取同一個（可能很大的）Excel 檔案，從而極大地提升整體效能。

### 2. 詳細組件分析 (Detailed Component Analysis)

*   **Class: `SafeWorkbookCache`**: 核心快取類別，其設計非常周全。
    *   **LRU 策略**: 使用 `collections.OrderedDict` 實現了「最久未使用演算法 (LRU)」。當快取達到最大容量時，它會自動移除最久沒有被存取過的項目。
    *   **執行緒安全**: 使用 `threading.RLock()` 來確保在多執行緒環境下，對快取的讀寫操作不會發生衝突。
    *   **數據新鮮度保證**: 在返回一個快取的物件前，它會檢查原始檔案在磁碟上的「最後修改時間」。如果檔案已經被外部修改過，則快取會被判定為無效並被移除，強制重新從磁碟讀取最新版本。
    *   **記憶體管理**: 在從快取中移除一個工作簿物件時，會主動呼叫其 `close()` 方法並觸發垃圾回收，以幫助 Python 盡快釋放記憶體。
*   **`get_safe_global_cache()`**: 使用「單例模式 (Singleton Pattern)」來確保整個應用程式中只有一個快取實例，避免資源浪費。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: `os`, `time`, `threading`, `gc`, `collections`, `openpyxl`
*   **被匯入 (Imported By)**: `utils.openpyxl_resolver`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **優點**: 這是專案中另一個設計極其出色的模組，體現了對效能、執行緒安全和記憶體管理的深刻理解。

### 5. 待處理的 `import` 語句

*   **已修正**: `from openpyxl import load_workbook` 已被移至檔案頂部。