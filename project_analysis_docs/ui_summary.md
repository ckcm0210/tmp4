# `ui` 資料夾模組深度分析報告 (最終完整版)

## 引言

本文檔旨在對 `ui` (User Interface) 資料夾內所有「實際運作中的」Python 模組，進行深入且詳盡的功能、設計與互動分析。`ui` 資料夾是本專案的門面，它包含了所有使用者可見、可互動的視窗、按鈕和版面配置，是使用者體驗的核心。

---

## `modes/__init__.py`

### 1. 總體功用 (Overall Purpose)

此檔案將 `modes` 這個子資料夾定義為一個 Python「套件 (Package)」。這使得上層程式碼可以更優雅地引用其中的內容。它還明確定義了 `InspectMode` 為此套件的「公開介面 (Public API)」，這是一種良好的封裝實踐，能清晰地告訴其他開發者，這個套件主要是用來提供 `InspectMode` 這個功能的。

### 2. 詳細組件分析 (Detailed Component Analysis)

*   `from .inspect_mode import InspectMode`:
    *   **用途**: 這是 Python 的「相對匯入」。`.` 表示「從目前所在的資料夾」。這行程式碼的意思是：「從目前資料夾中的 `inspect_mode.py` 檔案裡，把 `InspectMode` 這個類別匯入進來」。
*   `__all__ = ['InspectMode']`:
    *   **用途**: 這是在定義當其他程式碼使用 `from ui.modes import *` 這種「萬用字元匯入」時，應該匯出哪些公開的名稱。這可以防止套件中未預期暴露的內部變數或函式被意外匯入，增加了程式碼的穩定性。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: `.inspect_mode` (匯入同資料夾下的 `inspect_mode.py`)
*   **被匯入 (Imported By)**: `main.py` (雖然 `main.py` 目前是直接匯入 `ui.modes.inspect_mode`，但 `__init__.py` 的存在使得未來可以將其簡化為 `from ui.modes import InspectMode`)

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **結構建議**: 這是 Python 套件管理的標準實踐，程式碼品質良好，無需修改。

---

## `modes/inspect_mode.py`

### 1. 總體功用 (Overall Purpose)

此模組定義了「檢查模式 (Inspect Mode)」的完整介面與邏輯。它提供了一個與「正常模式」完全不同的、更為簡潔和專注的介面。它的核心設計思想是「重用」而非「重寫」，它透過繼承 `WorksheetController` 來獲得基礎功能，然後以程式化的方式「隱藏」掉大部分在檢查模式中不需要的複雜 UI 元件（如篩選器、摘要按鈕等），最後再補上該模式專屬的功能（如「掃描選定儲存格」按鈕）。

### 2. 詳細組件分析 (Detailed Component Analysis)

*   **Class: `InspectMode`**: 
    *   **描述**: 一個簡單的啟動器類別。當 `main.py` 切換到檢查模式時，會建立這個類別的實例。它的唯一職責就是在 `__init__` 中建立 `InspectModeView`，從而啟動整個模式的 UI。
*   **Class: `InspectModeView`**: 
    *   **描述**: 負責搭建檢查模式的頂層 UI 骨架。它會建立一個可左右拖動的 `PanedWindow`，並在其中放入兩個 `SimplifiedWorksheetController`，從而實現雙窗格佈局。它還包含了顯示/隱藏右側面板的邏輯。
*   **Class: `SimplifiedWorksheetController(WorksheetController)`**: 
    *   **描述**: 這是此模組最核心的類別。它繼承了「正常模式」下的 `WorksheetController`，但對其進行了大量「客製化改造」以適應檢查模式。
    *   **方法**:
        *   `hide_unnecessary_elements()`: 透過遞迴遍歷所有子元件的方式，找出並隱藏（`grid_forget()` 或 `pack_forget()`）那些在檢查模式下不需要的按鈕和篩選框。
        *   `modify_layout_for_inspect_mode()`: 對佈局進行微調，例如調整列表高度，使其更緊湊。
        *   `add_scan_current_selection_button()`: 為此模式新增一個專屬的「掃描選定儲存格」按鈕。
        *   `scan_selected_cell()`: 這是點擊上述按鈕後的處理邏輯。它會連接到 Excel，獲取使用者當前選擇的儲存格，然後呼叫 `core.excel_scanner.refresh_data` 來只掃描這一個儲存格，實現快速分析。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: `tkinter`, `core.excel_scanner`, `core.worksheet_tree`, `ui.worksheet.controller`
*   **被匯入 (Imported By)**: `main.py`, `ui.modes.__init__`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **結構建議**: 「繼承後再隱藏」的作法雖然能快速重用程式碼，但它建立了一種脆弱的耦合關係。如果基礎 `WorksheetController` 的 UI 佈局在未來有較大改動（例如從 `grid` 改為 `pack`），這個模組的 `hide_unnecessary_elements` 方法就可能會失效。更穩健的長遠之計是「組合優於繼承」原則，可以建立一個不含 UI 的 `BaseWorksheetController` 來放置共享的業務邏輯，然後讓「完整版」和「簡化版」的控制器都去繼承它，並各自完全獨立地建立自己所需的 UI。

### 5. 待處理的 `import` 語句

*   此檔案的 `import` 語句在先前的操作中已全部修正。

---

## `summary_window.py`

### 1. 總體功用 (Overall Purpose)

此模組定義了 `SummaryWindow`，一個功能極其豐富的彈出視窗。它不僅僅是為了「顯示」摘要，而是整合了一整套圍繞「外部連結」的分析與處理工具，是專案中一個核心的互動中樞。

### 2. 詳細組件分析 (Detailed Component Analysis)

*   **Class: `SummaryWindow(tk.Toplevel)`**: 
    *   **描述**: 作為一個獨立的頂層視窗，它在被建立時就完成了大量的數據處理工作。
    *   **方法**:
        *   `__init__(...)`: 建構函式是本類別的核心。它接收從主面板傳來的公式數據，立即進行遍歷和正規表示式匹配，以找出所有獨一無二的外部連結。同時，它會建立一個「連結 -> 儲存格地址列表」的快取 (`link_to_addresses_cache`)，為後續的「Go to Excel」和「取代」功能提供數據支持。如果數據量過大，它還會顯示一個進度條，使用者體驗考慮周到。
        *   `show_summary_by_worksheet()` / `show_summary_by_workbook()`: 提供兩種不同的數據聚合視圖，讓使用者可以從「工作表」和「工作簿」兩個層級來審視外部連結。
        *   `browse_for_new_link()`: 提供一個檔案選擇對話框，輔助使用者建立符合格式的「新連結」字串。
        *   `on_link_select(...)`: 當使用者在列表中選擇一個連結時，此函式會被觸發，更新 UI 上的「舊連結」輸入框，並顯示該連結影響了多少個儲存格。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: `ui.visualizer` (呼叫視覺化圖表), `utils.excel_helpers` (呼叫取代和選取功能), `utils.range_optimizer` (呼叫範圍優化顯示)。
*   **被匯入 (Imported By)**: `core.worksheet_summary`, `ui.worksheet.controller` (由它們呼叫以彈出視窗)。

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **結構建議**: `__init__` 建構函式承擔了過多職責，包括 UI 建立、數據處理、快取建立等，應將其拆分為 `_setup_ui`, `_process_data`, `_bind_events` 等多個更小的私有方法。此外，「執行取代」按鈕的 `lambda` 函式傳遞了過多參數，顯示出高耦合，建議重構。

### 5. 待處理的 `import` 語句

*   無。

---

## `visualizer.py`

### 1. 總體功用 (Overall Purpose)

提供數據視覺化功能，使用 `matplotlib` 函式庫來繪製一個模擬的 Excel 工作表網格圖，並在高亮顯示受特定外部連結影響的所有儲存格，提供影響力「熱圖」。

### 2. 詳細組件分析 (Detailed Component Analysis)

*   **Class: `ChartVisualizer`**: 管理視覺化視窗和 `matplotlib` 圖表的主類別。
    *   `create_chart()`: 核心視覺化邏輯，使用 `matplotlib` 繪製矩形來代表儲存格和使用範圍。
    *   `export_chart()`: 提供將圖表儲存為圖片檔案的功能。
*   **Function: `show_visual_chart(...)`**: 從外部呼叫的啟動函式，負責建立 `ChartVisualizer` 實例。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: `tkinter`, `matplotlib`, `os`, `utils.range_optimizer`
*   **被匯入 (Imported By)**: `ui.summary_window`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **結構建議:** 程式碼結構良好，`ChartVisualizer` 類別很好地封裝了所有相關邏輯。

### 5. 待處理的 `import` 語句

*   無。

---

## `workspace_view.py`

### 1. 總體功用 (Overall Purpose)

提供一個「工作區管理器」UI，能列出所有當前開啟的 Excel 檔案，進行批次操作（如儲存、關閉、最小化、啟用），並能將當前檔案列表儲存為「工作區」以便後續重新載入。

### 2. 詳細組件分析 (Detailed Component Analysis)

*   **Class: `AccumulateListbox(tk.Listbox)`**: 一個自訂的 Listbox，支援更直觀的拖動選取功能。
*   **Class: `Workspace`**: 整個工作區頁面的主類別。
    *   `get_open_excel_files()`: 使用 `win32com` 獲取當前開啟的所有 Excel 檔案資訊。
    *   `show_names()`: 在獨立執行緒中執行 `get_open_excel_files` 以刷新列表，避免 UI 凍結。
    *   `save_workspace()` / `load_workspace()`: 實現儲存和載入工作區的核心功能。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: `tkinter`, `pythoncom`, `win32com.client`, `win32gui`, `win32con`, `time`, `openpyxl`, `os`, `datetime`, `threading`
*   **被匯入 (Imported By)**: `main`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **優點:** 正確地使用了多執行緒來處理耗時操作，並安全地將結果傳回主 UI 執行緒，這是非常好的實踐。
*   **結構建議:** `Workspace` 類別非常龐大，可以考慮將檔案操作等邏輯提取到獨立的輔助模組中。

### 5. 待處理的 `import` 語句

*   無。

---

## `worksheet/__init__.py`

### 1. 總體功用 (Overall Purpose)

將 `ui/worksheet` 資料夾標記為一個 Python 套件。

### 2. 詳細組件分析 (Detailed Component Analysis)

*   只包含一個文件字串，沒有可執行的程式碼。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: 無。
*   **被匯入 (Imported By)**: 多個模組（如 `core.formula_comparator`）從這個套件中匯入子模組。

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **結構建議:** 標準的 `__init__.py` 用法。

### 5. 待處理的 `import` 語句

*   無。

---

## `worksheet/controller.py`

### 1. 總體功用 (Overall Purpose)

MVC 架構中的「控制器」，負責管理單一工作表分析面板的狀態和業務邏輯。

### 2. 詳細組件分析 (Detailed Component Analysis)

*   **Class: `WorksheetController`**: 儲存所有與單一工作表相關的數據（如公式列表）和 UI 狀態（如篩選條件）。它實例化對應的 `WorksheetView` 和 `TabManager`。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: `tkinter`, `ui.worksheet.tab_manager`, `ui.worksheet.view`, `ui.summary_window`
*   **被匯入 (Imported By)**: `core.formula_comparator`, `core.dual_pane_controller`, `ui.modes.inspect_mode`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **結構建議:** 類別中狀態變數較多，未來可考慮將相關狀態組合到更小的資料物件中。

### 5. 待處理的 `import` 語句

*   此檔案的 `import` 語句在先前的操作中已全部修正。

---

## `worksheet/tab_manager.py`

### 1. 總體功用 (Overall Purpose)

定義 `TabManager` 類別，專門負責管理「詳細資訊」區域的多分頁介面，包括建立、關閉和狀態管理。

### 2. 詳細組件分析 (Detailed Component Analysis)

*   **Class: `TabManager`**: 封裝了所有與 `ttk.Notebook` 相關的操作。
    *   `create_detail_tab(...)`: 建立新分頁的核心方法。
    *   `_setup_tab_close_functionality(...)`: 為分頁設定多種關閉方式（右鍵、雙擊、中鍵），使用者體驗良好。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: `tkinter`
*   **被匯入 (Imported By)**: `ui.worksheet.controller`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **結構建議:** 程式碼結構良好，是一個很好的功能封裝範例。

### 5. 待處理的 `import` 語句

*   無。

---

## `worksheet/view.py`

### 1. 總體功用 (Overall Purpose)

MVC 架構中的「視圖」，負責建立和管理單一工作表分析面板的視覺 UI 元件。

### 2. 詳細組件分析 (Detailed Component Analysis)

*   **Class: `WorksheetView(ttk.Frame)`**: 作為一個 UI 元件的容器，但將所有具體的建立和綁定工作委派給了 `ui.worksheet_ui` 模組中的函式。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: `tkinter`, `ui.worksheet_ui`
*   **被匯入 (Imported By)**: `core.formula_comparator`, `ui.worksheet.controller`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **結構建議:** 將 UI 建立邏輯分離到 `worksheet_ui` 中是一種優秀的重構，使得 `WorksheetView` 類別非常乾淨。

### 5. 待處理的 `import` 語句

*   無。

---

## `worksheet_ui.py`

### 1. 總體功用 (Overall Purpose)

一個「UI 工廠」模組，提供一系列函式來建立和設定單一工作表面板的所有 UI 元件，並將它們的事件綁定到核心邏輯函式。

### 2. 詳細組件分析 (Detailed Component Analysis)

*   `create_ui_widgets(self)`: 建立所有 UI 元件的主函式。
*   `bind_ui_commands(self)`: 將 UI 元件的事件（如按鈕點擊）與 `core` 模組中的處理函式連接起來。
*   `_set_placeholder(...)` 等函式: 實現篩選輸入框中的佔位符文字功能。

### 3. 模組互動 (Interactions)

*   **匯入 (Imports)**: `tkinter` 以及大量來自 `core` 模組的函式。
*   **被匯入 (Imported By)**: `ui.worksheet.view`

### 4. 程式碼品質與改善建議 (Code Quality & Enhancement Suggestions)

*   **結構建議:** 將 UI 建立邏輯分離到此模組中是優秀的設計選擇。

### 5. 待處理的 `import` 語句

*   無。