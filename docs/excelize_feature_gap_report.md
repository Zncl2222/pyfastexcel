# Excelize 功能盤點與 pyfastexcel 封裝建議

盤點日期：2026-07-17

## 摘要

pyfastexcel 目前已經涵蓋高速產生 Excel 報表所需的主要骨架，包括一般寫入、串流寫入、樣式、工作表操作、表格、圖表與樞紐分析表。它目前最合適的定位不是 Excelize 的完整 Python binding，而是：

> 以高效能寫出為核心的 Excel 商務報表產生器。

基於這個定位，後續應優先補齊報表輸出時常用、能維持 write-only 與 streaming 架構的功能，而不是立即擴張到讀取或修改既有活頁簿。

目前專案使用 Excelize v2.9.0，官方最新穩定版為 v2.11.0。由於專案已使用 Go 1.25.11，符合 Excelize v2.11.0 的 Go 版本要求，建議先評估升級上游依賴，再增加新功能。

## 目前已有功能

pyfastexcel 現有封裝包含：

- 一般寫入與 StreamWriter 高速寫入
- 多工作表建立、刪除、重新命名與切換
- 儲存格值、公式與樣式
- 欄寬與列高
- 合併儲存格
- 工作表隱藏
- Freeze/Split panes
- AutoFilter
- Data Validation
- Comment
- Row/Column grouping
- Table
- Chart 與組合圖
- Pivot Table
- Workbook properties
- Workbook protection
- ZIP 壓縮層級及多工作表平行寫出

這個涵蓋範圍已超過單純的 StreamWriter binding；主要缺口集中在列印、視覺提示、連結及工作表保護等報表功能。

## 建議優先級

| 功能 | 使用價值 | 實作難度 | Streaming 相容性 | 建議優先級 |
| --- | --- | --- | --- | --- |
| 升級 Excelize v2.11.0 | 非常高 | 中 | 高 | P0 |
| Conditional Formatting | 非常高 | 中 | 高 | P0 |
| Hyperlink | 高 | 低 | 中至高 | P0 |
| Header / Footer | 高 | 低 | 高 | P0 |
| Page Layout / Margins | 高 | 低至中 | 高 | P0 |
| Worksheet Protection | 高 | 中 | 高 | P0 |
| Auto-fit Column Width | 高 | 中 | 視實作而定 | P1 |
| Defined Names / Print Area | 中高 | 低至中 | 高 | P1 |
| Picture / Image | 高 | 中至高 | 中 | P1 |
| Cell Rich Text | 中高 | 中 | 高 | P1 |
| Sparklines | 中高 | 中 | 高 | P1 |
| Sheet View / Properties | 中 | 低 | 高 | P1 |
| Chart / Pivot 新版欄位 | 中 | 中 | 高 | P1 |
| Shape / Form Controls | 低至中 | 高 | 中 | P2 |
| Slicer | 中 | 高 | 中 | P2 |
| VBA Project | 低至中 | 高 | 低至中 | P2 |

## 第一階段：最值得加入的功能

### 1. Conditional Formatting

這是目前最明顯的報表功能缺口。Excelize 支援的規則包括：

- 儲存格比較條件
- 文字與時間條件
- 重複值與唯一值
- Top / Bottom
- 空白與錯誤值
- 自訂公式
- Two/Three Color Scale
- Data Bar
- Icon Set

建議先支援最常用的 `cell`、`formula`、`duplicate`、`unique`、color scale、data bar 和 icon set。

建議 API：

```python
ws.add_conditional_format(
    "B2:B100",
    type="cell",
    criteria=">",
    value=100,
    style=CustomStyle(font_color="FF0000"),
)
```

此功能符合報表定位，也能沿用現有的 StyleManager。

### 2. Hyperlink

目前 URL 只能作為一般文字寫入。建議同時支援外部網址與活頁簿內部位置：

```python
ws.add_hyperlink(
    "A1",
    "https://example.com",
    display="Official site",
    tooltip="Open website",
)

ws.add_hyperlink("B1", "Sheet2!A1", link_type="Location")
```

Hyperlink 實作範圍小、使用率高，可以在串流內容 flush 後統一加入。

### 3. Header、Footer 與列印設定

商務報表常見但目前缺少的項目包括：

- Header/Footer
- 頁碼、總頁數、日期、檔名與工作表名稱
- Portrait/Landscape
- Paper size
- Fit to width/height
- Print gridlines/headings
- Page margins
- Print area
- Repeat rows/columns
- Page breaks

建議 API：

```python
ws.set_page_layout(
    orientation="landscape",
    paper_size=9,
    fit_to_width=1,
    fit_to_height=0,
)
ws.set_page_margins(left=0.25, right=0.25, top=0.5, bottom=0.5)
ws.set_header_footer(odd_header="&CMonthly Report", odd_footer="&CPage &P of &N")
ws.set_print_area("A1:H200")
ws.set_print_titles(rows="1:2")
```

Print Area 與 Print Titles 在 Excelize 底層使用 defined names，因此可以同時提供通用的 `define_name()` API。

### 4. Worksheet Protection

目前已有 Workbook Protection，以及 style 層級的 `locked`/`hidden`，但沒有真正保護工作表。這會讓 locked style 無法獨立產生完整效果。

建議 API：

```python
ws.protect(
    password="secret",
    select_locked_cells=True,
    select_unlocked_cells=True,
    format_cells=False,
    insert_rows=False,
)
```

這是既有 style protection 語意的自然補齊。

### 5. Auto-fit Column Width

Excelize v2.11.0 新增 `AutoFitColWidth`。建議提供：

```python
ws.auto_fit_columns("A:D")
```

此功能應設為 opt-in，避免大型串流報表因事後掃描資料而出現額外的時間與記憶體成本。另一種較符合 pyfastexcel 定位的做法，是在 append row 時同步維護每欄最大顯示寬度，最後一次設定所有欄寬。

## 第二階段：視覺與進階報表功能

### 6. Picture / Image

常見用途包括公司 Logo、QR Code、簽名與報表截圖。建議同時接受檔案路徑與 bytes：

```python
ws.add_picture("A1", "logo.png", scale_x=0.5, scale_y=0.5)
ws.add_picture_bytes("D5", image_bytes, extension=".png")
```

路徑模式較簡單，但需注意跨程序路徑及檔案生命週期；bytes 模式較穩定，卻需要擴充 wire protocol。大型圖片宜使用獨立 binary frame，不應塞入 metadata JSON。

### 7. Cell Rich Text

目前 Comment 與 Chart 已有 rich text 概念，但一般儲存格沒有對應封裝。建議將 RichTextRun 抽為共用模型：

```python
ws.set_rich_text(
    "A1",
    [
        RichTextRun(text="Revenue: ", bold=True),
        RichTextRun(text="$100", color="00AA00"),
    ],
)
```

### 8. Sparklines

Sparkline 比完整 Chart 輕量，適合在大量資料列中顯示趨勢：

```python
ws.add_sparkline(
    location="H2:H100",
    data_range="B2:G100",
    type="line",
    show_markers=True,
)
```

### 9. Sheet View / Properties

建議補充：

- Zoom scale
- Show gridlines
- Show row/column headers
- Show zeros
- Right-to-left
- Top-left visible cell
- Page layout view
- Default row height
- Tab color
- Code name

建議 API：

```python
ws.set_view(zoom=90, show_grid_lines=False)
ws.set_properties(tab_color="4472C4", default_row_height=18)
```

### 10. Chart 與 Pivot 新版能力

Excelize v2.10 至 v2.11 新增或補強了以下功能。

Chart：

- Data point formatting
- Marker border
- Legend font
- Per-series legend
- Line dash/fill
- Transparency
- Drop lines / High-low lines
- Formula-based chart title
- Title layout 與 line formatting
- Stock charts

Pivot：

- `ShowValuesAs`
- `SelectedItems`
- Classic layout 等新版欄位
- Slicer

其中 Pivot `ShowValuesAs` 和 Chart title 是較值得優先補齊的項目。

## 第三階段：有明確需求再做

以下功能由 Excelize 支援，但不建議目前優先實作：

- Shape
- Form Controls
- Slicer
- VBA Project
- Worksheet background image
- Chart sheet
- Custom properties

這些功能通常需要較多模型、binary 資料傳輸或跨軟體相容性測試，對一般資料匯出場景的使用率則較低。

## 暫不建議：讀取與修改既有 Excel

目前架構大致是：

```text
Python 建立描述資料
        ↓
Wire protocol
        ↓
Go 建立全新 workbook
        ↓
輸出 bytes/file
```

如果加入 `Workbook.open()`、讀取儲存格或修改既有活頁簿，將需要處理：

- 原始 workbook binary 傳送
- 未修改 OOXML 元件的完整保留
- VBA、external links、images 等 round-trip
- Python 端 lazy access
- Streaming read API
- 不受信任檔案的資源限制與安全問題

這會把套件從高速 writer 變成完整 Excel engine binding。除非產品方向確定需要 template editing，否則暫時不建議優先包裝：

- GetCellValue
- Rows/Cols iterator
- SearchSheet
- Formula calculation
- GetChart/GetPicture/GetPivotTable
- 各種以讀取既有檔案為前提的 get/delete API

## Excelize 升級評估

專案目前依賴：

```go
github.com/xuri/excelize/v2 v2.9.0
```

建議評估直接升級至 v2.11.0，而不是只升到 v2.10.1，原因包括：

- v2.11.0 含安全修正。
- 新增 `AutoFitColWidth`。
- Pivot 支援 `ShowValuesAs` 與 selected items。
- Chart title 支援公式與更多 layout。
- 專案目前 Go 版本已符合 v2.11.0 要求。
- v2.10 已修正 Pivot Table 在 Excel for Mac 可能損毀的問題。

但 v2.11.0 存在 Chart breaking changes：

- `Chart.Title` 改為 `ChartTitle`
- `ChartLineType` 改為 `LineType`
- `ChartLine` 改為 `LineOptions`

因此升級不應只修改 `go.mod`，還需要同步調整 Go chart mapping，並執行 chart、pivot、streaming 及輸出檔案回歸測試。

## 建議 Roadmap

### Phase 0：上游升級

1. Excelize v2.9.0 升至 v2.11.0。
2. 調整 Chart breaking changes。
3. 執行 Workbook、Chart、Pivot 與 StreamWriter regression tests。
4. 使用 Excel、LibreOffice 或 WPS 進行實際開檔驗證。

### Phase 1：報表基本能力

1. Conditional Formatting
2. Hyperlink
3. Worksheet Protection
4. Header/Footer
5. Page Layout/Margins
6. Print Area/Print Titles/Page Breaks
7. Sheet View
8. Auto-fit Columns

### Phase 2：視覺內容

1. Pictures：path 與 bytes
2. Cell Rich Text
3. Sparklines
4. Chart v2.11 欄位補齊
5. Pivot `ShowValuesAs`

### Phase 3：進階 Excel 元件

1. Slicer
2. Shape
3. Form Controls
4. VBA Project
5. Custom Properties

## 最終建議

如果近期只能選五個工作項目，建議依序為：

1. 升級 Excelize v2.11.0。
2. Conditional Formatting。
3. Page Layout、Header/Footer 與 Print Area。
4. Hyperlink。
5. Worksheet Protection。

完成以上項目後，pyfastexcel 會從高速資料寫出工具，進一步成為功能完整度相當高的高速商務報表產生器，同時仍能保持目前最有辨識度的效能定位。

## 參考資料

- [Excelize 官方文件](https://xuri.me/excelize/en/)
- [Excelize Workbook API](https://xuri.me/excelize/en/workbook.html)
- [Excelize Conditional Formatting API](https://xuri.me/excelize/en/utils.html)
- [Excelize Picture API](https://xuri.me/excelize/en/image.html)
- [Excelize v2.10.0 Release Notes](https://xuri.me/excelize/en/releases/v2.10.0.html)
- [Excelize v2.11.0 Release Notes](https://xuri.me/excelize/fr/releases/v2.11.0.html)
