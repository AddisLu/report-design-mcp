# ReportDesign MCP：讓 Copilot 照公司範本產出像樣的 PowerPoint 報告

像 Claude Design 一樣運作，有三個重點：**設計系統**（範本＋受眾規則＋文風），**結構化產出**（模型只寫內容，不碰版面），以及**手繪草稿當需求**。可以掛在 VS Code 的 GitHub Copilot（Agent 模式）上使用，可在 **Windows 11＋Python 3.8** 執行。

## 為什麼直接叫 Sonnet 寫會不像樣
一次 prompt 就要模型同時讀資料、算數字、寫內容、排版面，每一件都只做到一半。這套工具把工作拆開：

| 工作 | 誰負責 |
|---|---|
| 讀 Excel／CSV／Word／PPT／PDF、算數字 | Python（數字用 `{t1…}` 引用，模型無法捏造） |
| 依受眾決定重點與故事線、寫結論句 | 模型（受 `audiences/*.yaml` 規則約束） |
| 版面、字型、顏色、位置 | 公司範本（模型不碰座標） |
| 文字放不下、用語、空泛詞、數字來源 | `build_deck` 自動修正或提醒 |

整個流程只有 **2 次 tool call**：`get_design_context` → `build_deck`。

## 安裝（Python 3.8.20）
```bat
cd C:\Coding\report-design-mcp
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\python tests\smoke_test.py
```

## 掛到 GitHub Copilot（VS Code）
1. 用 VS Code 開啟這個資料夾。`.vscode/mcp.json` 已經設定好 `report-design` server。
2. Copilot Chat 切到 **Agent** 模式，在工具清單中確認有 `report-design` 的 7 個工具。
3. `.github/copilot-instructions.md` 會自動套用工作流程；`.github/prompts/` 裡有 3 個一鍵 prompt：`/report-from-sketch`、`/exec-report`、`/project-review-mgr`。

想在其他專案資料夾也能使用：把 `.vscode/mcp.json` 的路徑改成絕對路徑，例如 `C:/Coding/report-design-mcp/.venv/Scripts/python.exe`，再放到該專案的 `.vscode/` 或 VS Code 使用者設定中。

> 公司若用 policy 停用了 MCP：同一套引擎可以改用 `cli.py`，讓 Copilot 在終端機執行（見下方 CLI）。

## 第一次設定：換成公司範本
```bat
copy 公司範本.pptx design_system\templates\
.venv\Scripts\python cli.py inspect design_system\templates\公司範本.pptx --set-default
```
接著打開 `design_system/layout_catalog.json`，調整每個版型的 `use`（用途說明），模型會依這些說明挑選版型。

還可以依公司習慣修改：
- `design_system/audiences/*.yaml`：各層級的頁數上限、條列數、重點、要避免的內容、語氣
- `design_system/style_guide.md`：公司文風（模型會讀）
- `design_system/style_rules.yaml`：自動替換的用語、空泛詞、禁用詞
- `design_system/examples/`：放 1–2 份「好報告」的 spec 當範例，效果最好

## 用手寫板畫草稿
1. 產生底稿：在 Copilot 說「給我 8 頁草稿底稿」，或執行：
   ```bat
   .venv\Scripts\python cli.py canvas sketches\canvas.pdf --pages 8
   ```
   第 1 頁是**需求頁**，後面是 12×8 格線的**版面頁**。把 PDF 匯入手寫板 App 當底圖。
2. 畫法：
   - **黑筆**畫框，框內寫類型：`標題`、`結論`、`KPI`、`圖`（可加註「長條／折線」）、`表`、`重點`、`圖片`、`備註`
   - **紅筆**寫註解，例如「延遲的標紅」「只講有問題的」「用折線圖」。這些是給模型的指示，**不會印到投影片上**；要原樣放進投影片的文字請加引號或寫「原文」
   - 左上角寫頁面類型：`需求`／`×每個專案`（重複頁，依資料每個專案展開一頁）
   - 右上角寫頁碼 `P3`；也可以寫「插在 P2 後」
3. 匯出成**一份多頁 PDF**（或放圖片的資料夾），在 Copilot 說：
   > 照 sketches\q3.pdf 的草稿，用 data\Q3.xlsx 做給主管看的報告
4. 完成後 Copilot 會回報檔案路徑和**註解解讀表**（每條手寫註解被辨識成什麼、怎麼套用）。有讀錯的地方，直接說「第 2 頁那條是『延遲』不是『延期』」，它會用 `patch_deck` 只改那一頁。
5. 滿意的版型說「把這份草稿版型存成『月報-主管版』」，下次說「用月報-主管版」就不必再畫。

## 工具一覽
| 工具 | 用途 |
|---|---|
| `get_design_context` | 一次取得受眾規則、文風、版型、已存版型、資料摘要、spec 寫法、範例、草稿圖片 |
| `build_deck` | spec → 檢查＋自動修正＋產出 .pptx（同時存 .spec.json） |
| `patch_deck` | 只改指定頁，再重新產出 |
| `read_sketch` | 草稿頁數很多時分批讀取 |
| `get_sketch_canvas` | 產生格線底稿 PDF／PNG |
| `save_layout` | 把草稿版型存成可重複使用的版型 |
| `inspect_template` | 解析新範本、設為預設 |

## 資料來源
`.xlsx` `.xlsm` `.csv` `.tsv`（表格 → `t1`、`t2`…，可用 `{}` 引用）；`.docx` `.pptx` `.pdf` `.md` `.txt`（文字 → `d1`…；其中的表格與 PPT 圖表資料也會轉成 `t#`）。也可以直接給資料夾。CSV 會自動判斷 UTF-8／Big5（cp950）編碼。

## CLI（不經過 Copilot）
```bat
.venv\Scripts\python cli.py context --audience 主管 --sources data\Q3.xlsx --sketch sketches\q3.pdf
.venv\Scripts\python cli.py build design_system\examples\executive_example.json -o output\
.venv\Scripts\python cli.py patch output\Q3重點專案進度_executive.spec.json patch.json
```

## 第二版規劃
- OpenProject 資料來源（API v3：work packages、進度、逾期）
- `preview_deck`：透過 PowerPoint COM 匯出 PNG，讓模型看圖微調
