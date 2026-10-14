## spec 寫法（build_deck 唯一的輸入）

```json
{
  "title": "Q3 專案進度",                 // 也是檔名
  "audience": "executive",                // 見上方受眾
  "sources": ["data/Q3.xlsx"],            // 和 get_design_context 用的一樣，表格 id（t1…）才會一致
  "author": "專案管理部",                 // 選填：封面下方的署名
  "date": "2026/10/14",                   // 選填：封面日期
  "asset_dirs": ["assets", "D:/專案截圖"],  // 選填：collect_assets 找圖的資料夾
  "directive": "需求頁的整體說明（選填）",
  "facts": {"gap": "=t1[代號=B].目標進度 - t1[代號=B].實際進度"},   // 選填：具名數字
  "annotations": [{"page": 1, "raw": "需求頁的手寫", "applied": "怎麼套用"}],
  "slides": [ ... ]
}
```

### 每頁欄位
| 欄位 | 說明 |
|---|---|
| type | `"section"`＝章節頁（只要 title，可加 subtitle）；其他頁省略 |
| layout | 範本版型 key（見版型表）。**省略＝自動挑選**（最穩）。`"custom"`＝用 regions 照草稿排。`"已存版型/頁碼"`＝用已存版型＋fill |
| title | 頁標題。主管版請寫成「結論句」 |
| headline | 一句話結論，粗體放在內容區第一行（第 1 頁只有 title＋headline 時會用封面版型） |
| bullets | 條列：字串、`{"text":"…","level":1,"emphasis":"accent|alert"}`，或巢狀陣列＝下一層 |
| kpis | `[{"value":"{t1.交付率.mean}","label":"平均交付率","delta":"+3.2% vs 上季","good":"up","status":"注意"}]`；delta 以 +/− 或 ▲▼ 開頭會自動上綠／紅色，`good:"down"`＝越低越好；`status`（正常／注意／延遲）會改色條與數字顏色；`emphasis:"alert"` 數字變紅 |
| status | 紅黃綠燈：`[{"name":"供應鏈整合","status":"延遲","note":"供應商交期延誤 2 週","owner":"採購部 王經理"}]`，或從表格 `{"from":"t1","name":"專案","status":"狀態","note":"主要風險","owner":"部門","sort":"狀態"}`（依狀態排序＝延遲→注意→正常）；狀態可寫 正常/注意/延遲、綠/黃/紅 |
| source | 資料來源，顯示在頁尾左下（有資料的頁面建議都寫） |
| chart | 見下方 |
| table | 見下方 |
| image | 圖片／截圖：路徑字串，或物件（見下方「圖片與截圖」）；`.svg` 檔會轉成可編輯圖形 |
| images | 多張圖片的圖組：`{"items":[圖片…],"columns":3}` |
| gantt / timeline / process / svg | 甘特圖、時程表、流程、自己畫的 SVG（見下方「特殊圖表與繪圖」） |
| notes | 講者備註 |
| regions | layout 為 custom 時：`[{"type":"kpi","x":0,"y":1.5,"w":12,"h":2,"content":…,"directive":"手寫註解轉成的指示"}]` |
| fill | 用已存版型時：`{"r1": 內容, "r2": 內容}` |
| repeat | `{"over":"t1[狀態!=正常]","by":"代號"}`：每個值展開一頁；文字中 `{$key}`＝該值，`t1[@]`＝該值的資料列 |
| source_page | 來自草稿第幾頁 |
| annotations | `[{"raw":"手寫辨識原文","applied":"你怎麼套用"}]`：有手寫註解時必填 |
| page_type | `"requirements"`＝需求頁（不產生投影片） |

一頁只放 1 個 text（headline＋bullets）＋最多 1 個 kpis/status/chart/table/gantt/timeline/process/svg/image/images；需要更多區塊就改用 `custom`。

### 數字一律引用，不要自己算或抄
在任何文字中使用 `{…}`：
- `{t1[專案=B].交付率}`：單一儲存格（篩選後必須剛好 1 列）
- `{t1.營收.sum}`、`.mean` `.min` `.max` `.median` `.count` `.distinct` `.first` `.last`
- `{t1[部門=A][狀態=延遲].count}`：多重篩選；也可用 `!=` `>` `<` `>=` `<=` `~`（包含）
- `{t1.count}`：列數
- `{=t1[代號=B].目標進度 - t1[代號=B].實際進度}`：四則運算（欄名含空白時用反引號 `` `t1[代號=B].實際 金額` ``）
- `{gap}`：spec.facts 裡的具名數字
- 格式：`{… | .1%}`、`{… | ,.0f}`、`{… | +.1f}`（Python format）。百分比欄位預設就會顯示成 %
- 強調：`**強調色**`、`!!紅色警示!!`

### chart
- 從表格：`{"type":"column|bar|line|pie|stacked|stacked_bar|area|doughnut","from":"t1[部門=A]","category":"專案","values":["目標進度","實際進度"],"agg":"sum","sort":"-實際進度","limit":10,"title":"…","labels":true,"highlight":["供應鏈整合"],"highlight_color":"alert"}`
- `highlight`：要強調的類別（主色，其他變灰；`highlight_color:"alert"` 改紅色）；數列名稱含「目標／預算／計畫／去年」自動當灰色基準
- 自訂：`{"type":"line","categories":["Q1","Q2"],"series":[{"name":"營收","values":["{t2[季=Q1].營收}","{t2[季=Q2].營收}"]}]}`

### table
- 從表格：`{"from":"t1","columns":["專案","進度","狀態"],"sort":"-進度","limit":8,"formats":{"進度":".0%"},"highlight":{"column":"狀態","equals":"延遲"},"title":"表格小標（選填）"}`
- 自訂：`{"columns":["項目","說明"],"rows":[["A","…"]],"highlight_rows":[0]}`

### 特殊圖表與繪圖（都是原生 PowerPoint 圖形，可在 PowerPoint 裡再編輯）
- **gantt 甘特圖**：`{"today":"2026/10/14","tasks":[{"name":"主資料標準","start":"2026/09/01","end":"2026/11/20","progress":0.45,"status":"延遲","group":"資料中台","owner":"李承翰"}],"milestones":[{"name":"年度上線","date":"2027/01/20"}]}`
  - 也可從表格：`{"from":"t1","name":"工作項","start":"開始日","end":"到期日","progress":"完成度","status":"狀態","group":"專案","owner":"負責人","today":"today"}`（OpenProject 匯出的 CSV／Excel 直接可用）
  - progress 可寫 0.45、45% 或 45；status 延遲/注意 會把該列變紅/黃；group 自動分組並畫摘要括號；`today` 畫今天線
- **timeline 時程表**：`{"today":"2026/10/14","items":[{"date":"2026/06/01","label":"架構設計","note":"資安審查通過"},{"date":"2026/10/20","label":"財務模組上線","state":"current"}]}`
  - state：done 已完成／current 進行中／todo 未開始（有 today 會自動判斷）；status 延遲/注意 變紅/黃；`"when":"Q3"` 可取代日期顯示；`"even":true` 等距排列；5 個以上自動上下交錯
- **process 流程**：`{"steps":[{"title":"需求確認","note":"訪談 6 個事業部","owner":"採購部","state":"done"},{"title":"API 介接","state":"current","status":"延遲"}]}`
  - 沒寫 state 時用主色由深到淺；有 state 時 done＝主色、current＝輔色、todo＝灰
- **svg 自己畫**（架構圖、關係圖、矩陣、組織圖、示意圖…）：content 直接放 `<svg viewBox="0 0 870 390">…</svg>` 字串，或 `{"svg":"<svg…>","title":"…"}`、`{"path":"diagram.svg"}`
  - viewBox 用區塊實際大小（單位≈pt）：寬 ≈ w×73、高 ≈ h×59（w、h 是 12×8 格線大小；整頁內容區約 870×390），font-size 就等於 pt，建議 ≥12
  - 支援 rect（rx 圓角）、circle、ellipse、line、polyline、polygon、path、text、g（transform）、marker-end 箭頭、stroke-dasharray 虛線、opacity
  - **顏色用主題名稱**：`primary` `secondary` `alert` `good` `warn` `bad` `neutral-50…900` `chart-1…6` `white`（也接受 hex）；文字可用 `{t1…}` 引用數字
  - 不支援：濾鏡、遮罩、嵌入圖片、CSS `<style>`；text 不會自動換行（多行用多個 text 或 tspan dy）

### 圖片與截圖（原生裁切與效果，可在 PowerPoint 裡調整）
`{"path":"shots/op.png","style":"browser","url":"openproject.company.com/…","caption":"OpenProject 工作項目（10/14 擷取）","highlights":[{"x":0.14,"y":0.27,"w":0.84,"h":0.06,"note":"供應鏈可視化：完成度 52%"}],"zoom":{"x":0.6,"y":0.25,"w":0.3,"h":0.2},"crop":[0,0.05,0,0]}`
- style：`card`（預設：圓角＋細框＋柔和陰影）、`browser`（瀏覽器視窗外框＋網址列，適合系統截圖）、`shadow`、`plain`
- 截圖四周的單色白邊會自動裁掉（`"trim":false` 關閉）；`crop`：[左,上,右,下] 要裁掉的比例（0–1）或像素
- `highlights`：在截圖上框出重點（x,y,w,h 是原圖的比例 0–1 或像素），有 note 會自動編號並在右側列出說明
- `zoom`：把一小塊放大成子畫面（size＝子畫面寬度占圖片的比例，預設 0.42）
- `fit:"cover"`：填滿區塊（裁掉多餘部分）；解析度太低會提醒
- **還沒有圖時**：`{"need":"OpenProject 甘特圖截圖","search":["gantt","甘特"],"style":"browser"}` → collect_assets 會在素材資料夾找候選圖；直接 build 會放「待補圖片」框
- 要框重點前先用 `view_image` 看圖（有 0.1 格線），座標照格線讀

### region type（custom 用，座標是 12×8 格線，和草稿底圖相同）
`title`標題 · `headline`結論 · `bullets`重點 · `text`文字 · `kpi` · `chart`圖 · `table`表 · `status`狀態燈 · `gantt`甘特圖 · `timeline`時程表 · `process`流程 · `svg`示意圖 · `image`圖片／截圖 · `images`圖組 · `note`備註
content：title/headline/text/note → 字串；bullets → 陣列；kpi → kpis 陣列；chart/table/status/gantt/timeline/process/images → 同上物件；svg → SVG 字串；image → 路徑或物件

### 你不必處理的（build_deck 會自動做）
文字放不下會自動縮字（只會用字級表內的大小），仍放不下就拆成「（續）」頁；區塊多於版型內容區時自動改成自訂排版；字型、顏色、卡片、圖表與表格樣式、頁碼頁尾都由設計系統套用；大陸用語自動換成台灣用語。
