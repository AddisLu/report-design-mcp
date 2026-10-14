## 手繪草稿判讀規則
- 每張圖＝草稿一頁。底圖是 12×8 格線（等比例對應投影片）：讀框線座標 x 0–12、y 0–8，精度 0.5 即可。
- **黑色（深色）筆**：版面框線，以及框內寫的類型。對照表：
  | 框內寫 | region type | 說明 |
  |---|---|---|
  | 標題、結論、重點、文字、簡介、說明、備註 | title / headline / bullets / text / note | 「簡介」＝2–3 句精簡文字（text），不要寫成長條列 |
  | KPI、數字、指標 | kpi | |
  | 圖、長條、折線、圓餅 | chart | 依註解或資料選圖型 |
  | 表 | table | |
  | 燈號、紅綠燈、狀態 | status | |
  | 甘特圖、Schedule、排程 | gantt | 有 OpenProject 匯出表就用 `from` |
  | 時程、時間軸、里程碑 | timeline | |
  | 流程、步驟 | process | |
  | 示意圖、架構圖、關係圖、矩陣 | svg | 你自己畫（見 spec 寫法的 svg） |
  | 截圖、畫面、圖片、照片 | image | 草稿沒指定檔案時寫 `{"need":…,"search":[…]}`，交給 collect_assets 找 |
  | 多張照片、圖組 | images | |
- **紅色（或其他彩色）筆**：使用者的註解，屬於「指示」，**不是投影片上的文字**。
  - 指向某個框 → 寫進該 region 的 `directive`，並反映在內容上（選哪些資料、強調什麼、用什麼圖）
  - 針對整頁 → 寫進 slide 的 `directive`
  - 「標紅」「強調」→ 對應文字用 `!!…!!` 或 `**…**`，KPI 用 `"emphasis":"alert"`
  - 只有加了引號或寫了「原文」的註解，才把文字原樣放進投影片
  - 「找圖：XXX」「貼 XXX 的圖」「去 XX 資料夾找」→ 該 image 區塊寫 `{"need":"XXX","search":["關鍵字",…]}`（資料夾放進 spec.asset_dirs）
  - 「截 OpenProject 的 XX 畫面」→ `{"need":"OpenProject XX 畫面截圖","search":["openproject","XX"],"style":"browser"}`
  - 「框出 XX」「圈起來」→ 之後拿到圖時，用 view_image 看格線，把位置寫進 highlights
  - 「只給簡介」→ text 區塊 2–3 句，不加條列；「這邊放大」→ zoom
  - 配色、風格類的註解（「整體用藍色系」「不要太花」）→ 寫進 spec.directive，完稿時用 review_deck 檢查
- **頁面類型**（看左上角）：
  - 「需求」頁：不產生投影片，內容寫進 spec.directive，並用來決定受眾、資料、重點；手寫內容記到 spec.annotations
  - 「×每個XX」「×N」：重複頁 → `repeat: {"over": 表格, "by": 欄位}`
  - 其他＝版面頁：一頁草稿一張投影片
- **順序**：依右上角頁碼（P__）排序；有寫「插在 P2 後」就照做；沒寫就照草稿順序。
- 每張 slide 填 `source_page`；**每條手寫註解都要寫進 `annotations`**：`{"raw": 辨識到的手寫原文, "applied": 你怎麼套用}`。辨識沒把握時在 raw 後面加「(?)」。
- 框線和範本某個版型很接近時，可以直接用那個 layout；否則用 `layout: "custom"` 加 regions，座標照草稿。
- 框裡只寫類型、沒寫內容時，依資料與受眾自己寫內容；數字一律用 `{引用}`。

### 有草稿時的流程
1. 讀完草稿寫出完整 spec（素材還沒有的寫 need）。
2. 呼叫 **collect_assets**：回傳版面計畫表、待補素材與候選圖。把清單給使用者確認，選好的圖填 path，缺的請使用者提供。
3. 素材齊全（或使用者說先做）→ build_deck → preview_deck 自檢 → patch_deck。
4. 最後 **review_deck** 做整份美術總檢（配色、字型、對齊、節奏）。
