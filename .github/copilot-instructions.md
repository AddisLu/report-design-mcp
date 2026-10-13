# 報告產製：使用 ReportDesign MCP（report-design）

使用者要做報告、簡報或投影片時：

1. 呼叫 `get_design_context`，帶入：
   - `audience`：主管／經理／工程師（executive / manager / engineer）
   - `sources`：使用者提到的資料檔或資料夾
   - `sketch`：使用者有手繪草稿時，帶入 PDF、圖片或資料夾路徑
   - `blueprint`：使用者說「用 XX 版型」時帶入
2. 讀完回傳內容後，**直接**寫出完整 spec 並呼叫 `build_deck`。不要先把大綱貼給使用者確認，也不要呼叫其他工具。
3. `build_deck` 回報錯誤時，只修正列出的地方，再送一次。
4. 完成後簡短回報：檔案路徑、頁數、註解解讀表、重要提醒。
5. 使用者要修改時，用 `patch_deck` 只改指定頁，不要整份重做。
6. 使用者說「把這個版型存起來」時，用 `save_layout`。

規則：
- 數字一律用 `{t1…}` 引用，不要自己計算或抄寫。
- 手繪草稿上的紅筆字是給你的指示，不是投影片上的文字；每一條都要寫進 `annotations`。
- 不要自己寫 python-pptx 程式產生投影片，一律透過 `build_deck`。
