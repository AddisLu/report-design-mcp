---
mode: agent
description: 依手繪草稿與資料產出 PowerPoint 報告
---
請用 ReportDesign MCP，依我的手繪草稿製作報告。

- 草稿：${input:sketch:草稿 PDF、圖片或資料夾路徑}
- 資料：${input:sources:資料檔或資料夾，多個用逗號分隔}
- 受眾：${input:audience:主管／經理／工程師}

流程：先呼叫 get_design_context（帶 sketch、sources、audience），接著寫出完整 spec（還沒有的圖寫 need），呼叫 collect_assets 列出版面計畫與待補素材給我確認；
素材齊全後 build_deck，用 preview_deck 看圖自檢、需要時 patch_deck，最後 review_deck 做美術總檢。
草稿上每一條紅筆註解都要寫進 annotations，完成後附上註解解讀表。
