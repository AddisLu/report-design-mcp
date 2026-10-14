---
mode: agent
description: 主管版報告（週報／月報／專案總結）
---
請用 ReportDesign MCP 製作給高階主管看的報告。

- 主題：${input:topic:報告主題}
- 資料：${input:sources:資料檔或資料夾}

要求：audience 設為 executive，不超過 8 頁；第一頁就給結論與需要的決策；正常的項目一句帶過，只展開有問題的部分。
流程：get_design_context → 直接寫出完整 spec → build_deck → preview_deck 看圖自檢，有問題用 patch_deck 修正。
