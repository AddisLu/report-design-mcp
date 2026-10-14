---
mode: agent
description: 經理版專案檢討（進度、資源、風險、對策）
---
請用 ReportDesign MCP 製作給部門經理看的專案檢討。

- 專案／範圍：${input:topic:專案或範圍}
- 資料：${input:sources:資料檔或資料夾}

要求：audience 設為 manager。總覽頁放進度燈號；問題專案每案一頁（用 repeat）；每個問題都要寫對策、負責人與期限。
流程：get_design_context → 直接寫出完整 spec → build_deck → preview_deck 看圖自檢，有問題用 patch_deck 修正。
