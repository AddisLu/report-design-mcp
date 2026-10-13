# 公司範本放這裡

把公司的 `.pptx` 或 `.potx` 範本放在這個資料夾。系統會自動使用第一個範本（依檔名排序），也可以在 `design_system/config.json` 的 `template` 指定。

換上新範本後，建議執行一次：

```
python cli.py inspect design_system/templates/公司範本.pptx --set-default
```

執行後會更新 `design_system/layout_catalog.json`。你可以手動修改其中每個版型的 `use`（用途說明），模型會依這些說明挑選版型。

還沒放範本時，系統會自動產生並使用 `_builtin_16x9.pptx`。
