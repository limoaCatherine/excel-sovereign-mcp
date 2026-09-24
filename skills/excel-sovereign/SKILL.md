---
name: excel-sovereign-mcp
description: 用 workbook_read、workbook_apply、excel_exec 修改本地 xlsx 和 xlsm。服务端选择引擎并在同一次调用里截图。
---

# excel-sovereign-mcp

一次调用带工作簿绝对路径和 `ops` 数组。不要传 `engine` 或 `session_id`。同一文件的修改合成一次调用。

## 调用前

- 只使用 `.xlsx` 和 `.xlsm`。加密、IRM、`.xls`、`.xlsb` 会被拒绝。
- 桌面上已经打开的同一文件先保存并关闭。
- 公式写英文函数名，例如 `=SUM(A1:A10)`。数字格式用 Excel 格式码，例如 `0.00`。
- 不要用区域格式去刷 Excel 表的表头或表体。加行用 `table_append`。

## 工具

- `workbook_read`：值、公式、缓存值。`includeStyles` 的字段和 `format` 相同。默认最多 4000 格，用返回的 `nextRange` 继续读。不截图。
- `workbook_apply` 和 `excel_exec`：同一份 `ops`。结构修改、表、透视表、图表、Power Query、数据模型、VBA 会整批走 Excel。调用返回前已经截图，不要再单独要图。

## 返回

- `committed: true` 表示文件已经保存。不要把同一批插入、删除、重命名再发一次。
- `verified: false` 时只发送 `suggestedNextActions` 里的 `verify`。
- `calculated: true` 只表示这次已经把重算后的缓存写回文件。只改字体、填充、行高列宽时 `calc` 为 `skipped`。

## 常用 ops

值：`{"action":"set_values","sheet":"Sheet1","range":"A1","values":[[1,"=A1+1"]]}`

公式：`{"action":"set_formulas","sheet":"Sheet1","range":"B1","formulas":[["=SUM(A1:A10)"]]}`

名称或 LAMBDA：`{"action":"define_name","name":"AddOne","formula":"=LAMBDA(x,x+1)"}`

排版：`{"action":"format","sheet":"Sheet1","range":"A1:B1","bold":true,"fontColor":"#1F4E79","fillColor":"#D9D9D9","numberFormat":"0.00"}`

新表用 `layout`，不要手写上面的颜色。`profile` 只表示起稿习惯：`finance`（参数、勾稽、清单）、`analytics`（参数、事实表、清单）、`general`（游戏、制造或其他行业，块自选）。颜色不随行业变化。块有 `title`、`section`、`subsection`、`parameters`、`catalog`、`formula`、`fact`。`fact` 会建成 Excel 表，不要再对它发 `format`。改旧表时按现有的块改值，不把整张表重刷成另一个行业。

```json
{"action":"layout","sheet":"假设","profile":"finance","blocks":[
  {"kind":"title","text":"收入预测"},
  {"kind":"parameters","columns":[
    {"header":"项目","role":"label"},
    {"header":"值","role":"percent","input":true},
    {"header":"单位","role":"unit"},
    {"header":"说明","role":"text"}
  ],"rows":[["收入增速",0.08,"%",""],["税率","=利润!C3","%","跨表"]]}
]}
```

插行：`{"action":"insert_rows","sheet":"Sheet1","row":2,"count":1}`

重命名：`{"action":"rename_sheet","oldName":"Sheet1","newName":"参数"}`

## 阶段 4 动作

这些动作的其余参数放在 `args` 里，名字与 excelcli 一致：

- 表：`table_list` `table_create` `table_append` `table_resize` `table_rename` `table_delete` `table_set_style` `table_apply_filter` `table_clear_filters`
- 透视表：`pivot_list` `pivot_create_from_range` `pivot_create_from_table` `pivot_refresh` `pivot_delete`
- 图表：`chart_list` `chart_create_from_range` `chart_create_from_table` `chart_move` `chart_fit` `chart_delete`
- Power Query：`powerquery_list` `powerquery_create` `powerquery_update` `powerquery_refresh` `powerquery_refresh_all` `powerquery_delete`
- 数据模型：`table_add_to_data_model` `table_create_from_dax` `datamodel_list_tables` `datamodel_list_measures` `datamodel_create_measure` `datamodel_update_measure` `datamodel_evaluate` `datamodel_refresh`
- 条件格式：`conditional_format_add` `conditional_format_clear` `conditional_format_list`
- 数据验证：`validation_add` `validation_remove` `validation_get`
- 批注：`comment_set` `comment_get` `comment_clear` `threaded_comment_add`
- 超链接：`hyperlink_add` `hyperlink_remove`
- 冻结和隐藏：`freeze` `unfreeze` `sheet_hide` `sheet_show`
- VBA：`vba_list` `vba_view` `vba_import` `vba_update` `vba_run` `vba_delete`。只有这批调用含 VBA 时才允许宏运行。Excel 需要信任对 VBA 工程对象模型的访问。

单次写入超过 10 万格会返回 `too_large`，不会截断保存。
