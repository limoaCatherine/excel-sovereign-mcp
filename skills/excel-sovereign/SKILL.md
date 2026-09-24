---
name: excel-sovereign-mcp
description: 用六个短入口修改本地 xlsx 和 xlsm。服务端选择引擎，并在同一次调用里保存和截图。
---

# excel-sovereign-mcp

一次调用带工作簿绝对路径和 `ops` 数组。不要传 `engine` 或 `session_id`。同一文件的修改合成一次调用。

## 调用前

- 只使用 `.xlsx` 和 `.xlsm`。加密、IRM、`.xls`、`.xlsb` 会被拒绝。
- 桌面上已经打开的同一文件先保存并关闭。
- 公式写英文函数名，例如 `=SUM(A1:A10)`。数字格式用 Excel 格式码，例如 `0.00`。
- 不要用区域格式去刷 Excel 表的表头或表体。加行用 `table_append`。

## 工具

六个入口。参数都是工作簿路径和 `ops`，不要把别的模块的动作塞进同一次调用。同一模块、同一文件合成一次。服务端选择引擎。调用返回前已经截图，不要再单独要图。

- `workbook_read`：值、公式、缓存值。默认最多 4000 格，用 `nextRange` 继续。不截图。
- `workbook_apply`：值、公式、名称、排版、工作表结构、`layout`。
- `excel_table`：表、透视表、图表、切片器。
- `excel_model`：Power Query、数据模型、DAX。
- `excel_view`：条件格式、数据验证、批注、超链接、冻结、隐藏工作表。
- `excel_vba`：VBA。只有这次调用才允许宏。Excel 需要信任对 VBA 工程对象模型的访问。

动作名在对应工具的说明里。其余字段直接写在 op 上。送错工具会返回 `wrong_tool`，并给出该用的工具名。

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

单次写入超过 10 万格会返回 `too_large`，不会截断保存。
