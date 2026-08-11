# Power Query 手动重建指南

> 如果 Excel 文件中的 Power Query 连接失效，按以下步骤手动重建。每条 M 代码可直接复制粘贴。

---

## 前提条件

- Excel 2016+ 或 Microsoft 365（内置 Power Query）
- 数据文件位于 `data/` 目录下
- 所有 CSV 使用 UTF-8 BOM 编码（Excel 可正确识别中文）

---

## 步骤 1：导入销售出库数据

1. **数据** → **从文本/CSV** → 选择 `data/sales_orders.csv`
2. 预览确认中文正常显示 → **转换数据**（进入 Power Query 编辑器）
3. 在 Power Query 编辑器中：
   - 选中「开单日期」列 → **转换** → **数据类型** → **日期**
   - 选中「到期日」列 → **转换** → **数据类型** → **日期**
   - 选中「数量」「单价」「金额」「折扣率」「实收金额」→ **数据类型** → **小数**
4. **主页** → **关闭并上载至** → **表** → 目标位置：`📦 销售出库!$A$1`

### M 代码（高级编辑器直接粘贴）

```m
let
    Source = Csv.Document(File.Contents("C:\Users\...\data\sales_orders.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {
        {"开单日期", type date}, {"到期日", type date},
        {"数量", Int64.Type}, {"单价", type number}, {"金额", type number},
        {"折扣率", type number}, {"实收金额", type number}, {"账期天数", Int64.Type}
    })
in
    #"Changed Type"
```

---

## 步骤 2：导入银行回款数据

同上，选择 `data/bank_receipts.csv`。

### M 代码

```m
let
    Source = Csv.Document(File.Contents("C:\Users\...\data\bank_receipts.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {
        {"到账日期", type date}, {"金额", type number}
    })
in
    #"Changed Type"
```

---

## 步骤 3：导入 Sell-Through 数据

选择 `data/dealer_sell_through.csv`。

### M 代码

```m
let
    Source = Csv.Document(File.Contents("C:\Users\...\data\dealer_sell_through.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {
        {"报表月份", type date}, {"销售数量", Int64.Type}, {"单价", type number}, {"销售金额", type number}
    })
in
    #"Changed Type"
```

---

## 步骤 4：导入经销商主数据

选择 `data/dealer_list.csv`。

### M 代码

```m
let
    Source = Csv.Document(File.Contents("C:\Users\...\data\dealer_list.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {
        {"账期天数", Int64.Type}
    })
in
    #"Changed Type"
```

---

## 步骤 5：刷新数据

建立以上 4 个查询后：

1. **数据** → **全部刷新**（或 Ctrl+Alt+F5）
2. 确认「📦 销售出库」「💰 银行回款」「🏪 Sell-Through」「经销商主数据」四个表中的数据与 CSV 一致
3. 「🔗 经销商对账」和「📊 Dashboard」中的公式会自动重新计算

---

## 注意事项

- **路径问题**：如果 CSV 文件路径变更，需在 Power Query 编辑器中修改 `File.Contents()` 的路径参数
- **编码问题**：如果中文乱码，确认 CSV 文件保存为 UTF-8 BOM 格式
- **类型问题**：如果金额列显示为文本，在 Power Query 中将该列手动设为「小数」
- **数据更新**：将新的 CSV 文件替换旧文件后，刷新即可更新所有数据
