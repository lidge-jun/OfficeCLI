# Pivot Table Showcase

This demo consists of three files that work together:

- **pivot-tables.py** — Python script that calls `officecli` commands to generate the workbook. Each pivot table command is shown as a copyable shell command in the comments, then executed by the script. Read this to learn the exact `officecli add --type pivottable --prop ...` syntax.
- **pivot-tables.xlsx** — The generated workbook with 13 sheets. Open in Excel to see the rendered pivot tables. Use `officecli get` or `officecli query` to inspect programmatically.
- **pivot-tables.md** — This file. Maps each sheet in the xlsx to the feature it demonstrates and the command that created it.

## Regenerate

```bash
cd examples/excel
python3 pivot-tables.py
# → pivot-tables.xlsx
```

## Source Data

| Sheet | Rows | Columns | Purpose |
|-------|------|---------|---------|
| Sheet1 | 50 | Region, Category, Product, Quarter, Sales, Quantity, Cost, Channel, Priority, Date | English sales data spanning 2024-2025 |
| CNData | 12 | 地区, 品类, 销售额 | Chinese sales data for locale sort demo |

## Pivot Tables

### Sheet: 1-Sales Overview

The most feature-rich pivot. Tabular layout with 2-level row hierarchy crossed against quarterly columns. Three value fields where Cost is shown as percentage of row total. Dual page filters let users slice by Channel and Priority. Outer row labels repeat on every row.

```bash
officecli add pivot-tables.xlsx "/1-Sales Overview" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Region,Category \
  --prop cols=Quarter \
  --prop 'values=Sales:sum,Quantity:sum,Cost:sum:percent_of_row' \
  --prop 'filters=Channel,Priority' \
  --prop layout=tabular \
  --prop repeatlabels=true \
  --prop grandtotals=both \
  --prop subtotals=on \
  --prop sort=desc \
  --prop style=PivotStyleDark2
```

**Features:** `layout=tabular`, `repeatlabels=true`, dual `filters`, `values` with `percent_of_row`, `sort=desc`

### Sheet: 2-Market Share

Each region's share within each category, shown as column percentages. Outline layout provides expand/collapse grouping.

```bash
officecli add pivot-tables.xlsx "/2-Market Share" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Region \
  --prop cols=Category \
  --prop 'values=Sales:sum:percent_of_col' \
  --prop filters=Channel \
  --prop layout=outline \
  --prop grandtotals=both \
  --prop style=PivotStyleMedium4
```

**Features:** `layout=outline`, `values` with `percent_of_col`

### Sheet: 3-Product Deep Dive

Five value fields with three different aggregation functions on the same source column (Sales:sum, Sales:average, Sales:max). No column axis — values become column headers automatically.

```bash
officecli add pivot-tables.xlsx "/3-Product Deep Dive" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Category,Product \
  --prop 'values=Sales:sum,Sales:average,Sales:max,Quantity:sum,Cost:sum' \
  --prop filters=Region \
  --prop layout=tabular \
  --prop grandtotals=rows \
  --prop subtotals=on \
  --prop sort=desc \
  --prop style=PivotStyleMedium9
```

**Features:** 5 `values` fields, no `cols` (synthetic Values axis), `grandtotals=rows`

### Sheet: 4-Channel Analysis

Sales shown as percentage of the grand total — reveals each channel's global share across quarters. No page filters.

```bash
officecli add pivot-tables.xlsx "/4-Channel Analysis" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Channel \
  --prop cols=Quarter \
  --prop 'values=Sales:sum:percent_of_total,Quantity:sum' \
  --prop layout=outline \
  --prop grandtotals=both \
  --prop style=PivotStyleLight21
```

**Features:** `values` with `percent_of_total`, no `filters`

### Sheet: 5-Priority Matrix

Blank rows inserted after each outer group (Priority) for visual separation. Ascending sort puts High first.

```bash
officecli add pivot-tables.xlsx "/5-Priority Matrix" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Priority,Region \
  --prop cols=Category \
  --prop 'values=Sales:sum,Cost:sum:percent_of_row' \
  --prop filters=Channel \
  --prop layout=tabular \
  --prop blankrows=true \
  --prop grandtotals=both \
  --prop subtotals=on \
  --prop sort=asc \
  --prop style=PivotStyleDark6
```

**Features:** `blankrows=true`, `sort=asc`

### Sheet: 6-Compact 3-Level

Three-level row hierarchy (Region > Category > Product) in compact layout — all labels share one column with progressive indentation.

```bash
officecli add pivot-tables.xlsx "/6-Compact 3-Level" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Region,Category,Product \
  --prop 'values=Sales:sum,Quantity:sum' \
  --prop filters=Priority \
  --prop layout=compact \
  --prop grandtotals=both \
  --prop subtotals=on \
  --prop sort=desc \
  --prop style=PivotStyleMedium14
```

**Features:** `layout=compact`, 3-level `rows`

### Sheet: 7-No Subtotals

Flat tabular view with subtotals disabled. Only the bottom grand total row remains. Outer labels are repeated on every row since there are no subtotal rows to carry them.

```bash
officecli add pivot-tables.xlsx "/7-No Subtotals" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Region,Category \
  --prop cols=Quarter \
  --prop values=Sales:sum \
  --prop layout=tabular \
  --prop repeatlabels=true \
  --prop grandtotals=cols \
  --prop subtotals=off \
  --prop sort=asc \
  --prop style=PivotStyleLight1
```

**Features:** `subtotals=off`, `grandtotals=cols`, `repeatlabels=true`

### Sheet: 8-Date Grouping

Automatic date grouping from a date column. `Date:year` creates year buckets ("2024", "2025"), `Date:quarter` creates quarter sub-buckets ("2024-Q1", ...). Uses native Excel fieldGroup XML.

```bash
officecli add pivot-tables.xlsx "/8-Date Grouping" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop 'rows=Date:year,Date:quarter' \
  --prop 'values=Sales:sum,Cost:sum' \
  --prop filters=Region \
  --prop layout=outline \
  --prop grandtotals=both \
  --prop subtotals=on \
  --prop style=PivotStyleMedium7
```

**Features:** `rows` with `Date:year,Date:quarter` date grouping syntax

### Sheet: 9-Top 5 Products

Only the top 5 products by sales are shown. Grand totals are hidden entirely.

```bash
officecli add pivot-tables.xlsx "/9-Top 5 Products" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Product \
  --prop 'values=Sales:sum,Quantity:sum,Cost:sum' \
  --prop layout=tabular \
  --prop grandtotals=none \
  --prop topN=5 \
  --prop sort=desc \
  --prop style=PivotStyleDark1
```

**Features:** `topN=5`, `grandtotals=none`

### Sheet: 10-Ultimate

Every feature combined in one pivot table — the kitchen sink.

```bash
officecli add pivot-tables.xlsx "/10-Ultimate" --type pivottable \
  --prop source=Sheet1!A1:J51 \
  --prop rows=Region,Category \
  --prop cols=Quarter \
  --prop 'values=Sales:sum,Quantity:average,Cost:sum:percent_of_row' \
  --prop 'filters=Channel,Priority' \
  --prop layout=tabular \
  --prop repeatlabels=true \
  --prop blankrows=true \
  --prop grandtotals=rows \
  --prop subtotals=on \
  --prop sort=desc \
  --prop style=PivotStyleDark11
```

**Features:** `repeatlabels=true` + `blankrows=true` + dual `filters` + mixed aggregations + `grandtotals=rows`

### Sheet: 11-Chinese Locale

Chinese data with pinyin-order sorting and a custom grand total label. Demonstrates that field names, filter values, and captions all work with non-ASCII text.

```bash
officecli add pivot-tables.xlsx "/11-Chinese Locale" --type pivottable \
  --prop source=CNData!A1:C13 \
  --prop rows=地区,品类 \
  --prop values=销售额:sum \
  --prop layout=tabular \
  --prop grandtotals=both \
  --prop subtotals=on \
  --prop sort=locale \
  --prop grandTotalCaption=合计 \
  --prop style=PivotStyleMedium2
```

**Features:** `sort=locale` (pinyin: 华北 < 华东 < 华南 < 西南), `grandTotalCaption`

## Inspect the Generated File

```bash
# List all pivot tables
officecli query pivot-tables.xlsx pivottable

# Get details of a specific pivot
officecli get pivot-tables.xlsx "/1-Sales Overview/pivottable[1]"

# View rendered data as text
officecli view pivot-tables.xlsx text --sheet "1-Sales Overview"
```
