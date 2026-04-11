#!/usr/bin/env python3
"""
Pivot Table Showcase — generates pivot-tables.xlsx with 11 pivot tables.

Each pivot table demonstrates different officecli features.
See pivot-tables.md for a guide to each sheet in the generated file.

Usage:
  python3 pivot-tables.py
"""

import subprocess, sys, os, json, atexit

FILE = "pivot-tables.xlsx"

def cli(cmd):
    """Run: officecli <cmd>"""
    r = subprocess.run(f"officecli {cmd}", shell=True, capture_output=True, text=True)
    out = (r.stdout or "").strip()
    if out:
        for line in out.split("\n"):
            if line.strip():
                print(f"  {line.strip()}")
    if r.returncode != 0:
        err = (r.stderr or "").strip()
        if err and "UNSUPPORTED" not in err and "process cannot access" not in err:
            print(f"  ERROR: {err}")

if os.path.exists(FILE):
    os.remove(FILE)

cli(f'create "{FILE}"')
cli(f'open "{FILE}"')
atexit.register(lambda: cli(f'close "{FILE}"'))

# ==========================================================================
# Source data — batch is used here only for speed (500+ cell writes).
# ==========================================================================
print("\n--- Populating source data ---")

data_cmds = []
for j, h in enumerate(["Region","Category","Product","Quarter","Sales","Quantity","Cost","Channel","Priority","Date"]):
    data_cmds.append({"command":"set","path":f"/Sheet1/{'ABCDEFGHIJ'[j]}1","props":{"text":h}})

rows = [
    ("North","Electronics","Laptop","Q1",12500,45,7500,"Online","High","2025-01-15"),
    ("North","Electronics","Phone","Q1",8900,120,5340,"Retail","High","2025-02-10"),
    ("North","Electronics","Tablet","Q2",6200,38,3720,"Online","Medium","2025-04-20"),
    ("North","Electronics","Laptop","Q2",15800,55,9480,"Retail","High","2025-05-08"),
    ("North","Electronics","Phone","Q3",11200,150,6720,"Online","High","2025-07-12"),
    ("North","Electronics","Tablet","Q4",9500,62,5700,"Retail","Medium","2025-10-05"),
    ("North","Clothing","Jacket","Q1",4200,85,2100,"Retail","Low","2025-01-22"),
    ("North","Clothing","Shoes","Q2",5600,70,2800,"Online","Medium","2025-04-15"),
    ("North","Clothing","Hat","Q3",1800,110,900,"Retail","Low","2025-08-03"),
    ("North","Clothing","Jacket","Q4",7800,95,3900,"Online","High","2025-11-18"),
    ("North","Food","Coffee","Q1",2400,200,1200,"Retail","Low","2025-03-01"),
    ("North","Food","Snacks","Q2",1500,180,750,"Online","Low","2025-06-10"),
    ("North","Food","Juice","Q3",1900,160,950,"Retail","Medium","2025-09-20"),
    ("North","Food","Coffee","Q4",3200,220,1600,"Online","Medium","2025-12-01"),
    ("South","Electronics","Phone","Q1",18500,200,11100,"Online","High","2024-01-20"),
    ("South","Electronics","Laptop","Q2",22000,72,13200,"Retail","High","2024-05-15"),
    ("South","Electronics","Tablet","Q3",7800,48,4680,"Online","Medium","2024-08-22"),
    ("South","Electronics","Phone","Q4",14200,165,8520,"Retail","High","2024-11-30"),
    ("South","Clothing","Shoes","Q1",9200,110,4600,"Retail","Medium","2024-02-14"),
    ("South","Clothing","Jacket","Q2",6500,78,3250,"Online","Low","2024-06-01"),
    ("South","Clothing","Hat","Q3",3100,130,1550,"Retail","Low","2024-09-10"),
    ("South","Clothing","Shoes","Q4",8800,98,4400,"Online","Medium","2024-12-20"),
    ("South","Food","Juice","Q1",1800,240,900,"Retail","Low","2024-03-08"),
    ("South","Food","Coffee","Q2",3500,280,1750,"Online","Medium","2024-04-25"),
    ("South","Food","Snacks","Q3",2200,190,1100,"Retail","Low","2024-07-14"),
    ("South","Food","Juice","Q4",2800,210,1400,"Online","Medium","2024-10-18"),
    ("East","Electronics","Tablet","Q1",5400,35,3240,"Online","Medium","2025-02-28"),
    ("East","Electronics","Laptop","Q2",19500,65,11700,"Retail","High","2025-05-20"),
    ("East","Electronics","Phone","Q3",13800,180,8280,"Online","High","2025-08-15"),
    ("East","Electronics","Tablet","Q4",8200,52,4920,"Retail","Medium","2025-11-02"),
    ("East","Clothing","Hat","Q1",2800,140,1400,"Retail","Low","2025-01-05"),
    ("East","Clothing","Jacket","Q2",7200,60,3600,"Online","Medium","2025-06-18"),
    ("East","Clothing","Shoes","Q3",5500,88,2750,"Retail","Medium","2025-09-25"),
    ("East","Clothing","Hat","Q4",3600,105,1800,"Online","Low","2025-12-10"),
    ("East","Food","Snacks","Q1",1200,300,600,"Retail","Low","2025-03-15"),
    ("East","Food","Juice","Q2",2100,170,1050,"Online","Medium","2025-04-30"),
    ("East","Food","Coffee","Q3",2800,230,1400,"Retail","Medium","2025-07-22"),
    ("East","Food","Snacks","Q4",1600,250,800,"Online","Low","2025-10-28"),
    ("West","Electronics","Laptop","Q1",20500,68,12300,"Online","High","2024-01-10"),
    ("West","Electronics","Phone","Q2",16800,190,10080,"Retail","High","2024-04-05"),
    ("West","Electronics","Tablet","Q3",8900,55,5340,"Online","Medium","2024-08-12"),
    ("West","Electronics","Laptop","Q4",25000,82,15000,"Retail","High","2024-11-15"),
    ("West","Clothing","Jacket","Q1",11000,88,5500,"Retail","Medium","2024-02-22"),
    ("West","Clothing","Shoes","Q2",7500,95,3750,"Online","Medium","2024-05-30"),
    ("West","Clothing","Hat","Q3",4200,120,2100,"Retail","Low","2024-09-08"),
    ("West","Clothing","Jacket","Q4",13500,105,6750,"Online","High","2024-12-01"),
    ("West","Food","Coffee","Q1",4500,350,2250,"Online","Medium","2024-03-18"),
    ("West","Food","Snacks","Q2",2800,280,1400,"Online","Medium","2024-06-22"),
    ("West","Food","Juice","Q3",3200,260,1600,"Retail","Low","2024-07-30"),
    ("West","Food","Coffee","Q4",5800,400,2900,"Online","High","2024-10-25"),
]
C = "ABCDEFGHIJ"
for i, row in enumerate(rows):
    for j, val in enumerate(row):
        data_cmds.append({"command":"set","path":f"/Sheet1/{C[j]}{i+2}","props":{"text":str(val)}})

data_cmds.append({"command":"add","parent":"/","type":"sheet","props":{"name":"CNData"}})
for j, h in enumerate(["地区","品类","销售额"]):
    data_cmds.append({"command":"set","path":f"/CNData/{C[j]}1","props":{"text":h}})
for i, (r, c, s) in enumerate([
    ("华东","电子产品",18000),("华东","服装",9500),("华东","食品",4200),
    ("华南","电子产品",22000),("华南","服装",12000),("华南","食品",5800),
    ("华北","电子产品",15000),("华北","服装",7800),("华北","食品",3600),
    ("西南","电子产品",11000),("西南","服装",6500),("西南","食品",2900),
]):
    for j, val in enumerate([r, c, s]):
        data_cmds.append({"command":"set","path":f"/CNData/{C[j]}{i+2}","props":{"text":str(val)}})

cli(f'batch "{FILE}" --force --commands \'{json.dumps(data_cmds)}\'')

# ==========================================================================
# 11 Pivot Tables
#
# Each section below shows the exact officecli command in a comment block,
# then executes it. You can copy any command block and run it in a terminal.
# ==========================================================================

# --------------------------------------------------------------------------
# Sheet: 1-Sales Overview
#
# officecli add pivot-tables.xlsx "/1-Sales Overview" --type pivottable \
#   --prop source=Sheet1!A1:J51 \
#   --prop rows=Region,Category \
#   --prop cols=Quarter \
#   --prop 'values=Sales:sum,Quantity:sum,Cost:sum:percent_of_row' \
#   --prop 'filters=Channel,Priority' \
#   --prop layout=tabular \
#   --prop repeatlabels=true \
#   --prop grandtotals=both \
#   --prop subtotals=on \
#   --prop sort=desc \
#   --prop name=SalesOverview \
#   --prop style=PivotStyleDark2
#
# Features: tabular layout, 2-level rows, column axis, 3 value fields,
#   Cost as percent_of_row, dual page filters, repeat item labels, desc sort
# --------------------------------------------------------------------------
print("\n--- 1-Sales Overview ---")
cli(f'add "{FILE}" / --type sheet --prop name="1-Sales Overview"')
cli(f'add "{FILE}" "/1-Sales Overview" --type pivottable'
    f' --prop source=Sheet1!A1:J51'
    f' --prop rows=Region,Category'
    f' --prop cols=Quarter'
    f' --prop values=Sales:sum,Quantity:sum,Cost:sum:percent_of_row'
    f' --prop filters=Channel,Priority'
    f' --prop layout=tabular'
    f' --prop repeatlabels=true'
    f' --prop grandtotals=both'
    f' --prop subtotals=on'
    f' --prop sort=desc'
    f' --prop name=SalesOverview'
    f' --prop style=PivotStyleDark2')

# --------------------------------------------------------------------------
# Sheet: 2-Market Share
#
# officecli add pivot-tables.xlsx "/2-Market Share" --type pivottable \
#   --prop source=Sheet1!A1:J51 \
#   --prop rows=Region \
#   --prop cols=Category \
#   --prop 'values=Sales:sum:percent_of_col' \
#   --prop filters=Channel \
#   --prop layout=outline \
#   --prop grandtotals=both \
#   --prop name=MarketShare \
#   --prop style=PivotStyleMedium4
#
# Features: outline layout, percent_of_col (each region's share per category)
# --------------------------------------------------------------------------
print("\n--- 2-Market Share ---")
cli(f'add "{FILE}" / --type sheet --prop name="2-Market Share"')
cli(f'add "{FILE}" "/2-Market Share" --type pivottable'
    f' --prop source=Sheet1!A1:J51'
    f' --prop rows=Region'
    f' --prop cols=Category'
    f' --prop values=Sales:sum:percent_of_col'
    f' --prop filters=Channel'
    f' --prop layout=outline'
    f' --prop grandtotals=both'
    f' --prop name=MarketShare'
    f' --prop style=PivotStyleMedium4')

# --------------------------------------------------------------------------
# Sheet: 3-Product Deep Dive
#
# officecli add pivot-tables.xlsx "/3-Product Deep Dive" --type pivottable \
#   --prop source=Sheet1!A1:J51 \
#   --prop rows=Category,Product \
#   --prop 'values=Sales:sum,Sales:average,Sales:max,Quantity:sum,Cost:sum' \
#   --prop filters=Region \
#   --prop layout=tabular \
#   --prop grandtotals=rows \
#   --prop subtotals=on \
#   --prop sort=desc \
#   --prop name=ProductDeepDive \
#   --prop style=PivotStyleMedium9
#
# Features: 5 value fields (sum, average, max), no column axis — values
#   become column headers via synthetic "Values" axis, row grand totals only
# --------------------------------------------------------------------------
print("\n--- 3-Product Deep Dive ---")
cli(f'add "{FILE}" / --type sheet --prop name="3-Product Deep Dive"')
cli(f'add "{FILE}" "/3-Product Deep Dive" --type pivottable'
    f' --prop source=Sheet1!A1:J51'
    f' --prop rows=Category,Product'
    f' --prop values=Sales:sum,Sales:average,Sales:max,Quantity:sum,Cost:sum'
    f' --prop filters=Region'
    f' --prop layout=tabular'
    f' --prop grandtotals=rows'
    f' --prop subtotals=on'
    f' --prop sort=desc'
    f' --prop name=ProductDeepDive'
    f' --prop style=PivotStyleMedium9')

# --------------------------------------------------------------------------
# Sheet: 4-Channel Analysis
#
# officecli add pivot-tables.xlsx "/4-Channel Analysis" --type pivottable \
#   --prop source=Sheet1!A1:J51 \
#   --prop rows=Channel \
#   --prop cols=Quarter \
#   --prop 'values=Sales:sum:percent_of_total,Quantity:sum' \
#   --prop layout=outline \
#   --prop grandtotals=both \
#   --prop name=ChannelTrend \
#   --prop style=PivotStyleLight21
#
# Features: percent_of_total (global share), no filters
# --------------------------------------------------------------------------
print("\n--- 4-Channel Analysis ---")
cli(f'add "{FILE}" / --type sheet --prop name="4-Channel Analysis"')
cli(f'add "{FILE}" "/4-Channel Analysis" --type pivottable'
    f' --prop source=Sheet1!A1:J51'
    f' --prop rows=Channel'
    f' --prop cols=Quarter'
    f' --prop values=Sales:sum:percent_of_total,Quantity:sum'
    f' --prop layout=outline'
    f' --prop grandtotals=both'
    f' --prop name=ChannelTrend'
    f' --prop style=PivotStyleLight21')

# --------------------------------------------------------------------------
# Sheet: 5-Priority Matrix
#
# officecli add pivot-tables.xlsx "/5-Priority Matrix" --type pivottable \
#   --prop source=Sheet1!A1:J51 \
#   --prop rows=Priority,Region \
#   --prop cols=Category \
#   --prop 'values=Sales:sum,Cost:sum:percent_of_row' \
#   --prop filters=Channel \
#   --prop layout=tabular \
#   --prop blankrows=true \
#   --prop grandtotals=both \
#   --prop subtotals=on \
#   --prop sort=asc \
#   --prop name=PriorityMatrix \
#   --prop style=PivotStyleDark6
#
# Features: blankRows — empty line after each outer group for visual separation
# --------------------------------------------------------------------------
print("\n--- 5-Priority Matrix ---")
cli(f'add "{FILE}" / --type sheet --prop name="5-Priority Matrix"')
cli(f'add "{FILE}" "/5-Priority Matrix" --type pivottable'
    f' --prop source=Sheet1!A1:J51'
    f' --prop rows=Priority,Region'
    f' --prop cols=Category'
    f' --prop values=Sales:sum,Cost:sum:percent_of_row'
    f' --prop filters=Channel'
    f' --prop layout=tabular'
    f' --prop blankrows=true'
    f' --prop grandtotals=both'
    f' --prop subtotals=on'
    f' --prop sort=asc'
    f' --prop name=PriorityMatrix'
    f' --prop style=PivotStyleDark6')

# --------------------------------------------------------------------------
# Sheet: 6-Compact 3-Level
#
# officecli add pivot-tables.xlsx "/6-Compact 3-Level" --type pivottable \
#   --prop source=Sheet1!A1:J51 \
#   --prop rows=Region,Category,Product \
#   --prop 'values=Sales:sum,Quantity:sum' \
#   --prop filters=Priority \
#   --prop layout=compact \
#   --prop grandtotals=both \
#   --prop subtotals=on \
#   --prop sort=desc \
#   --prop name=Compact3Level \
#   --prop style=PivotStyleMedium14
#
# Features: compact layout — 3-level hierarchy in one indented column
# --------------------------------------------------------------------------
print("\n--- 6-Compact 3-Level ---")
cli(f'add "{FILE}" / --type sheet --prop name="6-Compact 3-Level"')
cli(f'add "{FILE}" "/6-Compact 3-Level" --type pivottable'
    f' --prop source=Sheet1!A1:J51'
    f' --prop rows=Region,Category,Product'
    f' --prop values=Sales:sum,Quantity:sum'
    f' --prop filters=Priority'
    f' --prop layout=compact'
    f' --prop grandtotals=both'
    f' --prop subtotals=on'
    f' --prop sort=desc'
    f' --prop name=Compact3Level'
    f' --prop style=PivotStyleMedium14')

# --------------------------------------------------------------------------
# Sheet: 7-No Subtotals
#
# officecli add pivot-tables.xlsx "/7-No Subtotals" --type pivottable \
#   --prop source=Sheet1!A1:J51 \
#   --prop rows=Region,Category \
#   --prop cols=Quarter \
#   --prop values=Sales:sum \
#   --prop layout=tabular \
#   --prop repeatlabels=true \
#   --prop grandtotals=cols \
#   --prop subtotals=off \
#   --prop sort=asc \
#   --prop name=FlatView \
#   --prop style=PivotStyleLight1
#
# Features: subtotals=off (flat view), grandtotals=cols (bottom row only),
#   repeatlabels=true (essential when subtotals off — otherwise outer labels vanish)
# --------------------------------------------------------------------------
print("\n--- 7-No Subtotals ---")
cli(f'add "{FILE}" / --type sheet --prop name="7-No Subtotals"')
cli(f'add "{FILE}" "/7-No Subtotals" --type pivottable'
    f' --prop source=Sheet1!A1:J51'
    f' --prop rows=Region,Category'
    f' --prop cols=Quarter'
    f' --prop values=Sales:sum'
    f' --prop layout=tabular'
    f' --prop repeatlabels=true'
    f' --prop grandtotals=cols'
    f' --prop subtotals=off'
    f' --prop sort=asc'
    f' --prop name=FlatView'
    f' --prop style=PivotStyleLight1')

# --------------------------------------------------------------------------
# Sheet: 8-Date Grouping
#
# officecli add pivot-tables.xlsx "/8-Date Grouping" --type pivottable \
#   --prop source=Sheet1!A1:J51 \
#   --prop 'rows=Date:year,Date:quarter' \
#   --prop 'values=Sales:sum,Cost:sum' \
#   --prop filters=Region \
#   --prop layout=outline \
#   --prop grandtotals=both \
#   --prop subtotals=on \
#   --prop name=DateGrouping \
#   --prop style=PivotStyleMedium7
#
# Features: automatic date grouping — Date:year creates "2024","2025" buckets,
#   Date:quarter creates "2024-Q1",... sub-buckets. Uses native Excel fieldGroup XML.
# --------------------------------------------------------------------------
print("\n--- 8-Date Grouping ---")
cli(f'add "{FILE}" / --type sheet --prop name="8-Date Grouping"')
cli(f'add "{FILE}" "/8-Date Grouping" --type pivottable'
    f' --prop source=Sheet1!A1:J51'
    f' --prop rows=Date:year,Date:quarter'
    f' --prop values=Sales:sum,Cost:sum'
    f' --prop filters=Region'
    f' --prop layout=outline'
    f' --prop grandtotals=both'
    f' --prop subtotals=on'
    f' --prop name=DateGrouping'
    f' --prop style=PivotStyleMedium7')

# --------------------------------------------------------------------------
# Sheet: 9-Top 5 Products
#
# officecli add pivot-tables.xlsx "/9-Top 5 Products" --type pivottable \
#   --prop source=Sheet1!A1:J51 \
#   --prop rows=Product \
#   --prop 'values=Sales:sum,Quantity:sum,Cost:sum' \
#   --prop layout=tabular \
#   --prop grandtotals=none \
#   --prop topN=5 \
#   --prop sort=desc \
#   --prop name=Top5Products \
#   --prop style=PivotStyleDark1
#
# Features: topN=5 (only top 5 products by first value field), grandtotals=none
# --------------------------------------------------------------------------
print("\n--- 9-Top 5 Products ---")
cli(f'add "{FILE}" / --type sheet --prop name="9-Top 5 Products"')
cli(f'add "{FILE}" "/9-Top 5 Products" --type pivottable'
    f' --prop source=Sheet1!A1:J51'
    f' --prop rows=Product'
    f' --prop values=Sales:sum,Quantity:sum,Cost:sum'
    f' --prop layout=tabular'
    f' --prop grandtotals=none'
    f' --prop topN=5'
    f' --prop sort=desc'
    f' --prop name=Top5Products'
    f' --prop style=PivotStyleDark1')

# --------------------------------------------------------------------------
# Sheet: 10-Ultimate
#
# officecli add pivot-tables.xlsx "/10-Ultimate" --type pivottable \
#   --prop source=Sheet1!A1:J51 \
#   --prop rows=Region,Category \
#   --prop cols=Quarter \
#   --prop 'values=Sales:sum,Quantity:average,Cost:sum:percent_of_row' \
#   --prop 'filters=Channel,Priority' \
#   --prop layout=tabular \
#   --prop repeatlabels=true \
#   --prop blankrows=true \
#   --prop grandtotals=rows \
#   --prop subtotals=on \
#   --prop sort=desc \
#   --prop name=UltimatePivot \
#   --prop style=PivotStyleDark11
#
# Features: ALL features combined — tabular + repeatLabels + blankRows +
#   dual filters + 3 mixed-aggregation values + row-only grand totals
# --------------------------------------------------------------------------
print("\n--- 10-Ultimate ---")
cli(f'add "{FILE}" / --type sheet --prop name="10-Ultimate"')
cli(f'add "{FILE}" "/10-Ultimate" --type pivottable'
    f' --prop source=Sheet1!A1:J51'
    f' --prop rows=Region,Category'
    f' --prop cols=Quarter'
    f' --prop values=Sales:sum,Quantity:average,Cost:sum:percent_of_row'
    f' --prop filters=Channel,Priority'
    f' --prop layout=tabular'
    f' --prop repeatlabels=true'
    f' --prop blankrows=true'
    f' --prop grandtotals=rows'
    f' --prop subtotals=on'
    f' --prop sort=desc'
    f' --prop name=UltimatePivot'
    f' --prop style=PivotStyleDark11')

# --------------------------------------------------------------------------
# Sheet: 11-Chinese Locale
#
# officecli add pivot-tables.xlsx "/11-Chinese Locale" --type pivottable \
#   --prop source=CNData!A1:C13 \
#   --prop rows=地区,品类 \
#   --prop values=销售额:sum \
#   --prop layout=tabular \
#   --prop grandtotals=both \
#   --prop subtotals=on \
#   --prop sort=locale \
#   --prop grandTotalCaption=合计 \
#   --prop name=ChineseLocale \
#   --prop style=PivotStyleMedium2
#
# Features: sort=locale (Chinese pinyin: 华北 < 华东 < 华南 < 西南),
#   grandTotalCaption=合计 (custom grand total label)
# --------------------------------------------------------------------------
print("\n--- 11-Chinese Locale ---")
cli(f'add "{FILE}" / --type sheet --prop name="11-Chinese Locale"')
cli(f'add "{FILE}" "/11-Chinese Locale" --type pivottable'
    f' --prop source=CNData!A1:C13'
    f' --prop rows=地区,品类'
    f' --prop values=销售额:sum'
    f' --prop layout=tabular'
    f' --prop grandtotals=both'
    f' --prop subtotals=on'
    f' --prop sort=locale'
    f' --prop grandTotalCaption=合计'
    f' --prop name=ChineseLocale'
    f' --prop style=PivotStyleMedium2')

print(f"\nDone! Generated: {FILE}")
print("  13 sheets (Sheet1 + CNData + 11 pivot tables)")
