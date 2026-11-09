# 🧹 Excel Power Query Data Cleaning, Star Schema & Pizza Sales Dashboard

![Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Power Query](https://img.shields.io/badge/Power%20Query%20(M)-0078D4?style=for-the-badge)
![Power Pivot](https://img.shields.io/badge/Power%20Pivot%20%2F%20DAX-F2C811?style=for-the-badge&logoColor=black)
![Status](https://img.shields.io/badge/Status-Completed-success?style=for-the-badge)

An end-to-end **Excel BI** project built entirely inside Microsoft Excel:

1. **ETL with Power Query (M)** — combining a folder of monthly order files and cleaning messy Orders, Products and Customers sources.
2. **Data modeling with Power Pivot** — a star schema with a Calendar table, relationships and DAX measures.
3. **Reporting** — pivot tables feeding an interactive **Pizza Sales Dashboard** with slicers and a timeline.

---

## 📑 Table of Contents

- [Project Overview](#-project-overview)
- [Repository Structure](#-repository-structure)
- [Data Sources](#-data-sources)
- [Data Quality Issues Found](#-data-quality-issues-found)
- [Power Query Cleaning Steps](#-power-query-cleaning-steps)
  - [Orders](#1-orders-table-folder-of-monthly-files)
  - [Products](#2-products-table)
  - [Customers](#3-customers-table)
- [Data Model (Star Schema)](#-data-model-star-schema)
- [Pivot Tables](#-pivot-tables)
- [Pizza Sales Dashboard](#-pizza-sales-dashboard)
- [Key Insights](#-key-insights)
- [How to Use / Reproduce](#-how-to-use--reproduce)
- [Skills Demonstrated](#-skills-demonstrated)
- [Author](#-author)

---

## 🔍 Project Overview

| Item | Details |
|------|---------|
| **Tool** | Microsoft Excel (Power Query, Power Pivot, PivotTables, Slicers, Timeline) |
| **Language** | Power Query M, DAX |
| **Sources** | Folder of monthly order CSVs, `Products.csv`, `Customers.txt` |
| **Model** | Star schema: 1 fact table (`Orders`) + 3 dimensions (`Products`, `Customers`, `Calendar`) |
| **Output** | Clean, refreshable tables + interactive Excel dashboard |

The goal was to take raw, inconsistent data (mixed date formats, blank keys, duplicates, text-encoded prices, inconsistent casing and spacing) and turn it into a **refreshable, analysis-ready model** — then communicate the results through a one-page dashboard.

---

## 📁 Repository Structure

```
.
├── README.md
├── power-query/
│   ├── Orders.m        # Combine folder + clean & rebuild dates + new OrderID
│   ├── Products.m      # Dedupe, proper case, trim, fix price text → currency
│   └── Customers.m     # Dedupe, trim, proper case, split name, clean region
└── images/
    ├── data-model.png          # Power Pivot diagram view
    ├── dashboard.png           # Final Pizza Sales Dashboard
    ├── pivot-tables.png        # Pivot tables behind the dashboard
    └── top-bottom-pivots.png   # Top 5 / Bottom 5 pivot tables
```

> **Note:** the `Source` paths inside the `.m` files point to the author's local machine. Update them to your own folder/file paths before refreshing.

---

## 🗂 Data Sources

| Source | Type | Columns |
|--------|------|---------|
| `Orders/` | Folder of monthly CSV files (file name contains the month number, e.g. `Orders_1.csv`) | `OrderID`, `CustomerID`, `ProductID`, `OrderDate`, `Quantity` |
| `Products.csv` | CSV | `ProductID`, `ProductName`, `Category`, `Price` |
| `Customers.txt` | Comma-delimited text | `CustomerID`, `CustomerName`, `Region` |

---

## ⚠️ Data Quality Issues Found

| Table | Issue | Fix |
|-------|-------|-----|
| Orders | Data split across many monthly files | `Folder.Files` + custom *Transform File* function to combine |
| Orders | Missing `ProductID` / `Quantity` | Filtered out null / blank rows |
| Orders | **Mixed date formats** (`/` and `-` separators, year first or last, day/month order varies) | Normalized separator, split into parts, rebuilt Year / Month / Day with logic (see below) |
| Orders | `OrderID` unreliable after cleaning | Dropped and regenerated with an Index column |
| Products | Duplicate `ProductID`s | `Table.Distinct` on `ProductID` |
| Products | Inconsistent casing & extra spaces | `Text.Proper` + `Text.Trim` |
| Products | Price stored as text (`$`, `"Two Hundred"`) | Replaced text values, removed `$`, converted to **Currency** |
| Customers | Duplicate `CustomerID`s | `Table.Distinct` on `CustomerID` |
| Customers | Full name in one column, messy casing | Trim → Proper case → split into `Firstname` / `Lastname` |
| Customers | Region with stray spaces / casing | Removed spaces, Proper case, trimmed |

---

## 🛠 Power Query Cleaning Steps

### 1) Orders table (folder of monthly files)

**Approach**

1. **Combine files** – load every file from the `Orders` folder (hidden files excluded) and expand them with the auto-generated *Transform File* function. The file name is kept as `Source.Name`.
2. **Set types** and **remove invalid rows** where `ProductID` or `Quantity` is null/blank.
3. **Normalize dates** – replace `/` with `-`, then split `OrderDate` into three parts: `OrderDate.1`, `OrderDate.2`, `OrderDate.3`.
4. **Find the Year** – whichever part is `>= 2000` (first or last position).
5. **Find the Month from the file name** – extract the text between `_` and `.` in `Source.Name` (e.g. `Orders_7.csv` → `7`). This is the most reliable month source because each file holds exactly one month.
6. **Find the Day** – the first part that is `<= 31` and is **not** the month; if none qualifies, day = month (handles dates such as `5-5-2015`).
7. **Rebuild the date** – merge `Day-Month-Year` and convert to a proper `date` type using the `en-AU` (day-first) locale.
8. **Regenerate `OrderID`** – drop the original ID, add a 1-based index, and rename it `OrderID`.

**Key logic – Year & Day detection**

```m
// Year: take whichever split part looks like a 4-digit year
#"Added Conditional Column" = Table.AddColumn(#"Changed Type1", "Year", each
    if [OrderDate.1] >= 2000 then [OrderDate.1]
    else if [OrderDate.3] >= 2000 then [OrderDate.3]
    else "not found"),

// Month: taken from the file name, e.g. "Orders_7.csv" -> 7
#"Extracted Text Between Delimiters" = Table.TransformColumns(#"Added Conditional Column",
    {{"Source.Name", each Text.BetweenDelimiters(_, "_", "."), type text}}),

// Day: first part that is a valid day and is not the month
#"Added Custom" = Table.AddColumn(#"Reordered Columns", "Day", each
    if [OrderDate.1] <= 31 and [OrderDate.1] <> [Month] then [OrderDate.1]
    else if [OrderDate.2] <= 31 and [OrderDate.2] <> [Month] then [OrderDate.2]
    else if [OrderDate.3] <= 31 and [OrderDate.3] <> [Month] then [OrderDate.3]
    else [Month]),
```

<details>
<summary><b>📜 Full M code – Orders</b></summary>

```m
let
    Source = Folder.Files("C:\...\PowerQuery\Orders"),
    #"Filtered Hidden Files1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
    #"Invoke Custom Function1" = Table.AddColumn(#"Filtered Hidden Files1", "Transform File", each #"Transform File"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File", Table.ColumnNames(#"Transform File"(#"Sample File"))),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"Source.Name", type text}, {"OrderID", Int64.Type}, {"CustomerID", Int64.Type}, {"ProductID", Int64.Type}, {"OrderDate", type text}, {"Quantity", Int64.Type}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type", each [ProductID] <> null and [ProductID] <> ""),
    #"Filtered Rows1" = Table.SelectRows(#"Filtered Rows", each [Quantity] <> null and [Quantity] <> ""),
    #"Replaced Value" = Table.ReplaceValue(#"Filtered Rows1","/","-",Replacer.ReplaceText,{"OrderDate"}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Replaced Value", "OrderDate", Splitter.SplitTextByDelimiter("-", QuoteStyle.Csv), {"OrderDate.1", "OrderDate.2", "OrderDate.3"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"OrderDate.1", Int64.Type}, {"OrderDate.2", Int64.Type}, {"OrderDate.3", Int64.Type}}),
    #"Added Conditional Column" = Table.AddColumn(#"Changed Type1", "Year", each if [OrderDate.1] >= 2000 then [OrderDate.1] else if [OrderDate.3] >= 2000 then [OrderDate.3] else "not found"),
    #"Extracted Text Between Delimiters" = Table.TransformColumns(#"Added Conditional Column", {{"Source.Name", each Text.BetweenDelimiters(_, "_", "."), type text}}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Extracted Text Between Delimiters",{{"Source.Name", Int64.Type}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type2",{{"Source.Name", "Month"}}),
    #"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns",{"OrderID", "CustomerID", "ProductID", "OrderDate.1", "OrderDate.2", "OrderDate.3", "Quantity", "Year", "Month"}),
    #"Added Custom" = Table.AddColumn(#"Reordered Columns", "Day", each if [OrderDate.1]<=31 and [OrderDate.1]<>[Month] then [OrderDate.1]
else if [OrderDate.2]<=31 and [OrderDate.2]<>[Month] then [OrderDate.2]
else if [OrderDate.3]<=31 and [OrderDate.3]<>[Month] then [OrderDate.3]
else 
[Month]),
    #"Merged Columns" = Table.CombineColumns(Table.TransformColumnTypes(#"Added Custom", {{"Day", type text}, {"Month", type text}, {"Year", type text}}, "en-AU"),{"Day", "Month", "Year"},Combiner.CombineTextByDelimiter("-", QuoteStyle.None),"New Date"),
    #"Removed Columns" = Table.RemoveColumns(#"Merged Columns",{"OrderDate.1", "OrderDate.2", "OrderDate.3"}),
    #"Changed Type3" = Table.TransformColumnTypes(#"Removed Columns",{{"New Date", type date}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Changed Type3",{"OrderID"}),
    #"Added Index" = Table.AddIndexColumn(#"Removed Columns1", "Index", 1, 1, Int64.Type),
    #"Reordered Columns1" = Table.ReorderColumns(#"Added Index",{"Index", "CustomerID", "ProductID", "Quantity", "New Date"}),
    #"Renamed Columns2" = Table.RenameColumns(#"Reordered Columns1",{{"Index", "OrderID"}})
in
    #"Renamed Columns2"
```

</details>

**Output columns:** `OrderID`, `CustomerID`, `ProductID`, `Quantity`, `New Date`

---

### 2) Products table

**Steps**

1. Load `Products.csv` and promote headers.
2. Set types (`Price` kept as text at first because it contains non-numeric values).
3. **Remove duplicates** on `ProductID`.
4. `ProductName` and `Category` → **Proper case** + **Trim**.
5. Fix `Price`: replace `"Two Hundred"` → `200`, remove `$`, convert to **Currency**.

<details>
<summary><b>📜 Full M code – Products</b></summary>

```m
let
    Source = Csv.Document(File.Contents("C:\...\PowerQuery\Products.csv"),[Delimiter=",", Columns=4, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"ProductID", Int64.Type}, {"ProductName", type text}, {"Category", type text}, {"Price", type text}}),
    #"Removed Duplicates" = Table.Distinct(#"Changed Type", {"ProductID"}),
    #"Capitalized Each Word" = Table.TransformColumns(#"Removed Duplicates",{{"ProductName", Text.Proper, type text}}),
    #"Trimmed Text" = Table.TransformColumns(#"Capitalized Each Word",{{"ProductName", Text.Trim, type text}}),
    #"Capitalized Each Word1" = Table.TransformColumns(#"Trimmed Text",{{"Category", Text.Proper, type text}}),
    #"Trimmed Text1" = Table.TransformColumns(#"Capitalized Each Word1",{{"Category", Text.Trim, type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Trimmed Text1","Two Hundred","200",Replacer.ReplaceText,{"Price"}),
    #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","$","",Replacer.ReplaceText,{"Price"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Replaced Value1",{{"Price", Currency.Type}})
in
    #"Changed Type1"
```

</details>

**Output columns:** `ProductID`, `ProductName`, `Category`, `Price`

---

### 3) Customers table

**Steps**

1. Load `Customers.txt` (comma-delimited) and promote headers.
2. Set types and **remove duplicates** on `CustomerID`.
3. `CustomerName` → **Trim** + **Proper case**, then **split at the first space** into `Firstname` and `Lastname`.
4. `Region` → remove spaces, **Proper case**, **Trim**.

<details>
<summary><b>📜 Full M code – Customers</b></summary>

```m
let
    Source = Csv.Document(File.Contents("C:\...\PowerQuery\Customers.txt"),[Delimiter=",", Columns=3, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"CustomerID", Int64.Type}, {"CustomerName", type text}, {"Region", type text}}),
    #"Removed Duplicates" = Table.Distinct(#"Changed Type", {"CustomerID"}),
    #"Trimmed Text" = Table.TransformColumns(#"Removed Duplicates",{{"CustomerName", Text.Trim, type text}}),
    #"Capitalized Each Word" = Table.TransformColumns(#"Trimmed Text",{{"CustomerName", Text.Proper, type text}}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Capitalized Each Word", "CustomerName", Splitter.SplitTextByEachDelimiter({" "}, QuoteStyle.Csv, false), {"CustomerName.1", "CustomerName.2"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"CustomerName.1", type text}, {"CustomerName.2", type text}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type1",{{"CustomerName.1", "Firstname"}, {"CustomerName.2", "Lastname"}}),
    #"Replaced Value" = Table.ReplaceValue(#"Renamed Columns"," ","",Replacer.ReplaceText,{"Region"}),
    #"Capitalized Each Word1" = Table.TransformColumns(#"Replaced Value",{{"Region", Text.Proper, type text}}),
    #"Trimmed Text1" = Table.TransformColumns(#"Capitalized Each Word1",{{"Region", Text.Trim, type text}})
in
    #"Trimmed Text1"
```

</details>

**Output columns:** `CustomerID`, `Firstname`, `Lastname`, `Region`

---

## ⭐ Data Model (Star Schema)

The cleaned tables were loaded into the **Power Pivot Data Model** and related in a star schema:

<img width="991" height="583" alt="Screenshot 2026-11-09 221413" src="https://github.com/user-attachments/assets/261fcfd6-4b04-4eff-9c79-184c188c703c" />


| Relationship | Type |
|--------------|------|
| `Products[ProductID]` → `Orders[ProductID]` | One-to-many (1 : \*) |
| `Customers[CustomerID]` → `Orders[CustomerID]` | One-to-many (1 : \*) |
| `Calendar[Date]` → `Orders[New Date]` | One-to-many (1 : \*) |

**Calendar table columns:** `Date`, `Year`, `Month Number`, `Month`, `MMM-YYYY`, `Day Of Week Number`, `Day Of Week`, plus a **Date Hierarchy** (Year → Month → Date).

**DAX measures (in `Orders`):**

| Measure | Purpose |
|---------|---------|
| `Total Sales` | Revenue (quantity × product price) |
| `Total Quantity` | Total units sold |
| `Avg Sales per Order` | Average revenue per order |
| `Customers` | Number of distinct customers |

---

## 📊 Pivot Tables

All dashboard visuals are driven by pivot tables on a separate sheet:

- Sales % by **category** and by **size**
- KPI pivots: **Sum of Sales**, **Sum of Quantity**, **Count of Orders**, **Count of Customers**
- Order count by **month** and by **day of week**

- **Top 5** and **Bottom 5** pizzas (Top/Bottom N value filters)

---

## 🍕 Pizza Sales Dashboard

<img width="1394" height="523" alt="Sales Dashboard " src="https://github.com/user-attachments/assets/34ffa202-bcaa-462d-9d03-e8d531a90b38" />


**KPI cards**

| Total Sales | Total Customers | Total Orders | Total Quantity |
|:-----------:|:---------------:|:------------:|:--------------:|
| **$8,436** | **499** | **499** | **511** |

**Visuals**

- 🍩 **Total Sales per Category** – donut chart
- 🥧 **Sales Distribution by Pizza Size** – pie chart
- 📈 **Monthly Order Volume Peaks** – line chart
- 📊 **Weekly Order Distribution** – column chart
- 🏆 **Top 5 Best-Selling Pizzas** – bar chart
- 🔻 **Bottom 5 Least-Selling Pizzas** – bar chart

**Interactivity**

- Slicers for **pizza size** and **pizza category** (bilingual English/Arabic labels)
- **Timeline** slicer on order date (by quarter / month)
- All slicers connected to every pivot table, so the whole dashboard filters together

---

## 💡 Key Insights

- **Veggie** is the top category by sales share (**28.04%**), followed by Classic (26.96%) and Supreme (24.32%); **Chicken** is lowest (20.67%).
- **Large (L)** pizzas lead with **47%** of sales; **XL** is only **3%**.
- Order volume **peaks in June (54)** and August (51); **March is the lowest (30)**.
- **Friday** is the busiest day (**91** orders); **Wednesday** is the quietest (48).
- **The Classic Deluxe Pizza** is the best seller (459); **The Brie Carre Pizza** is the weakest (47).

---

## ▶️ How to Use / Reproduce

1. Clone or download this repository.
2. In Excel: **Data → Get Data → Launch Power Query Editor**.
3. For each table, create a **Blank Query → Advanced Editor** and paste the code from `power-query/`.
4. Update the `Source` paths to your local data location.
   - For **Orders**, the easiest route is **Get Data → From Folder → Combine & Transform**, which generates the `Transform File` / `Sample File` helper queries automatically; then replace the main query's steps with `Orders.m`.
5. **Close & Load To… → Only Create Connection** + **Add this data to the Data Model**.
6. In **Power Pivot → Diagram View**, create the Calendar table and the relationships shown above, then add the DAX measures.
7. Build pivot tables from the Data Model and connect slicers/timeline to all of them.
8. Drop new monthly files into the `Orders` folder and hit **Refresh All** — the whole pipeline updates.

---

## 🧠 Skills Demonstrated

- Combining multiple files from a folder with Power Query
- Writing and editing **M code** in the Advanced Editor
- Handling **mixed / inconsistent date formats** with custom logic
- Deriving data from **file metadata** (month from file name)
- Deduplication, text standardization, column splitting, type conversion
- Converting text-encoded numbers to proper **Currency** values
- **Star schema** modeling and relationships in Power Pivot
- Calendar / date table with hierarchy
- **DAX** measures
- Pivot tables, Top/Bottom N filters, slicers & timelines
- Dashboard design and data storytelling in Excel

---

## 👤 Author

**Mohamed Mahmoud Sobhy**

[![GitHub](https://img.shields.io/badge/GitHub-M--M--Sobhy-181717?style=flat&logo=github)](https://github.com/M-M-Sobhy)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Mohamed%20Sobhy-0A66C2?style=flat&logo=linkedin)](https://www.linkedin.com/in/mohamed-mahmoud-sobhy-668ba937a)
[![YouTube](https://img.shields.io/badge/YouTube-@M__Sob7y-FF0000?style=flat&logo=youtube)](https://www.youtube.com/@M_Sob7y)

⭐ If you found this project useful, consider giving it a star!
