# Warehouse Stock Control

An Excel workbook that tracks stock from goods-in to dispatch and shows what needs reordering. I built it to practise the stock skills asked for in entry-level warehouse and logistics jobs in Ireland.

**All products, prices and quantities are made up.**

[Open the workbook](warehouse_stock_control.xlsx)

## What it does

- Logs 2,133 sample stock movements for 30 products: receipts (IN), picks (OUT) and adjustments (ADJ)
- Works out on-hand stock for every product from its opening stock and movements
- Calculates each reorder point from average daily use, supplier lead time and safety stock
- Flags each product as OK, WATCH, REORDER or OUT OF STOCK, and suggests a quantity to order
- Ranks products with ABC analysis, to show which ones need the most attention
- Compares physical counts with system stock, showing the variance and count accuracy
- Summarises it all on a dashboard, with assumptions you can edit

## Excel skills used

`SUMIFS`, `COUNTIF`, `SUMIF`, `RANK`, nested `IF`, `ROUNDUP`, data validation, conditional formatting, filters, frozen panes and charts.

## The key formula

```
Reorder point = average daily use x lead time in days + safety stock
```

For example, a product selling 5 units a day with a 7-day lead time and 20 units of safety stock should be reordered at 55 units.

## Try it

1. Open the **Cycle Count** sheet and type counts in the yellow column.
2. Add a row to **Movements** and watch **Stock Status** update.
3. Change "Days of cover to order" on the **Dashboard**.

## Rebuilding the file

`build_workbook.py` creates the workbook with Python and `openpyxl`. Install `openpyxl`, then run:

```
python build_workbook.py
```
