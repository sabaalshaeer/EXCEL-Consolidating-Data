# EXCEL-Consolidating-Data
Open the workbook 
    open the Consolidate.xlsx workbook.

Examine the quarterly sales files, noting that the Product column contains the same products, but not in the same order.
    Using the taskbar to view different quarterly sales workbooks, note that the products are not consistently arranged within the same rows.
    For example, note the location of product 4269:
    Q1 Sales.xlsx: row 2.
    Q2 Sales.xlsx: row 19.
    Q3 Sales.xlsx: row 18.
    Q4 Sales.xlsx: row 15.
    Note: All of the same products appear in both sheets, but they are listed in a different order.

Consolidate the quantity of products sold from each of the quarterly sales files.
    Switch back to the My Consolidate.xlsx file.
    Select the cell range B3:C23.
    Select Data→Consolidate.
    In the Consolidate dialog box, in the Function drop-down list box, verify that Sum is selected.
    Verify the cursor is in the Reference field.
    In the taskbar, switch to the Q1 Sales.xlsx workbook.
    On the Quarter1 worksheet, select the cell range A1:B21.
    In the Consolidate dialog box, verify the reference to '[Q1 Sales.xlsx]Quarter1'!$A$1:$B$21.
    Select Add.
    The reference is added to the All references list.

Repeat these steps to add references for each of the other quarterly sales workbooks.
    Select the text in the Reference text box, and press Delete.
    Note: Make sure you perform this step. If you do not remove the previous reference, Excel might not enable you to select another workbook.
    In the taskbar, switch to the next quarterly sales workbook.
    For example, if you previously added Q1 Sales.xlsx, now you should switch to Q2 Sales.xlsx.
    Select the cell range A1:B21 and in the Consolidate dialog box, select Add to add the reference to the All references list box.
    When you finish, four quarter references will be listed as shown.

Finish consolidation and confirm the results.
    In the Consolidate dialog box, check Top row, Left column, and Create links to source data as shown.
    Select OK.
    Switch to the My Consolidate.xlsx workbook.
    Observe that the quantities for each product have been consolidated from all quarterly sales workbooks.
    Note: You may need to widen one or more columns to get the data to display correctly.

View the details of the consolidated data.
    In the Outline pane, to the left of the column headings, select 2 to expand the outline and show details.
    Select cell D4 and observe the Formula Bar reference to Q1 Sales.xlsx.
    Examine the source for cells D5, D6, and D7.
    Save the My Consolidate.xlsx workbook, close all files, and exit Excel.

