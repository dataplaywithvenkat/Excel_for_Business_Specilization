# Intermediate I

Table of Contents
- [Multiple Worksheets](#Multiple-Worksheets)
- [3D Formulas](#3D-Formulas)
- [Linking Workbooks](#Linking-Workbooks)
- [Consolidating by Position](#Consolidating-by-Position)
- [Consolidating by Category (Reference)](#Consolidating-by-Category-(Reference))


# Multiple Worksheets

### 1. Adding a New Worksheet
1. Click the `+` button to the right of the worksheet tabs.
2. A new blank worksheet will be added to the right of the current sheet.

### 2. Moving a Worksheet
1. Click and hold the worksheet tab you want to move.
2. Drag it to the desired location.
3. Look for the black arrow indicator before releasing the mouse button to place the sheet.

### 3. Copying a Worksheet
#### Method 1: Using Right-Click Menu
1. Right-click on the worksheet tab you want to copy.
2. Select `Move or Copy`.
3. Choose the destination for the new copy and check `Create a copy`.
4. Click `OK`.

#### Method 2: Using Drag and Drop
1. Hold the `Ctrl` key.
2. Click and drag the worksheet tab to the desired location.
3. Release the mouse button to create a copy.

### 4. Deleting a Worksheet
1. Right-click on the worksheet tab you want to delete.
2. Select `Delete`.
3. Confirm the deletion (note: this action cannot be undone).

### 5. Renaming a Worksheet
1. Double-click on the worksheet tab.
2. Type the new name and press `Enter`.

### 6. Color-Coding Worksheet Tabs
1. Right-click on the worksheet tab you want to color-code.
2. Select `Tab Color`.
3. Choose the desired color.

### 7. Grouping Worksheets
1. Click on the first worksheet tab you want to group.
2. Hold the `Shift` key and click on the last worksheet tab to group consecutive sheets.
3. To make changes to all selected sheets simultaneously, edit any cell.
4. Click away from the grouped tabs to ungroup them.

### 8. Hiding and Unhiding Worksheets
1. To hide a worksheet:
   - Right-click on the worksheet tab.
   - Select `Hide`.

2. To unhide a worksheet:
   - Right-click on any worksheet tab.
   - Select `Unhide`.
   - Choose the worksheet to unhide and click `OK`.

# 3D Formulas


### 1. Conventional Method
1. Click in the cell where you want the total to appear.
2. Type `=` to start the formula.
3. Click on the cell in the first worksheet (e.g., Sean) that contains the value you want to include.
4. Type `+` and then click on the corresponding cell in the next worksheet (e.g., Uma).
5. Repeat for all relevant worksheets.
6. Press `Enter` to complete the formula.

**Note:** This method is time-consuming and prone to errors, especially with many worksheets.

### 2. Using 3D Formulas
A more efficient way to sum values across multiple worksheets is by using 3D formulas.

#### Example: Summing Miles Driven
1. Click in the cell where you want the total to appear.
2. Type `=SUM(` to start the SUM function.
3. Click on the first worksheet tab (e.g., Sean) and then click on the cell you want to sum (e.g., C7).
4. Hold the `Shift` key and click on the last worksheet tab (e.g., Carlos). This selects all worksheets between Sean and Carlos.
5. Ensure your formula reads something like `=SUM(Sean:Carlos!C7)`.
6. Press `Enter` to complete the formula.

### 3. Copying Formulas Using Fill Handle
1. After entering the 3D formula, use the fill handle to copy the formula across adjacent cells.
2. Double-click the fill handle to copy the formula down through the column.

#### Example: Summing Lodging and Meals
1. Click in the cell where you want the total to appear (e.g., C17).
2. Type `=SUM(`.
3. Click on the first worksheet tab (e.g., Sean) and then click on the cell you want to sum (e.g., C17).
4. Hold the `Shift` key and click on the last worksheet tab (e.g., Carlos).
5. Press `Enter`.
6. Use the fill handle to copy the formula across and down.

### 4. Copying Formulas for Identical Structures
1. If the structure of the workbooks is identical, you can copy and paste formulas using `Ctrl+C` and `Ctrl+V`.

### Limitations and Considerations
- **Worksheet Order:** Ensure that worksheets remain in the same order if using 3D formulas. Moving sheets may exclude some data from the calculation.
- **Identical Structure:** The workbooks you are summing must have an identical structure for the 3D formula to work correctly.

# Linking Workbooks

### 1. Preparing Your Workbooks
1. **Open Workbooks**: Ensure you have opened `W1_LinkingWorkbooks.xlsx` and closed any other open workbooks.
2. **Open Multiple Workbooks**:
   - Press `Ctrl+O`.
   - Navigate to where your workbooks are stored.
   - Select `W1_ExpensesQ1.xlsx`, hold `Shift`, select Q3, and click `Open`.

### 2. Arrange Workbooks for Easy Viewing
1. **Minimize Ribbon**: Double-click the Home tab to minimize the ribbon.
2. **Arrange All Workbooks**:
   - Use the "Tell me what you want to do" feature by typing "Arrange All".
   - Select "Tile" and click "OK" to view all workbooks simultaneously.

### 3. Linking Data Between Workbooks
1. **Create a Link**:
   - Click in the cell where you want the total to appear (e.g., `C7`).
   - Type `=` to start the formula.
   - Switch to the Q1 workbook, click the cell with the value to link (e.g., `C7`).
   - Excel automatically uses absolute references (e.g., `$C$7`). Press `F4` three times to remove the absolute reference.
   - Press `Tab` to move to the next cell.

2. **Repeat for Other Workbooks**:
   - Type `=` in the target cell.
   - Switch to the Q2 workbook, click the cell with the value to link, and press `F4` three times.
   - Press `Tab` to move to the next cell.
   - Repeat for Q3.

### 4. Verify and Update Links
1. **Test Links**:
   - Change a value in one of the source workbooks (e.g., change Carlos' miles driven to 400 in Q1).
   - Check if the linked cell in the summary workbook updates automatically.
   - Undo the change to revert to the original value.

2. **Copy Formulas**:
   - Select the linked cells and use the fill handle to drag and copy the formulas across adjacent cells.
   - Use "Fill without formatting" to maintain original formatting.

### 5. Managing Workbook Links
1. **Edit Links**:
   - Go to the Data tab and select "Edit Links".
   - Here, you can see all linked workbooks.
   - If a link is broken, click "Change Source" to update the link.

2. **Break Links**:
   - Only break links if necessary, as it is irreversible.
   - Select the link to break and click "Break Link".
   - Confirm the warning. The cell formula will be replaced by the last known value.

# Consolidating by Position

### 1. Preparing Your Workbooks
1. **Open Workbooks**: Ensure you have the necessary workbooks open (`W1_LinkingWorkbooks.xlsx`, and the regional workbooks for Melbourne, Perth, and Sydney).
2. **Open Multiple Workbooks**:
   - Press `Ctrl+O`.
   - Navigate to where your workbooks are stored.
   - Select the Melbourne, Perth, and Sydney workbooks using `Ctrl` and click `Open`.

### 2. Arrange Workbooks for Easy Viewing
1. **Minimize Ribbon**: Double-click the Home tab to minimize the ribbon.
2. **Arrange All Workbooks**:
   - Go to the View tab and select "Arrange All".
   - Choose "Tile" and click "OK" to view all workbooks simultaneously.

### 3. Using the Consolidate Tool
1. **Start Consolidation**:
   - Click in the cell where you want the consolidation to begin (e.g., `C7`).
   - Go to the Data tab, find the Data Tools group, and click on "Consolidate".

2. **Configure Consolidation**:
   - In the Consolidate dialog, select the function you want to use (e.g., Sum).
   - Click in the reference box.
   - Select the range from the first workbook (e.g., Perth, `C7:F28`), then click "Add".
   - Repeat this for the other workbooks (Melbourne and Sydney), ensuring you click "Add" after each selection.

3. **Complete Consolidation**:
   - Once all ranges are added, click "OK".
   - The data from the selected workbooks will be summed and placed into the new workbook. Note that this is a snapshot; there are no formulas, just values.

### 4. Updating Consolidated Data
1. **Manual Update**:
   - If the source data changes, the consolidated data will not update automatically. 
   - To update, click back into the cell where the consolidation starts (e.g., `C7`).
   - Go to Data > Data Tools > Consolidate, and click "OK" again.

### 5. Creating Linked Consolidations
1. **Enable Links**:
   - Duplicate the worksheet to preserve the original data (hold `Ctrl` and drag the worksheet tab).
   - Clear the data in the new sheet.
   - Run the consolidation again, but this time, check the box "Create links to source data" before clicking "OK".

2. **Manage Linked Consolidations**:
   - The linked consolidation will create an outline with plus signs (+) next to rows.
   - Clicking the plus sign will expand the details, showing data pulled from each workbook.
   - Note that linked consolidations cannot be undone easily. To remove, select affected cells, go to Data > Outline > Ungroup, and then "Clear Outline".

### 6. Important Considerations
1. **Identical Workbook Structure**:
   - Ensure all workbooks have an identical structure when using Consolidate by Position.
2. **Potential Issues with Links**:
   - Consider the implications of linked data, as it can be more complex and harder to manage.


# Consolidating by Category (Reference)


### 1. Preparing Your Workbooks
1. **Open Workbooks**: Ensure you have the necessary workbooks open (`W1_LinkingWorkbooks.xlsx`, and the regional workbooks for Melbourne, Perth, and Sydney).
2. **Open Multiple Workbooks**:
   - Press `Ctrl+O`.
   - Navigate to where your workbooks are stored.
   - Select the Melbourne, Perth, and Sydney workbooks using `Ctrl` and click `Open`.

### 2. Arrange Workbooks for Easy Viewing
1. **Minimize Ribbon**: Double-click the Home tab to minimize the ribbon.
2. **Arrange All Workbooks**:
   - Go to the View tab and select "Arrange All".
   - Choose "Tile" and click "OK" to view all workbooks simultaneously.

### 3. Using the Consolidate Tool
1. **Start Consolidation**:
   - Click in the cell where you want the consolidation to begin (e.g., `C7`).
   - Go to the Data tab, find the Data Tools group, and click on "Consolidate".

2. **Configure Consolidation**:
   - In the Consolidate dialog, select the function you want to use (e.g., Sum).
   - Click in the reference box.
   - Select the range from the first workbook (e.g., Perth, `C7:F28`), then click "Add".
   - Repeat this for the other workbooks (Melbourne and Sydney), ensuring you click "Add" after each selection.

3. **Complete Consolidation**:
   - Once all ranges are added, click "OK".
   - The data from the selected workbooks will be summed and placed into the new workbook. Note that this is a snapshot; there are no formulas, just values.

### 4. Updating Consolidated Data
1. **Manual Update**:
   - If the source data changes, the consolidated data will not update automatically. 
   - To update, click back into the cell where the consolidation starts (e.g., `C7`).
   - Go to Data > Data Tools > Consolidate, and click "OK" again.

### 5. Creating Linked Consolidations
1. **Enable Links**:
   - Duplicate the worksheet to preserve the original data (hold `Ctrl` and drag the worksheet tab).
   - Clear the data in the new sheet.
   - Run the consolidation again, but this time, check the box "Create links to source data" before clicking "OK".

2. **Manage Linked Consolidations**:
   - The linked consolidation will create an outline with plus signs (+) next to rows.
   - Clicking the plus sign will expand the details, showing data pulled from each workbook.
   - Note that linked consolidations cannot be undone easily. To remove, select affected cells, go to Data > Outline > Ungroup, and then "Clear Outline".

### 6. Important Considerations
1. **Identical Workbook Structure**:
   - Ensure all workbooks have an identical structure when using Consolidate by Position.
2. **Potential Issues with Links**:
   - Consider the implications of linked data, as it can be more complex and harder to manage.

## Conclusion
The Consolidate tool is a powerful feature for combining data from multiple workbooks without the complications of linking. Practice using it to streamline your data management processes in collaborative environments.

### 1. Preparing Your Workbooks
1. **Open Workbooks**:
   - Press `Ctrl+O`.
   - Click Browse and select the leave days workbooks for Melbourne, Perth, and Sydney.
   - Click `Open`.

### 2. Arrange Workbooks for Easy Viewing
1. **Minimize Ribbon**: Double-click the Home tab to minimize the ribbon.
2. **Arrange All Workbooks**:
   - Go to the View tab and select "Arrange All".
   - Press "Enter" to tile the workbooks for simultaneous viewing.

### 3. Using the Consolidate Tool with Labels
1. **Start Consolidation**:
   - Click in the cell where the consolidation should begin (e.g., `B5`).
   - Go to the Data tab, find the Data Tools group, and click on "Consolidate".

2. **Configure Consolidation**:
   - In the Consolidate dialog, choose the function (e.g., Sum).
   - Click in the reference box.
   - Select the data range in the first workbook (e.g., Melbourne, `B5:G10`), and click "Add".
   - Repeat this for the other workbooks (Sydney, `B5:G11` and Perth, `B5:G9`), clicking "Add" after each selection.

3. **Use Labels for Consolidation**:
   - To consolidate by reference, check the box for "Left column" under "Use labels in".
   - Click "OK".

### 4. Verifying the Consolidation
1. **Check the Consolidated Data**:
   - Verify the sum for a sample department (e.g., Accounting).
   - Ensure the numbers match the expected totals from the individual workbooks.

### 5. Using Different Functions
1. **Calculate Averages**:
   - Scroll down and click in a new cell (e.g., `B15`).
   - Go to Data > Data Tools > Consolidate.
   - Choose "Average" instead of "Sum".
   - Ensure all other settings are retained, including the "Left column" labels.
   - Click "OK".

2. **Check the Average**:
   - Verify the average leave days for each department.







