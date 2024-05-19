# Intermediate I : Week 3

# Table of Contents

- [Introducing Named Ranges](#Introducing-Named-Ranges)
- [Creating Named Ranges](#Creating-Named-Ranges)
- [Managing Named Ranges](#Managing-Named-Ranges)
- [Named Ranges in Formulas](#Named-Ranges-in-Formulas)
- [Apply Names](#Apply-Names)

# Introducing Named Ranges
### Using Named Ranges in Excel for Calculations

1. **Without Named Ranges**:
   - Manually enter the formula, referencing cells directly.
     - For example: `=AnnualSalary * 0.09` (assuming the pension rate is 9%).

2. **Using Named Ranges**:
   - Assign meaningful names to cells or ranges to make formulas more understandable.
   - Click into the cell containing the value you want to name (e.g., pension rate).
   - Click on the Name Box (located to the left of the formula bar).
   - Type the desired name (e.g., `pension_rate`) and press Enter.
   - Use the defined name in your formula instead of cell references.
     - For example: `=AnnualSalary * pension_rate`.

### Naming Rules:
- Named ranges can contain letters, numbers, or underscores but no other characters (no hyphens, no spaces).
- They can be up to 255 characters long.
- Avoid using a cell reference as a named range to prevent conflicts with existing cell addresses (e.g., avoid naming a range "A1").
- Be cautious when naming ranges to avoid unintentionally using cell references (e.g., "top_five" might be interpreted as cell reference "TOP5").

### Benefits of Named Ranges:
- Provides meaningful names for cells or ranges, making formulas more readable and understandable.
- Eliminates the need for absolute cell references ($), improving formula clarity.
- Enhances efficiency, especially with large datasets or data on separate worksheets.



# Creating Named Ranges

### Process for Creating Named Ranges in Excel

1. **Using the Name Box**:
   - Select the range of cells you want to name (excluding headings).
   - Click into the Name Box.
   - Type the desired name for the range.
   - Press Enter to apply the name.

2. **Using Define Name**:
   - Click into any cell within the range you want to name (excluding headings).
   - Press Ctrl + Shift + Down Arrow to select the entire range.
   - Navigate to the Formulas tab on the Ribbon.
   - In the Defined Names group, click on Define Name.
   - Adjust the range if necessary in the Define Name dialog.
   - Choose the Scope for the named range.
   - Click OK to create the named range.

3. **Using Create from Selection**:
   - Select the entire range including the headings.
   - Press Ctrl + Shift + Down Arrow to select all data and headings.
   - Navigate to the Formulas tab on the Ribbon.
   - In the Defined Names group, click on Create from Selection or press Ctrl + Shift + F3.
   - Choose whether to name each column using the label in the top row or each row using the label in the left column.
   - Untick the option if you don't want to label rows.
   - Click OK to create the named ranges.

### Note:
- When using Create from Selection, include headings in the selection.
- The benefits of named ranges become apparent when using them in formulas and functions.


# Managing Named Ranges

### Using Name Manager in Excel

1. **Accessing Name Manager**:
   - Located on the Formulas tab in the Defined Names group.
   - Click on "Name Manager" to open the Name Manager dialog.

2. **Customizing Name Manager**:
   - Resize the dialog: Click and drag the bottom-right corner.
   - Adjust column widths: Click and drag between column headers.
   - Filter named ranges: Use the filter button to view specific scopes or names with errors.

3. **Managing Named Ranges**:
   - Deleting a named range:
     - Select the named range from the list and click "Delete".
     - Confirm the deletion and be cautious of potential formula errors.
   - Editing a named range:
     - Double-click or select the named range and click "Edit".
     - Modify the name or range reference as needed.
     - Note: Scope cannot be changed after creation.
   - Creating a new named range:
     - Click "New" to create a new named range.
     - Select cells in the workbook or adjust the range.
     - Choose the name and scope, then click "OK".

4. **Important Notes**:
   - Deleting or modifying named ranges can impact formulas referencing those names.
   - Be cautious when deleting or editing named ranges to avoid formula errors.
   - Ensure the correct scope is set when creating named ranges initially.


# Named Ranges in Formulas

### Utilizing Named Ranges in Excel Calculations

1. **Calculating Total Package Using Named Ranges**:
   - Enter the formula in the desired cell (e.g., P4): `=EmpID + Pension`.
   - Use named ranges instead of cell references (e.g., `EmpID` and `Pension`).
   - Press `Ctrl+Enter` to keep the cursor in the original cell after entering the formula.
   - Double-click the fill handle to copy the formula down.

2. **Benefits of Using Named Ranges**:
   - No need to use the mouse for selecting cells.
   - Ensures consistent and error-free formulas.
   - Increases efficiency, especially when copying formulas.

3. **Calculating Quick Facts Using Named Ranges**:
   - Use named ranges for calculations such as sum, average, minimum, and maximum.
   - Type the function (e.g., `=SUM`, `=AVERAGE`, `=MIN`, `=MAX`), then use named ranges instead of selecting cells manually.
   - Press `Tab` to select the highlighted named range.
   - Use the `F3` key to display a list of named ranges in the workbook, allowing quick selection without using the mouse.

4. **Benefits of Using Named Ranges in Calculations**:
   - Increases speed and efficiency by avoiding manual cell selection.
   - Simplifies formulas and improves readability.
   - Provides easy access to named ranges without navigating through the worksheet.

By incorporating named ranges into Excel calculations, users can streamline their workflow, improve formula accuracy, and enhance overall productivity.


# Apply Names

### Utilizing Apply Names in Excel Calculations

1. **Creating Named Ranges with Apply Names**:
   - Navigate to the cell containing the calculation.
   - Use the Create from Selection feature (`Ctrl+Shift+F3`) to quickly name the range.
   - Ensure that cell references are used in the calculation instead of named ranges initially.

2. **Using Apply Names Tool**:
   - Access the Apply Names tool from the Define Names group on the Formulas tab.
   - Click on Define Names and select Apply Names from the dropdown menu.
   - The Apply Names dialog will display a list of named ranges in the workbook.
   - Select the appropriate named range(s) to replace cell references in the calculation.
   - Click OK to apply the named ranges to the calculation.

3. **Benefits and Limitations**:
   - Apply Names replaces cell references with meaningful named ranges, enhancing readability.
   - However, it can only be used within the same worksheet as the named range.

4. **Exploring Use in Formula Feature**:
   - Access the Use in Formula feature from the dropdown menu under Define Names.
   - Use it as an alternative to the `F3` shortcut for selecting named ranges.
   - Explore additional options like Paste Names to document named ranges in the workbook.

By utilizing Apply Names, users can efficiently incorporate named ranges into existing calculations, enhancing clarity and simplifying formula management.



