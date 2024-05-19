# Intermediate I : Week 4

# Table of Contents

- [COUNT functions](#COUNT-functions)
- [Counting with Criteria (COUNTIFS)](#COUNTIFS)
- [Adding with Criteria (SUMIFS)](#SUMIFS)
- [Sparklines](#Sparklines)
- [Advanced Charting](#Advanced-Charting)
- [Trendlines](#Trendlines)

# COUNT functions

# Using Count Functions in Excel

Sometimes, instead of summing up values, we need to count them. Excel provides various count functions to help with this task. In this guide, we'll explore how to use the COUNT, COUNTA, and COUNTBLANK functions to analyze data.

## Scenario:
Uma has been tasked by the sales team to manage a large dataset. The team needs to understand several key metrics, including the number of orders shipped, the total number of orders placed, and instances where order priority was not entered.

## Steps:

1. **Name Ranges:**
   - To start, it's essential to name our ranges for easier reference.
   - Select the dataset by clicking somewhere within it and pressing `Ctrl + A`.
   - Navigate to the `Formulas` tab and click on `Create from Selection`.
   - Ensure only the top row is checked and click `OK`.

2. **Count Shipped Orders:**
   - Determine the number of orders that have been shipped.
   - Click on cell `O3` and enter the formula `=COUNT()`.
   - Tab to select the `COUNT` function, which counts cells containing numeric values.
   - Excel will automatically count each cell in Column O with a number and ignore empty cells.

3. **Count Placed Orders:**
   - Calculate the total number of orders placed.
   - Click on cell `B3` and enter the formula `=COUNTA()`.
   - Specify the range `order number` to include all non-empty cells, both numerical and alphanumeric.

4. **Identify Missing Order Priorities:**
   - Determine instances where order priority was not entered.
   - Click on cell `J3` and enter the formula `=COUNTBLANK()`.
   - Choose the range `Order Priority` to count the blank cells.

## Results:
- **Shipped Orders:** 1037
- **Placed Orders:** 1039
- **Missing Order Priorities:** 2

By leveraging count functions like `COUNT`, `COUNTA`, and `COUNTBLANK`, we can efficiently analyze data and generate useful summaries in Excel.


# COUNTIFS
# Using COUNTIFS Function in Excel

In Excel, the `COUNTIFS` function is incredibly useful for counting cells within a range based on multiple criteria. In this guide, we'll learn how to use `COUNTIFS` to summarize data based on specific conditions.

## Scenario:
Uma needs to analyze the breakdown of orders from different states and customer types in a sales dashboard template.

## Steps:

1. **Open Sales Dashboard Template:**
   - Navigate to the "Sales Dash" sheet, where we'll summarize our data.

2. **Understanding `COUNTIFS`:**
   - `COUNTIFS` is the function we'll use for this task.
   - Unlike `COUNTIF`, `COUNTIFS` can handle multiple criteria.

3. **Count Orders by State:**
   - To start, we'll count the number of orders in each of the three states: Victoria (VIC), New South Wales (NSW), and Western Australia (WA).
   - Use `COUNTIFS` function with two arguments: range and criteria.
   - Define the range as the state column in the dataset.
   - Supply the criteria (state names) to count cells meeting the specified condition.
   - **Syntax:** `=COUNTIFS(range1, criteria1, [range2, criteria2], ...)`

4. **Count Orders by Customer Type:**
   - Next, we'll determine the number of orders per customer type.
   - Apply `COUNTIFS` function to count orders based on different customer types.
   - Specify the range as the customer type column and criteria as the specific customer type.
   - **Syntax:** `=COUNTIFS(range1, criteria1, [range2, criteria2], ...)`

5. **Count Orders with Quantity Over 40:**
   - Finally, we'll identify the number of orders with quantities over 40.
   - Use `COUNTIFS` function with the criteria ">40" to count cells where the order quantity exceeds 40.
   - **Syntax:** `=COUNTIFS(range1, criteria1, [range2, criteria2], ...)`

## Results:
- **Orders by State:**
  - VIC: 289 orders
  - NSW: 646 orders
  - WA: 104 orders

- **Orders by Customer Type:**
  - Home Office: [Number of orders]
  - Corporate: [Number of orders]
  - Consumer: [Number of orders]

- **Orders with Quantity Over 40:** 238 orders

By utilizing the `COUNTIFS` function, Uma efficiently summarizes large datasets and obtains valuable insights. In the next section, we'll explore the `SUMIFS` function for summing ranges based on specified criteria.



# SUMIFS

# Sparklines

# Advanced Charting

# Trendlines




