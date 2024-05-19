# Intermediate I : Week 2

# Table of Contents

- [Combining Text (CONCAT, &)](#combining-text-concat-)
- [Changing Text Case (UPPER, LOWER, PROPER)](#changing-text-case-upper-lower-proper-)
- [Extracting Text (LEFT, MID, RIGHT)](#extracting-text-left-mid-right-)
- [Finding Text (FIND)](#finding-text-find-)
- [Date Calculations (NOW, TODAY, YEARFRAC)](#date-calculations-now-today-yearfrac-)

<!-- ## Combining Text (CONCAT, &)

Content for the "Combining Text" section goes here.

## Changing Text Case (UPPER, LOWER, PROPER)

Content for the "Changing Text Case" section goes here.

## Extracting Text (LEFT, MID, RIGHT)

Content for the "Extracting Text" section goes here.

## Finding Text (FIND)

Content for the "Finding Text" section goes here.

## Date Calculations (NOW, TODAY, YEARFRAC)

Content for the "Date Calculations" section goes here.


Table of Contents
- [Combining Text (CONCAT, &)](#Combining-Text-(CONCAT,&))
- [Changing Text Case (UPPER, LOWER, PROPER)](#Changing-Text-Case-(UPPER,LOWER,PROPER))
- [Extracting Text (LEFT, MID, RIGHT)](#Extracting-Text-(LEFT,MID,RIGHT))
- [Finding Text (FIND)](#Finding-Text-(FIND))
- [Date Calculations (NOW, TODAY, YEARFRAC)](#Date-Calculations-(NOW,TODAY,YEARFRAC))


-->

# Combining Text (CONCAT, &)

### Improving Employee Database

To enhance the current employee database, Uma has been tasked with listing the full name of each employee in column D and adding all employees' email addresses in column E. This task can be accomplished using the CONCAT function in Excel. However, for users with older versions of Excel or working on Mac, the CONCATENATE function can be used, which operates similarly to CONCAT.

### Using CONCAT Function

The CONCAT function concatenates or links together different text strings. It requires providing exact instructions, including the elements to be concatenated and any necessary separators. Here's how to use it:

1. Type an equal sign in the cell where the result should appear.
2. Start typing the first three letters of the CONCAT function name, then use the arrow keys to select it.
3. Define the first value, followed by a comma.
4. Include any necessary separators, such as spaces or punctuation, within quotation marks.
5. Repeat for each element to be concatenated.
6. Press Enter to see the result.
7. Double click the fill handle in the cell to copy the formula down.

#### CONCAT Function:
```excel
=CONCAT("First Name", ", ", "Last Name")
```

### Using Ampersand Sign

Alternatively, the Ampersand (&) sign can be used as a quick and simple way to link text together. This method involves directly combining text strings without using a specific function. Here's how to use it:

1. Type an equal sign in the cell where the result should appear.
2. Combine text strings by using the Ampersand sign (&).
3. Encase any fixed text elements, such as domain names, in quotation marks.
4. Press Enter to see the result.
5. Double click the fill handle in the cell to copy the formula down.

### Ampersand Sign Function

```excel
="First Name" & "." & "Last Name" & "@pushpin.com"
```

Both the CONCAT function and the Ampersand sign offer efficient ways to join data from separate columns. Remember to review the syntax for each approach and encase text elements in quotation marks when necessary. Additionally, check the Toolbox for more details on the syntax and consider using functions to fix any remaining issues in the database.


# Changing Text Case (UPPER, LOWER, PROPER)

1. **UPPER Function**:
   - Syntax: `UPPER(text)`
   - `UPPER("hello")` returns "HELLO".
   - Description: Converts all letters in a text string to uppercase.

2. **LOWER Function**:
   - Syntax: `LOWER(text)`
   - `LOWER("WORLD")` returns "world".
   - Description: Converts all letters in a text string to lowercase.

3. **PROPER Function**:
   - Syntax: `PROPER(text)`
   - `PROPER("john doe")` returns "John Doe"
   - Description: Capitalizes the first letter in each word of a text string and converts the rest of the letters to lowercase.


# Extracting Text (LEFT, MID, RIGHT)

1. **LEFT Function**:
   - Syntax: `LEFT(text, num_chars)`
   - Description: Extracts a specified number of characters from the left side of a text string.
   - Example: 
     - `=LEFT("Location", 2)` extracts the first 2 characters from the text "Location", resulting in "Lo".

2. **RIGHT Function**:
   - Syntax: `RIGHT(text, num_chars)`
   - Description: Extracts a specified number of characters from the right side of a text string.
   - Example:
     - `=RIGHT("Location", 4)` extracts the last 4 characters from the text "Location", resulting in "tion".

3. **MID Function**:
   - Syntax: `MID(text, start_num, num_chars)`
   - Description: Extracts a specified number of characters from a text string, starting at a specified position.
   - Example:
     - `=MID("Location", 5, 3)` extracts 3 characters from the text "Location", starting at the 5th position, resulting in "ati".




# Finding Text (FIND)

1. **FIND Function**:
   - Syntax: `FIND(find_text, within_text, [start_num])`
   - Description: Finds the starting position of a specified text within another text string. Optionally, you can specify a start position for the search.
   - Example:
     - `=FIND(" ", K4)` finds the position of the first space within the text in cell K4.
     - `=FIND(" ", K4, 5)` finds the position of the first space within the text in cell K4, starting the search from the 5th character.


# Date Calculations (NOW, TODAY, YEARFRAC)

1. **NOW Function**:
   - Syntax: `NOW()`
   - Description: Returns the current date and time.
   - Example: `=NOW()` returns the current date and time.

2. **TODAY Function**:
   - Syntax: `TODAY()`
   - Description: Returns the current date.
   - Example: `=TODAY()` returns the current date.

3. **YEARFRAC Function**:
   - Syntax: `YEARFRAC(start_date, end_date, [basis])`
   - Description: Calculates the fraction of a year between two dates.
   - Example: `=YEARFRAC(F4, TODAY())` calculates the number of years between the start date (cell F4) and today's date.


