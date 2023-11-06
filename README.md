# CMS-Organizer-
# Excel VBA Macros

This repository contains a collection of Excel VBA macros designed to perform various data processing tasks on spreadsheet data.

## Macros Description

### `DummyCMSREAD`

This macro processes data in a worksheet named "Dummy." It performs the following tasks:

- **Column B**: Truncates alphanumeric content to the first two characters.
- **Column C**: If there are more than six alphanumeric characters, the excess is moved to Column D.
- **Column D**: If there are more than three alphanumeric characters, the excess is moved to Column E.
- **Column E**: If there are more than six alphanumeric characters, the first six remain, and the rest are moved to Column D, after a confirmation via MsgBox if Column D is not empty.

The script optimizes performance by disabling screen updating and automatic calculations.

### `AlphanumericCount`

A function that returns the count of alphanumeric characters in a given string.

### `OnlyAlphanumeric`

A function that extracts and returns only the alphanumeric characters from a given string.

### `CopyTwoDigitCode`

This macro copies a two-digit code from Column B to the same column in the following rows, only if the subsequent rows in Column C contain a six-digit code.

### `CheckAndDeleteEmptyColumns`

Checks for and deletes any completely empty columns starting from column N to the last column with data in Row 1. If non-empty columns are found, it highlights them and selects the first non-empty cell.

### `FilterFormulaNAandExcludeBlanks`

Applies filters to an assumed data range from `A1:E{lastRowA}`:
- Column A: Filters to show cells with `#N/A` errors.
- Column E: Filters to exclude blank cells.

## Installation
