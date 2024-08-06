# Excel Processing Scripts

This repository contains a collection of Node.js scripts for processing and manipulating Excel files. These scripts are designed to work together to transform raw data into a specific format for temperature and humidity monitoring.

## Files and Their Purposes

1. **cellValues.js**
   - Purpose: Defines cell values and styles for the final Excel output.
   - Used by: excel.js

2. **copy.js**
   - Purpose: Reads Excel files from a source directory, processes the data, and copies it to a new workbook.
   - Creates a consolidated file with temperature and humidity data from multiple sensors.

3. **create.js**
   - Purpose: Similar to copy.js, but creates separate files for temperature and humidity data.
   - Uses the createTemperatureFile and createHumidityFile functions from excel.js.

4. **excel.js**
   - Purpose: Contains functions to create formatted Excel files for temperature and humidity data.
   - Exports createTemperatureFile and createHumidityFile functions.
   - Used by: create.js

5. **interval.js**
   - Purpose: Processes Excel files from a source directory, removes random data points, and saves the modified files.
   - Used to create a subset of the original data.

6. **keep-required.js**
   - Purpose: Processes Excel files, keeps only the required date range of data, and saves the modified files.
   - Filters data to match a specific target date and time range.

## Flow of Operations

1. Raw data files are placed in the ./Sauce directory.
2. interval.js processes these files, removing some data points, and saves the results in ./Generated.
3. keep-required.js further processes the files in ./Generated, keeping only the required date range.
4. create.js (or copy.js) reads the processed files from ./Generated and creates the final output files in ./Final.

## Prerequisites

To use these scripts, you need:

1. Node.js installed on your system.
2. The following npm packages installed:
   - xlsx-populate
   - exceljs
   - dayjs
   - chalk (for colored console output)

You can install these packages using:

```
npm install xlsx-populate exceljs dayjs chalk
```

## Usage

1. Place your raw Excel files in the ./Sauce directory.
2. Run the scripts in the following order:

   ```
   node interval.js
   node keep-required.js
   node create.js
   ```

   (or use copy.js instead of create.js if you want a single consolidated file)

3. The final processed files will be in the ./Final directory.

## Notes

- Make sure you have the necessary read/write permissions for the directories used by the scripts.
- The scripts assume a specific structure for the input Excel files. Ensure your data follows the expected format.
- Modify the targetDate in keep-required.js if you need to change the date range of the kept data.
- Adjust the cell values and styles in cellValues.js if you need to change the formatting of the output files.

## Caution

These scripts modify Excel files. Always keep backups of your original data before running the scripts.
