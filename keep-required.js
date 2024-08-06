import fs from "fs";
import path from "path";
import XlsxPopulate from "xlsx-populate";
import dayjs from "dayjs";
import chalk from "chalk";

const sourceDir = "./Generated";

const targetDate = dayjs("2024-07-11 12:00:00");

const isDateClose = (date, target, toleranceMinutes = 8) => {
    return Math.abs(target.diff(date, "minute")) <= toleranceMinutes;
};

const processExcelFile = async (filePath) => {
    try {
        const workbook = await XlsxPopulate.fromFileAsync(filePath);
        const sheet = workbook.sheet(0);

        const usedRange = sheet.usedRange();
        const data = usedRange.value().slice(25);
        const lastRow = usedRange.endCell().rowNumber();

        for (let row = 26; row <= lastRow; row++) {
            sheet.cell(`A${row}`).value(null);
            sheet.cell(`B${row}`).value(null);
            sheet.cell(`C${row}`).value(null);
            sheet.cell(`D${row}`).value(null);
        }
        usedRange.endCell().rowNumber(700);
        console.log(data);

        const startIndex = data.findIndex((row) => {
            const cellValue = row[1];
            if (typeof cellValue === "number") {
                const cellDate = dayjs(XlsxPopulate.numberToDate(cellValue));
                return isDateClose(cellDate, targetDate);
            }
            if (typeof cellValue === "string") {
                return isDateClose(cellValue, targetDate);
            }
            return false;
        });

        if (startIndex === -1) {
            await workbook.toFileAsync(filePath);
            console.log(
                chalk.bgRed(
                    chalk.whiteBright(
                        `No matching date found in file: ${filePath}`
                    )
                )
            );
            return;
        }

        const matchingRows = data.slice(startIndex, startIndex + 673);

        matchingRows.forEach((row, rowIndex) => {
            row.forEach((cell, cellIndex) => {
                sheet.cell(rowIndex + 26, cellIndex + 1).value(cell);
            });
        });

        await workbook.toFileAsync(filePath);
        console.log(
            chalk.bgGreenBright(
                chalk.whiteBright(`Required data saved for: ${filePath}`)
            )
        );
    } catch (error) {
        console.error(`Error processing file ${filePath}: ${error.message}`);
    }
};

fs.readdir(sourceDir, (err, files) => {
    if (err) {
        return console.error(`Unable to scan directory: ${err}`);
    }

    files.forEach((file) => {
        const filePath = path.join(sourceDir, file);
        if (path.extname(file) === ".xlsx") {
            processExcelFile(filePath);
        }
    });
});
