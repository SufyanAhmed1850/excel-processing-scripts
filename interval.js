import fs from "fs";
import path from "path";
import XlsxPopulate from "xlsx-populate";

const sauceDir = "./Sauce";
const generatedDir = "./Generated";

if (!fs.existsSync(generatedDir)) {
    fs.mkdirSync(generatedDir);
}

const processExcelFile = async (filePath, outputDir) => {
    try {
        const workbook = await XlsxPopulate.fromFileAsync(filePath);
        const sheet = workbook.sheet(0);
        const usedRange = sheet.usedRange();
        const data = usedRange.value().slice(25);
        const lastRow = usedRange.endCell().rowNumber();

        const removeRandomArray = (arr) => {
            const result = [];
            for (let i = 0; i < arr.length; i += 3) {
                const group = arr.slice(i, i + 3);
                if (group.length > 0) {
                    const indexToRemove = Math.floor(
                        Math.random() * (group.length - 1)
                    );
                    group.splice(indexToRemove, 1);
                    result.push(...group);
                }
            }
            return result;
        };
        const modifiedData = removeRandomArray(data);

        for (let row = 26; row <= lastRow; row++) {
            sheet.cell(`A${row}`).value(null);
            sheet.cell(`B${row}`).value(null);
            sheet.cell(`C${row}`).value(null);
            sheet.cell(`D${row}`).value(null);
        }

        modifiedData.forEach((row, index) => {
            row.forEach((cell, cellIndex) => {
                sheet.cell(index + 26, cellIndex + 1).value(cell);
            });
        });

        const outputFilePath = path.join(outputDir, path.basename(filePath));
        await workbook.toFileAsync(outputFilePath);
        console.log(`Processed and saved: ${outputFilePath}`);
    } catch (error) {
        console.error(`Error processing file ${filePath}: ${error.message}`);
    }
};

fs.readdir(sauceDir, (err, files) => {
    if (err) {
        return console.error(`Unable to scan directory: ${err}`);
    }
    files.forEach((file) => {
        const filePath = path.join(sauceDir, file);
        if (path.extname(file) === ".xlsx" || path.extname(file) === ".xls") {
            processExcelFile(filePath, generatedDir);
        }
    });
});
