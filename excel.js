import ExcelJS from "exceljs";
import cellValues from "./cellValues.js";
import dayjs from "dayjs";

const startDate = "11-Jul-2024 12:00:13 PM";
const location = "Sampling Room Cephalosporin";

const createExcelFile = async (sensorData, dataType) => {
    const saveDirectory = `./Final/${
        dataType === "temperatures" ? "Temp" : "Hum"
    } ${location}.xlsx`;
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Sheet 1");

    // Set font for the entire sheet
    worksheet.eachRow((row) => {
        row.eachCell((cell) => {
            cell.font = { name: "Calibri", size: 10 };
        });
    });

    // Set font for the entire sheet
    for (let i = 1; i <= 12; i++) {
        worksheet.getRow(i).height = 17 / 1.2;
    }

    // Set column width
    worksheet.getColumn("A").width = 102 / 7;
    for (let i = 2; i <= 15; i++) {
        worksheet.getColumn(i).width = 45 / 7;
    }

    cellValues.forEach(({ cell, value, font, alignment, border }) => {
        if (cell === "A4") {
            worksheet.getCell(cell).value = value[dataType];
        } else {
            worksheet.getCell(cell).value = value;
        }
        worksheet.getCell(cell).font = font
            ? font
            : {
                  name: "Calibri",
                  size: 10,
                  bold: true,
              };
        if (alignment) worksheet.getCell(cell).alignment = alignment;
        if (border) allBorder(worksheet.getCell(cell));
    });

    setBorderOut(worksheet, 1, 2);
    setBorderOut(worksheet, 4, 6);
    worksheet.getCell("A" + 5).border = {
        left: { style: "thin", color: { argb: "000000" } },
    };
    worksheet.getCell("O" + 5).border = {
        right: { style: "thin", color: { argb: "000000" } },
    };
    setBorderOut(worksheet, 9, 11);
    worksheet.getCell("A" + 10).border = {
        left: { style: "thin", color: { argb: "000000" } },
    };
    worksheet.getCell("O" + 10).border = {
        right: { style: "thin", color: { argb: "000000" } },
    };

    let startRow = 14;
    let startFormulaRow = 16;
    let startColNum = 2;
    sensorData.forEach((logger, ind) => {
        let currentDate = dayjs(startDate);
        let startCol = String.fromCharCode(64 + startColNum);
        let logHead = worksheet.getCell(startCol + startRow);
        let logUnit = worksheet.getCell(startCol + (startRow + 1));
        logHead.value = logger.fileName.split(".")[0];
        logUnit.value = dataType === "temperatures" ? "[°C]" : "[%RH]";
        logHead.font = titleStyles.font;
        logUnit.font = titleStyles.font;
        logHead.alignment = titleStyles.alignment;
        logUnit.alignment = titleStyles.alignment;
        allBorder(logHead);
        allBorder(logUnit);
        logger[dataType].forEach((value, index) => {
            if (startColNum === 2) {
                if (index === 0) {
                    let locationCell = worksheet.getCell(
                        "A" + (startRow + index)
                    );
                    locationCell.value = location;
                    locationCell.font = {
                        name: "Calibri",
                        size: 8.5,
                        bold: true,
                    };
                    locationCell.alignment = {
                        horizontal: "center",
                        vertical: "middle",
                        wrapText: true,
                    };
                    allBorder(locationCell);
                }
                if (index === 1) {
                    let dateHeadCell = worksheet.getCell(
                        "A" + (startRow + index)
                    );
                    dateHeadCell.value = "Date/Time";
                    dateHeadCell.font = {
                        name: "Calibri",
                        size: 8,
                        bold: true,
                    };
                    dateHeadCell.alignment = {
                        horizontal: "center",
                        vertical: "middle",
                    };
                    allBorder(dateHeadCell);
                }
                let dateCell = worksheet.getCell("A" + (startRow + 2 + index));
                dateCell.value = currentDate.format("DD-MMM-YYYY hh:mm A");
                dateCell.alignment = {
                    horizontal: "center",
                    vertical: "middle",
                };
                dateCell.font = {
                    name: "Calibri",
                    size: 8,
                };
                allBorder(dateCell);
                currentDate = currentDate.add(15, "minutes");
            }
            let valueCell = worksheet.getCell(
                startCol + (startRow + 2 + index)
            );
            valueCell.value = value;
            valueCell.numFmt = "0.0";
            valueCell.font = logStyles.font;
            valueCell.alignment = logStyles.alignment;
            allBorder(valueCell);
        });
        startColNum = startColNum + 1;
        if (startColNum > 15 && sensorData[ind + 1]) {
            startColNum = 2;
            startRow += 677;
        }
        if (!sensorData[ind + 1]) {
            startColNum = 2;
            startRow += 677;
        }
    });

    sensorData.forEach((logger, ind) => {
        let startCol = String.fromCharCode(64 + startColNum);
        let logHead = worksheet.getCell(startCol + startRow);
        let minCell = worksheet.getCell(startCol + (startRow + 1));
        let maxCell = worksheet.getCell(startCol + (startRow + 2));
        let avgCell = worksheet.getCell(startCol + (startRow + 3));
        if (startColNum === 2) {
            let locationCell = worksheet.getCell("A" + startRow);
            let minHead = worksheet.getCell("A" + (startRow + 1));
            let maxHead = worksheet.getCell("A" + (startRow + 2));
            let avgHead = worksheet.getCell("A" + (startRow + 3));
            locationCell.value = location;
            minHead.value = "Minimum";
            maxHead.value = "Maximum";
            avgHead.value = "Average";
            locationCell.font = {
                name: "Calibri",
                size: 8.5,
                bold: true,
            };
            minHead.font = {
                name: "Calibri",
                size: 10,
                bold: true,
            };
            maxHead.font = {
                name: "Calibri",
                size: 10,
                bold: true,
            };
            avgHead.font = {
                name: "Calibri",
                size: 10,
                bold: true,
            };
            locationCell.alignment = {
                horizontal: "center",
                vertical: "middle",
                wrapText: true,
            };
            minHead.alignment = {
                horizontal: "center",
                vertical: "middle",
            };
            maxHead.alignment = {
                horizontal: "center",
                vertical: "middle",
            };
            avgHead.alignment = {
                horizontal: "center",
                vertical: "middle",
            };
            allBorder(locationCell);
            allBorder(minHead);
            allBorder(maxHead);
            allBorder(avgHead);
        }
        logHead.value = logger.fileName.split(".")[0];
        logHead.font = titleStyles.font;
        logHead.alignment = titleStyles.alignment;
        minCell.value = {
            formula: `MIN(${startCol + startFormulaRow}:${
                startCol + (startFormulaRow + 672)
            })`,
        };
        maxCell.value = {
            formula: `MAX(${startCol + startFormulaRow}:${
                startCol + (startFormulaRow + 672)
            })`,
        };
        avgCell.value = {
            formula: `AVERAGE(${startCol + startFormulaRow}:${
                startCol + (startFormulaRow + 672)
            })`,
        };
        minCell.font = {
            name: "Calibri",
            size: 8,
        };
        maxCell.font = {
            name: "Calibri",
            size: 8,
        };
        avgCell.font = {
            name: "Calibri",
            size: 8,
        };
        minCell.numFmt = "0.0";
        maxCell.numFmt = "0.0";
        avgCell.numFmt = "0.0";
        minCell.alignment = titleStyles.alignment;
        maxCell.alignment = titleStyles.alignment;
        avgCell.alignment = titleStyles.alignment;
        allBorder(logHead);
        allBorder(minCell);
        allBorder(maxCell);
        allBorder(avgCell);

        startColNum += 1;
        if (startColNum > 15 && sensorData[ind + 1]) {
            startColNum = 2;
            startRow += 5;
            startFormulaRow += 677;
        }
    });

    // Save the workbook
    await workbook.xlsx.writeFile(saveDirectory);
    console.log(`${saveDirectory} created successfully!`);
};

let titleStyles = {
    font: {
        name: "Calibri",
        size: 8,
        bold: true,
    },
    alignment: {
        horizontal: "center",
        vertical: "middle",
    },
};
let logStyles = {
    font: {
        name: "Calibri",
        size: 10,
    },
    alignment: {
        horizontal: "center",
        vertical: "middle",
    },
};

const allBorder = (cell) => {
    cell.border = {
        top: { style: "thin", color: { argb: "000000" } },
        right: { style: "thin", color: { argb: "000000" } },
        bottom: { style: "thin", color: { argb: "000000" } },
        left: { style: "thin", color: { argb: "000000" } },
    };
};

const setBorderOut = (worksheet, upRow, downRow) => {
    for (let i = 1; i <= 15; i++) {
        const col = String.fromCharCode(64 + i);
        if (i === 1) {
            worksheet.getCell(col + upRow).border = {
                top: { style: "thin", color: { argb: "000000" } },
                left: { style: "thin", color: { argb: "000000" } },
            };
            worksheet.getCell(col + downRow).border = {
                bottom: { style: "thin", color: { argb: "000000" } },
                left: { style: "thin", color: { argb: "000000" } },
            };
        } else if (i > 1 && i < 15) {
            worksheet.getCell(col + upRow).border = {
                top: { style: "thin", color: { argb: "000000" } },
            };
            worksheet.getCell(col + downRow).border = {
                bottom: { style: "thin", color: { argb: "000000" } },
            };
        } else if (i === 15) {
            worksheet.getCell(col + upRow).border = {
                top: { style: "thin", color: { argb: "000000" } },
                right: { style: "thin", color: { argb: "000000" } },
            };
            worksheet.getCell(col + downRow).border = {
                bottom: { style: "thin", color: { argb: "000000" } },
                right: { style: "thin", color: { argb: "000000" } },
            };
        }
    }
};

const createTemperatureFile = (sensorData) => {
    createExcelFile(sensorData, "temperatures");
};

const createHumidityFile = (sensorData) => {
    createExcelFile(sensorData, "humidities");
};

// Example usage
// createTemperatureFile(sensorData).catch(console.error);
// createHumidityFile(sensorData).catch(console.error);

export { createTemperatureFile, createHumidityFile };
