import XlsxPopulate from "xlsx-populate";
import fs from "fs";
import { createTemperatureFile, createHumidityFile } from "./excel.js";

const sauceDirectory = "./Generated/";
let sensorData = [];

fs.readdir(sauceDirectory, (err, files) => {
    if (err) {
        console.error("Error reading sauce directory:", err);
        return;
    }
    files.sort().forEach((file) => {
        if (file.endsWith(".xlsx")) {
            const filePath = sauceDirectory + file;
            processExcelFile(filePath, file, files);
        }
    });
});

const processExcelFile = async (source, file, files) => {
    const workbook = await XlsxPopulate.fromFileAsync(source);
    const sheet = workbook.sheet(0);
    let i = sensorData.length;
    sensorData[i] = {
        fileName: file.trim().split(" ")[0],
        temperatures: [],
        humidities: [],
    };
    console.log(file);

    const lastRow = 673;
    for (let j = 1; j <= lastRow; j++) {
        const temperature = sheet.cell(`C${j + 25}`).value();
        const humidity = sheet.cell(`D${j + 25}`).value();
        temperature && sensorData[i].temperatures.push(temperature);
        humidity && sensorData[i].humidities.push(humidity);
    }

    if (i + 1 === files.length) {
        handleEmptySensors();
        sortSensorData();
        // copyDataToNewWorkbook();
        createTemperatureFile(sensorData);
        createHumidityFile(sensorData);
    }
};

const handleEmptySensors = () => {
    const validSensorsWithTemperatures = sensorData.filter(
        (sensor) => sensor.temperatures.length
    );
    const validSensorsWithHumidities = sensorData.filter(
        (sensor) => sensor.humidities.length
    );

    sensorData.forEach((sensor) => {
        if (
            sensor.temperatures.length < 670 &&
            validSensorsWithTemperatures.length > 1
        ) {
            let randomSensor;
            do {
                let randomSensorNum = Math.floor(
                    Math.random() * validSensorsWithTemperatures.length
                );
                randomSensor = validSensorsWithTemperatures[randomSensorNum];
            } while (!randomSensor || !randomSensor.temperatures.length);

            let missingData = generateRandomNumber(
                randomSensor.temperatures,
                673
            );
            sensor.temperatures = missingData;
        }

        if (
            sensor.humidities.length < 670 &&
            validSensorsWithHumidities.length > 1
        ) {
            let randomSensor;
            do {
                let randomSensorNum = Math.floor(
                    Math.random() * validSensorsWithHumidities.length
                );
                randomSensor = validSensorsWithHumidities[randomSensorNum];
            } while (!randomSensor || !randomSensor.humidities.length);

            let missingData = generateRandomNumber(
                randomSensor.humidities,
                673
            );
            sensor.humidities = missingData;
        }
    });
};

const generateRandomNumber = (nums, count) => {
    let newNums = [];
    for (let i = 0; i < count; i++) {
        let genNum;
        let randomNum = (Math.random() * 0.2 - 0.1).toFixed(1);
        genNum = +(nums[i] + parseFloat(randomNum)).toFixed(1);
        newNums.push(genNum);
    }
    return newNums;
};

const sortSensorData = () => {
    sensorData.sort((a, b) => {
        // Extract the numeric part of the file names
        const getNumericPart = (fileName) => {
            const match = fileName.match(/\d+/);
            return match ? parseInt(match[0], 10) : 0;
        };

        // Extract the base name without the extension
        const aBaseName = a.fileName.split(".")[0];
        const bBaseName = b.fileName.split(".")[0];

        // Compare based on the numeric part
        const aNumber = getNumericPart(aBaseName);
        const bNumber = getNumericPart(bBaseName);

        // Sort by numeric value, then by the original file name
        if (aNumber !== bNumber) {
            return aNumber - bNumber;
        }
        return a.fileName.localeCompare(b.fileName);
    });
};

const getColumnLetter = (colIndex) => {
    let column = "";
    let dividend = colIndex + 1;
    while (dividend > 0) {
        let modulo = (dividend - 1) % 26;
        column = String.fromCharCode(65 + modulo) + column;
        dividend = Math.floor((dividend - modulo) / 26);
    }
    return column;
};

const copyDataToNewWorkbook = async () => {
    sensorData.forEach((logger, ind) => {
        let columnLetter = getColumnLetter(ind);
        logger.temperatures.forEach((val, i) => {
            if (i === 0) {
                const cell = temperatureSheet.cell(columnLetter + (i + 1));
                cell.value(logger.fileName.split(".")[0]);
                cell.style({
                    fill: "000000", // Black background
                    fontColor: "FFFFFF", // White text
                    horizontalAlignment: "center", // Center horizontally
                    verticalAlignment: "center", // Center vertically
                });
            }
        });
    });
};
