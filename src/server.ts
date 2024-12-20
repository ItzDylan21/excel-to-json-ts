import express from "express";
import multer from "multer";
import XLSX from "xlsx";
import path from "path";
import fs from "fs";

const app = express();
const upload = multer({ dest: "uploads/" });
const countries = require("i18n-iso-countries");

interface TableData {
  [key: string | number]: string | number | null;
}

// Type definition for column configurations
type Column =
  | string
  | {
      // Column name in the original data
      original?: string;
      // Column name in the output data
      translated: string;
      // Optional variations of the column in the Excel file
      variations?: string[];
      // Exclude the row if the column value is null
      excludeRowWhenNull?: boolean;
      // Optional settings for number columns
      isNumber?: boolean;
      // Optional settings for currency columns
      isCurrency?: boolean;
      // Optional function to format the column value
      format?: (value: string[]) => string | null;
      // Optional default value for the column
      defaultValue?: string;
    }
  | {
      // Resulting column name in the output data
      result: string;
      // Column names in the original data
      columns: string[];
      // Optional variations of the column in the Excel file
      variations?: { [columnName: string]: string[] };
      // Operation to perform on the data values
      operation: (values: number[][]) => number;
      // Optional default value for the column
      defaultValue?: number;
    };

app.use(express.static("public"));

// Handle file upload and processing
app.post("/upload", upload.single("file"), async (req, res) => {
  const filePath = req.file?.path;
  if (!filePath) {
    return res.status(400).send("No file uploaded.");
  }

  // Read the uploaded Excel file
  const workbook = XLSX.readFile(filePath, {
    cellFormula: true,
    cellNF: true,
    cellText: true,
  });

  // Get all sheets and trim any whitespace in their names
  const sheetNames = workbook.SheetNames.map((sheetName) => sheetName.trim());
  const evaluatedSheets: { [sheetName: string]: string[][] } = {};

  // Convert each sheet to JSON format
  sheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const jsonSheet = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      raw: false,
    }) as string[][];
    evaluatedSheets[sheetName] = jsonSheet;
  });

  // Extract settings from request body
  const groupBySheet = req?.body?.groupBySheet === "true";
  const headerRowIndex = parseInt(req?.body?.headerRowIndex) || 0;

  const jsonOutput = main(evaluatedSheets, groupBySheet, headerRowIndex);

  // Filter columns based on specified configurations
  // This is different for every use case (hardcoded)
  const filteredData = filterColumns(jsonOutput, [
    {
      original: "Productnaam",
      translated: "description",
      excludeRowWhenNull: true,
    },
    {
      original: "Eenheid",
      translated: "per",
    },
    {
      translated: "supplier",
      defaultValue: "Mastermate",
    },
    {
      result: "purchasePrice",
      columns: ["Nettoprijs", "Eenheid3"],
      operation: (values) => {
        const flattenedValues = values.flat();
        if (flattenedValues.length === 0) return 0;
        return Number(flattenedValues.reduce((a, b) => a / b).toFixed(2));
      },
    },
    {
      result: "retailPriceEx",
      columns: ["Nettoprijs", "Eenheid3"],
      operation: (values) => {
        const flattenedValues = values.flat();
        if (flattenedValues.length === 0) return 0;
        const purchasePrice = Number(
          flattenedValues.reduce((a, b) => a / b).toFixed(2)
        );
        const retailPriceEx = purchasePrice * 1.3;
        return Number(retailPriceEx.toFixed(2));
      },
    },
    {
      result: "purchasePriceVat",
      columns: ["Nettoprijs", "Eenheid3"],
      operation: (values) => {
        const flattenedValues = values.flat();
        if (flattenedValues.length === 0) return 0;
        const purchasePrice = Number(
          flattenedValues.reduce((a, b) => a / b).toFixed(2)
        );
        const purchasePriceVat = purchasePrice * 0.21;
        return purchasePriceVat;
      },
    },
    {
      result: "totalPurchasePrice",
      columns: ["Nettoprijs", "Eenheid3"],
      operation: (values) => {
        const flattenedValues = values.flat();
        if (flattenedValues.length === 0) return 0;
        const purchasePrice = Number(
          flattenedValues.reduce((a, b) => a / b).toFixed(2)
        );
        const purchasePriceVat = purchasePrice * 0.21;
        return purchasePriceVat + purchasePrice;
      },
    },
    {
      result: "retailPriceVat",
      columns: ["Nettoprijs", "Eenheid3"],
      operation: (values) => {
        const flattenedValues = values.flat();
        if (flattenedValues.length === 0) return 0;
        const purchasePrice = Number(
          flattenedValues.reduce((a, b) => a / b).toFixed(2)
        );
        const retailPriceEx = Number((purchasePrice * 1.3).toFixed(2));
        const retailPriceVat = retailPriceEx * 0.21;
        return retailPriceVat;
      },
    },
    {
      result: "totalPrice",
      columns: ["Nettoprijs", "Eenheid3"],
      operation: (values) => {
        const flattenedValues = values.flat();
        if (flattenedValues.length === 0) return 0;
        const purchasePrice = Number(
          flattenedValues.reduce((a, b) => a / b).toFixed(2)
        );
        const retailPriceEx = Number((purchasePrice * 1.3).toFixed(2));
        const retailPriceVat = retailPriceEx * 0.21;
        return retailPriceVat + retailPriceEx;
      },
    },
  ]);

  res.json(filteredData);

  // Save the output to a JSON file
  const jsonFilePath = path.join(__dirname, "output.json");
  fs.writeFile(jsonFilePath, JSON.stringify(filteredData), (err) => {});

  // Delete the uploaded file after processing
  fs.unlinkSync(filePath);
});

function main(
  valuesBySheet: { [sheetName: string]: string[][] },
  groupBySheet: boolean,
  headerRowIndex: number
): { [sheetName: string]: TableData[] } | TableData[] {
  if (groupBySheet) {
    const allData: { [sheetName: string]: TableData[] } = {};

    // Process each sheet into objects based on the header row
    Object.entries(valuesBySheet).forEach(([sheetName, values]) => {
      const objectArray: TableData[] = [];
      let objectKeys: string[] = [];

      for (let i = 0; i < values.length; i++) {
        if (i === headerRowIndex) {
          objectKeys = values[i].map((key) => key.trim());
          continue;
        }
        if (i <= headerRowIndex) continue;

        let object: { [key: string]: string | number | null } = {};
        for (let j = 0; j < objectKeys.length; j++) {
          object[objectKeys[j]] =
            values[i][j] !== undefined ? values[i][j] : null;
        }

        objectArray.push(object as TableData);
      }

      allData[sheetName] = objectArray;
    });

    return allData;
  } else {
    const allObjects: TableData[] = [];

    // Process data without grouping by sheets
    Object.entries(valuesBySheet).forEach(([sheetName, values]) => {
      let objectKeys: string[] = [];

      for (let i = 0; i < values.length; i++) {
        if (i === headerRowIndex) {
          objectKeys = values[i].map((key) => key.trim());
          continue;
        }
        if (i <= headerRowIndex) continue;

        let object: { [key: string]: string | number | null } = {};
        for (let j = 0; j < objectKeys.length; j++) {
          object[objectKeys[j]] =
            values[i][j] !== undefined ? values[i][j] : null;
        }

        allObjects.push(object);
      }
    });

    return allObjects;
  }
}

// Filter and transform columns according to provided configurations
function filterColumns(
  data: { [sheetName: string]: TableData[] } | TableData[],
  columns: Column[]
): { [sheetName: string]: Partial<TableData>[] } | Partial<TableData>[] {
  // Normalize columns to an array of objects
  const normalizedColumns = columns.map((column) => {
    if (typeof column === "string") {
      return { original: column, translated: column };
    } else {
      return column;
    }
  });

  const cleanCurrencyValue = (value: string | number | null): number => {
    if (typeof value === "string") {
      const cleanedValue = value.replace(/[^0-9.,-]/g, "").trim();
      return parseFloat(cleanedValue.replace(",", "."));
    }
    return Number(value);
  };

  // Filter and translate data
  const filterAndTranslateRow = (row: TableData): Partial<TableData> | null => {
    const filteredRow: Partial<TableData> = {};
    let excludeRow = false;

    normalizedColumns.forEach((column) => {
      if ("result" in column && "operation" in column) {
        // Handle calculated columns
        const values: number[][] = column.columns.map((col) => {
          const possibleNames = [col, ...(column.variations?.[col] || [])];
          return possibleNames
            .map((name) => {
              if (row[name] !== undefined && row[name] !== null) {
                return cleanCurrencyValue(row[name]);
              }
              return NaN;
            })
            .filter((value) => !isNaN(value));
        });

        const maxLength = Math.max(...values.map((v) => v.length));
        const groupedValues: number[][] = [];

        for (let i = 0; i < maxLength; i++) {
          const group: number[] = [];
          for (const value of values) {
            if (i < value.length) {
              group.push(value[i]);
            }
          }
          groupedValues.push(group);
        }

        const result = column.operation(groupedValues);
        filteredRow[column.result] = result || column.defaultValue || 0;
      } else {
        const possibleNames = [column.original, ...(column.variations || [])];
        const valuesToFormat: string[] = [];

        let valueIsValid = false;
        for (const name of possibleNames) {
          if (name !== undefined && row.hasOwnProperty(name)) {
            if (
              column.excludeRowWhenNull &&
              (row[name] === null || row[name] === undefined)
            ) {
              excludeRow = true; // Mark row for exclusion
            }
            let value = row[name];
            if (column.isNumber && column.isCurrency) {
              const cleanedValue = cleanCurrencyValue(value);
              if (!isNaN(cleanedValue)) {
                filteredRow[column.translated] = cleanedValue;
                valueIsValid = true;
              } else {
                excludeRow = true; // Exclude the row if it has invalid number data
              }
            } else if (column.isNumber) {
              filteredRow[column.translated] = value ? Number(value) : null;
              valueIsValid = true;
            }
            if (typeof value === "string") {
              valuesToFormat.push(value.trim());
            }
            if (valueIsValid) break;
          }
        }

        if (column.format) {
          const formattedValue = column.format(
            valuesToFormat.length > 0 ? valuesToFormat : [""]
          );
          filteredRow[column.translated] =
            formattedValue || column.defaultValue || null;
        } else if (!column.isNumber && !column.isCurrency) {
          filteredRow[column.translated] =
            valuesToFormat[valuesToFormat.length - 1] ||
            column.defaultValue ||
            null;
        }
      }
    });

    return excludeRow ? null : filteredRow;
  };

  // Process and filter the data either by individual sheet or without sheet grouping
  if (Array.isArray(data)) {
    return data
      .map(filterAndTranslateRow)
      .filter((row): row is Partial<TableData> => row !== null);
  } else {
    const filteredDataBySheet: { [sheetName: string]: Partial<TableData>[] } =
      {};
    Object.entries(data).forEach(([sheetName, sheetData]) => {
      filteredDataBySheet[sheetName] = sheetData
        .map(filterAndTranslateRow)
        .filter((row): row is Partial<TableData> => row !== null);
    });
    return filteredDataBySheet;
  }
}

// Start the Express server
const port = 3000;
app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
