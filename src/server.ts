import express from "express";
import multer from "multer";
import ExcelJS from "exceljs";
import XLSX from "xlsx";
import path from "path";
import fs from "fs";

const app = express();
const upload = multer({ dest: "uploads/" });

interface TableData {
  [key: string | number]: string | number | null;
}

type Column =
  | string
  | {
      original: string;
      translated: string;
      variations?: string[];
      excludeWhenNull?: boolean;
      isNumber?: boolean;
    }
  | {
      result: string;
      columns: string[];
      variations?: { [columnName: string]: string[] };
      operation: (values: number[][]) => number;
    };

app.use(express.static("public"));

app.post("/upload", upload.single("file"), async (req, res) => {
  const filePath = req.file?.path;
  if (!filePath) {
    return res.status(400).send("No file uploaded.");
  }

  // Read the XLS file with XLSX library
  const workbook = XLSX.readFile(filePath, {
    cellFormula: true,
    cellNF: true,
    cellText: true,
  });

  const sheetNames = workbook.SheetNames.map((sheetName) => sheetName.trim());
  const evaluatedSheets: { [sheetName: string]: string[][] } = {};

  // This will also resolve functions and formulas
  sheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const jsonSheet = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      raw: false,
    }) as string[][];
    evaluatedSheets[sheetName] = jsonSheet;
  });

  const groupBySheet = req?.body?.groupBySheet === "true";

  const jsonOutput = main(evaluatedSheets, groupBySheet);
  const filteredData = filterColumns(jsonOutput, [
    {
      original: "Algemene prijzen",
      variations: [
        "Algemene EPDM prijzen",
        "Materiaal:",
        "Materiaal",
        "Loodgieter en materialen",
        "Vaste prijzen lekdetectie",
        "Ventilatie prijzen vast",
        "Slotenmakenmaker plaats prijzen",
        "Materieel",
        "Onderaanneming",
      ],
      translated: "description",
      excludeWhenNull: true,
    },
    {
      original: "tarief per:",
      translated: "per",
      variations: ["Eenheid"],
    },
    {
      original: "verkoop prijs Excl.",
      translated: "retailPriceEx",
      variations: ["verkoop prijs", "verkoop prijs excl."],
      // excludeWhenNull: true,
      isNumber: true,
    },
    {
      original: "Inkoop materialen",
      translated: "purchasePrice",
      variations: ["Inkoop", "totaal inkoop"],
      //excludeWhenNull: true,
      isNumber: true,
    },
    {
      result: "purchasePriceVat",
      columns: ["Inkoop materialen"],
      variations: {
        "Inkoop materialen": ["Inkoop", "totaal inkoop"],
      },
      operation: (values) => {
        const flattenedValues = values.flat();
        if (flattenedValues.length === 0) return 0;
        return flattenedValues.reduce((a, b) => a + b * 0.21, 0);
      },
    },
    {
      result: "totalPurchasePrice",
      columns: ["Inkoop materialen"],
      variations: {
        "Inkoop materialen": ["Inkoop", "totaal inkoop"],
      },
      operation: (values) => {
        const flattenedValues = values.flat();
        if (flattenedValues.length === 0) return 0;
        const retailPrice = flattenedValues.reduce((a, b) => a + b, 0);
        const retailPriceVat = retailPrice * 0.21;
        return retailPrice + retailPriceVat;
      },
    },
    {
      result: "retailPriceVat",
      columns: ["verkoop prijs Excl."],
      variations: {
        "verkoop prijs Excl.": ["verkoop prijs excl.", "verkoop prijs"],
      },
      operation: (values) => {
        const flattenedValues = values.flat();
        if (flattenedValues.length === 0) return 0;
        return flattenedValues.reduce((a, b) => a + b * 0.21, 0);
      },
    },
    {
      result: "totalPrice",
      columns: ["verkoop prijs Excl."],
      variations: {
        "verkoop prijs Excl.": ["verkoop prijs excl.", "verkoop prijs"],
      },
      operation: (values) => {
        const flattenedValues = values.flat();
        if (flattenedValues.length === 0) return 0;
        const purchasePrice = flattenedValues.reduce((a, b) => a + b, 0);
        const purchasePriceVat = purchasePrice * 0.21;
        return purchasePrice + purchasePriceVat;
      },
    },
  ]);

  res.json(filteredData);

  const jsonFilePath = path.join(__dirname, "output.json");
  fs.writeFile(jsonFilePath, JSON.stringify(filteredData), (err) => {});

  fs.unlinkSync(filePath);
});

function main(
  valuesBySheet: { [sheetName: string]: string[][] },
  groupBySheet: boolean
): { [sheetName: string]: TableData[] } | TableData[] {
  if (groupBySheet) {
    const allData: { [sheetName: string]: TableData[] } = {};

    Object.entries(valuesBySheet).forEach(([sheetName, values]) => {
      const objectArray: TableData[] = [];
      let objectKeys: string[] = [];

      for (let i = 0; i < values.length; i++) {
        if (i === 0) {
          objectKeys = values[i].map((key) => key.trim());
          continue;
        }

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

    Object.entries(valuesBySheet).forEach(([sheetName, values]) => {
      let objectKeys: string[] = [];

      for (let i = 0; i < values.length; i++) {
        if (i === 0) {
          // Assuming the first row contains column headers
          objectKeys = values[i].map((key) => key.trim());
          continue;
        }

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

function filterColumns(
  data: { [sheetName: string]: TableData[] } | TableData[],
  columns: Column[]
): { [sheetName: string]: Partial<TableData>[] } | Partial<TableData>[] {
  // if (Array.isArray(data)) {
  //   console.log("Data is an array, not keyed by sheetName.");
  // } else {
  //   Object.keys(data).forEach((sheetName) => {
  //     data[sheetName].forEach((row, index) => {
  //       if (row.hasOwnProperty("Verkoopprijs Incl.")) {
  //         console.log(
  //           `Row ${index + 1} 'Verkoopprijs Incl.' contents:`,
  //           JSON.stringify(row["Verkoopprijs Incl."], null, 2) // Ensure pretty print is applied
  //         );
  //       }
  //     });
  //   });
  // }

  // Normalize columns to an array of objects
  const normalizedColumns = columns.map((column) => {
    if (typeof column === "string") {
      return { original: column, translated: column };
    } else {
      return column;
    }
  });

  // Function to clean and convert currency strings to numbers
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
        filteredRow[column.result] = result;
      } else {
        const possibleNames = [column.original, ...(column.variations || [])];
        let valueIsValid = false;

        for (const name of possibleNames) {
          if (row.hasOwnProperty(name)) {
            if (
              column.excludeWhenNull &&
              (row[name] === null || row[name] === undefined)
            ) {
              excludeRow = true; // Mark row for exclusion
            }
            if (column.isNumber) {
              const cleanedValue = cleanCurrencyValue(row[name]);
              if (!isNaN(cleanedValue)) {
                filteredRow[column.translated] = cleanedValue;
                valueIsValid = true;
              } else {
                excludeRow = true; // Exclude the row if it has invalid number data
              }
            } else {
              filteredRow[column.translated] = row[name];
              valueIsValid = true;
            }
            if (valueIsValid) break;
          }
        }
      }
    });

    return excludeRow ? null : filteredRow;
  };

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

const port = 3000;
app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
