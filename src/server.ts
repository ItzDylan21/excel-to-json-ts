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

type Column =
  | string
  | {
      original?: string;
      translated: string;
      variations?: string[];
      excludeRowWhenNull?: boolean;
      isNumber?: boolean;
      isCurrency?: boolean;
      format?: (value: string[]) => string | null;
      defaultValue?: string;
    }
  | {
      result: string;
      columns: string[];
      variations?: { [columnName: string]: string[] };
      operation: (values: number[][]) => number;
      defaultValue?: number;
    };

app.use(express.static("public"));

app.post("/upload", upload.single("file"), async (req, res) => {
  const filePath = req.file?.path;
  if (!filePath) {
    return res.status(400).send("No file uploaded.");
  }

  const workbook = XLSX.readFile(filePath, {
    cellFormula: true,
    cellNF: true,
    cellText: true,
  });

  const sheetNames = workbook.SheetNames.map((sheetName) => sheetName.trim());
  const evaluatedSheets: { [sheetName: string]: string[][] } = {};

  sheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const jsonSheet = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      raw: false,
    }) as string[][];
    evaluatedSheets[sheetName] = jsonSheet;
  });

  const groupBySheet = req?.body?.groupBySheet === "true";
  const headerRowIndex = parseInt(req?.body?.headerRowIndex) || 0;
  const jsonOutput = main(evaluatedSheets, groupBySheet, headerRowIndex);
  const filteredData = filterColumns(jsonOutput, [
    {
      original: "naam",
      translated: "name",
      excludeRowWhenNull: true,
    },
    {
      original: "telefoon",
      translated: "phonenumber",
    },
    {
      original: "emailadres",
      translated: "email",
      format: (values: string[]): string | null => {
        const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
        if (emailRegex.test(values[0])) {
          return values[0];
        } else {
          return null;
        }
      },
    },
    {
      original: "Btwnr",
      translated: "vatNumber",
    },
    {
      original: "KvKnummer",
      translated: "kvkNumber",
      isNumber: true,
    },
    {
      original: "adresregel1",
      variations: ["adresregel2", "land"],
      translated: "street",
      format: (values: string[]): string => {
        switch (values[2]) {
          case "SG":
          case "US": {
            if (values[1]) {
              const index = values[1].search(/\s[a-zA-Z]/);
              if (index === -1) return values[1];
              return values[1].substring(index + 1);
            } else {
              const index = values[0].search(/\s[a-zA-Z]/);
              if (index === -1) return values[0];
              return values[0].substring(index + 1);
            }
          }
          case "TW": {
            if (values[1]) {
              const index = values[1].search(/,\s/);
              if (index === -1) return values[1];
              return values[1].substring(0, index);
            } else {
              const index = values[0].search(/,\s/);
              if (index === -1) return values[0];
              return values[0].substring(0, index);
            }
          }
          default: {
            // Default for NL, CH, BE, DE, ES
            if (values[1]) {
              const index = values[1].search(/\s\d/);
              if (index === -1) return values[1];
              return values[1].substring(0, index);
            } else {
              const index = values[0].search(/\s\d/);
              if (index === -1) return values[0];
              return values[0].substring(0, index);
            }
          }
        }
      },
    },
    {
      original: "adresregel1",
      variations: ["adresregel2", "land"],
      translated: "housenumber",
      format: (values: string[]): string => {
        switch (values[2]) {
          case "SG":
          case "US": {
            if (values[1]) {
              const index = values[1].search(/\s[a-zA-Z]/);
              if (index === -1) return values[1];
              return values[1].substring(0, index);
            } else {
              const index = values[0].search(/\s[a-zA-Z]/);
              if (index === -1) return values[0];
              return values[0].substring(0, index);
            }
          }
          case "TW": {
            if (values[1]) {
              const index = values[1].search(/,\s/);
              if (index === -1) return values[1];
              return values[1].substring(index + 2);
            } else {
              const index = values[0].search(/,\s/);
              if (index === -1) return values[0];
              return values[0].substring(index + 2);
            }
          }
          default: {
            // Default for NL, CH, BE, DE, ES
            if (values[1]) {
              const index = values[1].search(/\s\d/);
              if (index === -1) return values[1];
              return values[1].substring(index + 1);
            } else {
              const index = values[0].search(/\s\d/);
              if (index === -1) return values[0];
              return values[0].substring(index + 1);
            }
          }
        }
      },
    },
    {
      original: "postcode",
      translated: "zipCode",
    },
    {
      original: "plaats",
      translated: "city",
    },
    {
      original: "land",
      translated: "country",
      format: (values: string[]): string | null => {
        const country = countries.getName(values[0], "nl");
        return country || null;
      },
    },
    {
      original: "emailadres",
      translated: "invoiceEmail",
      format: (values: string[]): string | null => {
        const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
        if (emailRegex.test(values[0])) {
          return values[0];
        } else {
          return null;
        }
      },
    },
    {
      original: "adresregel1",
      variations: ["adresregel2", "land"],
      translated: "invoiceStreet",
      format: (values: string[]): string => {
        switch (values[2]) {
          case "SG":
          case "US": {
            if (values[1]) {
              const index = values[1].search(/\s[a-zA-Z]/);
              if (index === -1) return values[1];
              return values[1].substring(index + 1);
            } else {
              const index = values[0].search(/\s[a-zA-Z]/);
              if (index === -1) return values[0];
              return values[0].substring(index + 1);
            }
          }
          case "TW": {
            if (values[1]) {
              const index = values[1].search(/,\s/);
              if (index === -1) return values[1];
              return values[1].substring(0, index);
            } else {
              const index = values[0].search(/,\s/);
              if (index === -1) return values[0];
              return values[0].substring(0, index);
            }
          }
          default: {
            // Default for NL, CH, BE, DE, ES
            if (values[1]) {
              const index = values[1].search(/\s\d/);
              if (index === -1) return values[1];
              return values[1].substring(0, index);
            } else {
              const index = values[0].search(/\s\d/);
              if (index === -1) return values[0];
              return values[0].substring(0, index);
            }
          }
        }
      },
    },
    {
      original: "adresregel1",
      variations: ["adresregel2", "land"],
      translated: "invoiceHousenumber",
      format: (values: string[]): string => {
        switch (values[2]) {
          case "SG":
          case "US": {
            if (values[1]) {
              const index = values[1].search(/\s[a-zA-Z]/);
              if (index === -1) return values[1];
              return values[1].substring(0, index);
            } else {
              const index = values[0].search(/\s[a-zA-Z]/);
              if (index === -1) return values[0];
              return values[0].substring(0, index);
            }
          }
          case "TW": {
            if (values[1]) {
              const index = values[1].search(/,\s/);
              if (index === -1) return values[1];
              return values[1].substring(index + 2);
            } else {
              const index = values[0].search(/,\s/);
              if (index === -1) return values[0];
              return values[0].substring(index + 2);
            }
          }
          default: {
            // Default for NL, CH, BE, DE, ES
            if (values[1]) {
              const index = values[1].search(/\s\d/);
              if (index === -1) return values[1];
              return values[1].substring(index + 1);
            } else {
              const index = values[0].search(/\s\d/);
              if (index === -1) return values[0];
              return values[0].substring(index + 1);
            }
          }
        }
      },
    },
    {
      original: "postcode",
      translated: "invoiceZipCode",
    },
    {
      original: "plaats",
      translated: "invoiceCity",
    },
    {
      original: "land",
      translated: "invoiceCountry",
      format: (values: string[]): string | null => {
        const country = countries.getName(values[0], "nl");
        return country || null;
      },
    },
    {
      original: "BEHEERDER",
      translated: "managerId",
      isNumber: true,
    },
    {
      original: "CODE",
      translated: "type",
      format: (values: string[]): string | null => {
        switch (values[0]) {
          case "1":
            return "VVE";
          case "2":
            return "COMPANY";
          case "3":
            return "PRIVATE";
          case "4":
            return "MANAGER";
          default:
            return null;
        }
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
  groupBySheet: boolean,
  headerRowIndex: number
): { [sheetName: string]: TableData[] } | TableData[] {
  if (groupBySheet) {
    const allData: { [sheetName: string]: TableData[] } = {};

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
        filteredRow[column.result] = result || column.defaultValue; // Apply defaultValue
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
