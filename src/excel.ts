// src/excel.ts

import * as ExcelJS from "exceljs";
import * as fs from "fs";
import * as path from "path";

/**
 * Creates an Excel workbook for a single entity, writing out multiple sheets if needed.
 *
 * @param {string} entityName - The name of the entity (used for the output file).
 * @param {{ sheetName: string; records: any[] }[]} sheetsData - An array of sheets to create; each includes a sheet name and an array of records.
 * @returns {Promise<void>} A promise that resolves when the workbook has been created and written to a file.
 * @throws Will throw an error if writing the workbook fails.
 */
export async function createWorkbookForEntity(
  entityName: string,
  sheetsData: { sheetName: string; records: any[] }[]
) {
  // 1) Ensure ./outputs folder exists
  const outputFolder = "./outputs";
  if (!fs.existsSync(outputFolder)) {
    fs.mkdirSync(outputFolder);
  }

  // 2) Create a new Workbook
  const workbook = new ExcelJS.Workbook();

  // 3) For each sheet, add records
  for (const { sheetName, records } of sheetsData) {
    // a) Create a new sheet
    const worksheet = workbook.addWorksheet(sheetName);

    if (records.length > 0) {
      // b) Define columns based on the keys of the first record
      //    Using `Object.keys(records[0])` as column headers
      const columns = Object.keys(records[0]).map((key) => ({
        header: key,
        key: key,
      }));

      worksheet.columns = columns;

      // c) Add each record as a row
      for (const record of records) {
        worksheet.addRow(record);
      }
    }
  }

  // 4) Write the workbook to file
  const outputPath = path.join(outputFolder, `${entityName}.xlsx`);

  try {
    await workbook.xlsx.writeFile(outputPath);
    console.log(`Workbook for ${entityName} created at ${outputPath}`);
  } catch (err) {
    console.error("Error writing workbook:", err);
  }
}
