// src/excel.ts

import * as ExcelJS from "exceljs";
import * as fs from "fs";
import * as path from "path";

/**
 * Creates a workbook for a single entity, writing out multiple sheets if needed.
 * @param entityName Name of the entity (used for the output file).
 * @param sheetsData An array of sheets to create; each includes a sheetName and array of records.
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
  const outputPath = path.join(outputFolder, `${entityName}1.xlsx`);

  try {
    await workbook.xlsx.writeFile(outputPath);
    console.log(`Workbook for ${entityName} created at ${outputPath}`);
  } catch (err) {
    console.error("Error writing workbook:", err);
  }
}
