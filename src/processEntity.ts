// src/processEntity.ts

import * as ExcelJS from "exceljs";
import { fetchEntityAttributes, transformAttribute } from "./entityColumns";
import { addRelationshipsSheet } from "./entityRelationships";
import { addFormsSheet } from "./entityForms";
import { addViewsSheet } from "./entityViews";
import { addBusinessRulesSheet } from "./entityBusinessRules";

export async function processEntityAll(
  accessToken: string,
  entityName: string
) {
  // 1) Create an in-memory workbook
  const workbook = new ExcelJS.Workbook();

  // 2) Add the Columns sheet
  //    a) fetch columns
  const orgUrl = "https://org0b26dba9.crm.dynamics.com/api/data/v9.2";
  const rawAttributes = await fetchEntityAttributes(
    entityName,
    accessToken,
    orgUrl
  );
  console.log(`Fetched ${rawAttributes.length} attributes for ${entityName}.`);

  //    b) transform
  const transformedCols = rawAttributes.map(transformAttribute);

  //    c) add "Columns" worksheet
  const columnsSheet = workbook.addWorksheet("Columns");
  if (transformedCols.length > 0) {
    const columns = Object.keys(transformedCols[0]).map((key) => ({
      header: key,
      key,
    }));
    columnsSheet.columns = columns;
    for (const row of transformedCols) {
      columnsSheet.addRow(row);
    }
  }

  // 3) Add the "Relationships" sheet
  await addRelationshipsSheet(workbook, entityName, accessToken, orgUrl);
  await addFormsSheet(workbook, entityName, accessToken, orgUrl);
  await addViewsSheet(workbook, entityName, accessToken, orgUrl);
  await addBusinessRulesSheet(workbook, entityName, accessToken, orgUrl);
  // 4) Save the final workbook
  await workbook.xlsx.writeFile(`./outputs/${entityName}1.xlsx`);
  console.log(`Wrote file: ./outputs/${entityName}.xlsx`);
}
