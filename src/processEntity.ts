// src/processEntity.ts

import * as ExcelJS from "exceljs";
import { addColumnsSheet } from "./entityColumns";
import { addRelationshipsSheet } from "./entityRelationships";
import { addFormsSheet } from "./entityForms";
import { addViewsSheet } from "./entityViews";
import { addBusinessRulesSheet } from "./entityBusinessRules";

/**
 * Orchestrates the process of fetching, transforming, and exporting all entity-related data (columns, relationships, forms, views, and business rules) to an Excel workbook.
 *
 * @param {string} accessToken - The OAuth2 access token for authentication.
 * @param {string} entityName - The name of the entity to process.
 * @returns {Promise<void>} A promise that resolves when the process is complete and the workbook is saved.
 * @throws Will throw an error if any step of the process fails.
 */
export async function processEntityAll(
  accessToken: string,
  entityName: string
) {
  //Create an in-memory workbook
  const workbook = new ExcelJS.Workbook();

  //Add the Columns sheet
  //    a) fetch columns
  const orgUrl = "https://org0b26dba9.crm.dynamics.com/api/data/v9.2";

  await addColumnsSheet(workbook, entityName, accessToken, orgUrl);
  await addRelationshipsSheet(workbook, entityName, accessToken, orgUrl);
  await addFormsSheet(workbook, entityName, accessToken, orgUrl);
  await addViewsSheet(workbook, entityName, accessToken, orgUrl);
  await addBusinessRulesSheet(workbook, entityName, accessToken, orgUrl);
  // 4) Save the final workbook
  await workbook.xlsx.writeFile(`./outputs/${entityName}.xlsx`);
  console.log(`Wrote file: ./outputs/${entityName}.xlsx`);
}
