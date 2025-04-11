// src/processEntity.ts

import * as ExcelJS from "exceljs";
import { addColumnsSheet } from "./entityColumns";
import {
  addRelationshipsSheet,
  fetchEntityRelationships,
  transformRelationship,
} from "./entityRelationships";
import { addFormsSheet } from "./entityForms";
import { addViewsSheet } from "./entityViews";
import { addBusinessRulesSheet } from "./entityBusinessRules";
import {
  generateRelationshipDiagram,
  DiagramFilterOptions,
} from "./generateRelationshipDiagram";

/**
 * Orchestrates fetching, transforming, and exporting all entity-related data to an Excel workbook
 * and generates a diagram for the entity's relationships.
 *
 * @param accessToken - OAuth2 access token.
 * @param entityName - The logical name of the entity (e.g., "account").
 */
export async function processEntityAll(
  accessToken: string,
  entityName: string
) {
  const workbook = new ExcelJS.Workbook();
  const orgUrl = "https://org0b26dba9.crm.dynamics.com/api/data/v9.2";

  await addColumnsSheet(workbook, entityName, accessToken, orgUrl);
  await addRelationshipsSheet(workbook, entityName, accessToken, orgUrl);
  await addFormsSheet(workbook, entityName, accessToken, orgUrl);
  await addViewsSheet(workbook, entityName, accessToken, orgUrl);
  await addBusinessRulesSheet(workbook, entityName, accessToken, orgUrl);

  // Write Excel output.
  const excelPath = `./outputs/${entityName}.xlsx`;
  await workbook.xlsx.writeFile(excelPath);
  console.log(`Wrote file: ${excelPath}`);

  // Generate the diagram for relationships.
  const rawRelationships = await fetchEntityRelationships(
    entityName,
    accessToken,
    orgUrl
  );
  const transformedRelationships = rawRelationships.map(transformRelationship);
  const filterOptions: DiagramFilterOptions = {
    // Uncomment and edit the filters if needed:
    // allowedTypes: ["OneToManyRelationship", "ManyToOneRelationship"],
    // allowedEntities: [entityName.toLowerCase()],
  };

  const diagramFileName = `${entityName}.png`;
  await generateRelationshipDiagram(
    transformedRelationships,
    diagramFileName,
    filterOptions
  );
  console.log(`Diagram created at ./diagrams/${diagramFileName}`);
}

// Run the process if this file is executed directly.
if (require.main === module) {
  const accessToken = "YOUR_ACCESS_TOKEN"; // Replace with a valid token.
  const entityName = "account"; // Change as needed.
  processEntityAll(accessToken, entityName).catch(console.error);
}
