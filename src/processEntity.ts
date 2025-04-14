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
 * Checks if the given relationship should be included based on custom filtering rules:
 * - Only include if either HasChanged is not null OR IsCustomRelationship is true.
 * - Exclude if "mssp" appears in the Schema Name, Entity Ref., or Referencing Entity.
 *
 * @param rel A Relationship object.
 * @returns true if the relationship passes the filter; false otherwise.
 */
function shouldIncludeRelationship(rel: any): boolean {
  const hasCustomOrChanged =
    rel.HasChanged !== null || rel.IsCustomRelationship === true;
  const msspCheck =
    rel["Schema Name"]?.toLowerCase().includes("mssp") ||
    rel["Entity Ref."]?.toLowerCase().includes("mssp") ||
    rel["Referencing Entity"]?.toLowerCase().includes("mssp");
  return hasCustomOrChanged && !msspCheck;
}

/**
 * Orchestrates fetching, transforming, and exporting all entity-related data to an Excel workbook
 * and generates separate diagrams for each relationship type.
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

  // Generate sheets for various metadata.
  await addColumnsSheet(workbook, entityName, accessToken, orgUrl);
  await addRelationshipsSheet(workbook, entityName, accessToken, orgUrl);
  await addFormsSheet(workbook, entityName, accessToken, orgUrl);
  await addViewsSheet(workbook, entityName, accessToken, orgUrl);
  await addBusinessRulesSheet(workbook, entityName, accessToken, orgUrl);

  // Write Excel output.
  const excelPath = `./outputs/${entityName}.xlsx`;
  await workbook.xlsx.writeFile(excelPath);
  console.log(`Wrote file: ${excelPath}`);

  // Generate relationship diagrams.
  const rawRelationships = await fetchEntityRelationships(
    entityName,
    accessToken,
    orgUrl
  );
  console.log(
    `Fetched ${rawRelationships.length} relationship records for ${entityName}.`
  );

  const transformedRelationships = rawRelationships.map(transformRelationship);

  // Define the relationship types to process.
  const relationshipTypes = [
    "OneToManyRelationship",
    "ManyToOneRelationship",
    "ManyToManyRelationship",
  ];

  // For each relationship type, apply our filtering rules and generate a diagram if any relationships remain.
  for (const type of relationshipTypes) {
    // First, filter by relationship type.
    const relationshipsOfType = transformedRelationships.filter(
      (rel) => rel.Type === type
    );

    // Then, apply our custom filtering.
    const filteredRelationships = relationshipsOfType.filter(
      shouldIncludeRelationship
    );

    // Log counts for debugging.
    console.log(`For entity "${entityName}", relationship type "${type}":`);
    console.log(`   Original count: ${relationshipsOfType.length}`);
    console.log(`   After filtering: ${filteredRelationships.length}`);

    if (filteredRelationships.length > 0) {
      const diagramFileName = `${entityName}-${type}.png`;

      // We define filterOptions here if needed by generateRelationshipDiagram.
      // Since we're already filtering manually, you could either pass an empty filter or simply let the function generate based on the passed array.
      const filterOptions: DiagramFilterOptions = {
        excludeDefault: false, // Not used here since we've already filtered manually.
      };

      await generateRelationshipDiagram(
        filteredRelationships,
        diagramFileName,
        filterOptions
      );
      console.log(
        `Diagram for ${type} relationships created at ./diagrams/${diagramFileName}`
      );
    } else {
      console.log(
        `No relationships of type ${type} remain for ${entityName} after filtering (check for mssp, HasChanged, and IsCustomRelationship).`
      );
    }
  }
}

// Run the process if this file is executed directly.
if (require.main === module) {
  const accessToken = "YOUR_ACCESS_TOKEN"; // Replace with a valid token.
  const entityName = "account"; // Change as needed.
  processEntityAll(accessToken, entityName).catch(console.error);
}
