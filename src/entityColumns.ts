// src/entityColumns.ts

import axios, { AxiosResponse } from "axios";
import * as ExcelJS from "exceljs";
import * as fs from "fs";
import * as path from "path";

/**
 * 1) Fetch all attributes for a given entity with paging
 */
export async function fetchEntityAttributes(
  entityName: string,
  accessToken: string,
  baseUrl: string
): Promise<any[]> {
  let allAttributes: any[] = [];
  let nextLink:
    | string
    | null = `${baseUrl}/EntityDefinitions(LogicalName='${entityName}')/Attributes`;

  while (nextLink) {
    try {
      const response: AxiosResponse<any> = await axios.get<any>(nextLink, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      });

      const data = response.data;
      // "value" will have the array of attributes
      const attributes = data.value || [];

      allAttributes = allAttributes.concat(attributes);

      // If there's a nextLink, we keep going; otherwise, break out
      if (data["@odata.nextLink"]) {
        nextLink = data["@odata.nextLink"];
      } else {
        nextLink = null;
      }
    } catch (error: any) {
      console.error(
        "Error fetching entity attributes:",
        error?.message || error
      );
      throw error;
    }
  }

  return allAttributes;
}

/**
 * 2) Transform a raw attribute object into a simplified object
 *    containing only the fields we care about (with new property names).
 */
export function transformAttribute(attribute: any): Record<string, any> {
  // Adjust the property order/logic as you wish
  return {
    Name: attribute.SchemaName || "",
    "Data Type": attribute.AttributeType || "",
    Entity: attribute.EntityLogicalName || "",
    Custom: attribute.IsCustomAttribute ?? "",
    "Primary ID": attribute.IsPrimaryId ?? "",
    Managed: attribute.IsManaged ?? "",
    Description: attribute?.Description?.LocalizedLabels?.[0]?.Label ?? "",
    Audited: attribute?.IsAuditEnabled?.Value ?? "",
    Customizable: attribute?.IsCustomizable?.Value ?? "",
    Required: attribute?.RequiredLevel?.Value ?? "",
    "Date Behavior": attribute?.DateTimeBehavior?.Value ?? "",
    Format: attribute?.FormatName?.Value ?? "",
  };
}

/**
 * 3) Write to Excel using exceljs
 */
async function writeAttributesToExcel(
  attributes: Record<string, any>[],
  entityName: string
) {
  // Ensure outputs folder
  const outputFolder = "./outputs";
  if (!fs.existsSync(outputFolder)) {
    fs.mkdirSync(outputFolder);
  }

  // Create workbook
  const workbook = new ExcelJS.Workbook();
  // Dynamically name the worksheet
  const worksheet = workbook.addWorksheet("Columns");

  if (attributes.length > 0) {
    // Set columns based on the keys of the first row
    const columns = Object.keys(attributes[0]).map((key) => ({
      header: key,
      key: key,
    }));
    worksheet.columns = columns;

    // Add rows
    for (const attr of attributes) {
      worksheet.addRow(attr);
    }
  }

  // Write file (again, dynamically name it)
  const outputPath = path.join(outputFolder, `${entityName}.xlsx`);
  await workbook.xlsx.writeFile(outputPath);
  console.log(`File written: ${outputPath}`);
}

/**
 * Orchestrator function to do it all for any entity (fetch + transform + excel).
 */
export async function processEntityColumns(
  accessToken: string,
  entityName: string,
  orgUrl: string
) {
  // 1) Fetch the raw attributes (with paging)
  const rawAttributes = await fetchEntityAttributes(
    entityName,
    accessToken,
    orgUrl
  );
  console.log(`Fetched ${rawAttributes.length} attributes for ${entityName}.`);

  // 2) Transform them
  const transformed = rawAttributes.map(transformAttribute);

  // 3) Write to Excel
  await writeAttributesToExcel(transformed, entityName);
}
