// src/entityColumns.ts

import axios, { AxiosResponse } from "axios";
import * as ExcelJS from "exceljs";

/**
 * Fetches all attributes for a given entity from Microsoft Dynamics 365, handling paging.
 */
export async function fetchEntityAttributes(
  entityName: string,
  accessToken: string,
  baseUrl: string
): Promise<any[]> {
  const attributesUrl = `${baseUrl}/EntityDefinitions(LogicalName='${entityName}')/Attributes`;
  let attributes: any[] = [];
  let nextLink: string | null = attributesUrl;

  while (nextLink) {
    const response: AxiosResponse<any> = await axios.get(nextLink, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });

    const data = response.data;
    attributes = attributes.concat(data.value || []);
    nextLink = data["@odata.nextLink"] || null;
  }

  return attributes;
}

/**
 * Transforms a raw attribute object into a simplified record with renamed fields.
 */
export function transformAttribute(attribute: any): Record<string, any> {
  return {
    "Display Name": attribute.SchemaName || "",
    "Data Type": attribute.AttributeType || "",
    Customizable: attribute.IsCustomizable?.Value || "",
    "Required Level": attribute.RequiredLevel?.Value || "",
  };
}

/**
 * Adds a worksheet for entity attributes (columns) to the given Excel workbook.
 */
export async function addColumnsSheet(
  workbook: ExcelJS.Workbook,
  entityName: string,
  accessToken: string,
  baseUrl: string
) {
  const rawAttributes = await fetchEntityAttributes(
    entityName,
    accessToken,
    baseUrl
  );
  console.log(`Fetched ${rawAttributes.length} attributes for ${entityName}.`);

  const transformedAttributes = rawAttributes.map(transformAttribute);

  const worksheet = workbook.addWorksheet("Columns");

  if (transformedAttributes.length > 0) {
    worksheet.columns = Object.keys(transformedAttributes[0]).map((key) => ({
      header: key,
      key,
    }));

    transformedAttributes.forEach((attr) => worksheet.addRow(attr));
  }
}
