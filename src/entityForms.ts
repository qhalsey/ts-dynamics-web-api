// src/entityForms.ts

import axios, { AxiosResponse } from "axios";
import * as ExcelJS from "exceljs";

/**
 * Reusable paging function: fetch all pages from a given URL
 */
async function fetchAllPages(url: string, accessToken: string): Promise<any[]> {
  let results: any[] = [];
  let nextLink: string | null = url;

  while (nextLink) {
    const response: AxiosResponse<any> = await axios.get<any>(nextLink, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });
    const data = response.data;
    const values = data.value || [];
    results = results.concat(values);

    if (data["@odata.nextLink"]) {
      nextLink = data["@odata.nextLink"];
    } else {
      nextLink = null;
    }
  }

  return results;
}

/**
 * Fetch all system forms for a given entity
 * @param entityName - The logical name of the entity (e.g., "account", "contact").
 * @param accessToken - The OAuth2 access token for authentication.
 * @param baseUrl - The base URL of the Dynamics 365 instance (e.g., "https://org0b26dba9.api.crm.dynamics.com/api/data/v9.1").
 * @return A promise that resolves to an array of form records.
 * @throws Will throw an error if the request fails.
 */
export async function fetchEntityForms(
  entityName: string,
  accessToken: string,
  baseUrl: string
): Promise<any[]> {
  // Implement this function to get system forms for the entity
  //GET https://your-org.crm.dynamics.com/api/data/v9.2/systemforms?$filter=objecttypecode eq 'account'
  // The specific fields that we want
  const selectedFields = [
    "formjson",
    "formactivationstate",
    "type",
    "description",
    "isdefault",
    "objecttypecode",
    "ismanaged",
    "name",
    "iscustomizable/Value", // Special handling for navigation property
  ].join(",");

  const formsUrl = `${baseUrl}/systemforms?$filter=objecttypecode eq '${entityName}'&$select=${selectedFields}`;

  try {
    const forms = await fetchAllPages(formsUrl, accessToken);
    return forms;
  } catch (error) {
    console.error(`Error fetching forms for ${entityName}:, error`);
    throw error;
  }
}

/**
 * Transform a raw form object into a simplified
 * record with just the fields we need, renamed appropriately.
 */
export function transformForm(form: any): Record<string, any> {
  return {
    JSON: form.formjson || "",
    "Activation State": form.formactivationstate || "",
    Type: form.type ?? "",
    Description: form.description || "",
    "Is Default": form.isdefault || false,
    "Object Type Code": form.objecttypecode || "",
    "Is Managed": form.ismanaged || false,
    Name: form.name || "",
    "Is Customizable": form["iscustomizable/Value"] || false,
  };
}

export async function addFormsSheet(
  workbook: ExcelJS.Workbook,
  entityName: string,
  accessToken: string,
  baseUrl: string
) {
  // 1) Fetch raw relationships
  const rawForms = await fetchEntityForms(entityName, accessToken, baseUrl);
  console.log(`Fetched ${rawForms.length} Form records for ${entityName}.`);

  // 2) Transform them
  const transformed = rawForms.map(transformForm);

  // 3) Create a new sheet named "Relationships"
  const worksheet = workbook.addWorksheet("Forms");

  if (transformed.length > 0) {
    // Set columns based on the keys of the first object
    const columns = Object.keys(transformed[0]).map((key) => ({
      header: key,
      key,
    }));
    worksheet.columns = columns;

    // Add each relationship as a row
    for (const row of transformed) {
      worksheet.addRow(row);
    }
  }
}
