// src/entityForms.ts

import axios, { AxiosResponse } from "axios";
import * as ExcelJS from "exceljs";

/**
 * Fetches all pages of data from a given URL, handling paging.
 *
 * @param {string} url - The initial URL to fetch data from.
 * @param {string} accessToken - The OAuth2 access token for authentication.
 * @returns {Promise<any[]>} A promise that resolves to an array of all fetched records.
 * @throws Will throw an error if the request fails.
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
 * Fetches all system forms for a given entity from Microsoft Dynamics 365.
 *
 * @param {string} entityName - The logical name of the entity (e.g., "account", "contact").
 * @param {string} accessToken - The OAuth2 access token for authentication.
 * @param {string} baseUrl - The base URL of the Dynamics 365 instance (e.g., "https://org0b26dba9.api.crm.dynamics.com/api/data/v9.1").
 * @returns {Promise<any[]>} A promise that resolves to an array of form records.
 * @throws Will throw an error if the request fails.
 */
export async function fetchEntityForms(
  entityName: string,
  accessToken: string,
  baseUrl: string
): Promise<any[]> {
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
 * Transforms a raw form object into a simplified record with renamed fields.
 *
 * @param {any} form - The raw form object to transform.
 * @returns {Record<string, any>} A transformed object with key-value pairs for the form.
 */
export function transformForm(form: any): Record<string, any> {
  const formActivationStates: Record<number, string> = {
    0: "Inactive",
    1: "Active",
  };

  const formComponentStates: Record<number, string> = {
    0: "Published",
    1: "Unpublished",
    2: "Deleted",
    3: "Deleted Unpublished",
  };

  const formPresentationStates: Record<number, string> = {
    0: "ClassicForm",
    1: "AirForm",
    2: "ConvertedICForm",
  };

  const formTypes: Record<number, string> = {
    0: "Dashboard",
    1: "AppointmentBook",
    2: "Main",
    3: "MiniCampaignBO",
    4: "Preview",
    5: "Mobile - Express",
    6: "Quick View Form",
    7: "Quick Create",
    8: "Dialog",
    9: "Task Flow Form",
    10: "InteractionCentricDashboard",
    11: "Card",
    12: "Main - Interactive Experience",
    13: "Conceptual Dashboard",
    100: "Other",
    101: "MainBackup",
    102: "AppointmentBookBackup",
    103: "Power BI Dashboard",
  };

  return {
    Name: form.name || "",
    Entity: form.objecttypecode || "",
    "Activation State":
      formActivationStates[form.formactivationstate] || "Unknown",
    "Component State": formComponentStates[form.componentstate] || "Unknown",
    "Presentation State":
      formPresentationStates[form.formpresentationstate] || "Unknown",
    "Is Air Merged": form.isairmerged === 1 ? "Yes" : "No",
    "Form Type": formTypes[form.type] || "Unknown",
    Description: form.description || "",
    "Is Default": form.isdefault ? "Yes" : "No",
    "Is Managed": form.ismanaged ? "Yes" : "No",
    "Is Customizable": form["iscustomizable/Value"] ? "Yes" : "No",
  };
}

/**
 * Adds a worksheet for forms to the given Excel workbook.
 *
 * @param {ExcelJS.Workbook} workbook - The Excel workbook to add the worksheet to.
 * @param {string} entityName - The name of the entity whose forms are being added.
 * @param {string} accessToken - The OAuth2 access token for authentication.
 * @param {string} baseUrl - The base URL of the Dynamics 365 instance.
 * @returns {Promise<void>} A promise that resolves when the worksheet has been added.
 * @throws Will throw an error if fetching forms fails.
 */
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
