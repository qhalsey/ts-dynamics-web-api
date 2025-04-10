// src/entityViews.ts

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
    const response: AxiosResponse<any> = await axios.get(nextLink, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });

    results.push(...(response.data.value || []));
    nextLink = response.data["@odata.nextLink"] || null;
  }

  return results;
}

/**
 * Fetches system views (savedqueries) for a given entity from Microsoft Dynamics 365.
 *
 * @param {string} entityName - The logical name of the entity (e.g., "account", "contact").
 * @param {string} accessToken - The OAuth2 access token for authentication.
 * @param {string} baseUrl - The base URL of the Dynamics 365 instance.
 * @returns {Promise<any[]>} A promise that resolves to an array of view records.
 * @throws Will throw an error if the request fails.
 */
export async function fetchEntityViews(
  entityName: string,
  accessToken: string,
  baseUrl: string
): Promise<any[]> {
  const selectedFields = [
    "name",
    "description",
    "componentstate",
    // "returnedtypecode",
    // "fetchxml",
    // "layoutxml",
    // "layoutjson",
    "isdefault",
    "ismanaged",
    "iscustomizable/Value",
  ].join(",");

  const viewsUrl = `${baseUrl}/savedqueries?$filter=returnedtypecode eq '${entityName}'&$select=${selectedFields}`;

  try {
    const views = await fetchAllPages(viewsUrl, accessToken);
    return views;
  } catch (error) {
    console.error(`Error fetching views for ${entityName}:`, error);
    throw error;
  }
}

/**
 * Transforms a raw view object into a simplified record with renamed fields.
 *
 * @param {any} view - The raw view object to transform.
 * @returns {Record<string, any>} A transformed object with key-value pairs for the view.
 */
export function transformView(view: any): Record<string, any> {
  const formComponentStates: Record<number, string> = {
    0: "Published",
    1: "Unpublished",
    2: "Deleted",
    3: "Deleted Unpublished",
  };

  return {
    Name: view.name || "",
    Description: view.description || "",
    // "Entity Name": view.returnedtypecode || "",
    "Component State": formComponentStates[view["componentstate"]] || "",
    "Is Default": view.isdefault || false,
    "Is Managed": view.ismanaged || false,
    "Is Customizable": view["iscustomizable/Value"] ?? false,
    // "Fetch XML": view.fetchxml || "",
    // "Layout XML": view.layoutxml || "",
    // "Layout JSON": view.layoutjson || "",
  };
}

/**
 * Adds a worksheet for views to the given Excel workbook.
 *
 * @param {ExcelJS.Workbook} workbook - The Excel workbook to add the worksheet to.
 * @param {string} entityName - The name of the entity whose views are being added.
 * @param {string} accessToken - The OAuth2 access token for authentication.
 * @param {string} baseUrl - The base URL of the Dynamics 365 instance.
 * @returns {Promise<void>} A promise that resolves when the worksheet has been added.
 * @throws Will throw an error if fetching views fails.
 */
export async function addViewsSheet(
  workbook: ExcelJS.Workbook,
  entityName: string,
  accessToken: string,
  baseUrl: string
): Promise<void> {
  const rawViews = await fetchEntityViews(entityName, accessToken, baseUrl);
  console.log(`Fetched ${rawViews.length} View records for ${entityName}.`);

  const transformed = rawViews.map(transformView);

  const worksheet = workbook.addWorksheet("Views");

  if (transformed.length > 0) {
    worksheet.columns = Object.keys(transformed[0]).map((key) => ({
      header: key,
      key,
    }));

    for (const row of transformed) {
      worksheet.addRow(row);
    }
  }
}
