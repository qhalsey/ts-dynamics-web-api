// src/entityViews.ts

import axios, { AxiosResponse } from "axios";
import * as ExcelJS from "exceljs";

/**
 * Fetch all paginated data
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
 * Fetch system views (savedqueries) for an entity
 */
export async function fetchEntityViews(
  entityName: string,
  accessToken: string,
  baseUrl: string
): Promise<any[]> {
  const selectedFields = [
    "name",
    "description",
    "returnedtypecode",
    "fetchxml",
    "layoutxml",
    "layoutjson",
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
 * Transform a raw view object into simplified, portable structure
 */
export function transformView(view: any): Record<string, any> {
  return {
    Name: view.name || "",
    Description: view.description || "",
    "Entity Name": view.returnedtypecode || "",
    "Is Default": view.isdefault || false,
    "Is Managed": view.ismanaged || false,
    "Is Customizable": view["iscustomizable/Value"] ?? false,
    "Fetch XML": view.fetchxml || "",
    "Layout XML": view.layoutxml || "",
    "Layout JSON": view.layoutjson || "",
  };
}

/**
 * Add a "Views" worksheet to Excel workbook
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
