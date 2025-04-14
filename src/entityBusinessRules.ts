// src/entityBusinessRules.ts

import axios, { AxiosResponse } from "axios";
import { parseStringPromise } from "xml2js";
import * as ExcelJS from "exceljs";

import { BusinessRule } from "./types/crm"; // Assuming you have a types file for your interfaces

/**
 * Fetches business rules (category=2) for a given entity from Microsoft Dynamics 365.
 *
 * @param {string} entityName - The name of the entity to fetch business rules for.
 * @param {string} accessToken - The OAuth2 access token for authentication.
 * @param {string} baseUrl - The base URL of the Dynamics 365 instance.
 * @returns {Promise<BusinessRule[]>} A promise that resolves to an array of business rules.
 * @throws Will throw an error if the request fails.
 */
export async function fetchEntityBusinessRules(
  entityName: string,
  accessToken: string,
  baseUrl: string
): Promise<BusinessRule[]> {
  const filter = `primaryentity eq '${entityName}'`;
  const selectFields = [
    "name",
    "category",
    "type",
    "scope",
    "ismanaged",
    "iscustomizable/Value",
    "statecode",
    "statuscode",
    // "xaml",
  ].join(",");

  const url = `${baseUrl}/workflows?$filter=${encodeURIComponent(
    filter
  )}&$select=${selectFields}`;

  const response: AxiosResponse<any> = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json",
    },
  });

  if (response.status !== 200) {
    throw new Error(`Failed to fetch business rules: ${response.statusText}`);
  }

  if (!response.data || !response.data.value) {
    return [];
  }

  return (response.data.value || []).map((rule: any) => ({
    name: rule.name,
    // primaryentity: rule.primaryentity,
    // clientdata: rule.clientdata,
    scope: rule.scope,
    ismanaged: rule.ismanaged,
    iscustomizable: rule["iscustomizable/Value"] ?? false,
    statecode: rule.statecode,
    statuscode: rule.statuscode,
    type: rule.type,
    category: rule.category,
    // xaml: rule.xaml,
  }));
}

/**
 * Transforms a business rule object into a friendly format suitable for Excel export.
 *
 * @param {BusinessRule} rule - The business rule to transform.
 * @returns {Record<string, any>} A transformed object with key-value pairs for Excel export.
 */
export function transformBusinessRule(rule: BusinessRule): Record<string, any> {
  const businessRuleStatusCode: Record<number, string> = {
    1: "Draft",
    2: "Activated",
    3: "CompanyDLPViolation",
  };

  const businessRuleType: Record<number, string> = {
    0: "Business Flow",
    1: "Task Flow",
  };

  const businessRuleCategory: Record<number, string> = {
    0: "Workflow",
    1: "Dialog",
    2: "Business Rule",
    3: "Action",
    4: "Business Process Flow",
    5: "Modern Flow",
    6: "Desktop Flow",
    7: "AI Flow",
  };

  const businessRuleScope: Record<number, string> = {
    1: "User",
    2: "Business Unit",
    3: "Parent: Child Business Unit",
    4: "Organization",
  };

  const businessRuleComponentState: Record<number, string> = {
    0: "Published",
    1: "Unpublished",
    2: "Deleted",
    3: "Deleted Unpublished",
  };

  return {
    Name: rule.name,
    // "Entity Name": rule.primaryentity,
    Scope: businessRuleScope[rule.scope] || "",
    "Is Managed": rule.ismanaged,
    "Is Customizable": rule.iscustomizable,
    "State Code": businessRuleComponentState[rule.statecode] || "",
    "Status Code": businessRuleStatusCode[rule.statuscode] || "",
    Category: businessRuleCategory[rule.category] || "",
    Type: businessRuleType[rule.type] || "",
    // Logic: parseBusinessRuleXaml(rule.xaml) || "No XAML provided",
  };
}

/**
 * Adds a worksheet for business rules to the given Excel workbook.
 *
 * @param {ExcelJS.Workbook} workbook - The Excel workbook to add the worksheet to.
 * @param {string} entityName - The name of the entity whose business rules are being added.
 * @param {string} accessToken - The OAuth2 access token for authentication.
 * @param {string} baseUrl - The base URL of the Dynamics 365 instance.
 * @returns {Promise<void>} A promise that resolves when the worksheet has been added.
 * @throws Will throw an error if fetching business rules fails.
 */
export async function addBusinessRulesSheet(
  workbook: ExcelJS.Workbook,
  entityName: string,
  accessToken: string,
  baseUrl: string
): Promise<void> {
  const rules = await fetchEntityBusinessRules(
    entityName,
    accessToken,
    baseUrl
  );
  console.log(`Fetched ${rules.length} Business Rules for ${entityName}`);

  const transformed = rules.map(transformBusinessRule);
  const worksheet = workbook.addWorksheet("Business Rules");

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
