// src/entityBusinessRules.ts

import axios, { AxiosResponse } from "axios";
import * as ExcelJS from "exceljs";

import { BusinessRule } from "./types/crm"; // Assuming you have a types file for your interfaces

/**
 * Fetch business rules (category=2) for a given entity
 */
export async function fetchEntityBusinessRules(
  entityName: string,
  accessToken: string,
  baseUrl: string
): Promise<BusinessRule[]> {
  const filter = `category eq 2 and primaryentity eq '${entityName}'`;
  const selectFields = [
    "name",
    "description",
    "primaryentity",
    "xaml",
    "clientdata",
    "scope",
    "ismanaged",
    "iscustomizable/Value",
    "statecode",
    "statuscode",
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

  return (response.data.value || []).map((rule: any) => ({
    name: rule.name,
    description: rule.description,
    primaryentity: rule.primaryentity,
    xaml: rule.xaml,
    clientdata: rule.clientdata,
    scope: rule.scope,
    ismanaged: rule.ismanaged,
    iscustomizable: rule["iscustomizable/Value"] ?? false,
    statecode: rule.statecode,
    statuscode: rule.statuscode,
  }));
}

/**
 * Transform rule to friendly Excel format
 */
export function transformBusinessRule(rule: BusinessRule): Record<string, any> {
  return {
    Name: rule.name,
    Description: rule.description ?? "",
    "Entity Name": rule.primaryentity,
    Scope: rule.scope,
    "Is Managed": rule.ismanaged,
    "Is Customizable": rule.iscustomizable,
    "State Code": rule.statecode,
    "Status Code": rule.statuscode,
    "Business Logic (XAML)": rule.xaml ?? "",
    "Client Script": rule.clientdata ?? "",
  };
}

/**
 * Adds a sheet for Business Rules to the workbook
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
