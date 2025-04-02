// src/entityRelationships.ts

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
 * Fetch all relationships (OneToMany, ManyToOne, ManyToMany) for a given entity
 */
export async function fetchEntityRelationships(
  entityName: string,
  accessToken: string,
  baseUrl: string
): Promise<any[]> {
  const oneToManyUrl = `${baseUrl}/EntityDefinitions(LogicalName='${entityName}')/OneToManyRelationships`;
  const manyToOneUrl = `${baseUrl}/EntityDefinitions(LogicalName='${entityName}')/ManyToOneRelationships`;
  const manyToManyUrl = `${baseUrl}/EntityDefinitions(LogicalName='${entityName}')/ManyToManyRelationships`;

  // 1) Fetch each relationship type, with paging
  const oneToMany = await fetchAllPages(oneToManyUrl, accessToken);
  const manyToOne = await fetchAllPages(manyToOneUrl, accessToken);
  const manyToMany = await fetchAllPages(manyToManyUrl, accessToken);

  // 2) Combine them all into one array
  //    The "RelationshipType" property often is already set in the JSON
  //    (e.g. "OneToManyRelationship", etc.), but let's rely on it to distinguish them.
  const allRels = [...oneToMany, ...manyToOne, ...manyToMany];

  return allRels;
}

/**
 * Transform a raw relationship object into a simplified
 * record with just the fields we need, renamed appropriately.
 */
export function transformRelationship(rel: any): Record<string, any> {
  return {
    "Schema Name": rel.SchemaName || "",
    "Security Types": rel.SecurityTypes || "",
    Managed: rel.IsManaged ?? "",
    Type: rel.RelationshipType || "",
    "Attribute Ref.": rel.ReferencedAttribute || "",
    "Entity Ref.": rel.ReferencedEntity || "",
    "Referencing Attribute": rel.ReferencingAttribute || "",
    "Referencing Entity": rel.ReferencingEntity || "",
    Hierarchical: rel.IsHierarchical ?? "",
    // RelationshipBehavior -> "Behavior"
    Behavior: rel.RelationshipBehavior ?? "",
    // IsCustomizable.Value -> "Customizable"
    Customizable: rel?.IsCustomizable?.Value ?? "",
    // AssociatedMenuConfiguration.Behavior -> "Menu Behavior"
    "Menu Behavior": rel?.AssociatedMenuConfiguration?.Behavior ?? "",
    // AssociatedMenuConfiguration.IsCustomizable -> "Menu Customization"
    "Menu Customization":
      rel?.AssociatedMenuConfiguration?.IsCustomizable ?? "",
    // CascadeConfiguration.* -> various
    Assign: rel?.CascadeConfiguration?.Assign ?? "",
    Delete: rel?.CascadeConfiguration?.Delete ?? "",
    Archive: rel?.CascadeConfiguration?.Archive ?? "",
    Merge: rel?.CascadeConfiguration?.Merge ?? "",
    Reparent: rel?.CascadeConfiguration?.Reparent ?? "",
    Share: rel?.CascadeConfiguration?.Share ?? "",
    Unshare: rel?.CascadeConfiguration?.Unshare ?? "",
    RollupView: rel?.CascadeConfiguration?.RollupView ?? "",
  };
}

export async function addRelationshipsSheet(
  workbook: ExcelJS.Workbook,
  entityName: string,
  accessToken: string,
  baseUrl: string
) {
  // 1) Fetch raw relationships
  const rawRelationships = await fetchEntityRelationships(
    entityName,
    accessToken,
    baseUrl
  );
  console.log(
    `Fetched ${rawRelationships.length} relationship records for ${entityName}.`
  );

  // 2) Transform them
  const transformed = rawRelationships.map(transformRelationship);

  // 3) Create a new sheet named "Relationships"
  const worksheet = workbook.addWorksheet("Relationships");

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
