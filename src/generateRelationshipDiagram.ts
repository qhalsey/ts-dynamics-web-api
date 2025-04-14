// src/generateRelationshipDiagram.ts

import fs from "fs";
import path from "path";
import { exec } from "child_process";
import { Relationship } from "./types";

/**
 * Options for filtering which relationships to include in the diagram.
 * (Using excludeDefault means "only include relationships if HasChanged != null OR IsCustomRelationship === true")
 */
export interface DiagramFilterOptions {
  /**
   * When set to true, only include relationships that have been changed or are marked as custom.
   */
  excludeDefault?: boolean;
}

/**
 * Determines whether to include a relationship in the diagram.
 *
 * A relationship is included only if:
 * - It has been changed (HasChanged is not null) OR IsCustomRelationship is true.
 * AND
 * - Its "Schema Name", "Entity Ref.", and "Referencing Entity" do NOT start with "msdyn".
 *
 * @param rel A Relationship object.
 * @returns true if the relationship should be included; false otherwise.
 */
function includeRelationship(rel: Relationship): boolean {
  // Exclude if Schema Name starts with "msdyn"
  if (
    rel["Schema Name"] &&
    rel["Schema Name"].toLowerCase().startsWith("msdyn")
  ) {
    return false;
  }
  // Exclude if Entity Ref. starts with "msdyn"
  if (
    rel["Entity Ref."] &&
    rel["Entity Ref."].toLowerCase().startsWith("msdyn")
  ) {
    return false;
  }
  // Exclude if Referencing Entity starts with "msdyn"
  if (
    rel["Referencing Entity"] &&
    rel["Referencing Entity"].toLowerCase().startsWith("msdyn")
  ) {
    return false;
  }
  // Include if the relationship has been changed (HasChanged is not null)
  // OR if it is explicitly marked as a custom relationship.
  return rel.HasChanged !== null || rel.IsCustomRelationship === true;
}

/**
 * Generates Graphviz DOT content based on an array of Relationship objects.
 * Additionally logs the count of relationships before and after filtering.
 *
 * @param relationships The relationships to render.
 * @param filterOptions Optional filter criteria.
 * @returns The DOT language representation of the diagram.
 */
function generateDot(
  relationships: Relationship[],
  filterOptions?: DiagramFilterOptions
): string {
  // Log total relationships before filtering.
  console.log(`Total relationships before filtering: ${relationships.length}`);

  // Apply our filter: if excludeDefault is true, keep only relationships that
  // either have been changed (HasChanged !== null) or are marked as custom,
  // and filter out any whose Schema Name, Entity Ref., or Referencing Entity starts with "msdyn".
  if (filterOptions && filterOptions.excludeDefault) {
    relationships = relationships.filter(includeRelationship);
  }

  // Log total relationships after filtering.
  console.log(`Total relationships after filtering: ${relationships.length}`);

  // Build the DOT graph content.
  let dot = "digraph Relationships {\n";
  dot += "  rankdir=LR;\n";
  dot +=
    '  node [shape=box, style="filled,rounded", fillcolor="#EFEFEF", fontname="Helvetica"];\n\n';

  relationships.forEach((rel) => {
    const source = rel["Referencing Entity"] || "Unknown";
    const target = rel["Entity Ref."] || "Unknown";

    // Build a label that shows the Schema Name and Relationship Type.
    // Append [Custom] if the relationship is marked as custom,
    // and [Changed] if HasChanged is not null.
    const labelParts = [rel["Schema Name"], `(${rel.Type})`];
    if (rel.IsCustomRelationship === true) {
      labelParts.push("[Custom]");
    }
    if (rel.HasChanged !== null) {
      labelParts.push("[Changed]");
    }
    const label = labelParts.join(" ");

    dot += `  "${source}" -> "${target}" [label="${label}"];\n`;
  });

  dot += "}";
  return dot;
}

/**
 * Generates a relationship diagram as a PNG image.
 *
 * @param relationships An array of Relationship objects.
 * @param outputFileName The output file name (e.g., "account.png").
 * @param filterOptions Optional filtering options.
 * @returns A Promise that resolves when the diagram image is successfully saved.
 */
export function generateRelationshipDiagram(
  relationships: Relationship[],
  outputFileName: string,
  filterOptions?: DiagramFilterOptions
): Promise<void> {
  return new Promise((resolve, reject) => {
    try {
      const diagramsDir = path.join(__dirname, "..", "diagrams");
      if (!fs.existsSync(diagramsDir)) {
        fs.mkdirSync(diagramsDir, { recursive: true });
      }

      const dotContent = generateDot(relationships, filterOptions);
      const tempDotPath = path.join(diagramsDir, "temp_relationships.dot");
      fs.writeFileSync(tempDotPath, dotContent, "utf8");

      const outputFilePath = path.join(diagramsDir, outputFileName);
      exec(
        `dot -Tpng -o "${outputFilePath}" "${tempDotPath}"`,
        (error, stdout, stderr) => {
          fs.unlinkSync(tempDotPath);
          if (error) {
            console.error(`Error generating diagram: ${error.message}`);
            return reject(error);
          }
          console.log(`Diagram successfully saved to ${outputFilePath}`);
          resolve();
        }
      );
    } catch (err) {
      reject(err);
    }
  });
}
