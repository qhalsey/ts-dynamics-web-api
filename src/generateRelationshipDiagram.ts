// src/generateRelationshipDiagram.ts

import fs from "fs";
import path from "path";
import { exec } from "child_process";
import { Relationship } from "./types";

/**
 * Options for filtering which relationships to include in the diagram.
 */
export interface DiagramFilterOptions {
  allowedTypes?: string[];
  allowedEntities?: string[];
}

/**
 * Generates Graphviz DOT content based on an array of Relationship objects.
 *
 * @param relationships - The relationships to render.
 * @param filterOptions - Optional filter criteria to limit the diagram output.
 * @returns A string containing the DOT language representation of the diagram.
 */
function generateDot(
  relationships: Relationship[],
  filterOptions?: DiagramFilterOptions
): string {
  if (filterOptions) {
    relationships = relationships.filter((rel) => {
      if (
        filterOptions.allowedTypes &&
        !filterOptions.allowedTypes.includes(rel.Type)
      ) {
        return false;
      }
      if (filterOptions.allowedEntities) {
        const allowed = filterOptions.allowedEntities.map((entity) =>
          entity.toLowerCase()
        );
        if (
          !allowed.includes(rel["Entity Ref."].toLowerCase()) &&
          !allowed.includes(rel["Referencing Entity"].toLowerCase())
        ) {
          return false;
        }
      }
      return true;
    });
  }

  let dot = "digraph Relationships {\n";
  dot += "  rankdir=LR;\n";
  dot +=
    '  node [shape=box, style="filled,rounded", fillcolor="#EFEFEF", fontname="Helvetica"];\n\n';

  relationships.forEach((rel) => {
    const source = rel["Referencing Entity"] || "Unknown";
    const target = rel["Entity Ref."] || "Unknown";
    const label = `${rel["Schema Name"]} (${rel.Type})`;
    dot += `  "${source}" -> "${target}" [label="${label}"];\n`;
  });

  dot += "}";
  return dot;
}

/**
 * Generates a relationship diagram as an image file.
 *
 * @param relationships - An array of Relationship objects.
 * @param outputFileName - The name of the output file (e.g., "account.png").
 * @param filterOptions - Optional filters to limit relationships shown.
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
