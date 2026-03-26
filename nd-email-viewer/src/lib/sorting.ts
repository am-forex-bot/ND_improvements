import { NDDocument, SortColumn } from "./types";

/**
 * Multi-column sort — the thing NetDocuments can't do.
 * Sorts by primary column first, then secondary, tertiary, etc.
 */
export function multiColumnSort(
  documents: NDDocument[],
  sortColumns: SortColumn[]
): NDDocument[] {
  if (sortColumns.length === 0) return documents;

  return [...documents].sort((a, b) => {
    for (const { field, direction } of sortColumns) {
      const comparison = compareField(a, b, field);
      if (comparison !== 0) {
        return direction === "asc" ? comparison : -comparison;
      }
    }
    return 0;
  });
}

function getFieldValue(doc: NDDocument, field: string): string | number {
  switch (field) {
    case "name":
    case "subject":
      return doc.emailAttributes?.subject || doc.name;
    case "from":
    case "sender":
      return extractName(doc.emailAttributes?.from || doc.createdBy);
    case "to":
      return doc.emailAttributes?.to?.map(extractName).join(", ") || "";
    case "sentDate":
      return doc.emailAttributes?.sentDate || doc.createdDate;
    case "createdDate":
      return doc.createdDate;
    case "modifiedDate":
      return doc.modifiedDate;
    case "extension":
      return doc.extension;
    case "size":
      return doc.size;
    case "createdBy":
      return doc.createdBy;
    case "modifiedBy":
      return doc.modifiedBy;
    case "version":
      return doc.version;
    case "client":
      return doc.client || "";
    case "matter":
      return doc.matter || "";
    default:
      return "";
  }
}

function compareField(a: NDDocument, b: NDDocument, field: string): number {
  const valA = getFieldValue(a, field);
  const valB = getFieldValue(b, field);

  if (typeof valA === "number" && typeof valB === "number") {
    return valA - valB;
  }

  return String(valA).localeCompare(String(valB), undefined, {
    sensitivity: "base",
  });
}

function extractName(emailStr: string): string {
  // "Andrew Meston <andrew@firm.com>" -> "Andrew Meston"
  const match = emailStr.match(/^([^<]+)/);
  return match ? match[1].trim() : emailStr;
}
