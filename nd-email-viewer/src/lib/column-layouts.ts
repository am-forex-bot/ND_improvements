import { ColumnLayout, FolderType, UserPreferences } from "./types";

/**
 * Default column layouts per folder type.
 * This is the fix for ND's "every folder uses the same layout" problem.
 * Email folders automatically get email columns. Doc folders get doc columns.
 */

export const EMAIL_COLUMNS: ColumnLayout = {
  name: "Email Layout",
  folderType: "emails",
  columns: [
    { field: "from", label: "From", width: 180, sortable: true, visible: true },
    { field: "to", label: "To", width: 180, sortable: true, visible: true },
    { field: "subject", label: "Subject", width: 300, sortable: true, visible: true },
    { field: "sentDate", label: "Sent", width: 150, sortable: true, visible: true },
    { field: "size", label: "Size", width: 80, sortable: true, visible: true },
    { field: "client", label: "Client", width: 140, sortable: true, visible: false },
    { field: "matter", label: "Matter", width: 140, sortable: true, visible: false },
  ],
};

export const DOCUMENT_COLUMNS: ColumnLayout = {
  name: "Document Layout",
  folderType: "documents",
  columns: [
    { field: "name", label: "Name", width: 300, sortable: true, visible: true },
    { field: "extension", label: "Type", width: 70, sortable: true, visible: true },
    { field: "modifiedDate", label: "Modified", width: 150, sortable: true, visible: true },
    { field: "modifiedBy", label: "Modified By", width: 150, sortable: true, visible: true },
    { field: "createdDate", label: "Created", width: 150, sortable: true, visible: true },
    { field: "createdBy", label: "Created By", width: 150, sortable: true, visible: false },
    { field: "version", label: "Version", width: 70, sortable: true, visible: true },
    { field: "size", label: "Size", width: 80, sortable: true, visible: true },
    { field: "client", label: "Client", width: 140, sortable: true, visible: false },
    { field: "matter", label: "Matter", width: 140, sortable: true, visible: false },
  ],
};

export const MIXED_COLUMNS: ColumnLayout = {
  name: "Mixed Layout",
  folderType: "mixed",
  columns: [
    { field: "name", label: "Name", width: 280, sortable: true, visible: true },
    { field: "extension", label: "Type", width: 70, sortable: true, visible: true },
    { field: "from", label: "From", width: 150, sortable: true, visible: true },
    { field: "modifiedDate", label: "Modified", width: 150, sortable: true, visible: true },
    { field: "modifiedBy", label: "Modified By", width: 150, sortable: true, visible: true },
    { field: "size", label: "Size", width: 80, sortable: true, visible: true },
  ],
};

const PREFS_KEY = "nd-viewer-preferences";

export function getLayoutForFolderType(folderType: FolderType): ColumnLayout {
  const prefs = loadPreferences();
  if (prefs.defaultLayouts[folderType]) {
    return prefs.defaultLayouts[folderType];
  }
  switch (folderType) {
    case "emails":
      return EMAIL_COLUMNS;
    case "documents":
      return DOCUMENT_COLUMNS;
    case "mixed":
      return MIXED_COLUMNS;
  }
}

export function getLayoutForFolder(
  folderId: string,
  folderType: FolderType
): ColumnLayout {
  const prefs = loadPreferences();
  // Check for folder-specific override first
  if (prefs.folderOverrides[folderId]) {
    return prefs.folderOverrides[folderId];
  }
  return getLayoutForFolderType(folderType);
}

export function saveDefaultLayout(
  folderType: FolderType,
  layout: ColumnLayout
): void {
  const prefs = loadPreferences();
  prefs.defaultLayouts[folderType] = layout;
  savePreferences(prefs);
}

export function saveFolderLayout(
  folderId: string,
  layout: ColumnLayout
): void {
  const prefs = loadPreferences();
  prefs.folderOverrides[folderId] = layout;
  savePreferences(prefs);
}

function loadPreferences(): UserPreferences {
  if (typeof window === "undefined")
    return { defaultLayouts: {} as Record<FolderType, ColumnLayout>, folderOverrides: {} };
  try {
    const stored = localStorage.getItem(PREFS_KEY);
    if (stored) return JSON.parse(stored);
  } catch {
    // Ignore
  }
  return {
    defaultLayouts: {} as Record<FolderType, ColumnLayout>,
    folderOverrides: {},
  };
}

function savePreferences(prefs: UserPreferences): void {
  if (typeof window === "undefined") return;
  localStorage.setItem(PREFS_KEY, JSON.stringify(prefs));
}
