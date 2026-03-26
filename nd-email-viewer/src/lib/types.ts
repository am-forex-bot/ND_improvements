// NetDocuments document types
export type DocType = "document" | "email";
export type FolderType = "documents" | "emails" | "mixed";

export interface EmailAttributes {
  from: string;
  to: string[];
  cc?: string[];
  sentDate: string; // ISO date
  subject: string;
}

export interface NDDocument {
  id: string;
  name: string;
  type: DocType;
  extension: string;
  size: number; // bytes
  createdDate: string;
  modifiedDate: string;
  createdBy: string;
  modifiedBy: string;
  version: number;
  // Email-specific (only populated when type === "email")
  emailAttributes?: EmailAttributes;
  // Profile attributes
  client?: string;
  matter?: string;
}

export interface NDFolder {
  id: string;
  name: string;
  parentId: string | null;
  folderType: FolderType; // We detect this from contents
  children?: NDFolder[];
  documentCount: number;
}

export interface NDWorkspace {
  id: string;
  name: string;
  client: string;
  matter: string;
  folders: NDFolder[];
}

export interface NDCabinet {
  id: string;
  name: string;
  workspaces: NDWorkspace[];
}

// Sorting
export type SortDirection = "asc" | "desc";

export interface SortColumn {
  field: string;
  direction: SortDirection;
}

// Column layout configuration
export interface ColumnDef {
  field: string;
  label: string;
  width: number; // px
  sortable: boolean;
  visible: boolean;
}

export interface ColumnLayout {
  name: string;
  folderType: FolderType;
  columns: ColumnDef[];
}

// User preferences
export interface UserPreferences {
  defaultLayouts: Record<FolderType, ColumnLayout>;
  folderOverrides: Record<string, ColumnLayout>; // folderId -> layout
}
