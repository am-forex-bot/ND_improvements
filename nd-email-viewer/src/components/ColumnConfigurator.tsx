"use client";

import { ColumnLayout, FolderType } from "@/lib/types";
import { saveDefaultLayout } from "@/lib/column-layouts";
import { Eye, EyeOff, Save, X } from "lucide-react";

interface ColumnConfiguratorProps {
  layout: ColumnLayout;
  folderType: FolderType;
  onLayoutChange: (layout: ColumnLayout) => void;
  onClose: () => void;
}

export function ColumnConfigurator({
  layout,
  folderType,
  onLayoutChange,
  onClose,
}: ColumnConfiguratorProps) {
  const toggleColumn = (field: string) => {
    const updated: ColumnLayout = {
      ...layout,
      columns: layout.columns.map((c) =>
        c.field === field ? { ...c, visible: !c.visible } : c
      ),
    };
    onLayoutChange(updated);
  };

  const handleSaveAsDefault = () => {
    saveDefaultLayout(folderType, layout);
  };

  const folderTypeLabel =
    folderType === "emails"
      ? "Email"
      : folderType === "documents"
      ? "Document"
      : "Mixed";

  return (
    <div className="bg-white border-b border-gray-200 px-4 py-3 shadow-sm">
      <div className="flex items-center justify-between mb-2">
        <div className="flex items-center gap-2">
          <h3 className="text-sm font-semibold text-gray-700">
            Column Configuration
          </h3>
          <span className="text-xs px-2 py-0.5 bg-blue-50 text-blue-600 rounded-full">
            {folderTypeLabel} folder
          </span>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={handleSaveAsDefault}
            className="flex items-center gap-1 px-2 py-1 text-xs bg-blue-600 text-white rounded hover:bg-blue-700 transition-colors"
          >
            <Save size={12} />
            Save as default for {folderTypeLabel} folders
          </button>
          <button
            onClick={onClose}
            className="p-1 hover:bg-gray-100 rounded"
          >
            <X size={16} className="text-gray-500" />
          </button>
        </div>
      </div>

      <div className="flex flex-wrap gap-2">
        {layout.columns.map((col) => (
          <button
            key={col.field}
            onClick={() => toggleColumn(col.field)}
            className={`flex items-center gap-1.5 px-2.5 py-1 rounded text-xs font-medium transition-colors ${
              col.visible
                ? "bg-blue-50 text-blue-700 border border-blue-200"
                : "bg-gray-50 text-gray-400 border border-gray-200"
            }`}
          >
            {col.visible ? <Eye size={12} /> : <EyeOff size={12} />}
            {col.label}
          </button>
        ))}
      </div>

      <p className="text-xs text-gray-400 mt-2">
        Column layout is auto-detected from folder type. Save as default to
        apply this layout to all {folderTypeLabel.toLowerCase()} folders. No more
        switching filters every time you open a folder.
      </p>
    </div>
  );
}
