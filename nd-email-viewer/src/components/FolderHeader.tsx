"use client";

import { NDFolder, NDWorkspace, FolderType } from "@/lib/types";
import { Mail, Folder, FolderOpen, RefreshCw } from "lucide-react";

interface FolderHeaderProps {
  folder: NDFolder;
  workspace: NDWorkspace;
  detectedType: FolderType;
  onRefresh: () => void;
}

export function FolderHeader({
  folder,
  workspace,
  detectedType,
  onRefresh,
}: FolderHeaderProps) {
  const typeLabel =
    detectedType === "emails"
      ? "Email Folder"
      : detectedType === "documents"
      ? "Document Folder"
      : "Mixed Folder";

  const TypeIcon =
    detectedType === "emails" ? Mail : detectedType === "documents" ? Folder : FolderOpen;

  return (
    <div className="flex items-center justify-between px-4 py-3 bg-white border-b border-gray-200">
      <div className="flex items-center gap-3">
        <TypeIcon size={20} className="text-blue-600" />
        <div>
          <h1 className="text-base font-semibold text-gray-900">
            {folder.name}
          </h1>
          <div className="flex items-center gap-2 text-xs text-gray-500">
            <span>{workspace.name}</span>
            <span>·</span>
            <span>{workspace.client}</span>
            <span>·</span>
            <span className={`px-1.5 py-0.5 rounded-full font-medium ${
              detectedType === "emails"
                ? "bg-blue-50 text-blue-600"
                : "bg-amber-50 text-amber-600"
            }`}>
              {typeLabel} — layout auto-applied
            </span>
          </div>
        </div>
      </div>
      <button
        onClick={onRefresh}
        className="flex items-center gap-1.5 px-3 py-1.5 text-sm text-gray-600 hover:bg-gray-100 rounded transition-colors"
        title="Refresh folder contents"
      >
        <RefreshCw size={14} />
        Refresh
      </button>
    </div>
  );
}
