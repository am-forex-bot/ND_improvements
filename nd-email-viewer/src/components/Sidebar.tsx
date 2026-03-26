"use client";

import { NDCabinet, NDFolder, NDWorkspace } from "@/lib/types";
import {
  ChevronDown,
  ChevronRight,
  Folder,
  FolderOpen,
  Mail,
  FileText,
  Building2,
  Briefcase,
} from "lucide-react";
import { useState } from "react";

interface SidebarProps {
  cabinet: NDCabinet;
  selectedFolderId: string | null;
  onFolderSelect: (folder: NDFolder, workspace: NDWorkspace) => void;
  onDragOver: (e: React.DragEvent, folderId: string) => void;
  onDrop: (e: React.DragEvent, folderId: string) => void;
}

export function Sidebar({
  cabinet,
  selectedFolderId,
  onFolderSelect,
  onDragOver,
  onDrop,
}: SidebarProps) {
  const [expandedWorkspaces, setExpandedWorkspaces] = useState<Set<string>>(
    new Set([cabinet.workspaces[0]?.id])
  );
  const [dropTargetId, setDropTargetId] = useState<string | null>(null);

  const toggleWorkspace = (wsId: string) => {
    setExpandedWorkspaces((prev) => {
      const next = new Set(prev);
      if (next.has(wsId)) next.delete(wsId);
      else next.add(wsId);
      return next;
    });
  };

  const FolderIcon = ({ folder, isSelected }: { folder: NDFolder; isSelected: boolean }) => {
    if (folder.folderType === "emails") {
      return <Mail size={16} className={isSelected ? "text-white" : "text-blue-300"} />;
    }
    if (isSelected) {
      return <FolderOpen size={16} className="text-white" />;
    }
    return <Folder size={16} className="text-yellow-400" />;
  };

  return (
    <div className="w-64 min-w-64 h-full bg-[#1c2833] text-gray-200 overflow-y-auto flex flex-col">
      {/* Header */}
      <div className="px-4 py-3 border-b border-gray-700">
        <div className="flex items-center gap-2">
          <Building2 size={18} className="text-blue-400" />
          <span className="font-semibold text-sm">{cabinet.name}</span>
        </div>
      </div>

      {/* Workspace tree */}
      <div className="flex-1 py-2">
        {cabinet.workspaces.map((ws) => (
          <div key={ws.id}>
            {/* Workspace header */}
            <button
              onClick={() => toggleWorkspace(ws.id)}
              className="w-full flex items-center gap-1.5 px-3 py-1.5 hover:bg-[#2c3e50] text-left text-sm"
            >
              {expandedWorkspaces.has(ws.id) ? (
                <ChevronDown size={14} className="text-gray-400 shrink-0" />
              ) : (
                <ChevronRight size={14} className="text-gray-400 shrink-0" />
              )}
              <Briefcase size={15} className="text-blue-400 shrink-0" />
              <span className="truncate font-medium">{ws.name}</span>
            </button>

            {/* Folders */}
            {expandedWorkspaces.has(ws.id) && (
              <div className="ml-4">
                {ws.folders.map((folder) => {
                  const isSelected = selectedFolderId === folder.id;
                  const isDropTarget = dropTargetId === folder.id;

                  return (
                    <button
                      key={folder.id}
                      onClick={() => onFolderSelect(folder, ws)}
                      onDragOver={(e) => {
                        e.preventDefault();
                        setDropTargetId(folder.id);
                        onDragOver(e, folder.id);
                      }}
                      onDragLeave={() => setDropTargetId(null)}
                      onDrop={(e) => {
                        setDropTargetId(null);
                        onDrop(e, folder.id);
                      }}
                      className={`w-full flex items-center gap-1.5 px-3 py-1.5 text-left text-sm rounded-sm mx-1 transition-colors ${
                        isSelected
                          ? "bg-blue-600 text-white"
                          : isDropTarget
                          ? "bg-green-800 text-white outline-dashed outline-1 outline-green-400"
                          : "hover:bg-[#2c3e50]"
                      }`}
                    >
                      <FolderIcon folder={folder} isSelected={isSelected} />
                      <span className="truncate">{folder.name}</span>
                      <span className="ml-auto text-xs opacity-60">
                        {folder.documentCount}
                      </span>
                    </button>
                  );
                })}
              </div>
            )}
          </div>
        ))}
      </div>

      {/* Footer */}
      <div className="px-4 py-2 border-t border-gray-700 text-xs text-gray-500">
        <div className="flex items-center gap-1">
          <FileText size={12} />
          <span>ND Email Viewer — Mock Mode</span>
        </div>
      </div>
    </div>
  );
}
