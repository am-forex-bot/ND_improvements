"use client";

import { NDDocument, SortColumn, ColumnLayout, FolderType } from "@/lib/types";
import { multiColumnSort } from "@/lib/sorting";
import { SortIndicator } from "./SortIndicator";
import { ColumnConfigurator } from "./ColumnConfigurator";
import {
  Mail,
  FileText,
  FileSpreadsheet,
  File,
  GripVertical,
  Settings2,
  ArrowUpDown,
  X,
} from "lucide-react";
import { useState, useCallback, useMemo, useRef } from "react";

interface DocumentTableProps {
  documents: NDDocument[];
  layout: ColumnLayout;
  sortColumns: SortColumn[];
  folderType: FolderType;
  onSortChange: (sortColumns: SortColumn[]) => void;
  onLayoutChange: (layout: ColumnLayout) => void;
  onDragStart: (e: React.DragEvent, doc: NDDocument) => void;
  onDragEnd: () => void;
  isDropTarget: boolean;
  onDragOver: (e: React.DragEvent) => void;
  onDrop: (e: React.DragEvent) => void;
  newDocIds: Set<string>;
}

function formatDate(iso: string): string {
  const d = new Date(iso);
  return d.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  }) + " " + d.toLocaleTimeString("en-GB", {
    hour: "2-digit",
    minute: "2-digit",
  });
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function extractName(emailStr: string): string {
  const match = emailStr.match(/^([^<]+)/);
  return match ? match[1].trim() : emailStr;
}

function getFileIcon(doc: NDDocument) {
  if (doc.type === "email") return <Mail size={14} className="text-blue-500 shrink-0" />;
  switch (doc.extension) {
    case "pdf":
      return <File size={14} className="text-red-500 shrink-0" />;
    case "xlsx":
    case "xls":
      return <FileSpreadsheet size={14} className="text-green-600 shrink-0" />;
    default:
      return <FileText size={14} className="text-blue-700 shrink-0" />;
  }
}

function getCellValue(doc: NDDocument, field: string): string {
  switch (field) {
    case "from":
    case "sender":
      return extractName(doc.emailAttributes?.from || doc.createdBy);
    case "to":
      return doc.emailAttributes?.to?.map(extractName).join(", ") || "";
    case "subject":
      return doc.emailAttributes?.subject || doc.name;
    case "name":
      return doc.name;
    case "sentDate":
      return formatDate(doc.emailAttributes?.sentDate || doc.createdDate);
    case "createdDate":
      return formatDate(doc.createdDate);
    case "modifiedDate":
      return formatDate(doc.modifiedDate);
    case "extension":
      return doc.extension.toUpperCase();
    case "size":
      return formatSize(doc.size);
    case "createdBy":
      return doc.createdBy;
    case "modifiedBy":
      return doc.modifiedBy;
    case "version":
      return `v${doc.version}`;
    case "client":
      return doc.client || "";
    case "matter":
      return doc.matter || "";
    default:
      return "";
  }
}

export function DocumentTable({
  documents,
  layout,
  sortColumns,
  folderType,
  onSortChange,
  onLayoutChange,
  onDragStart,
  onDragEnd,
  isDropTarget,
  onDragOver,
  onDrop,
  newDocIds,
}: DocumentTableProps) {
  const [showColumnConfig, setShowColumnConfig] = useState(false);
  const [draggingDocId, setDraggingDocId] = useState<string | null>(null);
  const tableRef = useRef<HTMLDivElement>(null);

  const visibleColumns = layout.columns.filter((c) => c.visible);

  const sortedDocs = useMemo(
    () => multiColumnSort(documents, sortColumns),
    [documents, sortColumns]
  );

  const handleColumnClick = useCallback(
    (field: string, e: React.MouseEvent) => {
      // Shift+click adds secondary sort; regular click replaces
      if (e.shiftKey) {
        const existing = sortColumns.findIndex((s) => s.field === field);
        if (existing !== -1) {
          // Toggle direction
          const next = [...sortColumns];
          next[existing] = {
            ...next[existing],
            direction: next[existing].direction === "asc" ? "desc" : "asc",
          };
          onSortChange(next);
        } else {
          onSortChange([...sortColumns, { field, direction: "asc" }]);
        }
      } else {
        const existing = sortColumns.find((s) => s.field === field);
        if (existing) {
          onSortChange([
            {
              field,
              direction: existing.direction === "asc" ? "desc" : "asc",
            },
          ]);
        } else {
          onSortChange([{ field, direction: "asc" }]);
        }
      }
    },
    [sortColumns, onSortChange]
  );

  const handleDragStart = (e: React.DragEvent, doc: NDDocument) => {
    setDraggingDocId(doc.id);
    onDragStart(e, doc);
  };

  const handleDragEnd = () => {
    setDraggingDocId(null);
    onDragEnd();
  };

  return (
    <div className="flex flex-col h-full">
      {/* Toolbar */}
      <div className="flex items-center justify-between px-4 py-2 bg-gray-50 border-b border-gray-200">
        <div className="flex items-center gap-3">
          <span className="text-sm text-gray-600">
            {documents.length} item{documents.length !== 1 ? "s" : ""}
          </span>

          {/* Active sorts display */}
          {sortColumns.length > 0 && (
            <div className="flex items-center gap-1.5">
              <ArrowUpDown size={13} className="text-gray-400" />
              {sortColumns.map((sc, i) => (
                <span
                  key={sc.field}
                  className="inline-flex items-center gap-1 px-2 py-0.5 bg-blue-50 text-blue-700 rounded text-xs font-medium"
                >
                  {i + 1}. {visibleColumns.find((c) => c.field === sc.field)?.label || sc.field}{" "}
                  {sc.direction === "asc" ? "↑" : "↓"}
                  <button
                    onClick={() =>
                      onSortChange(sortColumns.filter((_, idx) => idx !== i))
                    }
                    className="hover:text-red-500 ml-0.5"
                  >
                    <X size={10} />
                  </button>
                </span>
              ))}
              {sortColumns.length > 0 && (
                <button
                  onClick={() => onSortChange([])}
                  className="text-xs text-gray-400 hover:text-red-500"
                >
                  Clear
                </button>
              )}
            </div>
          )}
        </div>

        <div className="flex items-center gap-2">
          <span className="text-xs text-gray-400">
            Shift+click header for multi-sort
          </span>
          <button
            onClick={() => setShowColumnConfig(!showColumnConfig)}
            className="flex items-center gap-1 px-2 py-1 text-xs text-gray-600 hover:bg-gray-200 rounded transition-colors"
            title="Configure columns"
          >
            <Settings2 size={14} />
            Columns
          </button>
        </div>
      </div>

      {/* Column configurator */}
      {showColumnConfig && (
        <ColumnConfigurator
          layout={layout}
          folderType={folderType}
          onLayoutChange={onLayoutChange}
          onClose={() => setShowColumnConfig(false)}
        />
      )}

      {/* Table */}
      <div
        ref={tableRef}
        className={`flex-1 overflow-auto ${
          isDropTarget ? "drop-target-active" : ""
        }`}
        onDragOver={onDragOver}
        onDrop={onDrop}
      >
        <table className="w-full border-collapse text-sm">
          <thead className="sticky top-0 z-10">
            <tr className="bg-gray-100 border-b border-gray-300">
              {/* Drag handle column */}
              <th className="w-8 px-1 py-2" />
              {/* Icon column */}
              <th className="w-8 px-1 py-2" />
              {visibleColumns.map((col) => (
                <th
                  key={col.field}
                  style={{ width: col.width, minWidth: col.width }}
                  className={`px-3 py-2 text-left font-semibold text-gray-700 select-none ${
                    col.sortable
                      ? "cursor-pointer hover:bg-gray-200 transition-colors"
                      : ""
                  }`}
                  onClick={
                    col.sortable ? (e) => handleColumnClick(col.field, e) : undefined
                  }
                >
                  <span className="inline-flex items-center">
                    {col.label}
                    <SortIndicator field={col.field} sortColumns={sortColumns} />
                  </span>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {sortedDocs.map((doc, i) => {
              const isNew = newDocIds.has(doc.id);
              const isDragging = draggingDocId === doc.id;

              return (
                <tr
                  key={doc.id}
                  draggable
                  onDragStart={(e) => handleDragStart(e, doc)}
                  onDragEnd={handleDragEnd}
                  className={`border-b border-gray-100 transition-all duration-300 ${
                    isDragging
                      ? "opacity-40"
                      : isNew
                      ? "bg-green-50 animate-pulse"
                      : i % 2 === 0
                      ? "bg-white"
                      : "bg-gray-50/50"
                  } hover:bg-blue-50 cursor-default`}
                >
                  <td className="px-1 py-1.5 cursor-grab active:cursor-grabbing">
                    <GripVertical
                      size={14}
                      className="text-gray-300 hover:text-gray-500"
                    />
                  </td>
                  <td className="px-1 py-1.5">{getFileIcon(doc)}</td>
                  {visibleColumns.map((col) => (
                    <td
                      key={col.field}
                      className="px-3 py-1.5 truncate"
                      style={{ maxWidth: col.width }}
                      title={getCellValue(doc, col.field)}
                    >
                      {getCellValue(doc, col.field)}
                    </td>
                  ))}
                </tr>
              );
            })}
          </tbody>
        </table>

        {documents.length === 0 && (
          <div className="flex items-center justify-center h-40 text-gray-400 text-sm">
            No documents in this folder
          </div>
        )}
      </div>
    </div>
  );
}
