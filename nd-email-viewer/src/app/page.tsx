"use client";

import { useState, useCallback, useRef } from "react";
import { Sidebar } from "@/components/Sidebar";
import { DocumentTable } from "@/components/DocumentTable";
import { FolderHeader } from "@/components/FolderHeader";
import { DropOverlay } from "@/components/DropOverlay";
import {
  mockCabinet,
  getFolderDocuments,
  addDocumentToFolder,
  removeDocumentFromFolder,
} from "@/lib/mock-data";
import { getLayoutForFolder } from "@/lib/column-layouts";
import {
  NDFolder,
  NDWorkspace,
  NDDocument,
  SortColumn,
  ColumnLayout,
} from "@/lib/types";
import { Mail, MousePointerClick } from "lucide-react";

export default function Home() {
  const [selectedFolder, setSelectedFolder] = useState<NDFolder | null>(null);
  const [selectedWorkspace, setSelectedWorkspace] =
    useState<NDWorkspace | null>(null);
  const [documents, setDocuments] = useState<NDDocument[]>([]);
  const [sortColumns, setSortColumns] = useState<SortColumn[]>([]);
  const [currentLayout, setCurrentLayout] = useState<ColumnLayout | null>(null);
  const [isDropTarget, setIsDropTarget] = useState(false);
  const [newDocIds, setNewDocIds] = useState<Set<string>>(new Set());
  const draggedDoc = useRef<NDDocument | null>(null);
  const dragSourceFolder = useRef<string | null>(null);

  const loadFolder = useCallback((folder: NDFolder, workspace: NDWorkspace) => {
    setSelectedFolder(folder);
    setSelectedWorkspace(workspace);
    setNewDocIds(new Set());

    const docs = getFolderDocuments(folder.id);
    setDocuments(docs);

    // Auto-detect folder type and apply correct layout
    const layout = getLayoutForFolder(folder.id, folder.folderType);
    setCurrentLayout(layout);

    // Auto-apply sensible default sort for the folder type
    if (folder.folderType === "emails") {
      // Default: sort by sender asc, then date desc — the thing ND can't do!
      setSortColumns([
        { field: "from", direction: "asc" },
        { field: "sentDate", direction: "desc" },
      ]);
    } else {
      // Documents: sort by modified date desc
      setSortColumns([{ field: "modifiedDate", direction: "desc" }]);
    }
  }, []);

  const handleRefresh = useCallback(() => {
    if (selectedFolder && selectedWorkspace) {
      loadFolder(selectedFolder, selectedWorkspace);
    }
  }, [selectedFolder, selectedWorkspace, loadFolder]);

  // Drag and drop handlers
  const handleDocDragStart = useCallback(
    (e: React.DragEvent, doc: NDDocument) => {
      draggedDoc.current = doc;
      dragSourceFolder.current = selectedFolder?.id || null;
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", doc.id);
    },
    [selectedFolder]
  );

  const handleDocDragEnd = useCallback(() => {
    draggedDoc.current = null;
    dragSourceFolder.current = null;
    setIsDropTarget(false);
  }, []);

  const handleTableDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
    setIsDropTarget(true);
  }, []);

  const handleTableDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setIsDropTarget(false);
      // Table drop = dropping within same folder, no action needed
    },
    []
  );

  const handleSidebarDragOver = useCallback(
    (e: React.DragEvent, _folderId: string) => {
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
    },
    []
  );

  const handleSidebarDrop = useCallback(
    (e: React.DragEvent, targetFolderId: string) => {
      e.preventDefault();
      const doc = draggedDoc.current;
      const sourceFolderId = dragSourceFolder.current;

      if (!doc || !sourceFolderId || sourceFolderId === targetFolderId) return;

      // Remove from source
      removeDocumentFromFolder(sourceFolderId, doc.id);

      // Add to target with new ID
      const newDoc = { ...doc, id: `${targetFolderId}-moved-${Date.now()}` };
      addDocumentToFolder(targetFolderId, newDoc);

      // If we're viewing the source folder, refresh it
      if (selectedFolder?.id === sourceFolderId) {
        setDocuments(getFolderDocuments(sourceFolderId));
      }

      // If we're viewing the target folder, refresh and highlight the new doc
      if (selectedFolder?.id === targetFolderId) {
        const updatedDocs = getFolderDocuments(targetFolderId);
        setDocuments(updatedDocs);
        setNewDocIds(new Set([newDoc.id]));
        // Clear highlight after 3 seconds
        setTimeout(() => setNewDocIds(new Set()), 3000);
      }

      draggedDoc.current = null;
      dragSourceFolder.current = null;
    },
    [selectedFolder]
  );

  const handleLayoutChange = useCallback((layout: ColumnLayout) => {
    setCurrentLayout(layout);
  }, []);

  return (
    <div className="flex h-screen overflow-hidden">
      <Sidebar
        cabinet={mockCabinet}
        selectedFolderId={selectedFolder?.id || null}
        onFolderSelect={loadFolder}
        onDragOver={handleSidebarDragOver}
        onDrop={handleSidebarDrop}
      />

      <div className="flex-1 flex flex-col overflow-hidden relative">
        {selectedFolder && selectedWorkspace && currentLayout ? (
          <>
            <FolderHeader
              folder={selectedFolder}
              workspace={selectedWorkspace}
              detectedType={selectedFolder.folderType}
              onRefresh={handleRefresh}
            />

            <DropOverlay
              visible={isDropTarget}
              folderName={selectedFolder.name}
            />

            <DocumentTable
              documents={documents}
              layout={currentLayout}
              sortColumns={sortColumns}
              folderType={selectedFolder.folderType}
              onSortChange={setSortColumns}
              onLayoutChange={handleLayoutChange}
              onDragStart={handleDocDragStart}
              onDragEnd={handleDocDragEnd}
              isDropTarget={isDropTarget}
              onDragOver={handleTableDragOver}
              onDrop={handleTableDrop}
              newDocIds={newDocIds}
            />
          </>
        ) : (
          <div className="flex-1 flex items-center justify-center bg-gray-50">
            <div className="text-center max-w-md">
              <div className="flex justify-center mb-4">
                <div className="w-16 h-16 rounded-full bg-blue-50 flex items-center justify-center">
                  <Mail size={28} className="text-blue-500" />
                </div>
              </div>
              <h2 className="text-xl font-semibold text-gray-800 mb-2">
                ND Email Viewer
              </h2>
              <p className="text-gray-500 mb-4">
                NetDocuments, but with sorting that actually works.
              </p>
              <div className="text-left bg-white rounded-lg p-4 shadow-sm border border-gray-200 text-sm space-y-2">
                <p className="font-medium text-gray-700 flex items-center gap-2">
                  <MousePointerClick size={16} className="text-blue-500" />
                  What this fixes:
                </p>
                <ul className="space-y-1 text-gray-600 ml-6 list-disc">
                  <li>
                    <strong>Multi-column sort</strong> — click a header, then
                    Shift+click another. Sort by sender THEN date. Finally.
                  </li>
                  <li>
                    <strong>Auto column layouts</strong> — email folders get
                    email columns, doc folders get doc columns. No more
                    switching filter every time.
                  </li>
                  <li>
                    <strong>Instant refresh on filing</strong> — drag a doc to
                    another folder, it sorts into position immediately. No page
                    refresh needed.
                  </li>
                  <li>
                    <strong>Save default layouts</strong> — set your preferred
                    columns once per folder type, they stick.
                  </li>
                </ul>
              </div>
              <p className="text-xs text-gray-400 mt-4">
                Select a folder from the sidebar to get started.
              </p>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
