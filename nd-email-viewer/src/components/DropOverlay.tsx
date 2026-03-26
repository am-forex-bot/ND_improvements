"use client";

import { Upload } from "lucide-react";

interface DropOverlayProps {
  visible: boolean;
  folderName: string;
}

export function DropOverlay({ visible, folderName }: DropOverlayProps) {
  if (!visible) return null;

  return (
    <div className="absolute inset-0 bg-green-50/90 border-2 border-dashed border-green-500 rounded-lg z-20 flex items-center justify-center pointer-events-none">
      <div className="text-center">
        <Upload size={40} className="text-green-600 mx-auto mb-2" />
        <p className="text-green-700 font-semibold">
          Drop to file into {folderName}
        </p>
        <p className="text-green-600 text-sm">
          Document will auto-sort into position
        </p>
      </div>
    </div>
  );
}
