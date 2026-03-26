"use client";

import { SortColumn } from "@/lib/types";
import { ArrowUp, ArrowDown } from "lucide-react";

interface SortIndicatorProps {
  field: string;
  sortColumns: SortColumn[];
}

export function SortIndicator({ field, sortColumns }: SortIndicatorProps) {
  const index = sortColumns.findIndex((s) => s.field === field);
  if (index === -1) return null;

  const sort = sortColumns[index];
  return (
    <span className="inline-flex items-center gap-0.5 ml-1">
      {sort.direction === "asc" ? (
        <ArrowUp size={12} className="text-blue-600" />
      ) : (
        <ArrowDown size={12} className="text-blue-600" />
      )}
      {sortColumns.length > 1 && (
        <span className="text-[10px] font-bold text-blue-600 bg-blue-100 rounded-full w-3.5 h-3.5 flex items-center justify-center">
          {index + 1}
        </span>
      )}
    </span>
  );
}
