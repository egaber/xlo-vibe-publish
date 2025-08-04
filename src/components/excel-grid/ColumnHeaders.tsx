import { Selection } from "@/types/cellTypes";
import { forwardRef } from "react";

interface ColumnHeadersProps {
  selection?: Selection;
  onColumnSelect: (colIndex: number) => void;
}

const GRID_COLS = 39;
const GRID_ROWS = 100;

const getColumnName = (index: number): string => {
  let result = '';
  while (index >= 0) {
    result = String.fromCharCode(65 + (index % 26)) + result;
    index = Math.floor(index / 26) - 1;
  }
  return result;
};

export const ColumnHeaders = forwardRef<HTMLDivElement, ColumnHeadersProps>(({ selection, onColumnSelect }, ref) => {
  const isColSelected = (colIndex: number): boolean => {
    if (!selection) return false;
    const { start, end } = selection;
    const minCol = Math.min(start.col, end.col);
    const maxCol = Math.max(start.col, end.col);
    
    // Check if this column is within the selected range
    return colIndex >= minCol && colIndex <= maxCol;
  };

  return (
    <div ref={ref} className="flex bg-white border-b border-gray-300 overflow-x-auto scrollbar-hide" style={{ scrollbarWidth: 'none', msOverflowStyle: 'none' }}>
      {/* Top-left corner */}
      <div className="w-12 h-6 bg-gray-100 border-r border-gray-300 flex items-center justify-center text-xs font-medium flex-shrink-0 sticky left-0 z-10"></div>
      
      {/* Column Headers Container */}
      <div className="flex" style={{ minWidth: 'max-content' }}>
        {Array.from({ length: GRID_COLS }, (_, colIndex) => (
          <div
            key={colIndex}
            className={`w-20 h-6 border-r border-gray-300 flex items-center justify-center text-xs font-medium cursor-pointer select-none flex-shrink-0 ${
              isColSelected(colIndex)
                ? 'text-[#127d42] border-b-2 border-b-[#127d42]'
                : 'bg-gray-100 hover:bg-gray-200 text-gray-600 border-b border-b-gray-300'
            }`}
            style={{
              backgroundColor: isColSelected(colIndex) ? '#caead8' : undefined
            }}
            onClick={() => onColumnSelect(colIndex)}
          >
            {getColumnName(colIndex)}
          </div>
        ))}
      </div>
    </div>
  );
});

ColumnHeaders.displayName = 'ColumnHeaders';
