import { Selection } from "@/types/cellTypes";
import { forwardRef } from "react";

interface ColumnHeadersProps {
  selection?: Selection;
  onColumnSelect: (colIndex: number) => void;
  onSelectAll?: () => void;
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

export const ColumnHeaders = forwardRef<HTMLDivElement, ColumnHeadersProps>(({ selection, onColumnSelect, onSelectAll }, ref) => {
  const isColSelected = (colIndex: number): boolean => {
    if (!selection) return false;
    const { start, end } = selection;
    const minCol = Math.min(start.col, end.col);
    const maxCol = Math.max(start.col, end.col);
    
    // Check if this column is within the selected range
    return colIndex >= minCol && colIndex <= maxCol;
  };

  // Check if all cells are selected
  const isAllSelected = (): boolean => {
    if (!selection) return false;
    const { start, end } = selection;
    return start.row === 0 && start.col === 0 && 
           end.row === GRID_ROWS - 1 && end.col === GRID_COLS - 1;
  };

  return (
    <div ref={ref} className="flex bg-white border-b border-gray-300 overflow-x-auto scrollbar-hide" style={{ scrollbarWidth: 'none', msOverflowStyle: 'none' }}>
      {/* Top-left corner - Select All Triangle */}
      <div 
        className={`w-12 h-6 border-r border-gray-300 flex items-center justify-center text-xs font-medium flex-shrink-0 sticky left-0 z-10 cursor-pointer select-none ${
          isAllSelected() 
            ? 'bg-[#caead8] text-[#127d42]' 
            : 'bg-gray-100 hover:bg-gray-200 text-gray-600'
        }`}
        onClick={onSelectAll}
        title="Select All"
      >
        <svg 
          width="8" 
          height="8" 
          viewBox="0 0 8 8" 
          fill="currentColor"
          className="rotate-45"
        >
          <polygon points="0,0 8,0 0,8" />
        </svg>
      </div>
      
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
