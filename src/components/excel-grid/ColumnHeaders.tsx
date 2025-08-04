import { Selection } from "@/types/cellTypes";
import { forwardRef, useState, useRef, useCallback, useEffect } from "react";

interface ColumnHeadersProps {
  selection?: Selection;
  onColumnSelect: (colIndex: number) => void;
  onSelectAll?: () => void;
  columnWidths?: number[];
  onColumnWidthChange?: (columnIndex: number, width: number) => void;
}

const GRID_COLS = 39;
const GRID_ROWS = 100;
const DEFAULT_COLUMN_WIDTH = 80;
const MIN_COLUMN_WIDTH = 20;

const getColumnName = (index: number): string => {
  let result = '';
  while (index >= 0) {
    result = String.fromCharCode(65 + (index % 26)) + result;
    index = Math.floor(index / 26) - 1;
  }
  return result;
};

export const ColumnHeaders = forwardRef<HTMLDivElement, ColumnHeadersProps>(({ 
  selection, 
  onColumnSelect, 
  onSelectAll, 
  columnWidths = Array(GRID_COLS).fill(DEFAULT_COLUMN_WIDTH),
  onColumnWidthChange 
}, ref) => {
  const [isResizing, setIsResizing] = useState(false);
  const [resizingColumn, setResizingColumn] = useState<number | null>(null);
  const [startX, setStartX] = useState(0);
  const [startWidth, setStartWidth] = useState(0);
  const containerRef = useRef<HTMLDivElement>(null);
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

  const handleResizeStart = useCallback((colIndex: number, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    
    setIsResizing(true);
    setResizingColumn(colIndex);
    setStartX(e.clientX);
    setStartWidth(columnWidths[colIndex] || DEFAULT_COLUMN_WIDTH);
    
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
  }, [columnWidths]);

  const handleMouseMove = useCallback((e: MouseEvent) => {
    if (!isResizing || resizingColumn === null || !onColumnWidthChange) return;
    
    const deltaX = e.clientX - startX;
    const newWidth = Math.max(MIN_COLUMN_WIDTH, startWidth + deltaX);
    
    onColumnWidthChange(resizingColumn, newWidth);
  }, [isResizing, resizingColumn, startX, startWidth, onColumnWidthChange]);

  const handleMouseUp = useCallback(() => {
    setIsResizing(false);
    setResizingColumn(null);
    
    document.body.style.cursor = '';
    document.body.style.userSelect = '';
  }, []);

  // Attach global mouse events for resizing
  useEffect(() => {
    if (isResizing) {
      document.addEventListener('mousemove', handleMouseMove);
      document.addEventListener('mouseup', handleMouseUp);
      
      return () => {
        document.removeEventListener('mousemove', handleMouseMove);
        document.removeEventListener('mouseup', handleMouseUp);
      };
    }
  }, [isResizing, handleMouseMove, handleMouseUp]);

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
      <div ref={containerRef} className="flex" style={{ minWidth: 'max-content' }}>
        {Array.from({ length: GRID_COLS }, (_, colIndex) => {
          const columnWidth = columnWidths[colIndex] || DEFAULT_COLUMN_WIDTH;
          
          return (
            <div
              key={colIndex}
              className={`h-6 border-r border-gray-300 flex items-center justify-center text-xs font-medium cursor-pointer select-none flex-shrink-0 relative ${
                isColSelected(colIndex)
                  ? 'text-[#127d42] border-b-2 border-b-[#127d42]'
                  : 'bg-gray-100 hover:bg-gray-200 text-gray-600 border-b border-b-gray-300'
              }`}
              style={{
                width: `${columnWidth}px`,
                minWidth: `${columnWidth}px`,
                backgroundColor: isColSelected(colIndex) ? '#caead8' : undefined
              }}
              onClick={() => onColumnSelect(colIndex)}
            >
              {getColumnName(colIndex)}
              
              {/* Resize handle */}
              <div
                className="absolute right-0 top-0 w-1 h-full cursor-col-resize hover:bg-blue-400 opacity-0 hover:opacity-100 transition-opacity"
                onMouseDown={(e) => handleResizeStart(colIndex, e)}
                title="Resize column"
              />
            </div>
          );
        })}
      </div>
    </div>
  );
});

ColumnHeaders.displayName = 'ColumnHeaders';
