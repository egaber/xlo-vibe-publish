import { useState, useRef, useEffect, useCallback } from "react";
import { evaluateFormula, FormulaContext } from "@/utils/formulaEngine";
import { CellData, Selection } from "@/types/cellTypes";
import { getCellStyles, formatCellValue } from "@/utils/cellFormatting";

interface FormulaReference {
  id: string;
  range: string;
  color: string;
}

interface ExcelGridProps {
  onCellSelect: (cellRef: string, value: string) => void;
  cellData: Record<string, CellData>;
  onCellUpdate: (cellRef: string, data: CellData) => void;
  isFormulaBuildingMode?: boolean;
  formulaReferences?: FormulaReference[];
  rangeSelectionStart?: string | null;
  onCellClickInFormulaMode?: (cellRef: string, isRangeSelection?: boolean) => void;
  onSelectionChange?: (selection: Selection) => void;
}

const GRID_ROWS = 100;
const GRID_COLS = 39; // Extended to AM (A=1, B=2, ..., Z=26, AA=27, AB=28, ..., AM=39)

const getColumnName = (index: number): string => {
  let result = '';
  while (index >= 0) {
    result = String.fromCharCode(65 + (index % 26)) + result;
    index = Math.floor(index / 26) - 1;
  }
  return result;
};

const getCellRef = (row: number, col: number): string => {
  return `${getColumnName(col)}${row + 1}`;
};

const parseSelection = (selection: Selection): string => {
  const startRef = getCellRef(selection.start.row, selection.start.col);
  if (selection.start.row === selection.end.row && selection.start.col === selection.end.col) {
    return startRef;
  }
  const endRef = getCellRef(selection.end.row, selection.end.col);
  return `${startRef}:${endRef}`;
};

export const ExcelGrid = ({ 
  onCellSelect, 
  cellData, 
  onCellUpdate, 
  isFormulaBuildingMode = false,
  formulaReferences = [],
  rangeSelectionStart = null,
  onCellClickInFormulaMode,
  onSelectionChange
}: ExcelGridProps) => {
  const [selection, setSelection] = useState<Selection>({
    start: { row: 0, col: 0 },
    end: { row: 0, col: 0 }
  });
  const [isSelecting, setIsSelecting] = useState(false);
  const [editingCell, setEditingCell] = useState<{ row: number; col: number } | null>(null);
  const [editValue, setEditValue] = useState("");
  const gridRef = useRef<HTMLDivElement>(null);
  const editInputRef = useRef<HTMLInputElement>(null);

  const isInSelection = useCallback((row: number, col: number): boolean => {
    const minRow = Math.min(selection.start.row, selection.end.row);
    const maxRow = Math.max(selection.start.row, selection.end.row);
    const minCol = Math.min(selection.start.col, selection.end.col);
    const maxCol = Math.max(selection.start.col, selection.end.col);
    
    return row >= minRow && row <= maxRow && col >= minCol && col <= maxCol;
  }, [selection]);

  const isRowSelected = useCallback((row: number): boolean => {
    const minRow = Math.min(selection.start.row, selection.end.row);
    const maxRow = Math.max(selection.start.row, selection.end.row);
    return row >= minRow && row <= maxRow;
  }, [selection]);

  const isColSelected = useCallback((col: number): boolean => {
    const minCol = Math.min(selection.start.col, selection.end.col);
    const maxCol = Math.max(selection.start.col, selection.end.col);
    return col >= minCol && col <= maxCol;
  }, [selection]);

  // Helper function to check if cell is in formula reference
  const isInFormulaReference = useCallback((cellRef: string): FormulaReference | null => {
    for (const ref of formulaReferences) {
      if (ref.range.includes(':')) {
        // Range reference
        const [start, end] = ref.range.split(':');
        const startCell = parseCellRefString(start);
        const endCell = parseCellRefString(end);
        const currentCell = parseCellRefString(cellRef);
        
        if (startCell && endCell && currentCell) {
          if (currentCell.row >= Math.min(startCell.row, endCell.row) &&
              currentCell.row <= Math.max(startCell.row, endCell.row) &&
              currentCell.col >= Math.min(startCell.col, endCell.col) &&
              currentCell.col <= Math.max(startCell.col, endCell.col)) {
            return ref;
          }
        }
      } else {
        // Single cell reference
        if (ref.range === cellRef) {
          return ref;
        }
      }
    }
    return null;
  }, [formulaReferences]);

  // Helper function to parse cell reference string
  const parseCellRefString = (cellRef: string): { row: number; col: number } | null => {
    const match = cellRef.match(/^([A-Z]+)(\d+)$/);
    if (!match) return null;
    
    const colStr = match[1];
    const rowStr = match[2];
    
    let col = 0;
    for (let i = 0; i < colStr.length; i++) {
      col = col * 26 + (colStr.charCodeAt(i) - 65 + 1);
    }
    col -= 1; // Convert to 0-based
    
    const row = parseInt(rowStr) - 1; // Convert to 0-based
    
    return { col, row };
  };

  const getRangeBorders = useCallback((row: number, col: number) => {
    if (!isInSelection(row, col)) return '';
    
    const minRow = Math.min(selection.start.row, selection.end.row);
    const maxRow = Math.max(selection.start.row, selection.end.row);
    const minCol = Math.min(selection.start.col, selection.end.col);
    const maxCol = Math.max(selection.start.col, selection.end.col);
    
    const isTopEdge = row === minRow;
    const isBottomEdge = row === maxRow;
    const isLeftEdge = col === minCol;
    const isRightEdge = col === maxCol;
    
    let borderClasses = '';
    
    if (isTopEdge) borderClasses += ' border-t-2 border-t-[#127d42]';
    if (isBottomEdge) borderClasses += ' border-b-2 border-b-[#127d42]';
    if (isLeftEdge) borderClasses += ' border-l-2 border-l-[#127d42]';
    if (isRightEdge) borderClasses += ' border-r-2 border-r-[#127d42]';
    
    return borderClasses;
  }, [selection, isInSelection]);

  const handleCellClick = (row: number, col: number, isShiftClick: boolean = false) => {
    if (editingCell) {
      finishEditing();
    }

    const cellRef = getCellRef(row, col);

    // Handle formula building mode
    if (isFormulaBuildingMode && onCellClickInFormulaMode) {
      onCellClickInFormulaMode(cellRef, isShiftClick);
      return;
    }

    if (isShiftClick && selection.start) {
      setSelection({
        start: selection.start,
        end: { row, col }
      });
    } else {
      setSelection({
        start: { row, col },
        end: { row, col }
      });
    }

    const cellInfo = cellData[cellRef] || { value: "" };
    onCellSelect(cellRef, cellInfo.value);
  };

  const handleCellDoubleClick = (row: number, col: number) => {
    const cellRef = getCellRef(row, col);
    const cellInfo = cellData[cellRef] || { value: "" };
    setEditingCell({ row, col });
    setEditValue(cellInfo.formula || cellInfo.value);
    setTimeout(() => editInputRef.current?.focus(), 0);
  };

  const handleMouseDown = (row: number, col: number) => {
    setIsSelecting(true);
    handleCellClick(row, col);
  };

  const handleMouseOver = (row: number, col: number) => {
    if (isSelecting) {
      setSelection(prev => ({
        start: prev.start,
        end: { row, col }
      }));
    }
  };

  const handleMouseUp = () => {
    setIsSelecting(false);
  };

  const finishEditing = () => {
    if (editingCell) {
      const cellRef = getCellRef(editingCell.row, editingCell.col);
      const isFormula = editValue.startsWith('=');
      
      let displayValue = editValue;
      
      if (isFormula) {
        // Create formula context for evaluation
        const formulaContext: FormulaContext = {
          getCellValue: (ref: string) => {
            const cellInfo = cellData[ref];
            return cellInfo ? cellInfo.value : '';
          },
          setCellValue: (ref: string, value: string) => {
            // This would be used for more complex formulas that modify other cells
            // For now, we'll just implement the getter
          }
        };
        
        // Evaluate the formula
        displayValue = evaluateFormula(editValue, formulaContext);
      }
      
      onCellUpdate(cellRef, {
        value: displayValue,
        formula: isFormula ? editValue : undefined
      });
      
      setEditingCell(null);
      setEditValue("");
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (editingCell) {
      if (e.key === 'Enter') {
        finishEditing();
      } else if (e.key === 'Escape') {
        setEditingCell(null);
        setEditValue("");
      }
      return;
    }

    // Navigation keys
    const { row, col } = selection.start;
    let newRow = row;
    let newCol = col;

    switch (e.key) {
      case 'ArrowUp':
        newRow = Math.max(0, row - 1);
        break;
      case 'ArrowDown':
        newRow = Math.min(GRID_ROWS - 1, row + 1);
        break;
      case 'ArrowLeft':
        newCol = Math.max(0, col - 1);
        break;
      case 'ArrowRight':
        newCol = Math.min(GRID_COLS - 1, col + 1);
        break;
      case 'Enter':
        handleCellDoubleClick(row, col);
        return;
      case 'F2':
        handleCellDoubleClick(row, col);
        return;
      default:
        // Start editing if a printable character is pressed
        if (e.key.length === 1 && !e.ctrlKey && !e.altKey) {
          setEditingCell({ row, col });
          setEditValue(e.key);
          setTimeout(() => editInputRef.current?.focus(), 0);
        }
        return;
    }

    if (newRow !== row || newCol !== col) {
      e.preventDefault();
      handleCellClick(newRow, newCol, e.shiftKey);
    }
  };

  useEffect(() => {
    const currentSelection = parseSelection(selection);
    const cellRef = getCellRef(selection.start.row, selection.start.col);
    const cellInfo = cellData[cellRef] || { value: "" };
    onCellSelect(currentSelection, cellInfo.value);
    onSelectionChange?.(selection);
  }, [selection, cellData, onCellSelect, onSelectionChange]);

  useEffect(() => {
    document.addEventListener('mouseup', handleMouseUp);
    return () => document.removeEventListener('mouseup', handleMouseUp);
  }, []);

  return (
    <div 
      className="flex-1 overflow-auto bg-white focus:outline-none" 
      tabIndex={0}
      onKeyDown={handleKeyDown}
      ref={gridRef}
    >
      <div className="inline-block min-w-full">
        {/* Header Row */}
        <div className="flex sticky top-0 z-10">
          {/* Top-left corner */}
          <div className="w-12 h-6 bg-gray-100 border-r border-b border-gray-300 flex items-center justify-center text-xs font-medium sticky left-0 z-20"></div>
          
          {/* Column Headers */}
          {Array.from({ length: GRID_COLS }, (_, colIndex) => (
            <div
              key={colIndex}
              className={`w-20 h-6 border-r border-b border-gray-300 flex items-center justify-center text-xs font-medium cursor-pointer select-none ${
                isColSelected(colIndex)
                  ? 'bg-[#c9e9d7] text-[#127d42] border-b-[#127d42]'
                  : 'bg-gray-100 hover:bg-gray-200 text-gray-600'
              }`}
              onClick={() => {
                // Select entire column
                setSelection({
                  start: { row: 0, col: colIndex },
                  end: { row: GRID_ROWS - 1, col: colIndex }
                });
              }}
            >
              {getColumnName(colIndex)}
            </div>
          ))}
        </div>

        {/* Data Rows */}
        {Array.from({ length: GRID_ROWS }, (_, rowIndex) => (
          <div key={rowIndex} className="flex">
            {/* Row Header */}
            <div
              className={`w-12 h-6 border-r border-b border-gray-300 flex items-center justify-end pr-2 text-sm font-medium cursor-pointer select-none sticky left-0 z-10 ${
                isRowSelected(rowIndex)
                  ? 'bg-[#c9e9d7] text-[#127d42] border-r-[#127d42]'
                  : 'bg-gray-100 hover:bg-gray-200 text-gray-600'
              }`}
              onClick={() => {
                // Select entire row
                setSelection({
                  start: { row: rowIndex, col: 0 },
                  end: { row: rowIndex, col: GRID_COLS - 1 }
                });
              }}
            >
              {rowIndex + 1}
            </div>

            {/* Data Cells */}
            {Array.from({ length: GRID_COLS }, (_, colIndex) => {
              const cellRef = getCellRef(rowIndex, colIndex);
              const cellInfo = cellData[cellRef] || { value: "" };
              const isSelected = isInSelection(rowIndex, colIndex);
              const isActiveCell = selection.start.row === rowIndex && selection.start.col === colIndex;
              const isEditing = editingCell?.row === rowIndex && editingCell?.col === colIndex;
              const rangeBorders = getRangeBorders(rowIndex, colIndex);
              
              // Check if cell is in formula reference
              const formulaRef = isInFormulaReference(cellRef);
              const isRangeStart = rangeSelectionStart === cellRef;
              
              // Get cell formatting styles
              const cellStyles = getCellStyles(cellInfo.format);
              const formattedValue = formatCellValue(cellInfo.value, cellInfo.format);
              
              // Determine background color and styles
              let bgColor = 'bg-white hover:bg-gray-50';
              let cellStyle: React.CSSProperties = { ...cellStyles };
              
              if (isSelected) {
                bgColor = 'bg-[#e8f2ec]';
              } else if (formulaRef) {
                // Override with formula reference styling
                cellStyle = {
                  ...cellStyle,
                  backgroundColor: `rgba(${parseInt(formulaRef.color.slice(1, 3), 16)}, ${parseInt(formulaRef.color.slice(3, 5), 16)}, ${parseInt(formulaRef.color.slice(5, 7), 16)}, 0.3)`,
                  borderColor: formulaRef.color,
                  borderWidth: '2px',
                  borderStyle: 'solid'
                };
              } else if (isRangeStart) {
                bgColor = 'bg-blue-100';
              }

              return (
                <div
                  key={colIndex}
                  className={`w-20 h-6 border-r border-b border-gray-300 relative cursor-cell ${bgColor} ${rangeBorders}`}
                  style={cellStyle}
                  onMouseDown={() => handleMouseDown(rowIndex, colIndex)}
                  onMouseOver={() => handleMouseOver(rowIndex, colIndex)}
                  onDoubleClick={() => handleCellDoubleClick(rowIndex, colIndex)}
                >
                  {isEditing ? (
                    <input
                      ref={editInputRef}
                      value={editValue}
                      onChange={(e) => setEditValue(e.target.value)}
                      onBlur={finishEditing}
                      onKeyDown={(e) => {
                        e.stopPropagation();
                        if (e.key === 'Enter') {
                          finishEditing();
                        } else if (e.key === 'Escape') {
                          setEditingCell(null);
                          setEditValue("");
                        }
                      }}
                      className="w-full h-full px-1 text-xs border-0 outline-none bg-transparent"
                      style={cellStyles}
                      autoFocus
                    />
                  ) : (
                    <div 
                      className="w-full h-full px-1 text-xs flex items-center overflow-hidden"
                      style={cellStyles}
                    >
                      {formattedValue}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        ))}
      </div>
    </div>
  );
};
