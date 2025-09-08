import { useState, useRef, useEffect, useCallback } from "react";
import { evaluateFormula, FormulaContext } from "@/utils/formulaEngine";
import { CellData, Selection, MultiSelection } from "@/types/cellTypes";
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
  onMultiSelectionChange?: (multiSelection: MultiSelection) => void;
  onColumnSelect?: (colIndex: number) => void;
  onSelectAll?: () => void;
  externalSelection?: Selection;
  onScroll?: (scrollLeft: number) => void;
  columnWidths?: number[];
  onColumnWidthChange?: (columnIndex: number, width: number) => void;
}

const GRID_ROWS = 100;
const GRID_COLS = 39; // Extended to AM (A=1, B=2, ..., Z=26, AA=27, AB=28, ..., AM=39)
const DEFAULT_COLUMN_WIDTH = 80; // Default column width in pixels

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

// Helper function to calculate text width
const getTextWidth = (text: string, font: string): number => {
  const canvas = document.createElement('canvas');
  const context = canvas.getContext('2d');
  if (context) {
    context.font = font;
    return context.measureText(text).width;
  }
  return 0;
};

// Helper function to check if a cell should spill over
const shouldSpillOver = (
  text: string, 
  cellWidth: number, 
  fontSize: number, 
  fontFamily: string,
  rowIndex: number,
  colIndex: number,
  cellData: Record<string, CellData>,
  columnWidths: number[]
): { shouldSpill: boolean; spillCells: number } => {
  if (!text || text.trim() === '') return { shouldSpill: false, spillCells: 0 };
  
  const font = `${fontSize}px ${fontFamily}`;
  const textWidth = getTextWidth(text, font);
  const availableWidth = cellWidth - 8; // Account for padding
  
  if (textWidth <= availableWidth) {
    return { shouldSpill: false, spillCells: 0 };
  }
  
  // Calculate how many cells we need to spill into
  let spillCells = 0;
  let totalWidth = availableWidth;
  
  // Check adjacent cells to the right
  for (let i = colIndex + 1; i < GRID_COLS && totalWidth < textWidth; i++) {
    const adjacentCellRef = getCellRef(rowIndex, i);
    const adjacentCellInfo = cellData[adjacentCellRef];
    
    // Stop if adjacent cell has content
    if (adjacentCellInfo && adjacentCellInfo.value && adjacentCellInfo.value.trim() !== '') {
      break;
    }
    
    spillCells++;
    totalWidth += (columnWidths[i] || DEFAULT_COLUMN_WIDTH);
  }
  
  return { shouldSpill: spillCells > 0, spillCells };
};

export const ExcelGrid = ({ 
  onCellSelect, 
  cellData, 
  onCellUpdate, 
  isFormulaBuildingMode = false,
  formulaReferences = [],
  rangeSelectionStart = null,
  onCellClickInFormulaMode,
  onSelectionChange,
  onMultiSelectionChange,
  onColumnSelect,
  onSelectAll,
  externalSelection,
  onScroll,
  columnWidths = Array(GRID_COLS).fill(DEFAULT_COLUMN_WIDTH),
  onColumnWidthChange
}: ExcelGridProps) => {
  const [selection, setSelection] = useState<Selection>({
    start: { row: 0, col: 0 },
    end: { row: 0, col: 0 }
  });
  const [multiSelection, setMultiSelection] = useState<MultiSelection>({
    primary: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
    additional: []
  });
  const [isSelecting, setIsSelecting] = useState(false);
  const [editingCell, setEditingCell] = useState<{ row: number; col: number } | null>(null);
  const [editValue, setEditValue] = useState("");
  const gridRef = useRef<HTMLDivElement>(null);
  const editInputRef = useRef<HTMLInputElement>(null);

  const isInSelection = useCallback((row: number, col: number): boolean => {
    // Check primary selection
    const minRow = Math.min(multiSelection.primary.start.row, multiSelection.primary.end.row);
    const maxRow = Math.max(multiSelection.primary.start.row, multiSelection.primary.end.row);
    const minCol = Math.min(multiSelection.primary.start.col, multiSelection.primary.end.col);
    const maxCol = Math.max(multiSelection.primary.start.col, multiSelection.primary.end.col);
    
    if (row >= minRow && row <= maxRow && col >= minCol && col <= maxCol) {
      return true;
    }
    
    // Check additional selections
    for (const sel of multiSelection.additional) {
      const selMinRow = Math.min(sel.start.row, sel.end.row);
      const selMaxRow = Math.max(sel.start.row, sel.end.row);
      const selMinCol = Math.min(sel.start.col, sel.end.col);
      const selMaxCol = Math.max(sel.start.col, sel.end.col);
      
      if (row >= selMinRow && row <= selMaxRow && col >= selMinCol && col <= selMaxCol) {
        return true;
      }
    }
    
    return false;
  }, [multiSelection]);

  const isRowSelected = useCallback((row: number): boolean => {
    // Check primary selection
    const primaryMinRow = Math.min(multiSelection.primary.start.row, multiSelection.primary.end.row);
    const primaryMaxRow = Math.max(multiSelection.primary.start.row, multiSelection.primary.end.row);
    
    if (row >= primaryMinRow && row <= primaryMaxRow) {
      return true;
    }
    
    // Check additional selections
    for (const sel of multiSelection.additional) {
      const selMinRow = Math.min(sel.start.row, sel.end.row);
      const selMaxRow = Math.max(sel.start.row, sel.end.row);
      
      if (row >= selMinRow && row <= selMaxRow) {
        return true;
      }
    }
    
    return false;
  }, [multiSelection]);

  const isColSelected = useCallback((col: number): boolean => {
    // Check primary selection
    const primaryMinCol = Math.min(multiSelection.primary.start.col, multiSelection.primary.end.col);
    const primaryMaxCol = Math.max(multiSelection.primary.start.col, multiSelection.primary.end.col);
    
    if (col >= primaryMinCol && col <= primaryMaxCol) {
      return true;
    }
    
    // Check additional selections
    for (const sel of multiSelection.additional) {
      const selMinCol = Math.min(sel.start.col, sel.end.col);
      const selMaxCol = Math.max(sel.start.col, sel.end.col);
      
      if (col >= selMinCol && col <= selMaxCol) {
        return true;
      }
    }
    
    return false;
  }, [multiSelection]);

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
    
    // Check if cell is in primary selection
    const primaryMinRow = Math.min(multiSelection.primary.start.row, multiSelection.primary.end.row);
    const primaryMaxRow = Math.max(multiSelection.primary.start.row, multiSelection.primary.end.row);
    const primaryMinCol = Math.min(multiSelection.primary.start.col, multiSelection.primary.end.col);
    const primaryMaxCol = Math.max(multiSelection.primary.start.col, multiSelection.primary.end.col);
    
    const isInPrimary = row >= primaryMinRow && row <= primaryMaxRow && 
                        col >= primaryMinCol && col <= primaryMaxCol;
    
    if (isInPrimary) {
      const isTopEdge = row === primaryMinRow;
      const isBottomEdge = row === primaryMaxRow;
      const isLeftEdge = col === primaryMinCol;
      const isRightEdge = col === primaryMaxCol;
      
      let borderClasses = 'dancing-ants';
      
      if (isTopEdge) borderClasses += ' border-t-2 border-t-[#127d42]';
      if (isBottomEdge) borderClasses += ' border-b-2 border-b-[#127d42]';
      if (isLeftEdge) borderClasses += ' border-l-2 border-l-[#127d42]';
      if (isRightEdge) borderClasses += ' border-r-2 border-r-[#127d42]';
      
      return borderClasses;
    }
    
    // Check additional selections
    for (const sel of multiSelection.additional) {
      const selMinRow = Math.min(sel.start.row, sel.end.row);
      const selMaxRow = Math.max(sel.start.row, sel.end.row);
      const selMinCol = Math.min(sel.start.col, sel.end.col);
      const selMaxCol = Math.max(sel.start.col, sel.end.col);
      
      if (row >= selMinRow && row <= selMaxRow && col >= selMinCol && col <= selMaxCol) {
        const isTopEdge = row === selMinRow;
        const isBottomEdge = row === selMaxRow;
        const isLeftEdge = col === selMinCol;
        const isRightEdge = col === selMaxCol;
        
        let borderClasses = '';
        
        if (isTopEdge) borderClasses += ' border-t-2 border-t-[#127d42]';
        if (isBottomEdge) borderClasses += ' border-b-2 border-b-[#127d42]';
        if (isLeftEdge) borderClasses += ' border-l-2 border-l-[#127d42]';
        if (isRightEdge) borderClasses += ' border-r-2 border-r-[#127d42]';
        
        return borderClasses;
      }
    }
    
    return '';
  }, [multiSelection, isInSelection]);

  const handleCellClick = (row: number, col: number, isShiftClick: boolean = false, isCtrlClick: boolean = false) => {
    if (editingCell) {
      finishEditing();
    }

    const cellRef = getCellRef(row, col);

    // Handle formula building mode
    if (isFormulaBuildingMode && onCellClickInFormulaMode) {
      onCellClickInFormulaMode(cellRef, isShiftClick);
      return;
    }

    if (isCtrlClick) {
      // Add to multi-selection
      const newSelection: Selection = {
        start: { row, col },
        end: { row, col }
      };
      
      setMultiSelection(prev => ({
        primary: newSelection,
        additional: [...prev.additional, prev.primary]
      }));
      
      setSelection(newSelection);
    } else if (isShiftClick && selection.start) {
      // Extend current selection
      const newSelection: Selection = {
        start: selection.start,
        end: { row, col }
      };
      
      setSelection(newSelection);
      setMultiSelection({
        primary: newSelection,
        additional: []
      });
    } else {
      // Single cell selection - clear multi-selection
      const newSelection: Selection = {
        start: { row, col },
        end: { row, col }
      };
      
      setSelection(newSelection);
      setMultiSelection({
        primary: newSelection,
        additional: []
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

  const handleMouseDown = (row: number, col: number, e: React.MouseEvent) => {
    setIsSelecting(true);
    handleCellClick(row, col, e.shiftKey, e.ctrlKey);
  };

  const handleMouseOver = (row: number, col: number) => {
    if (isSelecting) {
      const newSelection: Selection = {
        start: selection.start,
        end: { row, col }
      };
      
      setSelection(newSelection);
      
      // Update primary selection if we're dragging (not ctrl-clicking)
      if (multiSelection.additional.length === 0 || !multiSelection.additional.some(sel => 
        sel.start.row === selection.start.row && sel.start.col === selection.start.col)) {
        setMultiSelection(prev => ({
          ...prev,
          primary: newSelection
        }));
      }
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
          e.preventDefault(); // Prevent the character from being typed twice
          setEditingCell({ row, col });
          setEditValue(e.key);
          setTimeout(() => {
            if (editInputRef.current) {
              editInputRef.current.focus();
              // Move cursor to the end of the input
              editInputRef.current.setSelectionRange(1, 1);
            }
          }, 0);
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
    onMultiSelectionChange?.(multiSelection);
  }, [selection, multiSelection, cellData, onCellSelect, onSelectionChange, onMultiSelectionChange]);

  useEffect(() => {
    document.addEventListener('mouseup', handleMouseUp);
    return () => document.removeEventListener('mouseup', handleMouseUp);
  }, []);

  // Handle GPT responses
  useEffect(() => {
    const handleGptResponse = (event: CustomEvent) => {
      const { requestId, result, prompt } = event.detail;
      
      // Find cells with GPT formulas that match this requestId or prompt
      Object.entries(cellData).forEach(([cellRef, data]) => {
        if (data.formula && data.formula.includes('GPT(') && data.value && 
            (data.value.includes(requestId) || 
             (data.formula.includes('"' + prompt + '"') || data.formula.includes("'" + prompt + "'")))) {
          // Update the cell with the GPT response
          onCellUpdate(cellRef, {
            value: result,
            formula: data.formula
          });
        }
      });
    };

    window.addEventListener('gpt-response', handleGptResponse as EventListener);
    return () => window.removeEventListener('gpt-response', handleGptResponse as EventListener);
  }, [cellData, onCellUpdate]);

  return (
    <div 
      className="h-full overflow-auto bg-white focus:outline-none select-none" 
      tabIndex={0}
      onKeyDown={handleKeyDown}
      onScroll={(e) => onScroll?.(e.currentTarget.scrollLeft)}
      ref={gridRef}
    >
      <div className="inline-block min-w-full">
        {/* Data Rows */}
        {Array.from({ length: GRID_ROWS }, (_, rowIndex) => (
          <div key={rowIndex} className="flex">
            {/* Row Header */}
            <div
              className={`w-12 h-6 border-r border-b border-gray-300 flex items-center justify-end pr-2 text-sm font-medium cursor-pointer select-none sticky left-0 z-10 ${
                isRowSelected(rowIndex)
                  ? 'text-[#127d42] border-r-2 border-r-[#127d42]'
                  : 'bg-gray-100 hover:bg-gray-200 text-gray-600'
              }`}
              style={{
                backgroundColor: isRowSelected(rowIndex) ? '#caead8' : undefined
              }}
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
              const isActiveCell = multiSelection.primary.start.row === rowIndex && 
                                 multiSelection.primary.start.col === colIndex;
              const isEditing = editingCell?.row === rowIndex && editingCell?.col === colIndex;
              const rangeBorders = getRangeBorders(rowIndex, colIndex);
              
              // Check if this is the bottom-right cell of the primary selection
              const maxRow = Math.max(multiSelection.primary.start.row, multiSelection.primary.end.row);
              const maxCol = Math.max(multiSelection.primary.start.col, multiSelection.primary.end.col);
              const isBottomRightCell = rowIndex === maxRow && colIndex === maxCol;
              
              // Check if cell is in formula reference
              const formulaRef = isInFormulaReference(cellRef);
              const isRangeStart = rangeSelectionStart === cellRef;
              
              // Get cell formatting styles
              const cellStyles = getCellStyles(cellInfo.format);
              const formattedValue = formatCellValue(cellInfo.value, cellInfo.format);
              
              // Check for spillover
              const fontSize = (cellStyles.fontSize as number) || 14;
              const fontFamily = (cellStyles.fontFamily as string) || '"Aptos Narrow (Body)", "Segoe UI", system-ui, sans-serif';
              const currentCellWidth = columnWidths[colIndex] || DEFAULT_COLUMN_WIDTH;
              const { shouldSpill, spillCells } = shouldSpillOver(
                formattedValue, 
                currentCellWidth,
                fontSize,
                fontFamily,
                rowIndex,
                colIndex,
                cellData,
                columnWidths
              );
              
              // Check if this cell is being spilled into by a cell to the left
              let isSpilledInto = false;
              let spillSourceCol = -1;
              for (let i = colIndex - 1; i >= 0; i--) {
                const leftCellRef = getCellRef(rowIndex, i);
                const leftCellInfo = cellData[leftCellRef] || { value: "" };
                if (leftCellInfo.value && leftCellInfo.value.trim() !== '') {
                  const leftCellStyles = getCellStyles(leftCellInfo.format);
                  const leftFontSize = (leftCellStyles.fontSize as number) || 14;
                  const leftFontFamily = (leftCellStyles.fontFamily as string) || '"Aptos Narrow (Body)", "Segoe UI", system-ui, sans-serif';
                  const leftFormattedValue = formatCellValue(leftCellInfo.value, leftCellInfo.format);
                  const leftCellWidth = columnWidths[i] || DEFAULT_COLUMN_WIDTH;
                  
                  const { shouldSpill: leftShouldSpill, spillCells: leftSpillCells } = shouldSpillOver(
                    leftFormattedValue,
                    leftCellWidth,
                    leftFontSize,
                    leftFontFamily,
                    rowIndex,
                    i,
                    cellData,
                    columnWidths
                  );
                  
                  if (leftShouldSpill && (i + leftSpillCells) >= colIndex) {
                    isSpilledInto = true;
                    spillSourceCol = i;
                    break;
                  }
                }
                // If we hit a cell with content, stop looking
                if (leftCellInfo.value && leftCellInfo.value.trim() !== '') {
                  break;
                }
              }
              
              // Determine background color and styles
              let bgColor = 'bg-white hover:bg-gray-50';
              let cellStyle: React.CSSProperties = { ...cellStyles };
              
              // If this cell is being spilled into, remove its background
              if (isSpilledInto && !isSelected) {
                bgColor = 'bg-transparent hover:bg-gray-50';
                // Remove background color from spilled content
                cellStyle = {
                  ...cellStyle,
                  backgroundColor: 'transparent'
                };
              }
              
              if (isSelected && isActiveCell) {
                // Only the active cell (top-left of selection) is transparent
                bgColor = 'bg-transparent';
              } else if (isSelected) {
                // Other selected cells have selection background color
                cellStyle = {
                  ...cellStyle,
                  backgroundColor: '#e8f2ec'
                };
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

              const renderCellWidth = columnWidths[colIndex] || DEFAULT_COLUMN_WIDTH;
              
              return (
                <div
                  key={colIndex}
                  className={`h-6 border-r border-b border-gray-300 relative cursor-cell select-none ${bgColor} ${rangeBorders}`}
                  style={{ 
                    ...cellStyle,
                    width: `${renderCellWidth}px`,
                    minWidth: `${renderCellWidth}px`,
                    flexShrink: 0
                  }}
                  onMouseDown={(e) => handleMouseDown(rowIndex, colIndex, e)}
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
                      className="w-full h-full px-1 border-0 outline-none bg-transparent"
                      style={{
                        ...cellStyles,
                        fontSize: cellStyles.fontSize || '14px',
                        fontFamily: cellStyles.fontFamily || '"Aptos Narrow (Body)", "Segoe UI", system-ui, sans-serif',
                        display: 'flex',
                        alignItems: 'center',
                        boxSizing: 'border-box',
                        margin: 0,
                        padding: '0 4px',
                        lineHeight: '1.2'
                      }}
                      autoFocus
                    />
                  ) : (
                    <div 
                      className="w-full h-full px-1 flex items-center overflow-hidden select-none relative"
                      style={{
                        ...cellStyles,
                        fontSize: cellStyles.fontSize || '14px',
                        fontFamily: cellStyles.fontFamily || '"Aptos Narrow (Body)", "Segoe UI", system-ui, sans-serif',
                        boxSizing: 'border-box',
                        margin: 0,
                        padding: '0 4px',
                        lineHeight: '1.2'
                      }}
                    >
                      {/* Normal cell content */}
                      {formattedValue && !isSpilledInto && (
                        <span 
                          style={{
                            display: 'block',
                            whiteSpace: 'nowrap',
                            overflow: shouldSpill ? 'visible' : 'hidden',
                            position: shouldSpill ? 'absolute' : 'static',
                            left: shouldSpill ? '4px' : 'auto',
                            top: shouldSpill ? '50%' : 'auto',
                            transform: shouldSpill ? 'translateY(-50%)' : 'none',
                            width: shouldSpill ? (() => {
                              let spillWidth = currentCellWidth - 8;
                              for (let i = 1; i <= spillCells; i++) {
                                spillWidth += (columnWidths[colIndex + i] || DEFAULT_COLUMN_WIDTH);
                              }
                              return `${spillWidth}px`;
                            })() : 'auto',
                            zIndex: shouldSpill ? 10 : 1,
                            backgroundColor: 'transparent'
                          }}
                        >
                          {formattedValue}
                        </span>
                      )}
                      
                      {/* Show content if this cell has its own content and is spilled into */}
                      {isSpilledInto && cellInfo.value && (
                        <span>{formattedValue}</span>
                      )}
                      {/* Fill handle - small square at bottom-right of selection */}
                      {isBottomRightCell && (
                        <div 
                          className="absolute bottom-0 right-0 w-[6px] h-[6px] bg-[#127d42] border border-white cursor-crosshair z-50"
                          style={{
                            transform: 'translate(4px, 4px)',
                            borderWidth: '0.5px'
                          }}
                        />
                      )}
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
