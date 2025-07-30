import { CellData, CellFormat, ClipboardData, Selection } from '../types/cellTypes';

// Helper function to get cell reference from row/col
export const getCellRef = (row: number, col: number): string => {
  const getColumnName = (index: number): string => {
    let result = '';
    while (index >= 0) {
      result = String.fromCharCode(65 + (index % 26)) + result;
      index = Math.floor(index / 26) - 1;
    }
    return result;
  };
  return `${getColumnName(col)}${row + 1}`;
};

// Parse cell reference to row/col
export const parseCellRef = (cellRef: string): { row: number; col: number } | null => {
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

// Get all cell references in a selection
export const getSelectionCellRefs = (selection: Selection): string[] => {
  const refs: string[] = [];
  const minRow = Math.min(selection.start.row, selection.end.row);
  const maxRow = Math.max(selection.start.row, selection.end.row);
  const minCol = Math.min(selection.start.col, selection.end.col);
  const maxCol = Math.max(selection.start.col, selection.end.col);
  
  for (let row = minRow; row <= maxRow; row++) {
    for (let col = minCol; col <= maxCol; col++) {
      refs.push(getCellRef(row, col));
    }
  }
  
  return refs;
};

// Apply formatting to a cell
export const applyCellFormat = (
  cellData: CellData | undefined,
  newFormat: Partial<CellFormat>
): CellData => {
  const existingData = cellData || { value: '' };
  const existingFormat = existingData.format || {};
  
  // Deep merge formatting properties to prevent overriding
  const mergedFormat: CellFormat = {
    ...existingFormat,
    ...newFormat
  };
  
  // Special handling for nested objects
  if (newFormat.borders && existingFormat.borders) {
    mergedFormat.borders = {
      ...existingFormat.borders,
      ...newFormat.borders
    };
  }
  
  if (newFormat.numberFormat && existingFormat.numberFormat) {
    mergedFormat.numberFormat = {
      ...existingFormat.numberFormat,
      ...newFormat.numberFormat
    };
  }
  
  return {
    ...existingData,
    format: mergedFormat
  };
};

// Generate CSS styles from cell format
export const getCellStyles = (format?: CellFormat): React.CSSProperties => {
  if (!format) return {};
  
  const styles: React.CSSProperties = {};
  
  // Font properties
  if (format.fontFamily) styles.fontFamily = format.fontFamily;
  if (format.fontSize) styles.fontSize = `${format.fontSize}px`;
  if (format.bold) styles.fontWeight = 'bold';
  if (format.italic) styles.fontStyle = 'italic';
  
  // Text decoration
  const decorations: string[] = [];
  if (format.underline) decorations.push('underline');
  if (format.strikethrough) decorations.push('line-through');
  if (decorations.length > 0) styles.textDecoration = decorations.join(' ');
  
  // Colors
  if (format.textColor) styles.color = format.textColor;
  if (format.backgroundColor) styles.backgroundColor = format.backgroundColor;
  
  // Alignment
  if (format.horizontalAlign) styles.textAlign = format.horizontalAlign;
  if (format.verticalAlign) {
    styles.display = 'flex';
    styles.alignItems = format.verticalAlign === 'top' ? 'flex-start' : 
                       format.verticalAlign === 'bottom' ? 'flex-end' : 'center';
  }
  
  // Borders
  if (format.borders) {
    const borderColor = format.borders.color || '#000000';
    const borderStyle = format.borders.style || 'solid';
    
    if (format.borders.top) styles.borderTop = `1px ${borderStyle} ${borderColor}`;
    if (format.borders.right) styles.borderRight = `1px ${borderStyle} ${borderColor}`;
    if (format.borders.bottom) styles.borderBottom = `1px ${borderStyle} ${borderColor}`;
    if (format.borders.left) styles.borderLeft = `1px ${borderStyle} ${borderColor}`;
  }
  
  return styles;
};

// Format number value based on number format
export const formatCellValue = (value: string, format?: CellFormat): string => {
  if (!format?.numberFormat || !value || isNaN(Number(value))) {
    return value;
  }
  
  const numValue = Number(value);
  const { type, decimals = 2, useThousandsSeparator = false } = format.numberFormat;
  
  let formatted = '';
  
  switch (type) {
    case 'number':
      formatted = numValue.toFixed(decimals);
      break;
    case 'currency':
      formatted = new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'USD',
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals
      }).format(numValue);
      break;
    case 'percentage':
      formatted = `${(numValue * 100).toFixed(decimals)}%`;
      break;
    case 'accounting':
      if (numValue < 0) {
        formatted = `(${Math.abs(numValue).toFixed(decimals)})`;
      } else {
        formatted = numValue.toFixed(decimals);
      }
      break;
    default:
      return value;
  }
  
  // Add thousands separator if requested
  if (useThousandsSeparator && type !== 'currency') {
    const parts = formatted.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    formatted = parts.join('.');
  }
  
  return formatted;
};

// Create clipboard data from selection
export const createClipboardData = (
  cellData: Record<string, CellData>,
  selection: Selection,
  operation: 'copy' | 'cut' = 'copy'
): ClipboardData => {
  const cells: Record<string, CellData> = {};
  const cellRefs = getSelectionCellRefs(selection);
  
  cellRefs.forEach(ref => {
    if (cellData[ref]) {
      cells[ref] = { ...cellData[ref] };
    }
  });
  
  return {
    cells,
    range: {
      startRow: Math.min(selection.start.row, selection.end.row),
      startCol: Math.min(selection.start.col, selection.end.col),
      endRow: Math.max(selection.start.row, selection.end.row),
      endCol: Math.max(selection.start.col, selection.end.col)
    },
    operation
  };
};

// Apply clipboard data to target selection
export const applyClipboardData = (
  currentCellData: Record<string, CellData>,
  clipboardData: ClipboardData,
  targetSelection: Selection
): Record<string, CellData> => {
  const newCellData = { ...currentCellData };
  
  const sourceWidth = clipboardData.range.endCol - clipboardData.range.startCol + 1;
  const sourceHeight = clipboardData.range.endRow - clipboardData.range.startRow + 1;
  
  const targetWidth = Math.abs(targetSelection.end.col - targetSelection.start.col) + 1;
  const targetHeight = Math.abs(targetSelection.end.row - targetSelection.start.row) + 1;
  
  // Calculate how many times to repeat the clipboard data
  const repeatX = Math.max(1, Math.ceil(targetWidth / sourceWidth));
  const repeatY = Math.max(1, Math.ceil(targetHeight / sourceHeight));
  
  const targetStartRow = Math.min(targetSelection.start.row, targetSelection.end.row);
  const targetStartCol = Math.min(targetSelection.start.col, targetSelection.end.col);
  
  // Apply clipboard data
  for (let repeatRow = 0; repeatRow < repeatY; repeatRow++) {
    for (let repeatCol = 0; repeatCol < repeatX; repeatCol++) {
      Object.entries(clipboardData.cells).forEach(([sourceRef, cellData]) => {
        const sourceCoords = parseCellRef(sourceRef);
        if (!sourceCoords) return;
        
        const relativeRow = sourceCoords.row - clipboardData.range.startRow;
        const relativeCol = sourceCoords.col - clipboardData.range.startCol;
        
        const targetRow = targetStartRow + (repeatRow * sourceHeight) + relativeRow;
        const targetCol = targetStartCol + (repeatCol * sourceWidth) + relativeCol;
        
        // Check if target is within bounds
        if (targetRow >= 0 && targetCol >= 0 && 
            targetRow <= targetSelection.end.row && 
            targetCol <= targetSelection.end.col) {
          const targetRef = getCellRef(targetRow, targetCol);
          newCellData[targetRef] = { ...cellData };
        }
      });
    }
  }
  
  // If it was a cut operation, clear the source cells
  if (clipboardData.operation === 'cut') {
    Object.keys(clipboardData.cells).forEach(ref => {
      delete newCellData[ref];
    });
  }
  
  return newCellData;
};
