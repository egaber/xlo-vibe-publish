// Cell formatting types
export interface CellFormat {
  // Font properties
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  
  // Colors
  textColor?: string;
  backgroundColor?: string;
  
  // Alignment
  horizontalAlign?: 'left' | 'center' | 'right';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  
  // Borders
  borders?: {
    top?: boolean;
    right?: boolean;
    bottom?: boolean;
    left?: boolean;
    color?: string;
    style?: 'solid' | 'dashed' | 'dotted';
  };
  
  // Number formatting
  numberFormat?: {
    type: 'general' | 'number' | 'currency' | 'percentage' | 'accounting';
    decimals?: number;
    useThousandsSeparator?: boolean;
  };
}

export interface CellData {
  value: string;
  formula?: string;
  format?: CellFormat;
}

// Clipboard data
export interface ClipboardData {
  cells: Record<string, CellData>;
  range: {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
  };
  operation?: 'copy' | 'cut';
}

// Selection interface
export interface Selection {
  start: { row: number; col: number };
  end: { row: number; col: number };
}

// Ribbon actions
export interface RibbonActions {
  // Clipboard
  copy: () => void;
  paste: () => void;
  cut: () => void;
  
  // Font formatting
  toggleBold: () => void;
  toggleItalic: () => void;
  toggleUnderline: () => void;
  setFontFamily: (family: string) => void;
  setFontSize: (size: number) => void;
  increaseFontSize: () => void;
  decreaseFontSize: () => void;
  
  // Colors
  setTextColor: (color: string) => void;
  setBackgroundColor: (color: string) => void;
  
  // Alignment
  setHorizontalAlignment: (align: 'left' | 'center' | 'right') => void;
  setVerticalAlignment: (align: 'top' | 'middle' | 'bottom') => void;
  
  // Number formatting
  setNumberFormat: (format: CellFormat['numberFormat']) => void;
  increaseDecimals: () => void;
  decreaseDecimals: () => void;
  
  // Undo/Redo
  undo: () => void;
  redo: () => void;
}
