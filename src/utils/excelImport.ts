import * as XLSX from 'xlsx';
import { CellData, CellFormat } from '@/types/cellTypes';

export interface ImportedExcelData {
  sheets: {
    name: string;
    data: Record<string, CellData>;
  }[];
}

/**
 * Convert Excel column index to letter notation (0 -> A, 1 -> B, etc.)
 */
function columnIndexToLetter(index: number): string {
  let letter = '';
  while (index >= 0) {
    letter = String.fromCharCode(65 + (index % 26)) + letter;
    index = Math.floor(index / 26) - 1;
  }
  return letter;
}

/**
 * Convert Excel color to hex format
 */
function excelColorToHex(color: any): string { // eslint-disable-line @typescript-eslint/no-explicit-any
  if (!color) return '#000000';
  
  // Handle theme colors
  if (color.theme !== undefined) {
    // Excel theme colors (approximate values)
    const themeColors = [
      '#FFFFFF', '#000000', '#E7E6E6', '#44546A',
      '#5B9BD5', '#ED7D31', '#A5A5A5', '#FFC000',
      '#4472C4', '#70AD47'
    ];
    return themeColors[color.theme] || '#000000';
  }
  
  // Handle RGB colors
  if (color.rgb) {
    // Excel stores RGB as AARRGGBB (alpha, red, green, blue)
    const rgb = color.rgb;
    if (typeof rgb === 'string' && rgb.length >= 6) {
      // Skip alpha channel if present
      const hex = rgb.length === 8 ? rgb.slice(2) : rgb;
      return '#' + hex.toUpperCase();
    }
  }
  
  // Handle indexed colors
  if (color.indexed !== undefined) {
    // Excel indexed color palette (simplified)
    const indexedColors = [
      '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF',
      '#FFFF00', '#FF00FF', '#00FFFF', '#000000', '#FFFFFF',
      '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF',
      '#00FFFF', '#800000', '#008000', '#000080', '#808000',
      '#800080', '#008080', '#C0C0C0', '#808080', '#9999FF',
      '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080',
      '#0066CC', '#CCCCFF', '#000080', '#FF00FF', '#FFFF00',
      '#00FFFF', '#800080', '#800000', '#008080', '#0000FF',
      '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF',
      '#FF99CC', '#CC99FF', '#FFCC99', '#3366FF', '#33CCCC',
      '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699',
      '#969696', '#003366', '#339966', '#003300', '#333300',
      '#993300', '#993366', '#333399', '#333333'
    ];
    return indexedColors[color.indexed] || '#000000';
  }
  
  return '#000000';
}

/**
 * Extract cell styling from workbook styles
 */
function getCellStyle(cell: XLSX.CellObject, workbook: XLSX.WorkBook): any { // eslint-disable-line @typescript-eslint/no-explicit-any
  if (!cell.s || typeof cell.s !== 'number') return null;
  
  const styles = (workbook as any).Styles; // eslint-disable-line @typescript-eslint/no-explicit-any
  if (!styles) return null;
  
  const cellXfId = cell.s;
  const cellXf = styles.CellXf?.[cellXfId];
  if (!cellXf) return null;
  
  const result: any = {}; // eslint-disable-line @typescript-eslint/no-explicit-any
  
  // Extract font
  if (cellXf.fontId !== undefined && styles.Fonts) {
    const font = styles.Fonts[cellXf.fontId];
    if (font) {
      result.font = {
        name: font.name || 'Calibri',
        sz: font.sz || 11,
        bold: font.bold || false,
        italic: font.italic || false,
        underline: font.underline || false,
        strike: font.strike || false,
        color: font.color
      };
    }
  }
  
  // Extract fill
  if (cellXf.fillId !== undefined && styles.Fills) {
    const fill = styles.Fills[cellXf.fillId];
    if (fill) {
      result.fill = fill;
    }
  }
  
  // Extract border
  if (cellXf.borderId !== undefined && styles.Borders) {
    const border = styles.Borders[cellXf.borderId];
    if (border) {
      result.border = border;
    }
  }
  
  // Extract alignment
  if (cellXf.alignment) {
    result.alignment = cellXf.alignment;
  }
  
  // Extract number format
  if (cellXf.numFmtId !== undefined) {
    let numFmt = '';
    
    // Check built-in formats
    const builtInFormats: Record<number, string> = {
      0: 'General',
      1: '0',
      2: '0.00',
      3: '#,##0',
      4: '#,##0.00',
      9: '0%',
      10: '0.00%',
      11: '0.00E+00',
      12: '# ?/?',
      13: '# ??/??',
      14: 'mm-dd-yy',
      15: 'd-mmm-yy',
      16: 'd-mmm',
      17: 'mmm-yy',
      18: 'h:mm AM/PM',
      19: 'h:mm:ss AM/PM',
      20: 'h:mm',
      21: 'h:mm:ss',
      22: 'm/d/yy h:mm',
      37: '#,##0 ;(#,##0)',
      38: '#,##0 ;[Red](#,##0)',
      39: '#,##0.00;(#,##0.00)',
      40: '#,##0.00;[Red](#,##0.00)',
      41: '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
      42: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
      43: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
      44: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
      45: 'mm:ss',
      46: '[h]:mm:ss',
      47: 'mmss.0',
      48: '##0.0E+0',
      49: '@'
    };
    
    if (builtInFormats[cellXf.numFmtId]) {
      numFmt = builtInFormats[cellXf.numFmtId];
    } else if (styles.NumFmts) {
      // Check custom formats
      const customFormat = styles.NumFmts.find((fmt: any) => fmt.numFmtId === cellXf.numFmtId); // eslint-disable-line @typescript-eslint/no-explicit-any
      if (customFormat) {
        numFmt = customFormat.formatCode || '';
      }
    }
    
    result.numFmt = numFmt;
  }
  
  return result;
}

/**
 * Import an Excel file and convert it to our internal data format with full styling
 */
export async function importExcelFile(file: File): Promise<ImportedExcelData> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (!data) {
          reject(new Error('Failed to read file'));
          return;
        }
        
        // Parse the workbook with styles
        const workbook = XLSX.read(data, { 
          type: 'array',
          cellStyles: true,  // Enable cell styles parsing
          cellNF: true,      // Enable number format parsing
          cellHTML: false    // We don't need HTML
        });
        
        console.log('Workbook loaded:', workbook);
        console.log('Workbook Styles:', (workbook as any).Styles); // eslint-disable-line @typescript-eslint/no-explicit-any
        
        // Convert all sheets to our format
        const sheets = workbook.SheetNames.map(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          const cellData: Record<string, CellData> = {};
          
          // Get the range of the worksheet
          const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
          
          // Iterate through all cells in the range
          for (let row = range.s.r; row <= range.e.r; row++) {
            for (let col = range.s.c; col <= range.e.c; col++) {
              const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
              const cell = worksheet[cellAddress];
              
              if (cell) {
                // Convert to our cell reference format (A1, B2, etc.)
                const ourCellRef = `${columnIndexToLetter(col)}${row + 1}`;
                
                // Extract cell value and formula
                let value = '';
                let formula: string | undefined;
                
                // Handle different cell types
                if (cell.f) {
                  // Cell has a formula
                  formula = '=' + cell.f;
                  // Use formatted value if available, otherwise raw value
                  value = cell.w || (cell.v !== undefined ? String(cell.v) : '');
                } else if (cell.w) {
                  // Use formatted value
                  value = cell.w;
                } else if (cell.v !== undefined) {
                  // Use raw value
                  value = String(cell.v);
                }
                
                // Extract cell styling from workbook styles
                const cellStyle = getCellStyle(cell, workbook);
                console.log(`Cell ${ourCellRef} style:`, cellStyle);
                
                // Default format
                const cellFormat: CellFormat = {
                  fontFamily: 'Calibri',
                  fontSize: 11,
                  bold: false,
                  italic: false,
                  underline: false,
                  textColor: '#000000',
                  backgroundColor: 'transparent',
                  horizontalAlign: 'left',
                  verticalAlign: 'middle',
                  numberFormat: { type: 'general' }
                };
                
                if (cellStyle) {
                  // Extract font properties
                  if (cellStyle.font) {
                    cellFormat.fontFamily = cellStyle.font.name || 'Calibri';
                    cellFormat.fontSize = cellStyle.font.sz || 11;
                    cellFormat.bold = cellStyle.font.bold || false;
                    cellFormat.italic = cellStyle.font.italic || false;
                    cellFormat.underline = cellStyle.font.underline || false;
                    
                    if (cellStyle.font.color) {
                      cellFormat.textColor = excelColorToHex(cellStyle.font.color);
                    }
                  }
                  
                  // Extract fill/background color
                  if (cellStyle.fill) {
                    // Handle pattern fills
                    if (cellStyle.fill.fgColor) {
                      const bgColor = excelColorToHex(cellStyle.fill.fgColor);
                      // Don't set white as background (Excel default)
                      if (bgColor !== '#FFFFFF' && bgColor !== '#000000') {
                        cellFormat.backgroundColor = bgColor;
                      }
                    } else if (cellStyle.fill.bgColor) {
                      const bgColor = excelColorToHex(cellStyle.fill.bgColor);
                      if (bgColor !== '#FFFFFF' && bgColor !== '#000000') {
                        cellFormat.backgroundColor = bgColor;
                      }
                    }
                  }
                  
                  // Extract alignment
                  if (cellStyle.alignment) {
                    if (cellStyle.alignment.horizontal === 'center') {
                      cellFormat.horizontalAlign = 'center';
                    } else if (cellStyle.alignment.horizontal === 'right') {
                      cellFormat.horizontalAlign = 'right';
                    } else if (cell.t === 'n' && !cellStyle.alignment.horizontal) {
                      // Numbers are right-aligned by default in Excel
                      cellFormat.horizontalAlign = 'right';
                    }
                    
                    if (cellStyle.alignment.vertical === 'top') {
                      cellFormat.verticalAlign = 'top';
                    } else if (cellStyle.alignment.vertical === 'bottom') {
                      cellFormat.verticalAlign = 'bottom';
                    }
                  } else if (cell.t === 'n') {
                    // Default number alignment
                    cellFormat.horizontalAlign = 'right';
                  }
                  
                  // Extract number format
                  if (cellStyle.numFmt) {
                    cellFormat.numberFormat = detectNumberFormat(cellStyle.numFmt);
                  }
                }
                
                // Check for direct color properties on the cell itself
                const cellAny = cell as any; // eslint-disable-line @typescript-eslint/no-explicit-any
                if (cellAny.fgColor) {
                  const bgColor = excelColorToHex(cellAny.fgColor);
                  console.log(`Cell ${ourCellRef} has direct fgColor:`, cellAny.fgColor, 'converted to:', bgColor);
                  // fgColor is typically the background fill color in Excel
                  // Don't filter out yellow (#FFFF00) or other valid colors
                  if (bgColor && bgColor !== 'transparent') {
                    cellFormat.backgroundColor = bgColor;
                  }
                } else if (cellAny.bgColor) {
                  const bgColor = excelColorToHex(cellAny.bgColor);
                  console.log(`Cell ${ourCellRef} has direct bgColor:`, cellAny.bgColor, 'converted to:', bgColor);
                  if (bgColor && bgColor !== 'transparent' && bgColor !== '#FFFFFF') {
                    cellFormat.backgroundColor = bgColor;
                  }
                }
                
                // Also check if cell.s is an object with direct style properties
                if (cell.s && typeof cell.s === 'object') {
                  const directStyle = cell.s as any; // eslint-disable-line @typescript-eslint/no-explicit-any
                  
                  // Check for fill colors in cell.s
                  if (directStyle.fgColor) {
                    const bgColor = excelColorToHex(directStyle.fgColor);
                    console.log(`Cell ${ourCellRef} has cell.s.fgColor:`, directStyle.fgColor, 'converted to:', bgColor);
                    if (bgColor && bgColor !== 'transparent') {
                      cellFormat.backgroundColor = bgColor;
                    }
                  } else if (directStyle.bgColor) {
                    const bgColor = excelColorToHex(directStyle.bgColor);
                    console.log(`Cell ${ourCellRef} has cell.s.bgColor:`, directStyle.bgColor, 'converted to:', bgColor);
                    if (bgColor && bgColor !== 'transparent' && bgColor !== '#FFFFFF') {
                      cellFormat.backgroundColor = bgColor;
                    }
                  }
                }
                
                cellData[ourCellRef] = {
                  value,
                  formula,
                  format: cellFormat
                };
              }
            }
          }
          
          return {
            name: sheetName,
            data: cellData
          };
        });
        
        resolve({ sheets });
      } catch (error) {
        console.error('Error importing Excel file:', error);
        reject(error);
      }
    };
    
    reader.onerror = () => {
      reject(new Error('Failed to read file'));
    };
    
    // Read the file as ArrayBuffer
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Detect the number format from a format string
 */
function detectNumberFormat(fmt: string): CellFormat['numberFormat'] {
  if (!fmt || fmt === 'General') {
    return { type: 'general' };
  }
  
  // Check for percentage
  if (fmt.includes('%')) {
    // Count decimal places
    const decimals = (fmt.match(/0/g) || []).length - 1;
    return { type: 'percentage', decimals: Math.max(0, decimals) };
  }
  
  // Check for currency
  if (fmt.includes('$') || fmt.includes('€') || fmt.includes('£')) {
    const decimals = fmt.includes('.') ? (fmt.split('.')[1].match(/0/g) || []).length : 0;
    return { type: 'currency', decimals };
  }
  
  // Check for accounting
  if (fmt.includes('_') && (fmt.includes('$') || fmt.includes('€') || fmt.includes('£'))) {
    const decimals = fmt.includes('.') ? (fmt.split('.')[1].match(/0/g) || []).length : 0;
    return { type: 'accounting', decimals };
  }
  
  // Check for thousands separator
  if (fmt.includes(',')) {
    const decimals = fmt.includes('.') ? (fmt.split('.')[1].match(/0/g) || []).length : 0;
    return { type: 'number', decimals, useThousandsSeparator: true };
  }
  
  // Default number format
  if (fmt.includes('0')) {
    const decimals = fmt.includes('.') ? (fmt.split('.')[1].match(/0/g) || []).length : 0;
    return { type: 'number', decimals };
  }
  
  return { type: 'general' };
}

/**
 * Create a file input element and handle file selection
 */
export function openExcelFile(): Promise<File> {
  return new Promise((resolve, reject) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xls';
    
    input.onchange = (e) => {
      const target = e.target as HTMLInputElement;
      const file = target.files?.[0];
      
      if (file) {
        resolve(file);
      } else {
        reject(new Error('No file selected'));
      }
    };
    
    // Trigger file selection dialog
    input.click();
  });
}
