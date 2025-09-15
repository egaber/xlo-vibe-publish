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
 * Import an Excel file and convert it to our internal data format
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
        
        // Parse the workbook
        const workbook = XLSX.read(data, { type: 'array' });
        
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
                  value = cell.v !== undefined ? String(cell.v) : '';
                } else if (cell.v !== undefined) {
                  // Cell has a value
                  value = String(cell.v);
                }
                
                // Create cell data with basic formatting
                const cellFormat: CellFormat = {
                  fontFamily: 'Calibri',
                  fontSize: 11,
                  bold: cell.s?.font?.bold || false,
                  italic: cell.s?.font?.italic || false,
                  underline: cell.s?.font?.underline || false,
                  textColor: '#000000',
                  backgroundColor: 'transparent',
                  horizontalAlign: 'left',
                  verticalAlign: 'middle',
                  numberFormat: detectNumberFormat(cell)
                };
                
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
 * Detect the number format from an Excel cell
 */
function detectNumberFormat(cell: XLSX.CellObject): CellFormat['numberFormat'] {
  if (!cell.t || !cell.z) {
    return { type: 'general' };
  }
  
  // Check cell type
  if (cell.t === 'n') {
    // Number type
    const format = String(cell.z);
    
    // Check for percentage
    if (format.includes('%')) {
      return { type: 'percentage', decimals: 2 };
    }
    
    // Check for currency
    if (format.includes('$') || format.includes('€') || format.includes('£')) {
      return { type: 'currency', decimals: 2 };
    }
    
    // Check for accounting
    if (format.includes('_') && format.includes('$')) {
      return { type: 'accounting', decimals: 2 };
    }
    
    // Default number format
    return { type: 'number', decimals: 0 };
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
