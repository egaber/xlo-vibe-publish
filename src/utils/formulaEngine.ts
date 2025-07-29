// Formula engine for Excel-like calculations
export interface CellData {
  value: string;
  formula?: string;
}

export interface FormulaContext {
  getCellValue: (cellRef: string) => string;
  setCellValue: (cellRef: string, value: string) => void;
}

// Parse cell reference (e.g., "A1" -> {col: 0, row: 0})
export const parseCellRef = (cellRef: string): { col: number; row: number } | null => {
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

// Convert column index to Excel column name (0 -> "A", 25 -> "Z", 26 -> "AA")
export const getColumnName = (index: number): string => {
  let result = '';
  while (index >= 0) {
    result = String.fromCharCode(65 + (index % 26)) + result;
    index = Math.floor(index / 26) - 1;
  }
  return result;
};

// Parse range reference (e.g., "A1:B2" -> array of cell references)
export const parseRange = (rangeRef: string): string[] => {
  const parts = rangeRef.split(':');
  if (parts.length !== 2) return [rangeRef]; // Single cell
  
  const start = parseCellRef(parts[0]);
  const end = parseCellRef(parts[1]);
  
  if (!start || !end) return [rangeRef];
  
  const cells: string[] = [];
  for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
    for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col++) {
      cells.push(`${getColumnName(col)}${row + 1}`);
    }
  }
  
  return cells;
};

// Types for function arguments and return values
type FunctionArg = string | number;
type FunctionResult = string | number;

// Built-in functions
export const FUNCTIONS: Record<string, (args: FunctionArg[], context: FormulaContext) => FunctionResult> = {
  SUM: (args: FunctionArg[], context: FormulaContext) => {
    let total = 0;
    for (const arg of args) {
      if (typeof arg === 'string' && arg.includes(':')) {
        // Range reference
        const cells = parseRange(arg);
        for (const cellRef of cells) {
          const value = parseFloat(context.getCellValue(cellRef)) || 0;
          total += value;
        }
      } else if (typeof arg === 'string' && parseCellRef(arg)) {
        // Single cell reference
        const value = parseFloat(context.getCellValue(arg)) || 0;
        total += value;
      } else {
        // Literal value
        total += parseFloat(String(arg)) || 0;
      }
    }
    return total;
  },
  
  AVERAGE: (args: FunctionArg[], context: FormulaContext) => {
    let total = 0;
    let count = 0;
    for (const arg of args) {
      if (typeof arg === 'string' && arg.includes(':')) {
        // Range reference
        const cells = parseRange(arg);
        for (const cellRef of cells) {
          const value = parseFloat(context.getCellValue(cellRef));
          if (!isNaN(value)) {
            total += value;
            count++;
          }
        }
      } else if (typeof arg === 'string' && parseCellRef(arg)) {
        // Single cell reference
        const value = parseFloat(context.getCellValue(arg));
        if (!isNaN(value)) {
          total += value;
          count++;
        }
      } else {
        // Literal value
        const value = parseFloat(String(arg));
        if (!isNaN(value)) {
          total += value;
          count++;
        }
      }
    }
    return count > 0 ? total / count : 0;
  },
  
  COUNT: (args: FunctionArg[], context: FormulaContext) => {
    let count = 0;
    for (const arg of args) {
      if (typeof arg === 'string' && arg.includes(':')) {
        // Range reference
        const cells = parseRange(arg);
        for (const cellRef of cells) {
          const value = context.getCellValue(cellRef);
          if (value && !isNaN(parseFloat(value))) {
            count++;
          }
        }
      } else if (typeof arg === 'string' && parseCellRef(arg)) {
        // Single cell reference
        const value = context.getCellValue(arg);
        if (value && !isNaN(parseFloat(value))) {
          count++;
        }
      } else {
        // Literal value
        if (!isNaN(parseFloat(String(arg)))) {
          count++;
        }
      }
    }
    return count;
  },
  
  MAX: (args: FunctionArg[], context: FormulaContext) => {
    let max = -Infinity;
    for (const arg of args) {
      if (typeof arg === 'string' && arg.includes(':')) {
        // Range reference
        const cells = parseRange(arg);
        for (const cellRef of cells) {
          const value = parseFloat(context.getCellValue(cellRef));
          if (!isNaN(value)) {
            max = Math.max(max, value);
          }
        }
      } else if (typeof arg === 'string' && parseCellRef(arg)) {
        // Single cell reference
        const value = parseFloat(context.getCellValue(arg));
        if (!isNaN(value)) {
          max = Math.max(max, value);
        }
      } else {
        // Literal value
        const value = parseFloat(String(arg));
        if (!isNaN(value)) {
          max = Math.max(max, value);
        }
      }
    }
    return max === -Infinity ? 0 : max;
  },
  
  MIN: (args: FunctionArg[], context: FormulaContext) => {
    let min = Infinity;
    for (const arg of args) {
      if (typeof arg === 'string' && arg.includes(':')) {
        // Range reference
        const cells = parseRange(arg);
        for (const cellRef of cells) {
          const value = parseFloat(context.getCellValue(cellRef));
          if (!isNaN(value)) {
            min = Math.min(min, value);
          }
        }
      } else if (typeof arg === 'string' && parseCellRef(arg)) {
        // Single cell reference
        const value = parseFloat(context.getCellValue(arg));
        if (!isNaN(value)) {
          min = Math.min(min, value);
        }
      } else {
        // Literal value
        const value = parseFloat(String(arg));
        if (!isNaN(value)) {
          min = Math.min(min, value);
        }
      }
    }
    return min === Infinity ? 0 : min;
  },
  
  CONCATENATE: (args: FunctionArg[], context: FormulaContext) => {
    let result = '';
    for (const arg of args) {
      if (typeof arg === 'string' && parseCellRef(arg)) {
        // Single cell reference
        result += context.getCellValue(arg);
      } else {
        // Literal value
        result += String(arg);
      }
    }
    return result;
  },
  
  IF: (args: FunctionArg[], context: FormulaContext) => {
    if (args.length < 2) return '';
    
    const condition = args[0];
    const trueValue = args[1];
    const falseValue = args.length > 2 ? args[2] : '';
    
    // Evaluate condition
    let conditionResult = false;
    if (typeof condition === 'string' && parseCellRef(condition)) {
      const cellValue = context.getCellValue(condition);
      conditionResult = Boolean(cellValue && cellValue !== '0');
    } else {
      conditionResult = Boolean(condition);
    }
    
    return conditionResult ? trueValue : falseValue;
  }
};

// Simple expression parser for basic arithmetic
export const evaluateExpression = (expr: string, context: FormulaContext): string | number => {
  // Remove whitespace
  expr = expr.trim();
  
  // Handle cell references
  if (parseCellRef(expr)) {
    return context.getCellValue(expr);
  }
  
  // Handle numbers
  if (!isNaN(parseFloat(expr))) {
    return parseFloat(expr);
  }
  
  // Handle strings (quoted)
  if (expr.startsWith('"') && expr.endsWith('"')) {
    return expr.slice(1, -1);
  }
  
  // Handle simple arithmetic operations
  const operators = ['+', '-', '*', '/'];
  for (const op of operators) {
    const parts = expr.split(op);
    if (parts.length === 2) {
      const left = evaluateExpression(parts[0].trim(), context);
      const right = evaluateExpression(parts[1].trim(), context);
      
      const leftNum = parseFloat(String(left));
      const rightNum = parseFloat(String(right));
      
      if (!isNaN(leftNum) && !isNaN(rightNum)) {
        switch (op) {
          case '+': return leftNum + rightNum;
          case '-': return leftNum - rightNum;
          case '*': return leftNum * rightNum;
          case '/': return rightNum !== 0 ? leftNum / rightNum : '#DIV/0!';
        }
      }
    }
  }
  
  return expr;
};

// Parse function call (e.g., "SUM(A1,B1)" -> {name: "SUM", args: ["A1", "B1"]})
export const parseFunction = (expr: string): { name: string; args: string[] } | null => {
  const match = expr.match(/^([A-Z]+)\((.*)\)$/);
  if (!match) return null;
  
  const name = match[1];
  const argsStr = match[2];
  
  if (!argsStr.trim()) return { name, args: [] };
  
  // Simple argument parsing (doesn't handle nested functions yet)
  const args = argsStr.split(',').map(arg => arg.trim());
  
  return { name, args };
};

// Main formula evaluation function
export const evaluateFormula = (formula: string, context: FormulaContext): string => {
  try {
    // Remove leading =
    if (formula.startsWith('=')) {
      formula = formula.slice(1);
    }
    
    // Convert formula to uppercase for case-insensitive function names and cell references
    formula = normalizeFormula(formula);
    
    // Check if it's a function call
    const functionCall = parseFunction(formula);
    if (functionCall && FUNCTIONS[functionCall.name]) {
      const func = FUNCTIONS[functionCall.name];
      
      // Evaluate arguments
      const evaluatedArgs = functionCall.args.map(arg => {
        // If it's a cell reference or range, keep as string
        if (parseCellRef(arg) || arg.includes(':')) {
          return arg;
        }
        // Otherwise evaluate as expression
        return evaluateExpression(arg, context);
      });
      
      const result = func(evaluatedArgs, context);
      return String(result);
    }
    
    // Otherwise, evaluate as expression
    const result = evaluateExpression(formula, context);
    return String(result);
    
  } catch (error) {
    return '#ERROR!';
  }
};

// Function to normalize formula by converting function names and cell references to uppercase
export const normalizeFormula = (formula: string): string => {
  // Convert the entire formula to uppercase for function names and cell references
  // But preserve quoted strings in their original case
  let result = '';
  let inQuotes = false;
  let quoteChar = '';
  
  for (let i = 0; i < formula.length; i++) {
    const char = formula[i];
    
    if (!inQuotes && (char === '"' || char === "'")) {
      inQuotes = true;
      quoteChar = char;
      result += char;
    } else if (inQuotes && char === quoteChar) {
      inQuotes = false;
      quoteChar = '';
      result += char;
    } else if (inQuotes) {
      // Inside quotes, preserve original case
      result += char;
    } else {
      // Outside quotes, convert to uppercase
      result += char.toUpperCase();
    }
  }
  
  return result;
};
