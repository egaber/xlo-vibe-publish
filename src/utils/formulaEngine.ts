// Formula engine for Excel-like calculations

// Extend Window interface for GPT requests tracking
declare global {
  interface Window {
    gptRequests?: Map<string, {
      prompt: string;
      status: 'pending' | 'completed' | 'error';
      result?: string;
      timestamp: number;
    }>;
  }
}

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
          const cellValue = context.getCellValue(cellRef);
          const value = parseFloat(cellValue);
          if (cellValue && cellValue.trim() !== '' && isNaN(value)) {
            return '#ERROR!';
          }
          total += value || 0;
        }
      } else if (typeof arg === 'string' && parseCellRef(arg)) {
        // Single cell reference
        const cellValue = context.getCellValue(arg);
        const value = parseFloat(cellValue);
        if (cellValue && cellValue.trim() !== '' && isNaN(value)) {
          return '#ERROR!';
        }
        total += value || 0;
      } else {
        // Literal value
        const value = parseFloat(String(arg));
        if (isNaN(value)) {
          return '#ERROR!';
        }
        total += value;
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
  },

  GPT: (args: FunctionArg[], context: FormulaContext) => {
    if (args.length === 0) return '#ERROR!';
    
    // Get the prompt from the first argument
    let prompt = String(args[0]);
    
    // If it's a cell reference, get the cell value
    if (typeof args[0] === 'string' && parseCellRef(args[0])) {
      prompt = context.getCellValue(args[0]);
    }
    
    // Remove quotes if present
    if (prompt.startsWith('"') && prompt.endsWith('"')) {
      prompt = prompt.slice(1, -1);
    }
    
    if (!prompt || prompt.trim() === '') {
      return '#ERROR!';
    }
    
    // Create a unique identifier for this request
    const requestId = `gpt_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    
    // Store the request for tracking
    if (!window.gptRequests) {
      window.gptRequests = new Map();
    }
    
    // Check if we already have a pending request for this exact prompt
    const existingRequest = Array.from(window.gptRequests.entries()).find(
      ([id, data]) => data.prompt === prompt && data.status === 'pending'
    );
    
    if (existingRequest) {
      return `Loading... (${existingRequest[0]})`;
    }
    
    // Store this request
    window.gptRequests.set(requestId, {
      prompt,
      status: 'pending',
      timestamp: Date.now()
    });
    
    // Make async call to VS LM API
    setTimeout(async () => {
      try {
        console.log('Calling VS LM API with prompt:', prompt);
        
        const response = await fetch('http://localhost:3000/api/chat', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            messages: [
              {
                role: 'user',
                content: prompt
              }
            ]
          })
        });
        
        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const data = await response.json();
        const result = data.response || data.choices?.[0]?.message?.content || data.content || 'No response';
        
        // Update the request status
        if (window.gptRequests.has(requestId)) {
          window.gptRequests.set(requestId, {
            prompt,
            status: 'completed',
            result,
            timestamp: Date.now()
          });
        }
        
        console.log('GPT Response received:', result);
        
        // Dispatch event with the response
        window.dispatchEvent(new CustomEvent('gpt-response', { 
          detail: { requestId, result, prompt } 
        }));
        
      } catch (error) {
        console.error('Error calling VS LM API:', error);
        
        // Update the request status
        if (window.gptRequests.has(requestId)) {
          window.gptRequests.set(requestId, {
            prompt,
            status: 'error',
            result: '#API_ERROR!',
            timestamp: Date.now()
          });
        }
        
        window.dispatchEvent(new CustomEvent('gpt-response', { 
          detail: { requestId, result: '#API_ERROR!', prompt } 
        }));
      }
    }, 100); // Small delay to ensure UI updates
    
    return `Loading... (${requestId})`;
  }
};

// Enhanced expression parser for arithmetic with proper operator precedence
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
  
  // Handle parentheses first
  if (expr.includes('(') && expr.includes(')')) {
    return evaluateExpressionWithParentheses(expr, context);
  }
  
  // Handle arithmetic operations with proper precedence
  return evaluateArithmeticExpression(expr, context);
};

// Helper function to handle expressions with parentheses
const evaluateExpressionWithParentheses = (expr: string, context: FormulaContext): string | number => {
  // Find innermost parentheses
  let depth = 0;
  let start = -1;
  let end = -1;
  
  for (let i = 0; i < expr.length; i++) {
    if (expr[i] === '(') {
      if (depth === 0) {
        start = i;
      }
      depth++;
    } else if (expr[i] === ')') {
      depth--;
      if (depth === 0) {
        end = i;
        break;
      }
    }
  }
  
  if (start !== -1 && end !== -1) {
    const innerExpr = expr.substring(start + 1, end);
    const innerResult = evaluateExpression(innerExpr, context);
    const newExpr = expr.substring(0, start) + String(innerResult) + expr.substring(end + 1);
    return evaluateExpression(newExpr, context);
  }
  
  return evaluateArithmeticExpression(expr, context);
};

// Helper function to handle arithmetic with proper operator precedence
const evaluateArithmeticExpression = (expr: string, context: FormulaContext): string | number => {
  // First handle multiplication and division (left to right)
  expr = handleOperators(expr, ['*', '/'], context);
  
  // Then handle addition and subtraction (left to right)
  expr = handleOperators(expr, ['+', '-'], context);
  
  // If we end up with a single number, return it
  if (!isNaN(parseFloat(expr))) {
    return parseFloat(expr);
  }
  
  return expr;
};

// Helper function to handle specific operators
const handleOperators = (expr: string, operators: string[], context: FormulaContext): string => {
  for (const op of operators) {
    let operatorIndex = -1;
    
    // Find the operator (from left to right)
    for (let i = 0; i < expr.length; i++) {
      if (expr[i] === op) {
        // Make sure it's not a negative sign at the beginning
        if (op === '-' && i === 0) continue;
        // Make sure it's not a negative sign after another operator
        if (op === '-' && i > 0 && ['+', '-', '*', '/'].includes(expr[i-1])) continue;
        
        operatorIndex = i;
        break;
      }
    }
    
    if (operatorIndex !== -1) {
      const left = expr.substring(0, operatorIndex).trim();
      const right = expr.substring(operatorIndex + 1).trim();
      
      const leftResult = evaluateExpression(left, context);
      const rightResult = evaluateExpression(right, context);
      
      const leftNum = parseFloat(String(leftResult));
      const rightNum = parseFloat(String(rightResult));
      
      if (!isNaN(leftNum) && !isNaN(rightNum)) {
        let result: number;
        switch (op) {
          case '+': result = leftNum + rightNum; break;
          case '-': result = leftNum - rightNum; break;
          case '*': result = leftNum * rightNum; break;
          case '/': 
            if (rightNum === 0) return '#DIV/0!';
            result = leftNum / rightNum; 
            break;
          default: return expr;
        }
        
        // Replace the expression with the result and continue processing
        const newExpr = String(result) + expr.substring(operatorIndex + right.length + 1);
        return handleOperators(newExpr, operators, context);
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
