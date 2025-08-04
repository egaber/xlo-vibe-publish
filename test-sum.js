// Quick test to verify SUM function works with literal numbers
import { evaluateFormula } from './src/utils/formulaEngine.js';

const mockContext = {
  getCellValue: (ref) => '',
  setCellValue: (ref, value) => {}
};

// Test SUM with literal numbers
const result = evaluateFormula('=SUM(1,2)', mockContext);
console.log('=SUM(1,2) result:', result);

// Test SUM with more numbers
const result2 = evaluateFormula('=SUM(1,2,3,4)', mockContext);
console.log('=SUM(1,2,3,4) result:', result2);

// Test SUM with decimal numbers
const result3 = evaluateFormula('=SUM(1.5,2.5)', mockContext);
console.log('=SUM(1.5,2.5) result:', result3);
