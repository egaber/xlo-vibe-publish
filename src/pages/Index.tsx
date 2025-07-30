import { useState } from "react";
import ExcelTopBar from "@/components/excel-ribbon/ExcelTopBar";
import { ExcelRibbon } from "@/components/excel-ribbon/ExcelRibbon";
import { FormulaBar } from "@/components/excel-grid/FormulaBar";
import { ExcelGrid } from "@/components/excel-grid/ExcelGrid";
import { SheetTabBar, Sheet } from "@/components/excel-grid/SheetTabBar";
import { StatusBar } from "@/components/excel-grid/StatusBar";
import { evaluateFormula, FormulaContext } from "@/utils/formulaEngine";
import { CellData, ClipboardData, Selection, RibbonActions, CellFormat } from "@/types/cellTypes";
import { 
  getSelectionCellRefs, 
  applyCellFormat, 
  createClipboardData, 
  applyClipboardData,
  getCellRef 
} from "@/utils/cellFormatting";

interface SheetData {
  cellData: Record<string, CellData>;
  selectedCell: string;
  selectedCellValue: string;
}

interface FormulaReference {
  id: string;
  range: string;
  color: string;
}

const Index = () => {
  // Sheet management
  const [sheets, setSheets] = useState<Sheet[]>([
    { id: "sheet1", name: "Sheet1", isProtected: false, isVisible: true },
    { id: "sheet2", name: "Sheet2", isProtected: false, isVisible: true },
    { id: "sheet3", name: "Sheet3", isProtected: false, isVisible: true },
  ]);
  const [activeSheetId, setActiveSheetId] = useState("sheet1");
  
  // Sheet data storage
  const [sheetDataMap, setSheetDataMap] = useState<Record<string, SheetData>>({
    sheet1: { cellData: {}, selectedCell: "A1", selectedCellValue: "" },
    sheet2: { cellData: {}, selectedCell: "A1", selectedCellValue: "" },
    sheet3: { cellData: {}, selectedCell: "A1", selectedCellValue: "" },
  });

  // Formula building state
  const [isFormulaBuildingMode, setIsFormulaBuildingMode] = useState(false);
  const [formulaReferences, setFormulaReferences] = useState<FormulaReference[]>([]);
  const [rangeSelectionStart, setRangeSelectionStart] = useState<string | null>(null);
  
  // Clipboard state
  const [clipboardData, setClipboardData] = useState<ClipboardData | null>(null);
  
  // Selection state for ribbon operations
  const [currentSelection, setCurrentSelection] = useState<Selection>({
    start: { row: 0, col: 0 },
    end: { row: 0, col: 0 }
  });
  
  // Colors for formula references
  const referenceColors = [
    '#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', 
    '#DDA0DD', '#98D8C8', '#F7DC6F', '#BB8FCE', '#85C1E9'
  ];

  // Current sheet data
  const currentSheetData = sheetDataMap[activeSheetId] || { cellData: {}, selectedCell: "A1", selectedCellValue: "" };
  const selectedCell = currentSheetData.selectedCell;
  const selectedCellValue = currentSheetData.selectedCellValue;
  const cellData = currentSheetData.cellData;

  // Helper function to update current sheet data
  const updateCurrentSheetData = (updates: Partial<SheetData>) => {
    setSheetDataMap(prev => ({
      ...prev,
      [activeSheetId]: {
        ...prev[activeSheetId],
        ...updates
      }
    }));
  };

  const handleCellSelect = (cellRef: string, value: string) => {
    // Get the cell data to check if it has a formula
    const currentData = sheetDataMap[activeSheetId];
    const cellInfo = currentData.cellData[cellRef];
    
    // Show formula in formula bar if cell has formula, otherwise show value
    const formulaBarValue = cellInfo?.formula || value;
    
    updateCurrentSheetData({
      selectedCell: cellRef,
      selectedCellValue: formulaBarValue
    });
  };

  const handleCellUpdate = (cellRef: string, data: CellData) => {
    const currentData = sheetDataMap[activeSheetId];
    const updatedCellData = {
      ...currentData.cellData,
      [cellRef]: data
    };
    
    updateCurrentSheetData({
      cellData: updatedCellData,
      // Update the formula bar if this is the currently selected cell
      // Show formula if it exists, otherwise show the value
      ...(cellRef === currentData.selectedCell ? { 
        selectedCellValue: data.formula || data.value 
      } : {})
    });
  };

  const handleFormulaBarChange = (value: string) => {
    updateCurrentSheetData({
      selectedCellValue: value
    });
    
    // Check if we should enter or exit formula building mode
    const isFormula = value.startsWith('=');
    console.log('Formula bar change:', value, 'isFormula:', isFormula, 'current mode:', isFormulaBuildingMode);
    
    if (isFormula && !isFormulaBuildingMode) {
      console.log('Entering formula building mode');
      setIsFormulaBuildingMode(true);
      setFormulaReferences([]);
      setRangeSelectionStart(null);
    } else if (!isFormula && isFormulaBuildingMode) {
      console.log('Exiting formula building mode');
      setIsFormulaBuildingMode(false);
      setFormulaReferences([]);
      setRangeSelectionStart(null);
    }
  };

  // Formula building mode handlers
  const handleFormulaEditStart = () => {
    const value = selectedCellValue;
    const isFormula = value.startsWith('=');
    setIsFormulaBuildingMode(isFormula);
    setFormulaReferences([]);
    setRangeSelectionStart(null);
  };

  const handleFormulaEditEnd = () => {
    setIsFormulaBuildingMode(false);
    setFormulaReferences([]);
    setRangeSelectionStart(null);
  };

  const handleFormulaReferenceAdd = (cellRef: string, isRange: boolean = false) => {
    if (!isFormulaBuildingMode) return;

    const refId = `ref_${Date.now()}`;
    const color = referenceColors[formulaReferences.length % referenceColors.length];
    
    let range = cellRef;
    if (isRange && rangeSelectionStart && rangeSelectionStart !== cellRef) {
      range = `${rangeSelectionStart}:${cellRef}`;
    }

    const newRef: FormulaReference = {
      id: refId,
      range,
      color
    };

    setFormulaReferences(prev => [...prev, newRef]);
    
    // Add reference to formula
    const currentFormula = selectedCellValue;
    const newFormula = currentFormula + range;
    updateCurrentSheetData({
      selectedCellValue: newFormula
    });

    setRangeSelectionStart(null);
  };

  const handleCellClickInFormulaMode = (cellRef: string, isRangeSelection: boolean = false) => {
    if (!isFormulaBuildingMode) {
      console.log('Not in formula building mode, isFormulaBuildingMode:', isFormulaBuildingMode);
      return;
    }

    console.log('Cell clicked in formula mode:', cellRef, 'isRangeSelection:', isRangeSelection);

    if (isRangeSelection && !rangeSelectionStart) {
      setRangeSelectionStart(cellRef);
      console.log('Set range selection start:', cellRef);
    } else {
      handleFormulaReferenceAdd(cellRef, Boolean(rangeSelectionStart));
      console.log('Added formula reference:', cellRef);
    }
  };

  const handleApplyFormula = (value?: string) => {
    const valueToUse = value || selectedCellValue;
    const isFormula = valueToUse.startsWith('=');
    
    let displayValue = valueToUse;
    
    if (isFormula) {
      // Create formula context for evaluation
      const currentData = sheetDataMap[activeSheetId];
      const formulaContext: FormulaContext = {
        getCellValue: (ref: string) => {
          const cellInfo = currentData.cellData[ref];
          return cellInfo ? cellInfo.value : '';
        },
        setCellValue: (ref: string, value: string) => {
          // This would be used for more complex formulas that modify other cells
          // For now, we'll just implement the getter
        }
      };
      
      // Evaluate the formula
      displayValue = evaluateFormula(valueToUse, formulaContext);
    }
    
    handleCellUpdate(selectedCell, {
      value: displayValue,
      formula: isFormula ? valueToUse : undefined
    });
  };

  // Sheet management functions
  const handleSheetSelect = (sheetId: string) => {
    setActiveSheetId(sheetId);
  };

  const handleSheetRename = (sheetId: string, newName: string) => {
    setSheets(prev => prev.map(sheet => 
      sheet.id === sheetId ? { ...sheet, name: newName } : sheet
    ));
  };

  const handleSheetAdd = () => {
    const newSheetNumber = sheets.length + 1;
    const newSheetId = `sheet${Date.now()}`;
    const newSheet: Sheet = {
      id: newSheetId,
      name: `Sheet${newSheetNumber}`,
      isProtected: false,
      isVisible: true
    };

    setSheets(prev => [...prev, newSheet]);
    setSheetDataMap(prev => ({
      ...prev,
      [newSheetId]: { cellData: {}, selectedCell: "A1", selectedCellValue: "" }
    }));
    setActiveSheetId(newSheetId);
  };

  const handleSheetDelete = (sheetId: string) => {
    if (sheets.length <= 1) return; // Don't delete the last sheet
    
    const sheetIndex = sheets.findIndex(s => s.id === sheetId);
    setSheets(prev => prev.filter(sheet => sheet.id !== sheetId));
    
    // Remove sheet data
    setSheetDataMap(prev => {
      const { [sheetId]: deleted, ...rest } = prev;
      return rest;
    });

    // Switch to adjacent sheet if deleting active sheet
    if (sheetId === activeSheetId) {
      const remainingSheets = sheets.filter(s => s.id !== sheetId);
      const newActiveIndex = Math.min(sheetIndex, remainingSheets.length - 1);
      setActiveSheetId(remainingSheets[newActiveIndex]?.id || remainingSheets[0]?.id);
    }
  };

  const handleSheetDuplicate = (sheetId: string) => {
    const sourceSheet = sheets.find(s => s.id === sheetId);
    if (!sourceSheet) return;

    const newSheetId = `sheet${Date.now()}`;
    const newSheet: Sheet = {
      id: newSheetId,
      name: `${sourceSheet.name} (2)`,
      isProtected: sourceSheet.isProtected,
      isVisible: sourceSheet.isVisible
    };

    setSheets(prev => [...prev, newSheet]);
    
    // Copy sheet data
    const sourceData = sheetDataMap[sheetId];
    setSheetDataMap(prev => ({
      ...prev,
      [newSheetId]: {
        cellData: { ...sourceData?.cellData || {} },
        selectedCell: "A1",
        selectedCellValue: ""
      }
    }));
  };

  const handleSheetMove = (sheetId: string, direction: 'left' | 'right') => {
    const currentIndex = sheets.findIndex(s => s.id === sheetId);
    if (currentIndex === -1) return;

    const newIndex = direction === 'left' ? currentIndex - 1 : currentIndex + 1;
    if (newIndex < 0 || newIndex >= sheets.length) return;

    const newSheets = [...sheets];
    [newSheets[currentIndex], newSheets[newIndex]] = [newSheets[newIndex], newSheets[currentIndex]];
    setSheets(newSheets);
  };

  const handleSheetToggleProtection = (sheetId: string) => {
    setSheets(prev => prev.map(sheet => 
      sheet.id === sheetId ? { ...sheet, isProtected: !sheet.isProtected } : sheet
    ));
  };

  const handleSheetToggleVisibility = (sheetId: string) => {
    setSheets(prev => prev.map(sheet => 
      sheet.id === sheetId ? { ...sheet, isVisible: !sheet.isVisible } : sheet
    ));
  };

  // Ribbon Actions Implementation
  const ribbonActions: RibbonActions = {
    // Clipboard operations
    copy: () => {
      const data = createClipboardData(cellData, currentSelection, 'copy');
      setClipboardData(data);
    },

    paste: () => {
      if (!clipboardData) return;
      const newCellData = applyClipboardData(cellData, clipboardData, currentSelection);
      updateCurrentSheetData({ cellData: newCellData });
    },

    cut: () => {
      const data = createClipboardData(cellData, currentSelection, 'cut');
      setClipboardData(data);
    },

    // Font formatting
    toggleBold: () => {
      const cellRefs = getSelectionCellRefs(currentSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        const isBold = currentCell?.format?.bold;
        newCellData[ref] = applyCellFormat(currentCell, { bold: !isBold });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    toggleItalic: () => {
      const cellRefs = getSelectionCellRefs(currentSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        const isItalic = currentCell?.format?.italic;
        newCellData[ref] = applyCellFormat(currentCell, { italic: !isItalic });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    toggleUnderline: () => {
      const cellRefs = getSelectionCellRefs(currentSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        const isUnderline = currentCell?.format?.underline;
        newCellData[ref] = applyCellFormat(currentCell, { underline: !isUnderline });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    setFontFamily: (family: string) => {
      const cellRefs = getSelectionCellRefs(currentSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        newCellData[ref] = applyCellFormat(currentCell, { fontFamily: family });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    setFontSize: (size: number) => {
      const cellRefs = getSelectionCellRefs(currentSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        newCellData[ref] = applyCellFormat(currentCell, { fontSize: size });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    increaseFontSize: () => {
      const cellRefs = getSelectionCellRefs(currentSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        const currentSize = currentCell?.format?.fontSize || 11;
        newCellData[ref] = applyCellFormat(currentCell, { fontSize: currentSize + 1 });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    decreaseFontSize: () => {
      const cellRefs = getSelectionCellRefs(currentSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        const currentSize = currentCell?.format?.fontSize || 11;
        newCellData[ref] = applyCellFormat(currentCell, { fontSize: Math.max(8, currentSize - 1) });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    // Colors
    setTextColor: (color: string) => {
      const cellRefs = getSelectionCellRefs(currentSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        newCellData[ref] = applyCellFormat(currentCell, { textColor: color });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    setBackgroundColor: (color: string) => {
      const cellRefs = getSelectionCellRefs(currentSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        newCellData[ref] = applyCellFormat(currentCell, { backgroundColor: color });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    // Alignment
    setHorizontalAlignment: (align: 'left' | 'center' | 'right') => {
      const cellRefs = getSelectionCellRefs(currentSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        newCellData[ref] = applyCellFormat(currentCell, { horizontalAlign: align });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    setVerticalAlignment: (align: 'top' | 'middle' | 'bottom') => {
      const cellRefs = getSelectionCellRefs(currentSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        newCellData[ref] = applyCellFormat(currentCell, { verticalAlign: align });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    // Number formatting
    setNumberFormat: (format: CellFormat['numberFormat']) => {
      const cellRefs = getSelectionCellRefs(currentSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        newCellData[ref] = applyCellFormat(currentCell, { numberFormat: format });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    increaseDecimals: () => {
      const cellRefs = getSelectionCellRefs(currentSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        const currentFormat = currentCell?.format?.numberFormat;
        const currentDecimals = currentFormat?.decimals || 0;
        newCellData[ref] = applyCellFormat(currentCell, { 
          numberFormat: { 
            ...currentFormat,
            type: currentFormat?.type || 'number',
            decimals: currentDecimals + 1 
          }
        });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    decreaseDecimals: () => {
      const cellRefs = getSelectionCellRefs(currentSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        const currentFormat = currentCell?.format?.numberFormat;
        const currentDecimals = currentFormat?.decimals || 0;
        newCellData[ref] = applyCellFormat(currentCell, { 
          numberFormat: { 
            ...currentFormat,
            type: currentFormat?.type || 'number',
            decimals: Math.max(0, currentDecimals - 1) 
          }
        });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    // Undo/Redo (basic implementation for now)
    undo: () => {
      // TODO: Implement proper undo/redo with history stack
      console.log('Undo functionality needs implementation');
    },

    redo: () => {
      // TODO: Implement proper undo/redo with history stack
      console.log('Redo functionality needs implementation');
    }
  };

  // Handle selection updates from ExcelGrid
  const handleSelectionChange = (selection: Selection) => {
    setCurrentSelection(selection);
  };

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col">
      {/* Fixed header area - stays at top */}
      <div className="flex-shrink-0">
        <ExcelTopBar />
        <ExcelRibbon ribbonActions={ribbonActions} />
        <FormulaBar
          selectedCell={selectedCell}
          cellValue={selectedCellValue}
          onCellValueChange={handleFormulaBarChange}
          onApplyFormula={handleApplyFormula}
          onFormulaEditStart={handleFormulaEditStart}
          onFormulaEditEnd={handleFormulaEditEnd}
        />
      </div>
      
      {/* Scrollable content area */}
      <div className="flex-1 overflow-hidden">
        <ExcelGrid
          onCellSelect={handleCellSelect}
          cellData={cellData}
          onCellUpdate={handleCellUpdate}
          isFormulaBuildingMode={isFormulaBuildingMode}
          formulaReferences={formulaReferences}
          rangeSelectionStart={rangeSelectionStart}
          onCellClickInFormulaMode={handleCellClickInFormulaMode}
          onSelectionChange={handleSelectionChange}
        />
      </div>
      
      {/* Fixed footer - Sheet tabs */}
      <div className="flex-shrink-0">
        <SheetTabBar
          sheets={sheets}
          activeSheetId={activeSheetId}
          onSheetSelect={handleSheetSelect}
          onSheetRename={handleSheetRename}
          onSheetAdd={handleSheetAdd}
          onSheetDelete={handleSheetDelete}
          onSheetDuplicate={handleSheetDuplicate}
          onSheetMove={handleSheetMove}
          onSheetToggleProtection={handleSheetToggleProtection}
          onSheetToggleVisibility={handleSheetToggleVisibility}
        />
      </div>
      
      {/* Fixed status bar - always at bottom */}
      <div className="flex-shrink-0">
        <StatusBar />
      </div>
    </div>
  );
};

export default Index;
