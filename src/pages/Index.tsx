import { useState, useRef, useEffect, useMemo, useCallback } from "react";
import ExcelTopBar from "@/components/excel-ribbon/ExcelTopBar";
import { ExcelRibbon } from "@/components/excel-ribbon/ExcelRibbon";
import { ImportedExcelData } from "@/utils/excelImport";
import { FormulaBar } from "@/components/excel-grid/FormulaBar";
import { ExcelGrid } from "@/components/excel-grid/ExcelGrid";
import { ColumnHeaders } from "@/components/excel-grid/ColumnHeaders";
import { SheetTabBar, Sheet } from "@/components/excel-grid/SheetTabBar";
import { StatusBar } from "@/components/excel-grid/StatusBar";
import { evaluateFormula, FormulaContext } from "@/utils/formulaEngine";
import { CellData, ClipboardData, Selection, MultiSelection, RibbonActions, CellFormat } from "@/types/cellTypes";
import { 
  getSelectionCellRefs,
  getMultiSelectionCellRefs,
  applyCellFormat, 
  createClipboardData, 
  applyClipboardData,
  getCellRef,
  DEFAULT_CELL_FORMAT 
} from "@/utils/cellFormatting";
import { convertSvgImages } from "@/utils/svgIconUtils";
import { getBranchFromUrl, retrieveBranchState, cleanupOldBranches } from "@/utils/branchState";
import { useToast } from "@/hooks/use-toast";

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

interface HistoryEntry {
  cellData: Record<string, CellData>;
  selectedCell: string;
  selectedCellValue: string;
  timestamp: number;
}

const Index = () => {
  const { toast } = useToast();
  
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
  
  // Refs for scroll synchronization
  const columnHeadersRef = useRef<HTMLDivElement>(null);
  const gridContainerRef = useRef<HTMLDivElement>(null);
  
  // Clipboard state
  const [clipboardData, setClipboardData] = useState<ClipboardData | null>(null);
  
  // Selection state for ribbon operations
  const [currentSelection, setCurrentSelection] = useState<Selection>({
    start: { row: 0, col: 0 },
    end: { row: 0, col: 0 }
  });
  
  // Multi-selection state
  const [currentMultiSelection, setCurrentMultiSelection] = useState<MultiSelection>({
    primary: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
    additional: []
  });
  
  // Column width state
  const [columnWidths, setColumnWidths] = useState<number[]>(Array(39).fill(80));
  
  // Undo/Redo history state
  const [history, setHistory] = useState<Record<string, HistoryEntry[]>>({
    sheet1: [],
    sheet2: [],
    sheet3: [],
  });
  const [historyIndex, setHistoryIndex] = useState<Record<string, number>>({
    sheet1: -1,
    sheet2: -1,
    sheet3: -1,
  });
  const maxHistorySize = 50; // Limit history to prevent memory issues
  
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

  // Convert SVG images to inline SVGs for colorization support
  useEffect(() => {
    // Convert SVGs after initial render and when ribbon updates
    const timer = setTimeout(() => {
      convertSvgImages();
    }, 100); // Small delay to ensure DOM is ready

    return () => clearTimeout(timer);
  }, []);

  // Helper function to update current sheet data with history tracking
  const updateCurrentSheetData = useCallback((updates: Partial<SheetData>, addToHistory: boolean = true) => {
    // Save current state to history before updating
    if (addToHistory) {
      const currentData = sheetDataMap[activeSheetId];
      if (currentData) {
        const historyEntry: HistoryEntry = {
          cellData: { ...currentData.cellData },
          selectedCell: currentData.selectedCell,
          selectedCellValue: currentData.selectedCellValue,
          timestamp: Date.now()
        };
        
        setHistory(prev => {
          const sheetHistory = prev[activeSheetId] || [];
          const currentIndex = historyIndex[activeSheetId] ?? -1;
          
          // Remove any history after current index (when making a new change after undo)
          const newHistory = sheetHistory.slice(0, currentIndex + 1);
          newHistory.push(historyEntry);
          
          // Limit history size
          if (newHistory.length > maxHistorySize) {
            newHistory.shift();
          }
          
          return {
            ...prev,
            [activeSheetId]: newHistory
          };
        });
        
        setHistoryIndex(prev => ({
          ...prev,
          [activeSheetId]: Math.min((prev[activeSheetId] ?? -1) + 1, maxHistorySize - 1)
        }));
      }
    }
    
    setSheetDataMap(prev => ({
      ...prev,
      [activeSheetId]: {
        ...prev[activeSheetId],
        ...updates
      }
    }));
  }, [activeSheetId, sheetDataMap, setSheetDataMap, historyIndex, maxHistorySize]);

  // Scroll synchronization handler
  const handleGridScroll = (scrollLeft: number) => {
    if (columnHeadersRef.current) {
      columnHeadersRef.current.scrollLeft = scrollLeft;
    }
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
    
    // Ensure cell has default formatting if no format is specified
    const cellDataWithDefaults = {
      ...data,
      format: data.format ? { ...DEFAULT_CELL_FORMAT, ...data.format } : DEFAULT_CELL_FORMAT
    };
    
    const updatedCellData = {
      ...currentData.cellData,
      [cellRef]: cellDataWithDefaults
    };
    
    updateCurrentSheetData({
      cellData: updatedCellData,
      // Update the formula bar if this is the currently selected cell
      // Show formula if it exists, otherwise show the value
      ...(cellRef === currentData.selectedCell ? { 
        selectedCellValue: cellDataWithDefaults.formula || cellDataWithDefaults.value 
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
      formula: isFormula ? valueToUse : undefined,
      format: cellData[selectedCell]?.format || DEFAULT_CELL_FORMAT
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
  const ribbonActions: RibbonActions = useMemo(() => ({
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
      const cellRefs = getMultiSelectionCellRefs(currentMultiSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        const isBold = currentCell?.format?.bold;
        newCellData[ref] = applyCellFormat(currentCell, { bold: !isBold });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    toggleItalic: () => {
      const cellRefs = getMultiSelectionCellRefs(currentMultiSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        const isItalic = currentCell?.format?.italic;
        newCellData[ref] = applyCellFormat(currentCell, { italic: !isItalic });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    toggleUnderline: () => {
      const cellRefs = getMultiSelectionCellRefs(currentMultiSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        const isUnderline = currentCell?.format?.underline;
        newCellData[ref] = applyCellFormat(currentCell, { underline: !isUnderline });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    setFontFamily: (family: string) => {
      const cellRefs = getMultiSelectionCellRefs(currentMultiSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        newCellData[ref] = applyCellFormat(currentCell, { fontFamily: family });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    setFontSize: (size: number) => {
      const cellRefs = getMultiSelectionCellRefs(currentMultiSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        newCellData[ref] = applyCellFormat(currentCell, { fontSize: size });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    increaseFontSize: () => {
      const cellRefs = getMultiSelectionCellRefs(currentMultiSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        const currentSize = currentCell?.format?.fontSize || 14;
        newCellData[ref] = applyCellFormat(currentCell, { fontSize: currentSize + 1 });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    decreaseFontSize: () => {
      const cellRefs = getMultiSelectionCellRefs(currentMultiSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        const currentSize = currentCell?.format?.fontSize || 14;
        newCellData[ref] = applyCellFormat(currentCell, { fontSize: Math.max(8, currentSize - 1) });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    // Colors
    setTextColor: (color: string) => {
      const cellRefs = getMultiSelectionCellRefs(currentMultiSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        newCellData[ref] = applyCellFormat(currentCell, { textColor: color });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    setBackgroundColor: (color: string) => {
      const cellRefs = getMultiSelectionCellRefs(currentMultiSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        newCellData[ref] = applyCellFormat(currentCell, { backgroundColor: color });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    // Alignment
    setHorizontalAlignment: (align: 'left' | 'center' | 'right') => {
      const cellRefs = getMultiSelectionCellRefs(currentMultiSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        newCellData[ref] = applyCellFormat(currentCell, { horizontalAlign: align });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    setVerticalAlignment: (align: 'top' | 'middle' | 'bottom') => {
      const cellRefs = getMultiSelectionCellRefs(currentMultiSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        newCellData[ref] = applyCellFormat(currentCell, { verticalAlign: align });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    // Number formatting
    setNumberFormat: (format: CellFormat['numberFormat']) => {
      const cellRefs = getMultiSelectionCellRefs(currentMultiSelection);
      const newCellData = { ...cellData };
      
      cellRefs.forEach(ref => {
        const currentCell = newCellData[ref];
        newCellData[ref] = applyCellFormat(currentCell, { numberFormat: format });
      });
      
      updateCurrentSheetData({ cellData: newCellData });
    },

    increaseDecimals: () => {
      const cellRefs = getMultiSelectionCellRefs(currentMultiSelection);
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
      const cellRefs = getMultiSelectionCellRefs(currentMultiSelection);
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

    // Undo/Redo functionality
    undo: () => {
      const sheetHistory = history[activeSheetId] || [];
      const currentIndex = historyIndex[activeSheetId] ?? -1;
      
      if (currentIndex >= 0 && sheetHistory[currentIndex]) {
        // Restore from history
        const historyEntry = sheetHistory[currentIndex];
        updateCurrentSheetData({
          cellData: { ...historyEntry.cellData },
          selectedCell: historyEntry.selectedCell,
          selectedCellValue: historyEntry.selectedCellValue
        }, false); // Don't add to history when undoing
        
        // Move history index back
        setHistoryIndex(prev => ({
          ...prev,
          [activeSheetId]: Math.max(-1, currentIndex - 1)
        }));
      }
    },

    redo: () => {
      const sheetHistory = history[activeSheetId] || [];
      const currentIndex = historyIndex[activeSheetId] ?? -1;
      const nextIndex = currentIndex + 2; // +2 because we need to go forward from current
      
      if (nextIndex < sheetHistory.length && sheetHistory[nextIndex]) {
        // Restore from history
        const historyEntry = sheetHistory[nextIndex];
        updateCurrentSheetData({
          cellData: { ...historyEntry.cellData },
          selectedCell: historyEntry.selectedCell,
          selectedCellValue: historyEntry.selectedCellValue
        }, false); // Don't add to history when redoing
        
        // Move history index forward
        setHistoryIndex(prev => ({
          ...prev,
          [activeSheetId]: nextIndex
        }));
      }
    }
  }), [
    cellData,
    currentSelection,
    currentMultiSelection,
    clipboardData,
    activeSheetId,
    history,
    historyIndex,
    updateCurrentSheetData
  ]);

  // Keyboard shortcuts for undo/redo
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Check for Ctrl+Z (undo)
      if (e.ctrlKey && e.key === 'z' && !e.shiftKey) {
        e.preventDefault();
        ribbonActions.undo();
      }
      // Check for Ctrl+Y or Ctrl+Shift+Z (redo)
      else if ((e.ctrlKey && e.key === 'y') || (e.ctrlKey && e.shiftKey && e.key === 'z')) {
        e.preventDefault();
        ribbonActions.redo();
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [ribbonActions]); // Dependencies for undo/redo

  // Load shared branch state if URL contains branch parameters
  useEffect(() => {
    const branchInfo = getBranchFromUrl();
    if (branchInfo) {
      const { branchId, code } = branchInfo;
      
      try {
        const sharedState = retrieveBranchState(branchId);
        
        if (sharedState) {
          // Load the shared state
          setSheets(sharedState.sheets);
          setSheetDataMap(sharedState.sheetDataMap);
          setActiveSheetId(sharedState.activeSheetId);
          setColumnWidths(sharedState.columnWidths);
          
          toast({
            title: "Shared State Loaded",
            description: `Successfully loaded shared Excel state (Code: ${code})`,
            duration: 5000,
          });
          
          // Clean up URL parameters to avoid reloading on refresh
          const url = new URL(window.location.href);
          url.searchParams.delete('branch');
          url.searchParams.delete('code');
          window.history.replaceState({}, '', url.toString());
        } else {
          toast({
            title: "Shared State Not Found",
            description: "The shared Excel state could not be found or has expired.",
            variant: "destructive",
            duration: 5000,
          });
        }
      } catch (error) {
        console.error('Failed to load shared state:', error);
        toast({
          title: "Load Failed",
          description: "Failed to load the shared Excel state.",
          variant: "destructive",
          duration: 5000,
        });
      }
    }
    
    // Clean up old branches periodically
    cleanupOldBranches();
  }, [toast]); // Run once on mount

  // Handle selection updates from ExcelGrid
  const handleSelectionChange = (selection: Selection) => {
    setCurrentSelection(selection);
  };
  
  // Handle multi-selection updates from ExcelGrid
  const handleMultiSelectionChange = (multiSelection: MultiSelection) => {
    setCurrentMultiSelection(multiSelection);
  };

  // Handle column selection
  const handleColumnSelect = (colIndex: number) => {
    const GRID_ROWS = 100; // Should match the constant in ExcelGrid
    const newSelection: Selection = {
      start: { row: 0, col: colIndex },
      end: { row: GRID_ROWS - 1, col: colIndex }
    };
    setCurrentSelection(newSelection);
  };

  // Handle select all
  const handleSelectAll = () => {
    const GRID_ROWS = 100; // Should match the constant in ExcelGrid
    const GRID_COLS = 39; // Should match the constant in ExcelGrid
    const newSelection: Selection = {
      start: { row: 0, col: 0 },
      end: { row: GRID_ROWS - 1, col: GRID_COLS - 1 }
    };
    setCurrentSelection(newSelection);
  };

  const handleColumnWidthChange = (columnIndex: number, width: number) => {
    setColumnWidths(prev => {
      const newWidths = [...prev];
      newWidths[columnIndex] = width;
      return newWidths;
    });
  };

  // Handle Excel file import
  const handleFileOpen = (data: ImportedExcelData) => {
    console.log('Imported Excel data:', data);
    
    // Clear existing sheets
    const newSheets: Sheet[] = [];
    const newSheetDataMap: Record<string, SheetData> = {};
    
    // Create sheets from imported data
    data.sheets.forEach((importedSheet, index) => {
      const sheetId = `sheet${Date.now()}_${index}`;
      
      newSheets.push({
        id: sheetId,
        name: importedSheet.name,
        isProtected: false,
        isVisible: true
      });
      
      newSheetDataMap[sheetId] = {
        cellData: importedSheet.data,
        selectedCell: "A1",
        selectedCellValue: ""
      };
    });
    
    // Update state with imported data
    setSheets(newSheets);
    setSheetDataMap(newSheetDataMap);
    setActiveSheetId(newSheets[0]?.id || "sheet1");
    
    // Reset history for new sheets
    const newHistory: Record<string, HistoryEntry[]> = {};
    const newHistoryIndex: Record<string, number> = {};
    newSheets.forEach(sheet => {
      newHistory[sheet.id] = [];
      newHistoryIndex[sheet.id] = -1;
    });
    setHistory(newHistory);
    setHistoryIndex(newHistoryIndex);
  };

  // Handle clearing all content
  const handleClearContent = () => {
    // Reset to default sheets
    const defaultSheets: Sheet[] = [
      { id: "sheet1", name: "Sheet1", isProtected: false, isVisible: true },
      { id: "sheet2", name: "Sheet2", isProtected: false, isVisible: true },
      { id: "sheet3", name: "Sheet3", isProtected: false, isVisible: true },
    ];
    
    const defaultSheetDataMap: Record<string, SheetData> = {
      sheet1: { cellData: {}, selectedCell: "A1", selectedCellValue: "" },
      sheet2: { cellData: {}, selectedCell: "A1", selectedCellValue: "" },
      sheet3: { cellData: {}, selectedCell: "A1", selectedCellValue: "" },
    };
    
    // Reset all state
    setSheets(defaultSheets);
    setSheetDataMap(defaultSheetDataMap);
    setActiveSheetId("sheet1");
    
    // Reset history
    setHistory({
      sheet1: [],
      sheet2: [],
      sheet3: [],
    });
    setHistoryIndex({
      sheet1: -1,
      sheet2: -1,
      sheet3: -1,
    });
    
    // Reset column widths
    setColumnWidths(Array(39).fill(80));
    
    // Reset clipboard
    setClipboardData(null);
    
    // Reset formula building mode
    setIsFormulaBuildingMode(false);
    setFormulaReferences([]);
    setRangeSelectionStart(null);
    
    // Reset selections
    setCurrentSelection({
      start: { row: 0, col: 0 },
      end: { row: 0, col: 0 }
    });
    setCurrentMultiSelection({
      primary: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
      additional: []
    });
  };

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col">
      {/* Fixed header area - stays at top */}
      <div className="fixed top-0 left-0 right-0 z-50 bg-gray-50 flex-shrink-0">
        <ExcelTopBar />
        <ExcelRibbon 
          ribbonActions={ribbonActions} 
          onFileOpen={handleFileOpen} 
          onClearContent={handleClearContent}
          sheets={sheets}
          sheetDataMap={sheetDataMap}
          activeSheetId={activeSheetId}
          columnWidths={columnWidths}
        />
        <FormulaBar
          selectedCell={selectedCell}
          cellValue={selectedCellValue}
          onCellValueChange={handleFormulaBarChange}
          onApplyFormula={handleApplyFormula}
          onFormulaEditStart={handleFormulaEditStart}
          onFormulaEditEnd={handleFormulaEditEnd}
        />
        <ColumnHeaders 
          ref={columnHeadersRef}
          selection={currentSelection}
          multiSelection={currentMultiSelection}
          onColumnSelect={handleColumnSelect}
          onSelectAll={handleSelectAll}
          columnWidths={columnWidths}
          onColumnWidthChange={handleColumnWidthChange}
        />
      </div>
      
      {/* Scrollable content area */}
      <div ref={gridContainerRef} className="fixed top-[291px] left-0 right-0 bottom-0 overflow-hidden"> {/* Push div */}
        <ExcelGrid
          onCellSelect={handleCellSelect}
          cellData={cellData}
          onCellUpdate={handleCellUpdate}
          isFormulaBuildingMode={isFormulaBuildingMode}
          formulaReferences={formulaReferences}
          rangeSelectionStart={rangeSelectionStart}
          onCellClickInFormulaMode={handleCellClickInFormulaMode}
          onSelectionChange={handleSelectionChange}
          onMultiSelectionChange={handleMultiSelectionChange}
          onColumnSelect={handleColumnSelect}
          onSelectAll={handleSelectAll}
          externalSelection={currentSelection}
          onScroll={handleGridScroll}
          columnWidths={columnWidths}
          onColumnWidthChange={handleColumnWidthChange}
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
