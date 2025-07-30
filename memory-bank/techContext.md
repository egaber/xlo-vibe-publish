# Technical Context

## Project Structure
```
src/
├── components/
│   ├── excel-grid/          # Grid components
│   │   ├── ExcelGrid.tsx    # Main grid with cell management
│   │   ├── FormulaBar.tsx   # Formula input bar
│   │   ├── SheetTabBar.tsx  # Sheet tabs
│   │   └── StatusBar.tsx    # Status bar
│   ├── excel-ribbon/        # Ribbon UI components
│   │   ├── ExcelRibbon.tsx  # Main ribbon container
│   │   ├── RibbonGroup.tsx  # Group containers
│   │   ├── ColorPicker.tsx  # Color selection
│   │   └── [Various]Dropdown.tsx # Dropdown components
│   └── ui/                  # shadcn/ui components
├── utils/
│   └── formulaEngine.ts     # Formula evaluation
└── pages/
    └── Index.tsx            # Main application
```

## Key Interfaces

### CellData
```typescript
interface CellData {
  value: string;
  formula?: string;
}
```

### Cell Formatting (Needs Extension)
Current cell data only stores value and formula. Need to extend for:
- Font family, size, style (bold, italic, underline)
- Text and background colors
- Alignment (horizontal, vertical)
- Borders
- Number formatting

## State Management
- Centralized in Index.tsx
- Sheet-based data storage with `sheetDataMap`
- Selection tracking with start/end coordinates
- Formula building mode for references

## Current Limitations
1. CellData interface too simple - needs formatting properties
2. No clipboard operations implemented
3. Ribbon buttons have no event handlers
4. No styling persistence in cells
5. No undo/redo system

## Technologies
- React 18 with hooks for state management
- TypeScript for type safety
- Tailwind CSS for styling
- shadcn/ui for consistent UI components
- Custom formula engine for calculations
