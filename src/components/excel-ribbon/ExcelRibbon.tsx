import { useState } from "react";
import { Button } from "@/components/ui/button";
import { RibbonTab } from "./RibbonTab";
import { RibbonGroup } from "./RibbonGroup";
import { RibbonDropdown } from "./RibbonDropdown";
import { CompactRibbonGroup } from "./CompactRibbonGroup";
import { FontDropdown } from "./FontDropdown";
import { FontSizeDropdown } from "./FontSizeDropdown";
import { BorderDropdown } from "./BorderDropdown";
import { ColorPicker } from "./ColorPicker";
import { ClipboardDropdown } from "./ClipboardDropdown";
import { FontDropdownMobile } from "./FontDropdownMobile";
import { AlignmentDropdownMobile } from "./AlignmentDropdownMobile";
import { NumberDropdownMobile } from "./NumberDropdownMobile";
import { StylesDropdownMobile } from "./StylesDropdownMobile";
import { CellsDropdownMobile } from "./CellsDropdownMobile";
import { EditingDropdownMobile } from "./EditingDropdownMobile";
import { SensitivityDropdownMobile } from "./SensitivityDropdownMobile";
import { CopilotDropdownMobile } from "./CopilotDropdownMobile";
import { RibbonActions } from "@/types/cellTypes";
import { 
  UndoIcon,
  RedoIcon,
  ClipboardIcon,
  CopyIcon,
  CutIcon,
  FormatPainterIcon,
  BoldIcon,
  ItalicIcon,
  UnderlineIcon,
  DoubleUnderlineIcon,
  StrikeThroughIcon,
  AlignLeftIcon,
  AlignCenterIcon,
  AlignRightIcon,
  AlignJustifyIcon,
  AlignTopIcon,
  AlignMiddleIcon,
  AlignBottomIcon,
  AlignTextLeftIcon,
  AlignTextCenterIcon,
  AlignTextRightIcon,
  PlusIcon,
  MinusIcon,
  PercentIcon,
  HashIcon,
  TypeIcon,
  PaletteIcon,
  SettingsIcon,
  SearchIcon,
  CalculatorIcon,
  ChevronDownIcon,
  SaveIcon,
  ShareIcon,
  EyeIcon,
  FileSpreadsheetIcon,
  DecreaseIndentIcon,
  IncreaseIndentIcon,
  OrientationIcon,
  MergeIcon,
  BordersIcon,
  FillColorIcon,
  FontIncrementIcon,
  FontDecrementIcon,
  AccountingIcon,
  CommaIcon,
  DecreaseDecimalIcon,
  IncreaseDecimalIcon,
  ConditionalFormattingIcon,
  CopilotIcon,
  LargeSensitivityIcon,
  LargeCopilotIcon
} from "./icons";

interface ExcelRibbonProps {
  ribbonActions: RibbonActions;
}

export const ExcelRibbon = ({ ribbonActions }: ExcelRibbonProps) => {
  const [activeTab, setActiveTab] = useState("Home");

  const tabs = ["File", "Home", "Insert", "Page Layout", "Formulas", "Data", "Review", "View", "Help"];

  return (
    <div className="w-full bg-gray-50">
      {/* Tab Navigation */}
      <div className="flex items-center justify-between bg-gray-50">
        <div className="flex">
          {tabs.map((tab) => (
            <RibbonTab
              key={tab}
              active={activeTab === tab}
              onClick={() => setActiveTab(tab)}
            >
              {tab}
            </RibbonTab>
          ))}
        </div>
        
        {/* Right-aligned peripheral controls */}
        <div className="flex items-center gap-1 pr-2">
          {/* Share Button */}
          <Button 
            variant="ghost" 
            size="sm" 
            className="h-6 px-3 text-xs flex items-center gap-2 bg-green-700 text-white hover:bg-green-950 rounded-[4px]"
          >
            <img src="/icons/share icon.png" alt="Share" className="w-4 h-4" />
            <span>Share</span>
            <svg width="12" height="12" viewBox="0 0 12 12" className="ml-1">
              <path d="M6 9L1.5 4.5L3 3L6 6L9 3L10.5 4.5L6 9Z" fill="currentColor" />
            </svg>
          </Button>
        </div>
      </div>

      {/* Ribbon Content */}
      {activeTab === "Home" && (
        <>
          {/* Desktop View */}
          <div className="hidden lg:flex bg-white rounded-[10px] shadow-md mx-2 mb-2 px-2 py-2">
            {/* Undo/Redo Group */}
            <RibbonGroup title="Undo">
              <div className="flex flex-col gap-0.5">
                <Button variant="ghost" className="h-6 w-6 p-0 hover:bg-blue-50" onClick={ribbonActions.undo}>
                  <UndoIcon className="w-4 h-4" />
                </Button>
                <Button variant="ghost" className="h-6 w-6 p-0 hover:bg-blue-50" onClick={ribbonActions.redo}>
                  <RedoIcon className="w-4 h-4" />
                </Button>
              </div>
            </RibbonGroup>

            {/* Clipboard Group */}
            <RibbonGroup title="Clipboard">
              <div className="flex items-center gap-1">
                <div className="flex flex-col">
                  <Button variant="ghost" size="sm" className="h-14 w-14 flex-col p-1 hover:bg-blue-50" onClick={ribbonActions.paste}>
                    <ClipboardIcon className="w-8 h-9 mb-1" />
                    <span className="text-xs">Paste</span>
                  </Button>
                </div>
                <div className="flex flex-col gap-0.5">
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50" onClick={ribbonActions.cut}>
                    <CutIcon className="w-4 h-4" />
                  </Button>
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50" onClick={ribbonActions.copy}>
                    <CopyIcon className="w-4 h-4" />
                  </Button>
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50">
                    <FormatPainterIcon className="w-4 h-4" />
                  </Button>
                </div>
              </div>
            </RibbonGroup>

            {/* Font Group */}
            <RibbonGroup title="Font">
              <div className="flex flex-col gap-1">
                {/* Row 1: Font name and size */}
                <div className="flex items-center gap-1">
                  <select 
                    className="text-xs border rounded-[4px] px-2 py-1 w-28 h-6"
                    onChange={(e) => ribbonActions.setFontFamily(e.target.value)}
                  >
                    <option value="Calibri">Calibri</option>
                    <option value="Arial">Arial</option>
                    <option value="Times New Roman">Times New Roman</option>
                    <option value="Helvetica">Helvetica</option>
                    <option value="Georgia">Georgia</option>
                  </select>
                  <select 
                    className="text-xs border rounded px-2 py-1 w-14 h-6"
                    onChange={(e) => ribbonActions.setFontSize(parseInt(e.target.value))}
                  >
                    <option value="8">8</option>
                    <option value="9">9</option>
                    <option value="10">10</option>
                    <option value="11">11</option>
                    <option value="12">12</option>
                    <option value="14">14</option>
                    <option value="16">16</option>
                    <option value="18">18</option>
                    <option value="20">20</option>
                    <option value="24">24</option>
                  </select>
                </div>
                
                {/* Row 2: B, I, U, double underline, strikethrough, and font size controls */}
                <div className="flex items-center gap-0.5">
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50" onClick={ribbonActions.toggleBold}>
                    <BoldIcon className="w-4 h-5" />
                  </Button>
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50" onClick={ribbonActions.toggleItalic}>
                    <ItalicIcon className="w-4 h-5" />
                  </Button>
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50" onClick={ribbonActions.toggleUnderline}>
                    <UnderlineIcon className="w-4 h-7" />
                  </Button>
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50">
                    <DoubleUnderlineIcon className="w-6 h-6" />
                  </Button>
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50">
                    <StrikeThroughIcon className="w-5 h-5" />
                  </Button>
                  <div className="w-px h-4 bg-gray-300 mx-1" />
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50" onClick={ribbonActions.increaseFontSize}>
                    <FontIncrementIcon className="w-5 h-5" />
                  </Button>
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50" onClick={ribbonActions.decreaseFontSize}>
                    <FontDecrementIcon className="w-5 h-5" />
                  </Button>
                </div>
                
                {/* Row 3: Borders, fill color, text color */}
                <div className="flex items-center gap-0.5">
                  <Button variant="ghost" size="sm" className="h-6 w-14 p-0 hover:bg-blue-50 flex items-center justify-start pl-1 relative">
                    <BordersIcon className="w-6 h-6" />
                    <ChevronDownIcon className="w-3 h-3 absolute bottom-1 right-3" />
                  </Button>
                  <ColorPicker type="background" defaultColor="#FFFF00" onColorChange={ribbonActions.setBackgroundColor} />
                  <ColorPicker type="text" defaultColor="#FF0000" onColorChange={ribbonActions.setTextColor} />
                </div>
              </div>
            </RibbonGroup>

            {/* Alignment Group */}
            <RibbonGroup title="Alignment">
              <div className="flex flex-col gap-1">
                {/* Row 1: Align top, middle, bottom */}
                <div className="flex items-center gap-0.5">
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50" onClick={() => ribbonActions.setVerticalAlignment('top')}>
                    <AlignTopIcon className="w-5 h-5" />
                  </Button>
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50" onClick={() => ribbonActions.setVerticalAlignment('middle')}>
                    <AlignMiddleIcon className="w-5 h-5" />
                  </Button>
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50" onClick={() => ribbonActions.setVerticalAlignment('bottom')}>
                    <AlignBottomIcon className="w-5 h-5" />
                  </Button>
                </div>
                
                {/* Row 2: Align text left, center, right */}
                <div className="flex items-center gap-0.5">
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50" onClick={() => ribbonActions.setHorizontalAlignment('left')}>
                    <AlignTextLeftIcon className="w-5 h-5" />
                  </Button>
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50" onClick={() => ribbonActions.setHorizontalAlignment('center')}>
                    <AlignTextCenterIcon className="w-5 h-5" />
                  </Button>
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50" onClick={() => ribbonActions.setHorizontalAlignment('right')}>
                    <AlignTextRightIcon className="w-5 h-5" />
                  </Button>
                </div>
                
                {/* Row 3: Decrease indent, increase indent, orientation */}
                <div className="flex items-center gap-0.5">
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50">
                    <DecreaseIndentIcon className="w-4 h-4" />
                  </Button>
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50">
                    <IncreaseIndentIcon className="w-4 h-4" />
                  </Button>
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50">
                    <OrientationIcon className="w-4 h-4" />
                  </Button>
                </div>
              </div>
            </RibbonGroup>

            {/* Number Group */}
            <RibbonGroup title="Number">
              <div className="flex flex-col gap-1">
                {/* Row 1: Format dropdown */}
                <select 
                  className="text-xs border rounded px-2 py-1 w-24 h-6"
                  onChange={(e) => {
                    const formatType = e.target.value as 'general' | 'number' | 'currency' | 'percentage';
                    ribbonActions.setNumberFormat({ type: formatType, decimals: 2 });
                  }}
                >
                  <option value="general">General</option>
                  <option value="number">Number</option>
                  <option value="currency">Currency</option>
                  <option value="percentage">Percentage</option>
                </select>
                
                {/* Row 2: Accounting, Percentage, Comma, Decrease/Increase Decimal */}
                <div className="flex items-center gap-0.5">
                  <Button 
                    variant="ghost" 
                    size="sm" 
                    className="h-6 w-8 p-0 hover:bg-blue-50 flex items-center justify-start pl-1 relative"
                    onClick={() => ribbonActions.setNumberFormat({ type: 'accounting', decimals: 2 })}
                  >
                    <AccountingIcon className="w-6 h-6" />
                    <ChevronDownIcon className="w-2 h-2 absolute bottom-0 right-3" />
                  </Button>
                  <Button 
                    variant="ghost" 
                    size="sm" 
                    className="h-6 w-6 p-0 hover:bg-blue-50"
                    onClick={() => ribbonActions.setNumberFormat({ type: 'percentage', decimals: 2 })}
                  >
                    <PercentIcon className="w-4 h-4" />
                  </Button>
                  <Button 
                    variant="ghost" 
                    size="sm" 
                    className="h-6 w-6 p-0 hover:bg-blue-50"
                    onClick={() => ribbonActions.setNumberFormat({ type: 'number', decimals: 0, useThousandsSeparator: true })}
                  >
                    <CommaIcon className="w-6 h-7" />
                  </Button>
                  <Button 
                    variant="ghost" 
                    size="sm" 
                    className="h-6 w-6 p-0 hover:bg-blue-50"
                    onClick={ribbonActions.decreaseDecimals}
                  >
                    <DecreaseDecimalIcon className="w-8 h-8" />
                  </Button>
                  <Button 
                    variant="ghost" 
                    size="sm" 
                    className="h-6 w-6 p-0 hover:bg-blue-50"
                    onClick={ribbonActions.increaseDecimals}
                  >
                    <IncreaseDecimalIcon className="w-8 h-8" />
                  </Button>
                </div>
              </div>
            </RibbonGroup>

            {/* Styles Group */}
            <RibbonGroup title="Styles">
              <div className="flex flex-col gap-0.5">
                {/* Row 1: Conditional Formatting */}
                <Button variant="ghost" size="sm" className="h-6 w-40 p-1 text-xs justify-start hover:bg-blue-50 flex items-center relative">
                  <ConditionalFormattingIcon className="w-4 h-4 mr-2" />
                  <span>Conditional Formatting</span>
                  <ChevronDownIcon className="w-3 h-3 absolute bottom-0 right-1" />
                </Button>
                
                {/* Row 2: Format as Table */}
                <Button variant="ghost" size="sm" className="h-6 w-40 p-1 text-xs justify-start hover:bg-blue-50 flex items-center relative">
                  <HashIcon className="w-4 h-4 mr-2" />
                  <span>Format as Table</span>
                  <ChevronDownIcon className="w-3 h-3 absolute bottom-0 right-1" />
                </Button>
                
                {/* Row 3: Cell Styles */}
                <Button variant="ghost" size="sm" className="h-6 w-40 p-1 text-xs justify-start hover:bg-blue-50 flex items-center relative">
                  <PaletteIcon className="w-4 h-4 mr-2" />
                  <span>Cell Styles</span>
                  <ChevronDownIcon className="w-3 h-3 absolute bottom-0 right-1" />
                </Button>
              </div>
            </RibbonGroup>

            {/* Cells Group */}
            <RibbonGroup title="Cells">
              <div className="flex flex-col gap-0.5">
                <Button variant="ghost" size="sm" className="h-6 w-20 p-1 text-xs justify-start hover:bg-blue-50">
                  <PlusIcon className="w-5 h-5 mr-1" />
                  Insert
                </Button>
                <Button variant="ghost" size="sm" className="h-6 w-20 p-1 text-xs justify-start hover:bg-blue-50">
                  <MinusIcon className="w-5 h-5 mr-1" />
                  Delete
                </Button>
                <Button variant="ghost" size="sm" className="h-6 w-20 p-1 text-xs justify-start hover:bg-blue-50">
                  <PaletteIcon className="w-5 h-5 mr-1" />
                  Format
                </Button>
              </div>
            </RibbonGroup>

            {/* Editing Group */}
            <RibbonGroup title="Editing">
              <div className="flex flex-col gap-0.5">
                <Button variant="ghost" size="sm" className="h-6 w-24 p-1 text-xs justify-start hover:bg-blue-50">
                  <CalculatorIcon className="w-3 h-3 mr-1" />
                  AutoSum
                  <ChevronDownIcon className="w-2 h-2 ml-auto" />
                </Button>
                <Button variant="ghost" size="sm" className="h-6 w-24 p-1 text-xs justify-start hover:bg-blue-50">
                  <PaletteIcon className="w-3 h-3 mr-1" />
                  Fill
                  <ChevronDownIcon className="w-2 h-2 ml-auto" />
                </Button>
                <Button variant="ghost" size="sm" className="h-6 w-24 p-1 text-xs justify-start hover:bg-blue-50">
                  <SearchIcon className="w-3 h-3 mr-1" />
                  Find & Select
                  <ChevronDownIcon className="w-2 h-2 ml-auto" />
                </Button>
              </div>
            </RibbonGroup>

            {/* Sensitivity Group */}
            <RibbonGroup title="Sensitivity">
              <div className="flex flex-col items-center">
                <Button variant="ghost" size="sm" className="h-14 w-14 flex-col p-1 hover:bg-blue-50">
                  <LargeSensitivityIcon className="w-8 h-8 mb-1" />
                  <span className="text-xs">Sensitivity</span>
                </Button>
              </div>
            </RibbonGroup>

            {/* Copilot Group */}
            <RibbonGroup title="Copilot">
              <div className="flex flex-col items-center">
                <Button variant="ghost" size="sm" className="h-14 w-14 flex-col p-1 hover:bg-blue-50 relative">
                  <LargeCopilotIcon className="w-8 h-8 mb-1" />
                  <span className="text-xs">Copilot</span>
                  <ChevronDownIcon className="w-3 h-3 absolute bottom-0 right-1" />
                </Button>
              </div>
            </RibbonGroup>
          </div>

          {/* Mobile/Tablet Compact View */}
          <div className="flex lg:hidden bg-white rounded-md shadow-md mx-2 mb-2 px-2 py-2 gap-2">
            {/* Undo/Redo Group - Same as desktop */}
            <RibbonGroup title="Undo">
              <div className="flex flex-col gap-0.5">
                <Button variant="ghost" className="h-6 w-6 p-0 hover:bg-blue-50">
                  <UndoIcon className="w-4 h-4" />
                </Button>
                <Button variant="ghost" className="h-6 w-6 p-0 hover:bg-blue-50">
                  <RedoIcon className="w-4 h-4" />
                </Button>
              </div>
            </RibbonGroup>

            {/* Clipboard Group - Same as desktop */}
            <RibbonGroup title="Clipboard">
              <div className="flex items-center gap-1">
                <div className="flex flex-col">
                  <Button variant="ghost" size="sm" className="h-14 w-14 flex-col p-1 hover:bg-blue-50">
                    <ClipboardIcon className="w-8 h-9 mb-1" />
                    <span className="text-xs">Paste</span>
                  </Button>
                </div>
                <div className="flex flex-col gap-0.5">
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50">
                    <CutIcon className="w-4 h-4" />
                  </Button>
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50">
                    <CopyIcon className="w-4 h-4" />
                  </Button>
                  <Button variant="ghost" size="sm" className="h-6 w-6 p-0 hover:bg-blue-50">
                    <FormatPainterIcon className="w-4 h-4" />
                  </Button>
                </div>
              </div>
            </RibbonGroup>

            {/* Other groups as dropdowns with chevrons */}
            <FontDropdownMobile />
            <AlignmentDropdownMobile />
            <NumberDropdownMobile />
            <StylesDropdownMobile />
            <CellsDropdownMobile />
            <EditingDropdownMobile />
            <SensitivityDropdownMobile />
            <CopilotDropdownMobile />
          </div>
        </>
      )}
    </div>
  );
};
