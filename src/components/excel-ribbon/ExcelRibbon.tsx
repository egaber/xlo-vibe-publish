import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle, DialogTrigger } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { RibbonTab } from "./RibbonTab";
import { RibbonGroup } from "./RibbonGroup";
import { RibbonDropdown } from "./RibbonDropdown";
import { CompactRibbonGroup } from "./CompactRibbonGroup";
import { FontDropdown } from "./FontDropdown";
import { FontSizeDropdown } from "./FontSizeDropdown";
import { BorderDropdown } from "./BorderDropdown";
import { ColorPicker } from "./ColorPicker";
import { ClipboardDropdown } from "./ClipboardDropdown";
import { PasteDropdown } from "./PasteDropdown";
import { FontDropdownMobile } from "./FontDropdownMobile";
import { AlignmentDropdownMobile } from "./AlignmentDropdownMobile";
import { NumberDropdownMobile } from "./NumberDropdownMobile";
import { StylesDropdownMobile } from "./StylesDropdownMobile";
import { CellsDropdownMobile } from "./CellsDropdownMobile";
import { EditingDropdownMobile } from "./EditingDropdownMobile";
import { SensitivityDropdownMobile } from "./SensitivityDropdownMobile";
import { CopilotDropdownMobile } from "./CopilotDropdownMobile";
import { RibbonActions } from "@/types/cellTypes";
import { useSvgIconConversion } from "@/hooks/useSvgIconConversion";
import { 
  UndoIcon,
  RedoIcon,
  ClipboardIcon,
  CopyIcon,
  CutIcon,
  FormatPainterIcon,
  PasteWithCopilotIcon,
  PasteValuesIcon,
  PasteFormulasIcon,
  PasteFormattingIcon,
  PasteLinkIcon,
  PasteColumnWidthsIcon,
  PasteTransposeIcon,
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
  MergeCenterIcon,
  BordersIcon,
  FillColorIcon,
  FontIncrementIcon,
  FontDecrementIcon,
  WrapTextIcon,
  AccountingIcon,
  CommaIcon,
  DecreaseDecimalIcon,
  IncreaseDecimalIcon,
  ConditionalFormattingIcon,
  CopilotIcon,
  LargeSensitivityIcon,
  LargeCopilotIcon,
  AutoSumIcon,
  ClearIcon,
  SortIcon,
  FilterIcon
} from "./icons";

interface ExcelRibbonProps {
  ribbonActions: RibbonActions;
}

export const ExcelRibbon = ({ ribbonActions }: ExcelRibbonProps) => {
  const [activeTab, setActiveTab] = useState("Home");
  const [currentTextColor, setCurrentTextColor] = useState("#FF0000");
  const [currentBackgroundColor, setCurrentBackgroundColor] = useState("#FFFF00");
  const [testDialogOpen, setTestDialogOpen] = useState(false);
  const [taskPaneOpen, setTaskPaneOpen] = useState(false);
  const [testInputValue, setTestInputValue] = useState("");

  // Convert SVG icons when the ribbon re-renders
  useSvgIconConversion(undefined, [activeTab]);

  const tabs = ["File", "Home", "Insert", "Page Layout", "Formulas", "Data", "Review", "View", "Help"];

  const handleTextColorChange = (color: string) => {
    setCurrentTextColor(color);
    ribbonActions.setTextColor(color);
  };

  const handleBackgroundColorChange = (color: string) => {
    setCurrentBackgroundColor(color);
    ribbonActions.setBackgroundColor(color);
  };

  const handleTestSubmit = () => {
    // Handle the test input submission here
    console.log("Test input:", testInputValue);
    setTestDialogOpen(false);
    setTestInputValue("");
  };

  const handleTestCancel = () => {
    setTestDialogOpen(false);
    setTestInputValue("");
  };

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
<div className="flex items-center justify-end gap-4 pr-4">
          {/* Share Button */}
          <Button 
            variant="ghost" 
            size="sm" 
            className="h-6 px-3 text-xs flex items-center gap-2 bg-green-700 text-white hover:bg-green-950 rounded-[4px]"
          >
<ShareIcon className="w-4 h-4" />
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
          <div className="hidden lg:flex bg-white rounded-[10px] shadow-md mx-2 mb-2 px-2 py-2 h-[120px]">
            {/* Undo/Redo Group */}
            <RibbonGroup title="Undo">
<div className="flex flex-col gap-1">
                <Button variant="ribbon" size="ribbon-mini" onClick={ribbonActions.undo}>
<UndoIcon />
                </Button>
                <Button variant="ribbon" size="ribbon-mini" onClick={ribbonActions.redo}>
<RedoIcon />
                </Button>
              </div>
            </RibbonGroup>

            {/* Clipboard Group */}
            <RibbonGroup title="Clipboard">
              <div className="flex items-center gap-1">
                <div className="flex flex-col">
                  <PasteDropdown 
                    onPaste={ribbonActions.paste}
                    onPasteWithCopilot={() => console.log('Paste with Copilot')}
                    onPasteValuesOnly={() => console.log('Paste Values Only')}
                    onPasteFormulasOnly={() => console.log('Paste Formulas Only')}
                    onPasteFormattingOnly={() => console.log('Paste Formatting Only')}
                    onPasteLinkToSource={() => console.log('Link to Source')}
                    onPasteKeepColumnWidth={() => console.log('Keep Source Column Width')}
                    onPasteTranspose={() => console.log('Transpose Rows and Columns')}
                  />
                </div>
                <div className="flex flex-col gap-0.5">
                  <Button variant="ribbon" size="ribbon-mini" onClick={ribbonActions.cut}>
<CutIcon />
                  </Button>
                  <Button variant="ribbon" size="ribbon-mini" onClick={ribbonActions.copy}>
<CopyIcon />
                  </Button>
                  <Button variant="ribbon" size="ribbon-mini">
<FormatPainterIcon />
                  </Button>
                </div>
              </div>
            </RibbonGroup>

            {/* Font Group - Responsive 2-line (286.8px) or 3-line (193.2px) layout */}
            <RibbonGroup title="Font" className="group">
              {/* 2-line layout for wide screens */}
              <div className="hidden xl:block" style={{ width: '286.8px' }}>
                <div className="flex flex-col gap-0.5 h-[100px] justify-center py-2">
                  {/* Row 1: Font name, size, and font size controls */}
                  <div className="flex items-center gap-1">
                    <select 
                      className="text-xs border rounded-[4px] px-2 py-1 h-[24px]"
                      style={{ width: '126.4px' }}
                      onChange={(e) => ribbonActions.setFontFamily(e.target.value)}
                    >
                      <option value="Calibri">Calibri</option>
                      <option value="Arial">Arial</option>
                      <option value="Times New Roman">Times New Roman</option>
                      <option value="Helvetica">Helvetica</option>
                      <option value="Georgia">Georgia</option>
                    </select>
                    <select 
                      className="text-xs border rounded px-2 py-1 h-[24px]"
                      style={{ width: '56.4px' }}
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
                    <div className="flex items-center gap-0.5 ml-1">
                      <Button variant="ribbon" size="ribbon-mini" onClick={ribbonActions.increaseFontSize}>
                        <FontIncrementIcon />
                      </Button>
                      <Button variant="ribbon" size="ribbon-mini" onClick={ribbonActions.decreaseFontSize}>
                        <FontDecrementIcon />
                      </Button>
                    </div>
                  </div>
                  
                  {/* Row 2: Formatting buttons, borders, and colors */}
                  <div className="flex items-center gap-1">
                    <div className="flex items-center gap-0.5">
                      <Button variant="ribbon" size="ribbon-mini" onClick={ribbonActions.toggleBold}>
                        <BoldIcon />
                      </Button>
                      <Button variant="ribbon" size="ribbon-mini" onClick={ribbonActions.toggleItalic}>
                        <ItalicIcon />
                      </Button>
                      <Button variant="ribbon" size="ribbon-mini" onClick={ribbonActions.toggleUnderline}>
                        <UnderlineIcon />
                      </Button>
                      <Button variant="ribbon" size="ribbon-mini">
                        <DoubleUnderlineIcon />
                      </Button>
                      <Button variant="ribbon" size="ribbon-mini">
                        <StrikeThroughIcon />
                      </Button>
                    </div>
                    <div className="w-px h-4 bg-gray-300 mx-1" />
                    <Button variant="ghost" size="sm" className="h-6 w-10 p-0 hover:bg-blue-50 flex items-center justify-start pl-1 relative">
                      <BordersIcon className="w-4 h-4" />
                      <ChevronDownIcon className="w-3 h-3 absolute bottom-1 right-1" />
                    </Button>
                    <div className="relative">
                      <ColorPicker 
                        type="background" 
                        defaultColor="#FFFF00" 
                        selectedColor={currentBackgroundColor}
                        onColorChange={handleBackgroundColorChange}
                        onButtonClick={() => ribbonActions.setBackgroundColor(currentBackgroundColor)}
                      />
                    </div>
                    <div className="relative">
                      <ColorPicker 
                        type="text" 
                        defaultColor="#FF0000" 
                        selectedColor={currentTextColor}
                        onColorChange={handleTextColorChange}
                        onButtonClick={() => ribbonActions.setTextColor(currentTextColor)}
                      />
                    </div>
                  </div>
                </div>
              </div>
              
              {/* 3-line layout for medium screens */}
              <div className="block xl:hidden" style={{ width: '193.2px' }}>
                <div className="flex flex-col gap-0.5 h-[100px] justify-center py-2">
                  {/* Row 1: Font name, size, and font size controls */}
                  <div className="flex items-center gap-1">
                    <select 
                      className="text-xs border rounded-[4px] px-2 py-1 h-[24px]"
                      style={{ width: '90px' }}
                      onChange={(e) => ribbonActions.setFontFamily(e.target.value)}
                    >
                      <option value="Calibri">Calibri</option>
                      <option value="Arial">Arial</option>
                      <option value="Times New Roman">Times New Roman</option>
                      <option value="Helvetica">Helvetica</option>
                      <option value="Georgia">Georgia</option>
                    </select>
                    <select 
                      className="text-xs border rounded px-2 py-1 h-[24px]"
                      style={{ width: '40px' }}
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
                    <Button variant="ribbon" size="ribbon-mini" onClick={ribbonActions.increaseFontSize}>
                      <FontIncrementIcon />
                    </Button>
                    <Button variant="ribbon" size="ribbon-mini" onClick={ribbonActions.decreaseFontSize}>
                      <FontDecrementIcon />
                    </Button>
                  </div>
                  
                  {/* Row 2: B, I, U, double underline, strikethrough */}
                  <div className="flex items-center gap-0.5">
                    <Button variant="ribbon" size="ribbon-mini" onClick={ribbonActions.toggleBold}>
                      <BoldIcon />
                    </Button>
                    <Button variant="ribbon" size="ribbon-mini" onClick={ribbonActions.toggleItalic}>
                      <ItalicIcon />
                    </Button>
                    <Button variant="ribbon" size="ribbon-mini" onClick={ribbonActions.toggleUnderline}>
                      <UnderlineIcon />
                    </Button>
                    <Button variant="ribbon" size="ribbon-mini">
                      <DoubleUnderlineIcon />
                    </Button>
                    <Button variant="ribbon" size="ribbon-mini">
                      <StrikeThroughIcon />
                    </Button>
                  </div>
                  
                  {/* Row 3: Borders, fill color, text color */}
                  <div className="flex items-center gap-0.5">
                    <Button variant="ghost" size="sm" className="h-6 w-10 p-0 hover:bg-blue-50 flex items-center justify-start pl-1 relative">
                      <BordersIcon className="w-4 h-4" />
                      <ChevronDownIcon className="w-3 h-3 absolute bottom-1 right-1" />
                    </Button>
                    <div className="relative">
                      <ColorPicker 
                        type="background" 
                        defaultColor="#FFFF00" 
                        selectedColor={currentBackgroundColor}
                        onColorChange={handleBackgroundColorChange}
                        onButtonClick={() => ribbonActions.setBackgroundColor(currentBackgroundColor)}
                      />
                    </div>
                    <div className="relative">
                      <ColorPicker 
                        type="text" 
                        defaultColor="#FF0000" 
                        selectedColor={currentTextColor}
                        onColorChange={handleTextColorChange}
                        onButtonClick={() => ribbonActions.setTextColor(currentTextColor)}
                      />
                    </div>
                  </div>
                </div>
              </div>
            </RibbonGroup>

            {/* Alignment Group */}
            <RibbonGroup title="Alignment">
              <div className="flex items-start gap-2" style={{ width: '243px' }}>
                {/* Left side: Main alignment controls - takes remaining space */}
                <div className="flex flex-col gap-1 flex-1">
                  {/* Row 1: Align top, middle, bottom */}
                  <div className="flex items-center gap-0.5">
                    <Button variant="ribbon" size="ribbon-mini" onClick={() => ribbonActions.setVerticalAlignment('top')}>
                      <AlignTopIcon />
                    </Button>
                    <Button variant="ribbon" size="ribbon-mini" onClick={() => ribbonActions.setVerticalAlignment('middle')}>
                      <AlignMiddleIcon />
                    </Button>
                    <Button variant="ribbon" size="ribbon-mini" onClick={() => ribbonActions.setVerticalAlignment('bottom')}>
                      <AlignBottomIcon />
                    </Button>
                  </div>
                  
                  {/* Row 2: Align text left, center, right */}
                  <div className="flex items-center gap-0.5">
                    <Button variant="ribbon" size="ribbon-mini" onClick={() => ribbonActions.setHorizontalAlignment('left')}>
                      <AlignTextLeftIcon />
                    </Button>
                    <Button variant="ribbon" size="ribbon-mini" onClick={() => ribbonActions.setHorizontalAlignment('center')}>
                      <AlignTextCenterIcon />
                    </Button>
                    <Button variant="ribbon" size="ribbon-mini" onClick={() => ribbonActions.setHorizontalAlignment('right')}>
                      <AlignTextRightIcon />
                    </Button>
                  </div>
                  
                  {/* Row 3: Decrease indent, increase indent, orientation */}
                  <div className="flex items-center gap-0.5">
                    <Button variant="ribbon" size="ribbon-mini">
                      <DecreaseIndentIcon />
                    </Button>
                    <Button variant="ribbon" size="ribbon-mini">
                      <IncreaseIndentIcon />
                    </Button>
                    <Button variant="ribbon" size="ribbon-mini">
                      <OrientationIcon />
                    </Button>
                  </div>
                </div>
                
                {/* Vertical divider */}
                <div className="w-px h-16 bg-gray-300 self-start mt-1"></div>
                
                {/* Right side: Wrap text and merge & center - 140px */}
                <div className="flex flex-col gap-1" style={{ width: '140px' }}>
                  <Button 
                    variant="ghost" 
                    size="sm" 
                    className="h-6 w-full p-1 text-xs justify-start hover:bg-blue-50 flex items-center"
                    title="Wrap Text"
                  >
                    <WrapTextIcon className="w-4 h-4 mr-1" />
                    <span>Wrap Text</span>
                  </Button>
                  <Button 
                    variant="ghost" 
                    size="sm" 
                    className="h-6 w-full p-1 text-xs justify-start hover:bg-blue-50 flex items-center"
                    title="Merge & Center"
                  >
                    <MergeCenterIcon className="w-4 h-4 mr-1" />
                    <span>Merge & Center</span>
                  </Button>
                </div>
              </div>
            </RibbonGroup>

            {/* Number Group */}
            <RibbonGroup title="Number">
              <div className="flex flex-col gap-1" style={{ width: '150px' }}>
                {/* Row 1: Format dropdown */}
                <select 
                  className="text-xs border rounded px-2 py-1 w-full h-6"
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
                    variant="ribbon" 
                    size="ribbon-mini"
                    className="w-8 relative"
                    onClick={() => ribbonActions.setNumberFormat({ type: 'accounting', decimals: 2 })}
                  >
<AccountingIcon />
                    <ChevronDownIcon className="w-2 h-2 absolute bottom-0 right-1" />
                  </Button>
                  <Button 
                    variant="ribbon" 
                    size="ribbon-mini"
                    onClick={() => ribbonActions.setNumberFormat({ type: 'percentage', decimals: 2 })}
                  >
<PercentIcon />
                  </Button>
                  <Button 
                    variant="ribbon" 
                    size="ribbon-mini"
                    onClick={() => ribbonActions.setNumberFormat({ type: 'number', decimals: 0, useThousandsSeparator: true })}
                  >
<CommaIcon />
                  </Button>
                  <Button 
                    variant="ribbon" 
                    size="ribbon-mini"
                    onClick={ribbonActions.decreaseDecimals}
                  >
<DecreaseDecimalIcon />
                  </Button>
                  <Button 
                    variant="ribbon" 
                    size="ribbon-mini"
                    onClick={ribbonActions.increaseDecimals}
                  >
<IncreaseDecimalIcon />
                  </Button>
                </div>
              </div>
            </RibbonGroup>

            {/* Styles Group */}
            <RibbonGroup title="Styles">
              <div className="flex flex-col gap-0.5">
                {/* Row 1: Conditional Formatting */}
                <Button variant="ghost" size="sm" className="h-6 w-40 p-1 text-xs justify-start hover:bg-blue-50 flex items-center relative">
<ConditionalFormattingIcon className="w-5 h-5 mr-2" />
                  <span>Conditional Formatting</span>
                  <ChevronDownIcon className="w-3 h-3 absolute bottom-0 right-1" />
                </Button>
                
                {/* Row 2: Format as Table */}
                <Button variant="ghost" size="sm" className="h-6 w-40 p-1 text-xs justify-start hover:bg-blue-50 flex items-center relative">
<HashIcon className="w-5 h-5 mr-2" />
                  <span>Format as Table</span>
                  <ChevronDownIcon className="w-3 h-3 absolute bottom-0 right-1" />
                </Button>
                
                {/* Row 3: Cell Styles */}
                <Button variant="ghost" size="sm" className="h-6 w-40 p-1 text-xs justify-start hover:bg-blue-50 flex items-center relative">
<PaletteIcon className="w-5 h-5 mr-2" />
                  <span>Cell Styles</span>
                  <ChevronDownIcon className="w-3 h-3 absolute bottom-0 right-1" />
                </Button>
              </div>
            </RibbonGroup>

            {/* Editing Group */}
            <RibbonGroup title="Editing">
              <div className="flex items-center gap-2">
                {/* Left side: Two small controls stacked vertically */}
                <div className="flex flex-col gap-0.5">
                  {/* Top: AutoSum with dropdown */}
                  <Button 
                    variant="ribbon" 
                    size="ribbon-mini"
                    className="w-8 relative"
                    onClick={() => console.log('AutoSum')}
                  >
                    <AutoSumIcon />
                    <ChevronDownIcon className="w-2 h-2 absolute bottom-0 right-1" />
                  </Button>
                  
                  {/* Bottom: Clear/Eraser with dropdown */}
                  <Button 
                    variant="ribbon" 
                    size="ribbon-mini"
                    className="w-8 relative"
                    onClick={() => console.log('Clear')}
                  >
                    <ClearIcon />
                    <ChevronDownIcon className="w-2 h-2 absolute bottom-0 right-1" />
                  </Button>
                </div>
                
                {/* Right side: Two larger buttons arranged horizontally */}
                <div className="flex gap-1">
                  {/* Sort & Filter button */}
                  <Button 
                    variant="ghost" 
                    size="sm" 
                    className="h-14 w-16 flex-col p-1 hover:bg-blue-50 text-xs"
                    onClick={() => console.log('Sort & Filter')}
                  >
                    <div className="flex items-center mb-1">
                      <SortIcon className="w-4 h-4 mr-1" />
                      <FilterIcon className="w-4 h-4" />
                    </div>
                    <span>Sort &</span>
                    <span>Filter</span>
                  </Button>
                  
                  {/* Find & Select button */}
                  <Button 
                    variant="ghost" 
                    size="sm" 
                    className="h-14 w-16 flex-col p-1 hover:bg-blue-50 text-xs"
                    onClick={() => console.log('Find & Select')}
                  >
                    <SearchIcon className="w-5 h-5 mb-1" />
                    <span>Find &</span>
                    <span>Select</span>
                  </Button>
                </div>
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
<LargeSensitivityIcon className="w-9 h-9 mb-1" />
                  <span className="text-xs">Sensitivity</span>
                </Button>
              </div>
            </RibbonGroup>

            {/* Copilot Group */}
            <RibbonGroup title="Copilot">
              <div className="flex flex-col items-center">
                <Button 
                  variant="ghost" 
                  size="sm" 
                  className="h-14 w-14 flex-col p-1 hover:bg-blue-50 relative"
                  onClick={() => setTaskPaneOpen(true)}
                >
<LargeCopilotIcon className="w-9 h-9 mb-1" />
                  <span className="text-xs">Copilot</span>
                </Button>
              </div>
            </RibbonGroup>

            {/* Test Group */}
            <RibbonGroup title="Test">
              <div className="flex flex-col items-center">
                <Dialog open={testDialogOpen} onOpenChange={setTestDialogOpen}>
                  <DialogTrigger asChild>
                    <Button variant="ghost" size="sm" className="h-14 w-14 flex-col p-1 hover:bg-gray-50">
                      <div className="text-2xl mb-1">🧪</div>
                      <span className="text-xs">Test</span>
                    </Button>
                  </DialogTrigger>
                  <DialogContent className="sm:max-w-[425px] bg-white border-0 rounded-3xl shadow-lg">
                    <DialogHeader className="px-6 py-4">
                      <DialogTitle className="text-[#127d42] text-lg font-medium">Find and Replace</DialogTitle>
                      <DialogDescription className="text-gray-600 text-sm">
                        Enter your search criteria below.
                      </DialogDescription>
                    </DialogHeader>
                    <div className="px-6 py-2">
                      <div className="space-y-4">
                        <div className="flex items-center space-x-3">
                          <Label htmlFor="find-input" className="text-gray-700 text-sm font-medium min-w-[80px]">
                            Find what:
                          </Label>
                          <Input
                            id="find-input"
                            value={testInputValue}
                            onChange={(e) => setTestInputValue(e.target.value)}
                            className="flex-1 border-gray-300 rounded-xl focus:border-[#127d42] focus:ring-0 focus:outline-none selection:bg-pink-200"
                            placeholder="Enter search text..."
                          />
                        </div>
                        <div className="flex items-center space-x-3">
                          <Label htmlFor="replace-input" className="text-gray-700 text-sm font-medium min-w-[80px]">
                            Replace with:
                          </Label>
                          <Input
                            id="replace-input"
                            className="flex-1 border-gray-300 rounded-xl focus:border-[#127d42] focus:ring-0 focus:outline-none selection:bg-pink-200"
                            placeholder="Enter replacement text..."
                          />
                        </div>
                      </div>
                    </div>
                    <DialogFooter className="px-6 pb-6 pt-4 flex justify-end space-x-3">
                      <Button 
                        type="button" 
                        variant="outline" 
                        onClick={handleTestCancel}
                        className="border-gray-300 text-gray-700 hover:bg-gray-50 rounded-xl"
                      >
                        Cancel
                      </Button>
                      <Button 
                        type="submit" 
                        onClick={handleTestSubmit}
                        className="bg-[#127d42] hover:bg-[#0f6937] text-white rounded-xl border-0"
                      >
                        Find All
                      </Button>
                      <Button 
                        type="submit" 
                        onClick={handleTestSubmit}
                        className="bg-[#127d42] hover:bg-[#0f6937] text-white rounded-xl border-0"
                      >
                        Replace All
                      </Button>
                    </DialogFooter>
                  </DialogContent>
                </Dialog>
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
                  <PasteDropdown 
                    onPaste={ribbonActions.paste}
                    onPasteWithCopilot={() => console.log('Paste with Copilot')}
                    onPasteValuesOnly={() => console.log('Paste Values Only')}
                    onPasteFormulasOnly={() => console.log('Paste Formulas Only')}
                    onPasteFormattingOnly={() => console.log('Paste Formatting Only')}
                    onPasteLinkToSource={() => console.log('Link to Source')}
                    onPasteKeepColumnWidth={() => console.log('Keep Source Column Width')}
                    onPasteTranspose={() => console.log('Transpose Rows and Columns')}
                  />
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
      {taskPaneOpen && (
        <div className="fixed top-0 right-0 w-80 h-full bg-white shadow-lg border-l border-gray-200">
          <div className="p-4">
            <h2 className="text-lg font-semibold text-gray-800">Copilot</h2>
            <p className="text-sm text-gray-600">This is the task pane content.</p>
          </div>
          <button 
            className="absolute top-2 right-2 text-gray-500 hover:text-gray-700"
            onClick={() => setTaskPaneOpen(false)}
          >
            ✕
          </button>
        </div>
      )}
    </div>
  );
};
