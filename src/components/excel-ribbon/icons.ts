// Excel Ribbon Icons - Using SVG icons with colorization support
import * as React from "react";

// Custom icon component for SVG images with colorization support
const createSvgIconComponent = (iconPath: string, alt: string) => {
  return React.forwardRef<HTMLImageElement, React.ImgHTMLAttributes<HTMLImageElement>>(
    ({ className, ...props }, ref) => {
      return React.createElement("img", {
        ref,
        src: iconPath,
        alt,
        className: className || "w-3 h-3", // Smaller default size for better button fit
        ...props
      });
    }
  );
};

// Custom icon component for PNG images (fallback)
const createIconComponent = (iconPath: string, alt: string) => {
  return React.forwardRef<HTMLImageElement, React.ImgHTMLAttributes<HTMLImageElement>>(
    ({ className, ...props }, ref) => {
      return React.createElement("img", {
        ref,
        src: iconPath,
        alt,
        className: className || "w-4 h-4", // Use provided className or default to w-4 h-4
        ...props
      });
    }
  );
};

// Main SVG icons with colorization support
export const BoldIcon = createSvgIconComponent("/icons/svg/bold.16.svg", "Bold");
export const ItalicIcon = createSvgIconComponent("/icons/svg/italic.16.svg", "Italic");
export const UnderlineIcon = createSvgIconComponent("/icons/svg/underline.16.svg", "Underline");
export const DoubleUnderlineIcon = createSvgIconComponent("/icons/svg/doubleunderline.16.svg", "Double Underline");
export const StrikeThroughIcon = createSvgIconComponent("/icons/svg/strikethrough2.16.svg", "Strikethrough");

// Alignment icons - using authentic XLO SVG icons (matching index.html)
export const AlignLeftIcon = createSvgIconComponent("/icons/svg/aligntablecelltopleft.16.svg", "Align Left");
export const AlignCenterIcon = createSvgIconComponent("/icons/svg/aligntablecelltopcenter.16.svg", "Align Center");
export const AlignRightIcon = createSvgIconComponent("/icons/svg/aligntablecelltopright.16.svg", "Align Right");
export const AlignJustifyIcon = createSvgIconComponent("/icons/svg/aligncenter.16.svg", "Align Justify");
export const AlignTopIcon = createSvgIconComponent("/icons/svg/aligntablecelltopcenter.16.svg", "Align Top");
export const AlignMiddleIcon = createSvgIconComponent("/icons/svg/aligntablecellmiddlecenter.16.svg", "Align Middle");
export const AlignBottomIcon = createSvgIconComponent("/icons/svg/aligntablecellbottomcenter.16.svg", "Align Bottom");
export const AlignTextLeftIcon = createSvgIconComponent("/icons/svg/aligntablecelltopleft.16.svg", "Align Text Left");
export const AlignTextCenterIcon = createSvgIconComponent("/icons/svg/aligntablecelltopcenter.16.svg", "Align Text Center");
export const AlignTextRightIcon = createSvgIconComponent("/icons/svg/aligntablecelltopright.16.svg", "Align Text Right");

// Font size icons - matching index.html paths exactly
export const FontIncrementIcon = createSvgIconComponent("/icons/svg/growfont.16.svg", "Increase Font Size");
export const FontDecrementIcon = createSvgIconComponent("/icons/svg/shrinkfont.16.svg", "Decrease Font Size");

// Color and formatting icons
export const FillColorIcon = createSvgIconComponent("/icons/svg/fillcolorsplitdropdown.16.svg", "Fill Color");
export const FontColorIcon = createSvgIconComponent("/icons/svg/fontcolor.16.svg", "Font Color");
export const WrapTextIcon = createSvgIconComponent("/icons/svg/wraptext.16.svg", "Wrap Text");
export const IncreaseIndentIcon = createSvgIconComponent("/icons/svg/indent.16.svg", "Increase Indent");

// Clipboard and action icons - using 20px versions to match index.html
export const UndoIcon = createSvgIconComponent("/icons/svg/undo.20.svg", "Undo");
export const RedoIcon = createSvgIconComponent("/icons/svg/redo.20.svg", "Redo");
export const ClipboardIcon = createSvgIconComponent("/icons/svg/paste.20.svg", "Clipboard");
export const PasteIcon = createSvgIconComponent("/icons/svg/paste.20.svg", "Paste");
export const CopyIcon = createSvgIconComponent("/icons/svg/copy.20.svg", "Copy");
export const CutIcon = createSvgIconComponent("/icons/svg/cut.20.svg", "Cut");
export const FormatPainterIcon = createSvgIconComponent("/icons/svg/formatpainter.20.svg", "Format Painter");

// Paste Special icons
export const PasteWithCopilotIcon = createSvgIconComponent("/icons/svg/pastewithcopilot.16.svg", "Paste with Copilot");
export const PasteValuesIcon = createSvgIconComponent("/icons/svg/pastevalues.16.svg", "Paste Values Only");
export const PasteFormulasIcon = createSvgIconComponent("/icons/svg/pasteformulas.16.svg", "Paste Formulas Only");
export const PasteFormattingIcon = createSvgIconComponent("/icons/svg/pasteformatting.16.svg", "Paste Formatting Only");
export const PasteLinkIcon = createSvgIconComponent("/icons/svg/pastelink.16.svg", "Link to Source");
export const PasteColumnWidthsIcon = createSvgIconComponent("/icons/svg/pastewithcolumnwidths.16.svg", "Keep Source Column Width");
export const PasteTransposeIcon = createSvgIconComponent("/icons/svg/pastetranspose.16.svg", "Transpose Rows and Columns");
export const ConditionalFormattingIcon = createSvgIconComponent("/icons/svg/conditionalformatting.16.svg", "Conditional Formatting");
export const SaveIcon = createSvgIconComponent("/icons/svg/save.16.svg", "Save");
export const SearchIcon = createSvgIconComponent("/icons/svg/search.16.svg", "Search");

// Editing icons
export const AutoSumIcon = createSvgIconComponent("/icons/svg/autosum.16.svg", "AutoSum");
export const ClearIcon = createSvgIconComponent("/icons/svg/clear.16.svg", "Clear");
export const SortIcon = createSvgIconComponent("/icons/svg/sortup.16.svg", "Sort");
export const FilterIcon = createSvgIconComponent("/icons/svg/filter.16.svg", "Filter");

// Number formatting icons (using authentic XLO SVGs)
export const AccountingIcon = createSvgIconComponent("/icons/svg/accounting.16.svg", "Accounting");
export const CommaIcon = createSvgIconComponent("/icons/svg/commaformat.16.svg", "Comma Style");
export const PercentageIcon = createSvgIconComponent("/icons/svg/percentage.16.svg", "Percentage");
export const CurrencyIcon = createSvgIconComponent("/icons/svg/currencyformat.16.svg", "Currency");
export const DecreaseDecimalIcon = createSvgIconComponent("/icons/svg/decreasedecimals.16.svg", "Decrease Decimal");
export const IncreaseDecimalIcon = createSvgIconComponent("/icons/svg/increasedecimal.16.svg", "Increase Decimal");

// PNG icons (keeping for icons not yet available as SVG)
export const PlusIcon = createIconComponent("/icons/insert.png", "Insert");
export const MinusIcon = createIconComponent("/icons/delete.png", "Delete");
export const PaletteIcon = createIconComponent("/icons/format.png", "Format");
export const SettingsIcon = createIconComponent("/icons/cell styles.png", "Cell Styles");
export const HashIcon = createSvgIconComponent("/icons/svg/formatastable.16.svg", "Format as Table");
// Borders and merge icons (using authentic XLO SVGs)
export const BordersIcon = createSvgIconComponent("/icons/svg/bottomborder.16.svg", "Bottom Border");
export const MergeIcon = createSvgIconComponent("/icons/svg/mergecells.16.svg", "Merge Cells");
export const MergeCenterIcon = createSvgIconComponent("/icons/svg/mergecellsacross.16.svg", "Merge & Center");

// PNG icons (keeping for icons not yet available as SVG)
export const CopilotIcon = createIconComponent("/icons/copilot.png", "Copilot");
export const DecreaseIndentIcon = createIconComponent("/icons/decrease indent.png", "Decrease Indent");
export const OrientationIcon = createSvgIconComponent("/icons/svg/formatcellalignment.20.svg", "Orientation");

// Fallback to Lucide icons for icons not available as PNG or SVG
export { Percent as PercentIcon } from "lucide-react";
export { Type as TypeIcon } from "lucide-react";
export { Calculator as CalculatorIcon } from "lucide-react";
export { ChevronDown as ChevronDownIcon } from "lucide-react";
export { Share as ShareIcon } from "lucide-react";
export { Eye as EyeIcon } from "lucide-react";
export { FileSpreadsheet as FileSpreadsheetIcon } from "lucide-react";

// Large icons for desktop mode
export const LargeSensitivityIcon = createIconComponent("/icons/responsive/large_sensitivty.png", "Sensitivity");
export const LargeCopilotIcon = createIconComponent("/icons/responsive/large_copilot.png", "Copilot");
