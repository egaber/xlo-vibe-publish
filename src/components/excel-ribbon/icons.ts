// Excel Ribbon Icons - Using custom PNG icons from the icons folder
import * as React from "react";

// Custom icon component for PNG images
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

// Export custom PNG icons
export const UndoIcon = createIconComponent("/icons/undo.png", "Undo");
export const RedoIcon = createIconComponent("/icons/redo.png", "Redo");
export const ClipboardIcon = createIconComponent("/icons/paste.png", "Clipboard");
export const CopyIcon = createIconComponent("/icons/paste small.png", "Copy");
export const CutIcon = createIconComponent("/icons/cut.png", "Cut");
export const FormatPainterIcon = createIconComponent("/icons/format painter.png", "Format Painter");
export const BoldIcon = createIconComponent("/icons/Bold.png", "Bold");
export const ItalicIcon = createIconComponent("/icons/italic.png", "Italic");
export const AlignLeftIcon = createIconComponent("/icons/align left.png", "Align Left");
export const AlignCenterIcon = createIconComponent("/icons/align center.png", "Align Center");
export const AlignRightIcon = createIconComponent("/icons/align text right.png", "Align Right");
export const AlignJustifyIcon = createIconComponent("/icons/align text center.png", "Align Justify");
export const PlusIcon = createIconComponent("/icons/insert.png", "Insert");
export const MinusIcon = createIconComponent("/icons/delete.png", "Delete");
export const PaletteIcon = createIconComponent("/icons/format.png", "Format");
export const SettingsIcon = createIconComponent("/icons/cell styles.png", "Cell Styles");
export const HashIcon = createIconComponent("/icons/format as table.png", "Format as Table");

// Fallback to Lucide icons for icons not available as PNG
export { Underline as UnderlineIcon } from "lucide-react";
export { Percent as PercentIcon } from "lucide-react";
export { Type as TypeIcon } from "lucide-react";
export { Search as SearchIcon } from "lucide-react";
export { Calculator as CalculatorIcon } from "lucide-react";
export { ChevronDown as ChevronDownIcon } from "lucide-react";
export { Save as SaveIcon } from "lucide-react";
export { Share as ShareIcon } from "lucide-react";
export { Eye as EyeIcon } from "lucide-react";
export { FileSpreadsheet as FileSpreadsheetIcon } from "lucide-react";

// Additional custom icons available
export const BordersIcon = createIconComponent("/icons/borders.png", "Borders");
export const ConditionalFormattingIcon = createIconComponent("/icons/conditonal formatting.png", "Conditional Formatting");
export const CopilotIcon = createIconComponent("/icons/copilot.png", "Copilot");
export const DecreaseIndentIcon = createIconComponent("/icons/decrease indent.png", "Decrease Indent");
export const IncreaseIndentIcon = createIconComponent("/icons/increase indent.png", "Increase Indent");
export const MergeIcon = createIconComponent("/icons/merge.png", "Merge");
export const OrientationIcon = createIconComponent("/icons/orentation.png", "Orientation");
export const AlignBottomIcon = createIconComponent("/icons/align bottom.png", "Align Bottom");
export const AlignTopIcon = createIconComponent("/icons/alignt top.png", "Align Top");
export const FillColorIcon = createIconComponent("/icons/fill color.png", "Fill Color");
export const DoubleUnderlineIcon = createIconComponent("/icons/doubleUnderline.png", "Double Underline");
export const StrikeThroughIcon = createIconComponent("/icons/ab.png", "Strikethrough");
export const FontIncrementIcon = createIconComponent("/icons/biggerFont.png", "Increase Font Size");
export const FontDecrementIcon = createIconComponent("/icons/smallerfont.png", "Decrease Font Size");
export const AlignTextLeftIcon = createIconComponent("/icons/align left.png", "Align Text Left");
export const AlignTextCenterIcon = createIconComponent("/icons/align text center.png", "Align Text Center");
export const AlignTextRightIcon = createIconComponent("/icons/align text right.png", "Align Text Right");
export const AlignMiddleIcon = createIconComponent("/icons/align center.png", "Align Middle");
export const AccountingIcon = createIconComponent("/icons/accounting.png", "Accounting");
export const CommaIcon = createIconComponent("/icons/comma.png", "Comma Style");
export const DecreaseDecimalIcon = createIconComponent("/icons/decreaseDecimal.png", "Decrease Decimal");
export const IncreaseDecimalIcon = createIconComponent("/icons/increaseDecimal.png", "Increase Decimal");

// Large icons for desktop mode
export const LargeSensitivityIcon = createIconComponent("/icons/responsive/large_sensitivty.png", "Sensitivity");
export const LargeCopilotIcon = createIconComponent("/icons/responsive/large_copilot.png", "Copilot");
