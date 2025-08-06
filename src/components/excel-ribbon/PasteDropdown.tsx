import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { 
  ClipboardIcon,
  PasteIcon,
  PasteWithCopilotIcon,
  PasteValuesIcon,
  PasteFormulasIcon,
  PasteFormattingIcon,
  PasteLinkIcon,
  PasteColumnWidthsIcon,
  PasteTransposeIcon,
  ChevronDownIcon
} from "./icons";

interface PasteDropdownProps {
  onPaste?: () => void;
  onPasteWithCopilot?: () => void;
  onPasteValuesOnly?: () => void;
  onPasteFormulasOnly?: () => void;
  onPasteFormattingOnly?: () => void;
  onPasteLinkToSource?: () => void;
  onPasteKeepColumnWidth?: () => void;
  onPasteTranspose?: () => void;
}

export const PasteDropdown = ({
  onPaste,
  onPasteWithCopilot,
  onPasteValuesOnly,
  onPasteFormulasOnly,
  onPasteFormattingOnly,
  onPasteLinkToSource,
  onPasteKeepColumnWidth,
  onPasteTranspose
}: PasteDropdownProps) => {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <div className="flex flex-col relative">
      {/* Main Paste Button */}
      <Button 
        variant="ghost" 
        size="sm" 
        className="h-12 w-14 flex-col p-1 hover:bg-blue-50" 
        onClick={() => {
          onPaste?.();
        }}
      >
        <ClipboardIcon className="w-9 h-9 mb-1" />
        <span className="text-xs">Paste</span>
      </Button>
      
      {/* Dropdown Arrow beneath the button */}
      <Popover open={isOpen} onOpenChange={setIsOpen}>
        <PopoverTrigger asChild>
          <Button 
            variant="ghost" 
            size="sm" 
            className="h-2 w-14 p-0 hover:bg-blue-50 flex items-center justify-center mt-2"
          >
            <ChevronDownIcon className="w-3 h-3" />
          </Button>
        </PopoverTrigger>
        <PopoverContent className="w-[280px] p-2 bg-white border border-gray-200 shadow-lg mt-4" side="bottom" align="start" sideOffset={8}>
          <div className="space-y-1">
            {/* Regular Paste Options */}
            <button 
              className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors"
              onClick={() => {
                onPaste?.();
                setIsOpen(false);
              }}
            >
              <PasteIcon className="w-4 h-4" />
              <span>Paste</span>
            </button>
            <button 
              className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors"
              onClick={() => {
                onPasteWithCopilot?.();
                setIsOpen(false);
              }}
            >
              <PasteWithCopilotIcon className="w-4 h-4" />
              <span>Paste with Copilot</span>
            </button>
            
            {/* Divider */}
            <div className="border-t border-gray-200 my-2"></div>
            
            {/* Paste Special Header */}
            <div className="px-3 py-1 text-xs font-medium text-gray-600">Paste Special</div>
            
            {/* Paste Special Options */}
            <button 
              className="w-full flex items-center justify-between px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors"
              onClick={() => {
                onPasteValuesOnly?.();
                setIsOpen(false);
              }}
            >
              <div className="flex items-center gap-3">
                <PasteValuesIcon className="w-4 h-4" />
                <span>Values Only</span>
              </div>
              <span className="text-xs text-gray-500">Ctrl+Shift+V</span>
            </button>
            <button 
              className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors"
              onClick={() => {
                onPasteFormulasOnly?.();
                setIsOpen(false);
              }}
            >
              <PasteFormulasIcon className="w-4 h-4" />
              <span>Formulas Only</span>
            </button>
            <button 
              className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors"
              onClick={() => {
                onPasteFormattingOnly?.();
                setIsOpen(false);
              }}
            >
              <PasteFormattingIcon className="w-4 h-4" />
              <span>Formatting Only</span>
            </button>
            <button 
              className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors"
              onClick={() => {
                onPasteLinkToSource?.();
                setIsOpen(false);
              }}
            >
              <PasteLinkIcon className="w-4 h-4" />
              <span>Link to Source</span>
            </button>
            <button 
              className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors"
              onClick={() => {
                onPasteKeepColumnWidth?.();
                setIsOpen(false);
              }}
            >
              <PasteColumnWidthsIcon className="w-4 h-4" />
              <span>Keep Source Column Width</span>
            </button>
            <button 
              className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors"
              onClick={() => {
                onPasteTranspose?.();
                setIsOpen(false);
              }}
            >
              <PasteTransposeIcon className="w-4 h-4" />
              <span>Transpose Rows and Columns</span>
            </button>
          </div>
        </PopoverContent>
      </Popover>
    </div>
  );
};
