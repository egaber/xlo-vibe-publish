import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { 
  UndoIcon, 
  RedoIcon, 
  ClipboardIcon, 
  CopyIcon, 
  CutIcon, 
  FormatPainterIcon 
} from "./icons";

export const ClipboardDropdown = () => {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <Popover open={isOpen} onOpenChange={setIsOpen}>
      <PopoverTrigger asChild>
        <Button variant="ribbon" size="ribbon-lg" className="flex-col gap-1 min-w-[60px]">
          <ClipboardIcon />
          <span className="text-xs">Clipboard</span>
        </Button>
      </PopoverTrigger>
      <PopoverContent className="w-[200px] p-2 bg-white border border-gray-200 shadow-lg" align="start">
        <div className="space-y-1">
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <UndoIcon className="w-4 h-4" />
            <span className="ml-auto text-xs text-gray-500">Ctrl+Z</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <RedoIcon className="w-4 h-4" />
            <span className="ml-auto text-xs text-gray-500">Ctrl+Y</span>
          </button>
          <div className="border-t border-gray-200 my-2"></div>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <ClipboardIcon className="w-4 h-4" />
            <span>Clipboard</span>
            <span className="ml-auto">▶</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <CutIcon className="w-4 h-4" />
            <span>Cut</span>
            <span className="ml-auto text-xs text-gray-500">Ctrl+X</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <CopyIcon className="w-4 h-4" />
            <span>Copy</span>
            <span className="ml-auto text-xs text-gray-500">Ctrl+C</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <ClipboardIcon className="w-4 h-4" />
            <span>Paste</span>
            <span className="ml-auto text-xs text-gray-500">Ctrl+V</span>
          </button>
          <div className="border-t border-gray-200 my-2"></div>
          <div className="px-3 py-1 text-xs font-medium text-gray-600">Paste Special</div>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span className="w-4 h-4 flex items-center justify-center">📋</span>
            <span>Values only</span>
            <span className="ml-auto text-xs text-gray-500">Ctrl+Shift+V</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors text-gray-400">
            <span className="w-4 h-4 flex items-center justify-center">📝</span>
            <span>Formulas only</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors text-gray-400">
            <span className="w-4 h-4 flex items-center justify-center">🎨</span>
            <span>Formatting only</span>
          </button>
        </div>
      </PopoverContent>
    </Popover>
  );
};