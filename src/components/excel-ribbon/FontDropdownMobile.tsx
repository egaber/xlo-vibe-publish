import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { 
  TypeIcon, 
  BoldIcon, 
  ItalicIcon, 
  UnderlineIcon,
  PlusIcon,
  MinusIcon,
  ChevronDownIcon
} from "./icons";

export const FontDropdownMobile = () => {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <Popover open={isOpen} onOpenChange={setIsOpen}>
      <PopoverTrigger asChild>
        <Button variant="ribbon" size="ribbon-lg" className="flex-col gap-1 min-w-[60px] relative">
          <img src="/icons/responsive/large_font.png" alt="Font" className="w-10 h-10" />
          <span className="text-xs">Font</span>
          <ChevronDownIcon className="absolute bottom-0 right-0 transform scale-[0.3] origin-bottom-right" />
        </Button>
      </PopoverTrigger>
      <PopoverContent className="w-[200px] p-2 bg-white border border-gray-200 shadow-lg" align="start">
        <div className="space-y-1">
          <div className="px-3 py-1 text-xs font-medium text-gray-600">Font Name</div>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span>Aptos Narrow...</span>
            <span className="ml-auto">▼</span>
          </button>
          <div className="px-3 py-1 text-xs font-medium text-gray-600 mt-3">Font Size</div>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span>11</span>
            <span className="ml-auto">▼</span>
          </button>
          <div className="border-t border-gray-200 my-2"></div>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <PlusIcon className="w-4 h-4" />
            <span>Grow Font</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <MinusIcon className="w-4 h-4" />
            <span>Shrink Font</span>
          </button>
          <div className="border-t border-gray-200 my-2"></div>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <BoldIcon className="w-4 h-4" />
            <span>Bold</span>
            <span className="ml-auto text-xs text-gray-500">Ctrl+B</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <ItalicIcon className="w-4 h-4" />
            <span>Italic</span>
            <span className="ml-auto text-xs text-gray-500">Ctrl+I</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <UnderlineIcon className="w-4 h-4" />
            <span>Underline</span>
            <span className="ml-auto text-xs text-gray-500">Ctrl+U</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span className="w-4 h-4 flex items-center justify-center">S̶</span>
            <span>Strikethrough</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span className="w-4 h-4 flex items-center justify-center">U̲</span>
            <span>Double Underline</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span className="w-4 h-4 flex items-center justify-center">⊞</span>
            <span>Borders</span>
            <span className="ml-auto">▶</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span className="w-4 h-4 bg-yellow-400 border border-gray-300"></span>
            <span>Fill Color</span>
            <span className="ml-auto">▶</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span className="w-4 h-4 text-red-600 font-bold">A</span>
            <span>Font Color</span>
            <span className="ml-auto">▶</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <TypeIcon className="w-4 h-4" />
            <span>More Font Styles</span>
          </button>
        </div>
      </PopoverContent>
    </Popover>
  );
};