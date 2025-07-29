import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { SettingsIcon, ChevronDownIcon } from "./icons";

export const CellsDropdownMobile = () => {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <Popover open={isOpen} onOpenChange={setIsOpen}>
      <PopoverTrigger asChild>
        <Button variant="ribbon" size="ribbon-lg" className="flex-col gap-1 min-w-[60px] relative">
          <img src="/icons/responsive/large_cells.png" alt="Cells" className="w-10 h-10" />
          <span className="text-xs">Cells</span>
          <ChevronDownIcon className="absolute bottom-0 right-0 transform scale-[0.3] origin-bottom-right" />
        </Button>
      </PopoverTrigger>
      <PopoverContent className="w-[160px] p-2 bg-white border border-gray-200 shadow-lg" align="start">
        <div className="space-y-1">
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span className="w-4 h-4 flex items-center justify-center">➕</span>
            <span>Insert</span>
            <span className="ml-auto">▶</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span className="w-4 h-4 flex items-center justify-center">➖</span>
            <span>Delete</span>
            <span className="ml-auto">▶</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span className="w-4 h-4 flex items-center justify-center">🎨</span>
            <span>Format</span>
            <span className="ml-auto">▶</span>
          </button>
        </div>
      </PopoverContent>
    </Popover>
  );
};