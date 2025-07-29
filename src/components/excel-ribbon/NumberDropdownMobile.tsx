import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { HashIcon, ChevronDownIcon } from "./icons";

export const NumberDropdownMobile = () => {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <Popover open={isOpen} onOpenChange={setIsOpen}>
      <PopoverTrigger asChild>
        <Button variant="ribbon" size="ribbon-lg" className="flex-col gap-1 min-w-[60px] relative">
          <img src="/icons/responsive/large_number.png" alt="Number" className="w-10 h-10" />
          <span className="text-xs">Number</span>
          <ChevronDownIcon className="absolute bottom-0 right-0 transform scale-[0.3] origin-bottom-right" />
        </Button>
      </PopoverTrigger>
      <PopoverContent className="w-[160px] p-2 bg-white border border-gray-200 shadow-lg" align="start">
        <div className="space-y-1">
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span>General</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span>Number</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span>Currency</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span>Accounting</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span>Short Date</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span>Long Date</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span>Time</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span>Percentage</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span>Fraction</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span>Scientific</span>
          </button>
          <button className="w-full flex items-center gap-3 px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors">
            <span>Text</span>
          </button>
        </div>
      </PopoverContent>
    </Popover>
  );
};
