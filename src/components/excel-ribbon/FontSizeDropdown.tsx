import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { ChevronDownIcon } from "./icons";
import { cn } from "@/lib/utils";

interface FontSizeDropdownProps {
  selectedSize?: number;
  onSizeChange?: (size: number) => void;
}

const fontSizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 36, 40, 44, 48, 54, 60, 66, 72];

export const FontSizeDropdown = ({ selectedSize = 11, onSizeChange }: FontSizeDropdownProps) => {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <Popover open={isOpen} onOpenChange={setIsOpen}>
      <PopoverTrigger asChild>
        <Button variant="ribbon" size="ribbon-md" className="justify-between min-w-[50px]">
          <span className="text-xs">{selectedSize}</span>
          <ChevronDownIcon className="w-3 h-3" />
        </Button>
      </PopoverTrigger>
      <PopoverContent className="w-[80px] p-1 bg-white border border-gray-200 shadow-lg" align="start">
        <div className="max-h-[200px] overflow-y-auto">
          {fontSizes.map((size) => (
            <button
              key={size}
              className={cn(
                "w-full text-left px-3 py-1 text-sm hover:bg-blue-50 rounded-sm transition-colors",
                selectedSize === size && "bg-blue-100"
              )}
              onClick={() => {
                onSizeChange?.(size);
                setIsOpen(false);
              }}
            >
              {size}
            </button>
          ))}
        </div>
      </PopoverContent>
    </Popover>
  );
};