import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { ChevronDownIcon } from "./icons";
import { cn } from "@/lib/utils";

interface FontDropdownProps {
  selectedFont?: string;
  onFontChange?: (font: string) => void;
}

const fonts = [
  "Aptos Narrow",
  "Arial", 
  "Calibri",
  "Times New Roman",
  "Helvetica",
  "Verdana",
  "Georgia",
  "Comic Sans MS",
  "Impact",
  "Courier New"
];

export const FontDropdown = ({ selectedFont = "Aptos Narrow", onFontChange }: FontDropdownProps) => {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <Popover open={isOpen} onOpenChange={setIsOpen}>
      <PopoverTrigger asChild>
        <Button variant="ribbon" size="ribbon-md" className="justify-between min-w-[120px]">
          <span className="text-xs truncate">{selectedFont}</span>
          <ChevronDownIcon className="w-3 h-3" />
        </Button>
      </PopoverTrigger>
      <PopoverContent className="w-[200px] p-1 bg-white border border-gray-200 shadow-lg" align="start">
        <div className="max-h-[200px] overflow-y-auto">
          {fonts.map((font) => (
            <button
              key={font}
              className={cn(
                "w-full text-left px-3 py-2 text-sm hover:bg-blue-50 rounded-sm transition-colors",
                selectedFont === font && "bg-blue-100"
              )}
              style={{ fontFamily: font }}
              onClick={() => {
                onFontChange?.(font);
                setIsOpen(false);
              }}
            >
              {font}
            </button>
          ))}
        </div>
      </PopoverContent>
    </Popover>
  );
};