import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { ChevronDownIcon } from "./icons";
import { cn } from "@/lib/utils";

const borderOptions = [
  { name: "Bottom Border", icon: "─" },
  { name: "Top Border", icon: "⎺" },
  { name: "Left Border", icon: "│" },
  { name: "Right Border", icon: "│" },
  { name: "No Border", icon: "□" },
  { name: "All Borders", icon: "⊞" },
  { name: "Outside Borders", icon: "⊡" },
  { name: "Thick Box Border", icon: "⊠" },
  { name: "Bottom Double Border", icon: "═" },
  { name: "Thick Bottom Border", icon: "━" },
  { name: "Top and Bottom Border", icon: "⫾" },
  { name: "Top and Thick Bottom Border", icon: "⫿" },
  { name: "Top and Double Bottom Border", icon: "╤" }
];

export const BorderDropdown = () => {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <Popover open={isOpen} onOpenChange={setIsOpen}>
      <PopoverTrigger asChild>
        <Button variant="ribbon" size="ribbon-icon" className="relative">
          <span className="text-sm">⊞</span>
          <ChevronDownIcon className="w-2 h-2 absolute bottom-0 right-0" />
        </Button>
      </PopoverTrigger>
      <PopoverContent className="w-[180px] p-2 bg-white border border-gray-200 shadow-lg" align="start">
        <div className="grid grid-cols-3 gap-1 mb-3">
          {borderOptions.slice(0, 9).map((border, index) => (
            <button
              key={index}
              className="flex items-center justify-center w-8 h-8 hover:bg-blue-50 rounded-sm transition-colors border border-gray-200"
              title={border.name}
            >
              <span className="text-sm">{border.icon}</span>
            </button>
          ))}
        </div>
        <div className="border-t border-gray-200 pt-2 space-y-1">
          {borderOptions.slice(9).map((border, index) => (
            <button
              key={index + 9}
              className="w-full text-left px-2 py-1 text-xs hover:bg-blue-50 rounded-sm transition-colors flex items-center gap-2"
            >
              <span>{border.icon}</span>
              <span>{border.name}</span>
            </button>
          ))}
        </div>
        <div className="border-t border-gray-200 pt-2 mt-2">
          <button className="w-full text-left px-2 py-1 text-xs hover:bg-blue-50 rounded-sm transition-colors">
            🎨 Draw Border
          </button>
          <button className="w-full text-left px-2 py-1 text-xs hover:bg-blue-50 rounded-sm transition-colors">
            ⊞ Draw Border Grid
          </button>
          <button className="w-full text-left px-2 py-1 text-xs hover:bg-blue-50 rounded-sm transition-colors">
            🗑️ Erase Border
          </button>
          <button className="w-full text-left px-2 py-1 text-xs hover:bg-blue-50 rounded-sm transition-colors">
            🎨 Border Color ▶
          </button>
          <button className="w-full text-left px-2 py-1 text-xs hover:bg-blue-50 rounded-sm transition-colors">
            ═ Border Style ▶
          </button>
          <button className="w-full text-left px-2 py-1 text-xs hover:bg-blue-50 rounded-sm transition-colors">
            📝 More Border Styles
          </button>
        </div>
      </PopoverContent>
    </Popover>
  );
};