import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { ChevronDownIcon, FillColorIcon } from "./icons";
import { cn } from "@/lib/utils";

const themeColors = [
  "#FFFFFF", "#000000", "#E7E6E6", "#44546A", "#5B9BD5", "#70AD47", "#A5A5A5", "#FFC000", "#F79646", "#C55A11"
];

const standardColors = [
  "#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050", "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0"
];

interface ColorPickerProps {
  selectedColor?: string;
  onColorChange?: (color: string) => void;
  defaultColor?: string;
  type: "background" | "text";
}

export const ColorPicker = ({ selectedColor, onColorChange, defaultColor, type }: ColorPickerProps) => {
  const [isOpen, setIsOpen] = useState(false);
  const [currentColor, setCurrentColor] = useState(selectedColor || defaultColor);

  const handleColorSelect = (color: string) => {
    setCurrentColor(color);
    onColorChange?.(color);
    setIsOpen(false);
  };

  return (
    <div className="flex flex-col items-start">
      <Popover open={isOpen} onOpenChange={setIsOpen}>
        <PopoverTrigger asChild>
          <Button variant="ghost" size="sm" className="h-6 w-14 p-0 hover:bg-blue-50 flex items-center justify-start pl-1 relative">
            {type === "background" ? 
              <FillColorIcon className="w-5 h-5" /> : 
              <span className="text-s text-gray-500">A</span>
            }
            <ChevronDownIcon className="w-3 h-3 absolute bottom-0 right-3" />
          </Button>
        </PopoverTrigger>
        <PopoverContent className="w-[240px] p-3 bg-white border border-gray-200 shadow-lg" align="start">
          <div className="space-y-3">
            {/* Automatic Color */}
            <div>
              <button
                className="flex items-center gap-2 w-full p-2 hover:bg-blue-50 rounded-sm transition-colors"
                onClick={() => handleColorSelect("automatic")}
              >
                <div className="w-4 h-4 border border-gray-300 bg-white"></div>
                <span className="text-sm">Automatic</span>
              </button>
            </div>

            {/* Theme Colors */}
            <div>
              <div className="text-xs font-medium mb-2">Theme Colors</div>
              <div className="grid grid-cols-10 gap-1">
                {themeColors.map((color, index) => (
                  <button
                    key={index}
                    className="w-4 h-4 border border-gray-300 hover:scale-110 transition-transform"
                    style={{ backgroundColor: color }}
                    onClick={() => handleColorSelect(color)}
                  />
                ))}
              </div>
              {/* Tints of theme colors */}
              {[0.8, 0.6, 0.4, 0.2].map((opacity) => (
                <div key={opacity} className="grid grid-cols-10 gap-1 mt-1">
                  {themeColors.map((color, index) => (
                    <button
                      key={`${opacity}-${index}`}
                      className="w-4 h-4 border border-gray-300 hover:scale-110 transition-transform"
                      style={{ 
                        backgroundColor: color === "#FFFFFF" ? `rgba(0,0,0,${1-opacity})` : `${color}${Math.round(opacity * 255).toString(16).padStart(2, '0')}`
                      }}
                      onClick={() => handleColorSelect(`${color}${Math.round(opacity * 255).toString(16).padStart(2, '0')}`)}
                    />
                  ))}
                </div>
              ))}
            </div>

            {/* Standard Colors */}
            <div>
              <div className="text-xs font-medium mb-2">Standard Colors</div>
              <div className="grid grid-cols-10 gap-1">
                {standardColors.map((color, index) => (
                  <button
                    key={index}
                    className="w-4 h-4 border border-gray-300 hover:scale-110 transition-transform"
                    style={{ backgroundColor: color }}
                    onClick={() => handleColorSelect(color)}
                  />
                ))}
              </div>
            </div>

            {/* More Colors */}
            <div className="border-t border-gray-200 pt-2">
              <button className="w-full text-left px-2 py-1 text-xs hover:bg-blue-50 rounded-sm transition-colors flex items-center gap-2">
                🎨 More Colors...
              </button>
              <button className="w-full text-left px-2 py-1 text-xs hover:bg-blue-50 rounded-sm transition-colors flex items-center gap-2">
                💧 Eyedropper
              </button>
            </div>
          </div>
        </PopoverContent>
      </Popover>
      
      {/* Color indicator rectangle */}
      <div 
        className="w-5 h-2 mt-0 ml-1 border border-gray-300"
        style={{ backgroundColor: currentColor === "automatic" ? (type === "background" ? "#FFFF00" : "#FF0000") : currentColor }}
      />
    </div>
  );
};