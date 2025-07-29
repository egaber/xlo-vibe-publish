import { Button } from "@/components/ui/button";
import { ChevronDown } from "lucide-react";
import { cn } from "@/lib/utils";

interface RibbonDropdownProps {
  children: React.ReactNode;
  className?: string;
  size?: "ribbon-sm" | "ribbon-md" | "ribbon-lg";
}

export const RibbonDropdown = ({ children, className, size = "ribbon-md" }: RibbonDropdownProps) => {
  return (
    <Button 
      variant="ribbon" 
      size={size}
      className={cn("pl-2 pr-1", className)}
    >
      {children}
      <ChevronDown className="ml-1" />
    </Button>
  );
};