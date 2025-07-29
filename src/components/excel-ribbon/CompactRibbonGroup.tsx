import { cn } from "@/lib/utils";
import { Button } from "@/components/ui/button";
import { ChevronDownIcon } from "./icons";

interface CompactRibbonGroupProps {
  title: string;
  icon: React.ReactNode;
  className?: string;
}

export const CompactRibbonGroup = ({ title, icon, className }: CompactRibbonGroupProps) => {
  return (
    <Button 
      variant="ribbon" 
      size="ribbon-lg" 
      className={cn("flex-col gap-1 min-w-[60px]", className)}
    >
      {icon}
      <span className="text-xs">{title}</span>
      <ChevronDownIcon className="w-2 h-2" />
    </Button>
  );
};