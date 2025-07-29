import { cn } from "@/lib/utils";

interface RibbonTabProps {
  children: React.ReactNode;
  active?: boolean;
  onClick?: () => void;
  className?: string;
}

export const RibbonTab = ({ children, active = false, onClick, className }: RibbonTabProps) => {
  return (
    <button
      onClick={onClick}
      className={cn(
        "px-4 py-2 text-sm font-normal text-gray-700 border-b-2 transition-colors relative",
        "hover:bg-gray-100",
        active 
          ? "border-b-[#127d42] text-black font-bold rounded-t-lg" 
          : "border-b-transparent hover:border-b-gray-300",
        className
      )}
    >
      {children}
    </button>
  );
};