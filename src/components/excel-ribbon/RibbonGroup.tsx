import { cn } from "@/lib/utils";

interface RibbonGroupProps {
  title: string;
  children: React.ReactNode;
  className?: string;
}

export const RibbonGroup = ({ title, children, className }: RibbonGroupProps) => {
  return (
    <div className={cn("flex flex-col border-r border-gray-200 min-w-fit", className)}>
      <div className="flex items-center gap-1 px-3 py-2 min-h-[68px] flex-1">
        {children}
      </div>
      <div className="px-2 py-1 text-xs text-gray-600 text-center">
        {title}
      </div>
    </div>
  );
};