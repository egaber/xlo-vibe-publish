import { useState, useRef, useEffect } from "react";
import { Button } from "@/components/ui/button";
import { 
  ChevronLeft, 
  ChevronRight, 
  Menu, 
  Plus, 
  Lock, 
  Eye, 
  EyeOff,
  MoreHorizontal 
} from "lucide-react";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import {
  ContextMenu,
  ContextMenuContent,
  ContextMenuItem,
  ContextMenuTrigger,
} from "@/components/ui/context-menu";

export interface Sheet {
  id: string;
  name: string;
  isProtected?: boolean;
  isVisible?: boolean;
}

interface SheetTabBarProps {
  sheets: Sheet[];
  activeSheetId: string;
  onSheetSelect: (sheetId: string) => void;
  onSheetRename: (sheetId: string, newName: string) => void;
  onSheetAdd: () => void;
  onSheetDelete: (sheetId: string) => void;
  onSheetDuplicate: (sheetId: string) => void;
  onSheetMove: (sheetId: string, direction: 'left' | 'right') => void;
  onSheetToggleProtection: (sheetId: string) => void;
  onSheetToggleVisibility: (sheetId: string) => void;
}

export const SheetTabBar = ({
  sheets,
  activeSheetId,
  onSheetSelect,
  onSheetRename,
  onSheetAdd,
  onSheetDelete,
  onSheetDuplicate,
  onSheetMove,
  onSheetToggleProtection,
  onSheetToggleVisibility,
}: SheetTabBarProps) => {
  const [editingSheetId, setEditingSheetId] = useState<string | null>(null);
  const [editValue, setEditValue] = useState("");
  const [canScrollLeft, setCanScrollLeft] = useState(false);
  const [canScrollRight, setCanScrollRight] = useState(false);
  const tabsContainerRef = useRef<HTMLDivElement>(null);
  const editInputRef = useRef<HTMLInputElement>(null);

  const checkScrollability = () => {
    if (tabsContainerRef.current) {
      const container = tabsContainerRef.current;
      const canScrollL = container.scrollLeft > 0;
      const canScrollR = container.scrollLeft < (container.scrollWidth - container.clientWidth);
      setCanScrollLeft(canScrollL);
      setCanScrollRight(canScrollR);
    }
  };

  useEffect(() => {
    checkScrollability();
    const container = tabsContainerRef.current;
    if (container) {
      container.addEventListener('scroll', checkScrollability);
      return () => container.removeEventListener('scroll', checkScrollability);
    }
  }, [sheets]);

  useEffect(() => {
    if (editingSheetId && editInputRef.current) {
      editInputRef.current.focus();
      editInputRef.current.select();
    }
  }, [editingSheetId]);

  const scrollTabs = (direction: 'left' | 'right') => {
    if (tabsContainerRef.current) {
      const scrollAmount = 120;
      const newScrollLeft = tabsContainerRef.current.scrollLeft + 
        (direction === 'left' ? -scrollAmount : scrollAmount);
      tabsContainerRef.current.scrollTo({ left: newScrollLeft, behavior: 'smooth' });
    }
  };

  const scrollToFirstSheet = () => {
    if (tabsContainerRef.current) {
      tabsContainerRef.current.scrollTo({ left: 0, behavior: 'smooth' });
    }
  };

  const scrollToLastSheet = () => {
    if (tabsContainerRef.current) {
      tabsContainerRef.current.scrollTo({ 
        left: tabsContainerRef.current.scrollWidth, 
        behavior: 'smooth' 
      });
    }
  };

  const handleTabDoubleClick = (sheet: Sheet) => {
    setEditingSheetId(sheet.id);
    setEditValue(sheet.name);
  };

  const handleRenameSubmit = () => {
    if (editingSheetId && editValue.trim()) {
      onSheetRename(editingSheetId, editValue.trim());
    }
    setEditingSheetId(null);
    setEditValue("");
  };

  const handleRenameCancel = () => {
    setEditingSheetId(null);
    setEditValue("");
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      handleRenameSubmit();
    } else if (e.key === 'Escape') {
      handleRenameCancel();
    }
  };

  return (
    <div className="fixed bottom-0 left-0 right-0 z-50 flex items-center bg-gray-100 border-t border-gray-300 h-8 min-h-8">
      {/* Navigation Area */}
      <div className="flex items-center bg-gray-100 border-r border-gray-300">
        {/* Scroll to First */}
        <Button
          variant="ghost"
          size="sm"
          className={`h-6 w-6 p-0 m-0.5 hover:bg-gray-200 ${!canScrollLeft ? 'opacity-50 cursor-not-allowed' : ''}`}
          onClick={scrollToFirstSheet}
          disabled={!canScrollLeft}
          title="Ctrl+click to scroll to the first sheet"
        >
          <ChevronLeft className="h-3 w-3" />
        </Button>

        {/* Scroll to Last */}
        <Button
          variant="ghost"
          size="sm"
          className={`h-6 w-6 p-0 m-0.5 hover:bg-gray-200 ${!canScrollRight ? 'opacity-50 cursor-not-allowed' : ''}`}
          onClick={scrollToLastSheet}
          disabled={!canScrollRight}
          title="Ctrl+click to scroll to the last sheet"
        >
          <ChevronRight className="h-3 w-3" />
        </Button>

        {/* All Sheets Dropdown */}
        <DropdownMenu>
          <DropdownMenuTrigger asChild>
            <Button
              variant="ghost"
              size="sm"
              className="h-6 w-6 p-0 m-0.5 hover:bg-gray-200"
              title="All Sheets"
            >
              <Menu className="h-3 w-3" />
            </Button>
          </DropdownMenuTrigger>
          <DropdownMenuContent align="start" className="w-48">
            {sheets.map((sheet) => (
              <DropdownMenuItem
                key={sheet.id}
                onClick={() => onSheetSelect(sheet.id)}
                className={`flex items-center gap-2 ${
                  sheet.id === activeSheetId ? 'bg-blue-50 font-semibold' : ''
                }`}
              >
                <div className="flex items-center gap-1 flex-1">
                  {sheet.isProtected && <Lock className="h-3 w-3 text-gray-500" />}
                  {sheet.isVisible === false && <EyeOff className="h-3 w-3 text-gray-500" />}
                  <span className="truncate">{sheet.name}</span>
                </div>
              </DropdownMenuItem>
            ))}
          </DropdownMenuContent>
        </DropdownMenu>
      </div>

      {/* Tab Content Area */}
      <div className="flex-1 flex items-center overflow-hidden">
        <div 
          ref={tabsContainerRef}
          className="flex items-center overflow-x-auto scrollbar-none"
          style={{ scrollbarWidth: 'none', msOverflowStyle: 'none' }}
        >
          {sheets.map((sheet, index) => {
            const isActive = sheet.id === activeSheetId;
            const isEditing = editingSheetId === sheet.id;

            return (
              <ContextMenu key={sheet.id}>
                <ContextMenuTrigger>
                  <div
                    className={`
                      relative flex items-center h-7 px-3 cursor-pointer select-none whitespace-nowrap bg-gray-100 hover:bg-gray-150
                      ${isActive 
                        ? 'font-bold text-gray-900 border-b-2 border-b-[#127d42]' 
                        : 'text-gray-700 border-b-2 border-b-transparent'
                      }
                    `}
                    onClick={() => onSheetSelect(sheet.id)}
                    onDoubleClick={() => handleTabDoubleClick(sheet)}
                  >
                    <div className="flex items-center gap-1 min-w-0">
                      {sheet.isProtected && (
                        <Lock className="h-3 w-3 text-gray-500 flex-shrink-0" />
                      )}
                      {sheet.isVisible === false && (
                        <EyeOff className="h-3 w-3 text-gray-500 flex-shrink-0" />
                      )}
                      
                      {isEditing ? (
                        <input
                          ref={editInputRef}
                          value={editValue}
                          onChange={(e) => setEditValue(e.target.value)}
                          onBlur={handleRenameSubmit}
                          onKeyDown={handleKeyDown}
                          className="bg-white border border-blue-500 rounded px-1 text-xs h-5 min-w-16 max-w-32"
                          style={{ fontSize: '11px' }}
                        />
                      ) : (
                        <span className="text-xs truncate max-w-24" style={{ fontSize: '11px' }}>
                          {sheet.name}
                        </span>
                      )}
                    </div>
                  </div>
                </ContextMenuTrigger>
                
                <ContextMenuContent>
                  <ContextMenuItem onClick={() => handleTabDoubleClick(sheet)}>
                    Rename
                  </ContextMenuItem>
                  <ContextMenuItem onClick={() => onSheetDuplicate(sheet.id)}>
                    Duplicate
                  </ContextMenuItem>
                  <ContextMenuItem 
                    onClick={() => onSheetDelete(sheet.id)}
                    disabled={sheets.length <= 1}
                  >
                    Delete
                  </ContextMenuItem>
                  <ContextMenuItem onClick={() => onSheetMove(sheet.id, 'left')}>
                    Move Left
                  </ContextMenuItem>
                  <ContextMenuItem onClick={() => onSheetMove(sheet.id, 'right')}>
                    Move Right
                  </ContextMenuItem>
                  <ContextMenuItem onClick={() => onSheetToggleProtection(sheet.id)}>
                    {sheet.isProtected ? 'Unprotect Sheet' : 'Protect Sheet'}
                  </ContextMenuItem>
                  <ContextMenuItem onClick={() => onSheetToggleVisibility(sheet.id)}>
                    {sheet.isVisible === false ? 'Show Sheet' : 'Hide Sheet'}
                  </ContextMenuItem>
                </ContextMenuContent>
              </ContextMenu>
            );
          })}
          
          {/* Add New Sheet Button in tab area */}
          <div
            className="flex items-center h-7 px-2 cursor-pointer select-none bg-gray-100 hover:bg-gray-200 text-gray-600"
            onClick={onSheetAdd}
            title="New sheet"
          >
            <Plus className="h-3 w-3" />
          </div>
        </div>
      </div>
    </div>
  );
};
