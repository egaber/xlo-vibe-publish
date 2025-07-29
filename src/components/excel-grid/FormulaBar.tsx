import { useState, useRef, useEffect } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";

interface FormulaBarProps {
  selectedCell: string;
  cellValue: string;
  onCellValueChange: (value: string) => void;
  onApplyFormula: (value?: string) => void;
  onFormulaEditStart?: () => void;
  onFormulaEditEnd?: () => void;
}

export const FormulaBar = ({ selectedCell, cellValue, onCellValueChange, onApplyFormula, onFormulaEditStart, onFormulaEditEnd }: FormulaBarProps) => {
  const [isEditing, setIsEditing] = useState(false);
  const [tempValue, setTempValue] = useState(cellValue);
  const [isExpanded, setIsExpanded] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    setTempValue(cellValue);
  }, [cellValue]);

  const handleStartEdit = () => {
    setIsEditing(true);
    onFormulaEditStart?.();
    setTimeout(() => {
      if (isExpanded) {
        textareaRef.current?.focus();
      } else {
        inputRef.current?.focus();
      }
    }, 0);
  };

  const handleSubmit = () => {
    onCellValueChange(tempValue);
    setIsEditing(false);
    onFormulaEditEnd?.();
    onApplyFormula(tempValue);
  };

  const handleCancel = () => {
    setTempValue(cellValue);
    setIsEditing(false);
    onFormulaEditEnd?.();
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSubmit();
    } else if (e.key === 'Escape') {
      handleCancel();
    }
  };

  const toggleExpanded = () => {
    setIsExpanded(!isExpanded);
  };

  const insertFunction = () => {
    // Placeholder for function dialog
    setTempValue(tempValue + "=SUM()");
    handleStartEdit();
  };

  return (
    <div className="flex flex-col bg-gray-100 border-b border-gray-300">
      <div className="flex items-center px-2 py-1.5 gap-2">
        {/* Name Box */}
        <div className="flex items-center">
          <div className="flex items-center border border-gray-400 rounded bg-white">
            <input
              type="text"
              value={selectedCell}
              readOnly
              className="w-20 h-6 px-2 text-xs bg-white font-medium text-gray-700 cursor-default border-0 rounded-l"
              style={{ fontFamily: 'inherit' }}
            />
            <button
              type="button"
              className="h-6 w-5 flex items-center justify-center bg-white hover:bg-gray-50 rounded-r"
              aria-label="Name Box dropdown"
            >
              <svg width="8" height="4" viewBox="0 0 8 4" className="text-gray-600">
                <path
                  d="M0 0L4 4L8 0"
                  fill="currentColor"
                />
              </svg>
            </button>
          </div>
        </div>

        {/* Control Buttons */}
        <div className="flex items-center border border-gray-400 rounded bg-white">
          <button
            type="button"
            className={`h-6 w-6 flex items-center justify-center hover:bg-gray-50 rounded-l ${
              !isEditing ? 'opacity-50 cursor-not-allowed' : ''
            }`}
            onClick={handleCancel}
            disabled={!isEditing}
            aria-label="Cancel edit"
          >
            <span className={`text-sm font-bold ${isEditing ? 'text-red-600' : 'text-gray-400'}`}>✕</span>
          </button>
          
          <button
            type="button"
            className={`h-6 w-6 flex items-center justify-center hover:bg-gray-50 ${
              !isEditing ? 'opacity-50 cursor-not-allowed' : ''
            }`}
            onClick={handleSubmit}
            disabled={!isEditing}
            aria-label="Commit edit"
          >
            <span className={`text-sm font-bold ${isEditing ? 'text-green-600' : 'text-gray-400'}`}>✓</span>
          </button>
          
          <button
            type="button"
            className="h-6 w-8 flex items-center justify-center hover:bg-gray-50 rounded-r"
            onClick={insertFunction}
            aria-label="Insert function"
            title="Insert function"
          >
            <div className="flex items-center gap-0.5">
              <span className="text-xs font-bold text-gray-700 italic">fx</span>
              <svg width="6" height="3" viewBox="0 0 6 3" className="text-gray-600">
                <path d="M0 0L3 3L6 0" fill="currentColor" />
              </svg>
            </div>
          </button>
        </div>

        {/* Expand Button */}
        <button
          type="button"
          className="h-5 w-4 flex items-center justify-center hover:bg-gray-200 rounded ml-1"
          onClick={toggleExpanded}
          aria-label={isExpanded ? "Collapse formula bar" : "Expand formula bar"}
        >
          <svg 
            width="12" 
            height="12" 
            viewBox="0 0 12 12" 
            className={`text-gray-600 transition-transform ${isExpanded ? 'rotate-180' : ''}`}
          >
            <path d="M6 9L1.5 4.5L3 3L6 6L9 3L10.5 4.5L6 9Z" fill="currentColor" />
          </svg>
        </button>

        {/* Formula Input Area */}
        <div className="flex-1 min-w-0 flex items-center gap-1">
          <div className="flex-1">
            {isEditing ? (
              isExpanded ? (
                <textarea
                  ref={textareaRef}
                  value={tempValue}
                  onChange={(e) => setTempValue(e.target.value)}
                  onKeyDown={handleKeyDown}
                  onBlur={handleSubmit}
                  className="w-full h-20 p-2 text-sm border-2 border-black rounded bg-white resize-none focus:outline-none"
                  placeholder="Enter a value or formula"
                  style={{ fontFamily: 'inherit' }}
                />
              ) : (
                <div className="border-2 border-black rounded bg-white h-8">
                  <input
                    ref={inputRef}
                    type="text"
                    value={tempValue}
                    onChange={(e) => setTempValue(e.target.value)}
                    onKeyDown={handleKeyDown}
                    onBlur={handleSubmit}
                    className="w-full h-full px-2 text-sm border-0 focus:outline-none bg-white rounded"
                    placeholder="Enter a value or formula"
                    style={{ fontFamily: 'inherit' }}
                  />
                </div>
              )
            ) : (
              <div className="border border-gray-400 rounded bg-white h-8">
                <div
                  className="w-full h-full px-2 text-sm cursor-text flex items-center hover:bg-gray-50 rounded"
                  onClick={handleStartEdit}
                  style={{ fontFamily: 'inherit' }}
                >
                  {cellValue || <span className="text-gray-400">Enter a value or formula</span>}
                </div>
              </div>
            )}
          </div>
          
          {/* Formula Expand Button */}
          <button
            type="button"
            className="h-6 w-4 flex items-center justify-center hover:bg-gray-200 rounded"
            onClick={() => {
              setIsExpanded(!isExpanded);
              if (!isEditing) {
                handleStartEdit();
              }
            }}
            aria-label={isExpanded ? "Collapse formula input" : "Expand formula input to 3 rows"}
            title={isExpanded ? "Collapse formula input" : "Expand formula input to 3 rows"}
          >
            <svg 
              width="8" 
              height="4" 
              viewBox="0 0 8 4" 
              className={`text-gray-600 transition-transform ${isExpanded ? 'rotate-180' : ''}`}
            >
              <path d="M0 0L4 4L8 0" fill="currentColor" />
            </svg>
          </button>
        </div>
      </div>
      
      {/* Expanded Formula Area */}
      {isExpanded && !isEditing && (
        <div className="px-2 pb-2">
          <div className="border border-gray-400 rounded bg-white">
            <div
              className="w-full h-20 p-2 text-sm cursor-text hover:bg-gray-50 rounded flex items-start"
              onClick={handleStartEdit}
              style={{ fontFamily: 'inherit' }}
            >
              {cellValue || <span className="text-gray-400">Enter a value or formula</span>}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
