import React, { useState } from 'react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Star, Settings, Search, Cloud, Shield } from 'lucide-react';

const ExcelTopBar = () => {
  const [documentTitle, setDocumentTitle] = useState('Book44.xlsx');
  const [isFavorited, setIsFavorited] = useState(false);

  return (
    <div className="flex items-center bg-gray-50 px-4 py-2 min-h-12">
      {/* Left Section - Logo, Branding and Document Title */}
      <div className="flex items-center gap-3 flex-1 min-w-0">
        {/* Waffle Menu Icon */}
        <Button variant="ghost" size="sm" className="p-1 h-8 w-8">
          <svg className="w-5 h-5 text-gray-600" viewBox="0 0 18 18" fill="currentColor">
            {/* Row 1 */}
            <circle cx="3" cy="3" r="1.5" />
            <circle cx="9" cy="3" r="1.5" />
            <circle cx="15" cy="3" r="1.5" />
            {/* Row 2 */}
            <circle cx="3" cy="9" r="1.5" />
            <circle cx="9" cy="9" r="1.5" />
            <circle cx="15" cy="9" r="1.5" />
            {/* Row 3 */}
            <circle cx="3" cy="15" r="1.5" />
            <circle cx="9" cy="15" r="1.5" />
            <circle cx="15" cy="15" r="1.5" />
          </svg>
        </Button>
        
        {/* Excel Logo using favicon */}
        <div className="flex items-center gap-2">
          <img src="/FavIcon_Excel.ico" alt="Excel" className="w-6 h-6" />
          <span className="bg-yellow-400 text-black text-xs px-1 rounded font-semibold">DF</span>
        </div>

        {/* Document Title */}
        <div className="flex items-center gap-1">
          <Input
            value={documentTitle}
            onChange={(e) => setDocumentTitle(e.target.value)}
            className="border-none shadow-none font-medium bg-transparent text-sm w-24 px-1"
            placeholder="Document name"
          />
          
          <Button
            variant="ghost"
            size="sm"
            onClick={() => setIsFavorited(!isFavorited)}
            className="p-1 h-6 w-6"
          >
            <Star 
              className={`w-4 h-4 ${isFavorited ? 'fill-yellow-400 text-yellow-400' : 'text-gray-400'}`} 
            />
          </Button>

          {/* Sensitivity Badge */}
          <Button variant="ghost" size="sm" className="p-1 h-6 w-6 text-orange-500">
            <Shield className="w-4 h-4" />
          </Button>

          {/* Save Status */}
          <Button variant="ghost" size="sm" className="p-1 h-6 w-6 text-green-600">
            <Cloud className="w-4 h-4" />
          </Button>
        </div>
      </div>

      {/* Center Section - Search Box */}
      <div className="flex justify-center flex-1">
        <div className="relative max-w-md w-full">
          <Search className="w-4 h-4 absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400" />
          <Input
            placeholder="Search for tools, help, and more (Alt + Q)"
            className="pl-10 pr-4 py-1 text-sm bg-gray-50 border-gray-200 rounded-md"
          />
        </div>
      </div>

      {/* Right Section - Settings and Profile */}
      <div className="flex items-center gap-2 flex-1 justify-end">
        {/* Settings */}
        <Button variant="ghost" size="sm" className="p-2 h-8 w-8">
          <Settings className="w-4 h-4 text-gray-600" />
        </Button>

        {/* Profile Picture */}
        <div className="w-8 h-8 rounded-full bg-[#127d42] flex items-center justify-center text-white text-sm font-medium">
          EG
        </div>
      </div>
    </div>
  );
};

export default ExcelTopBar;
