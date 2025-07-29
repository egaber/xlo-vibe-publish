import { Button } from "@/components/ui/button";

export const StatusBar = () => {
  return (
    <div className="container-367 css-422 bg-gray-50 border-t border-gray-300 px-2 py-1">
      <div className="flex items-center justify-between">
        {/* Left Section */}
        <div id="leftStatusBarRegion-ControlsGroup" className="left-370 flex items-center gap-2">
          <Button 
            variant="ghost" 
            size="sm" 
            className="h-6 px-2 text-xs hover:bg-gray-100"
            aria-label="Workbook Statistics"
          >
            <span className="text-xs">Workbook Statistics</span>
          </Button>
          
          <Button 
            variant="ghost" 
            size="sm" 
            className="h-6 px-2 text-xs hover:bg-gray-100"
            aria-label="Page load statistics"
          >
            <span className="text-xs">Page load statistics</span>
          </Button>
        </div>

        {/* Center Section */}
        <div id="centerStatusBarRegion-ControlsGroup" className="center-375 flex items-center gap-4">
          <div id="EnvironmentName">
            <div className="root-408">
              <div className="label-409 text-xs text-gray-600" aria-hidden="true">
                Inner Ring (Fastfood) : FIE1
              </div>
            </div>
          </div>
          
          <div id="PhaseStatus">
            <div className="root-408">
              <div className="label-409 text-xs text-gray-600" aria-hidden="true">
                Phase: unchanged, Time: 0ms, 
              </div>
            </div>
          </div>
          
          <div id="NetworkInformation">
            <div className="root-408">
              <div className="label-409 text-xs text-gray-600" aria-hidden="true">
                Network: 2.5 Mbps, 100 ms
              </div>
            </div>
          </div>
        </div>

        {/* Right Section */}
        <div id="rightStatusBarRegion-ControlsGroup" className="right-376 flex items-center gap-1">
          {/* Loading Add-ins Button */}
          <Button 
            variant="ghost" 
            size="sm" 
            className="h-6 px-2 text-xs hover:bg-gray-100 flex items-center gap-1"
            aria-label="Loading Add-ins"
          >
            <svg width="16" height="16" viewBox="0,0,2048,2048" className="w-3 h-3">
              <path
                d="M1152 109q25-15 51 0l717 410q26 15 26 44v819q0 30-26 45l-717 409q-25 15-51 0l-143-81q15-31 15-65v-45l154 88 665-380v-760l-665-380-666 380v329q-41 0-76 20l-26 15v-394q0-30 25-44l717-410z m-256 1127l-359-205q-25-15-50 0l-359 205q-26 15-26 44v410q0 30 26 44l359 205q25 15 50 0l359-205q26-15 26-44v-410q0-30-26-44z m-691 424v-350l307-176 307 176v350l-307 175-307-175z"
                fill="currentColor"
              />
            </svg>
            <span className="text-xs">Loading Add-ins</span>
          </Button>

          {/* Query Labels */}
          <div id="QueryAverageLabel" className="text-xs text-gray-600">
            Average: 2
          </div>
          
          <div id="QueryCountLabel" className="text-xs text-gray-600">
            Count: 4
          </div>
          
          <div id="QuerySumLabel" className="text-xs text-gray-600">
            Sum: 8
          </div>

          {/* Customize Status Bar Button */}
          <Button 
            variant="ghost" 
            size="sm" 
            className="h-6 w-6 p-0 hover:bg-gray-100"
            aria-label="Customize Status Bar"
          >
            <svg width="12" height="12" viewBox="0 0 2048 2048" className="w-3 h-3">
              <path d="M896 1589 165 859l182-182 549 550 549-550 182 182z" fill="currentColor" />
            </svg>
          </Button>

          {/* Keyboard Shortcuts Button */}
          <Button 
            variant="ghost" 
            size="sm" 
            className="h-6 w-6 p-0 hover:bg-gray-100"
            aria-label="Keyboard Shortcuts"
          >
            <svg width="16" height="16" viewBox="0,0,2048,2048" className="w-3 h-3">
              <path
                d="M384 1344q0-27 19-45 19-19 45-19h1152q27 0 45 19 19 19 19 45 0 27-19 45-19 19-45 19h-1152q-27 0-45-19-19-19-19-45z m32-448q40 0 68-28 28-28 28-68 0-40-28-68-28-28-68-28-40 0-68 28-28 28-28 68 0 40 28 68 28 28 68 28z m864-96q0 40-28 68-28 28-68 28-40 0-68-28-28-28-28-68 0-40 28-68 28-28 68-28 40 0 68 28 28 28 28 68z m-480 96q40 0 68-28 28-28 28-68 0-40-28-68-28-28-68-28-40 0-68 28-28 28-28 68 0 40 28 68 28 28 68 28z m864-96q0 40-28 68-28 28-68 28-40 0-68-28-28-28-28-68 0-40 28-68 28-28 68-28 40 0 68 28 28 28 28 68z m-992 352q40 0 68-28 28-28 28-68 0-40-28-68-28-28-68-28-40 0-68 28-28 28-28 68 0 40 28 68 28 28 68 28z m480-96q0 40-28 68-28 28-68 28-40 0-68-28-28-28-28-68 0-40 28-68 28-28 68-28 40 0 68 28 28 28 28 68z m288 96q40 0 68-28 28-28 28-68 0-40-28-68-28-28-68-28-40 0-68 28-28 28-28 68 0 40 28 68 28 28 68 28z m-1312-544q0-93 66-158 66-66 158-66h1344q93 0 158 66 66 66 66 158v832q0 93-66 158-66 66-158 66h-1344q-93 0-158-66-66-66-66-158v-832z m224-96q-40 0-68 28-28 28-28 68v832q0 40 28 68 28 28 68 28h1344q40 0 68-28 28-28 28-68v-832q0-40-28-68-28-28-68-28h-1344z"
                fill="currentColor"
              />
            </svg>
          </Button>

          {/* Help Button */}
          <Button 
            variant="ghost" 
            size="sm" 
            className="h-6 w-6 p-0 hover:bg-gray-100"
            aria-label="Help Resources"
          >
            <svg width="16" height="16" viewBox="0,0,2048,2048" className="w-3 h-3">
              <path
                d="M1024 256q209 0 388 105 174 102 275 275 105 179 105 388 0 209-105 388-102 174-275 275-179 105-388 105-209 0-388-105-174-102-275-275-105-179-105-388 0-209 105-388 102-174 275-275 179-105 388-105z m0 128q-174 0-323 87-145 85-230 230-87 149-87 323 0 174 87 323 85 145 230 230 149 87 323 87 174 0 323-87 145-85 230-230 87-149 87-323 0-174-87-323-85-145-230-230-149-87-323-87z m0 960q40 0 68 28 28 28 28 68 0 40-28 68-28 28-68 28-40 0-68-28-28-28-28-68 0-40 28-68 28-28 68-28z m0-768q106 0 181 75 75 75 75 181 0 68-22 117-21 45-75 102l-33 34q-36 38-49 64-13 27-13 67 0 27-19 45-19 19-45 19-27 0-45-19-19-19-19-45 0-68 22-117 21-45 75-102l33-34q36-38 49-64 13-27 13-67 0-53-37-91-37-37-91-37-53 0-91 37-37 37-37 91 0 27-19 45-19 19-45 19-27 0-45-19-19-19-19-45 0-106 75-181 75-75 181-75z"
                fill="currentColor"
              />
            </svg>
          </Button>

          {/* Feedback Button */}
          <Button 
            variant="ghost" 
            size="sm" 
            className="h-6 w-6 p-0 hover:bg-gray-100"
            aria-label="Help Improve Office"
          >
            <svg width="16" height="16" viewBox="0 0 16 16" className="w-3 h-3">
              <path
                d="M9.5 1A1.5 1.5 0 0 0 8 2.5v2a1.5 1.5 0 0 0 1 1.414V7a.5.5 0 0 0 .82.384L11.48 6h2.02A1.5 1.5 0 0 0 15 4.5v-2A1.5 1.5 0 0 0 13.5 1zM9 2.5a.5.5 0 0 1 .5-.5h4a.5.5 0 0 1 .5.5v2a.5.5 0 0 1-.5.5h-2.2a.5.5 0 0 0-.32.116l-.98.816V5.5a.5.5 0 0 0-.5-.5.5.5 0 0 1-.5-.5zM3 6a2 2 0 1 1 4 0 2 2 0 0 1-4 0m2-1a1 1 0 1 0 0 2 1 1 0 0 0 0-2M2.5 9h5A1.5 1.5 0 0 1 9 10.5c0 1.116-.459 2.01-1.212 2.615C7.047 13.71 6.053 14 5 14s-2.047-.29-2.788-.885C1.46 12.51 1 11.616 1 10.5A1.5 1.5 0 0 1 2.5 9m5 1h-5a.5.5 0 0 0-.5.5c0 .817.325 1.423.838 1.835C3.364 12.757 4.12 13 5 13s1.636-.243 2.162-.665C7.675 11.923 8 11.317 8 10.5a.5.5 0 0 0-.5-.5"
                fill="currentColor"
              />
            </svg>
          </Button>

          {/* Zoom Out Button */}
          <Button 
            variant="ghost" 
            size="sm" 
            className="h-6 w-6 p-0 hover:bg-gray-100"
            aria-label="Zoom Out"
          >
            <svg width="16" height="16" viewBox="0,0,2048,2048" className="w-3 h-3">
              <path d="M 1877 1024 h -1877 v -171 h 1877 z" fill="currentColor" />
            </svg>
          </Button>

          {/* Zoom ComboBox */}
          <div id="ZoomComboBox" className="flex items-center">
            <select className="text-xs border border-gray-300 rounded px-1 py-0.5 bg-white hover:bg-gray-50 focus:outline-none focus:ring-1 focus:ring-blue-500">
              <option value="50%">50%</option>
              <option value="75%">75%</option>
              <option value="100%" selected>100%</option>
              <option value="125%">125%</option>
              <option value="150%">150%</option>
              <option value="200%">200%</option>
            </select>
          </div>

          {/* Zoom In Button */}
          <Button 
            variant="ghost" 
            size="sm" 
            className="h-6 w-6 p-0 hover:bg-gray-100"
            aria-label="Zoom In"
          >
            <svg width="16" height="16" viewBox="0,0,2048,2048" className="w-3 h-3">
              <path d="M 1877 1024 h -853 v 853 h -171 v -853 h -853 v -171 h 853 v -853 h 171 v 853 h 853 z" fill="currentColor" />
            </svg>
          </Button>

          {/* Full Screen Button */}
          <Button 
            variant="ghost" 
            size="sm" 
            className="h-6 w-6 p-0 hover:bg-gray-100"
            aria-label="Enter Full Screen"
          >
            <svg width="16" height="16" viewBox="0,0,2048,2048" className="w-3 h-3">
              <path
                d="M480 384q-40 0-68 28-28 28-28 68v224q0 27-19 45-19 19-45 19-27 0-45-19-19-19-19-45v-224q0-93 66-158 66-66 158-66h224q27 0 45 19 19 19 19 45 0 27-19 45-19 19-45 19h-224z m800-64q0-27 19-45 19-19 45-19h224q93 0 158 66 66 66 66 158v224q0 27-19 45-19 19-45 19-27 0-45-19-19-19-19-45v-224q0-40-28-68-28-28-68-28h-224q-27 0-45-19-19-19-19-45z m-960 960q27 0 45 19 19 19 19 45v224q0 40 28 68 28 28 68 28h224q27 0 45 19 19 19 19 45 0 27-19 45-19 19-45 19h-224q-93 0-158-66-66-66-66-158v-224q0-27 19-45 19-19 45-19z m1408 0q27 0 45 19 19 19 19 45v224q0 93-66 158-66 66-158 66h-224q-27 0-45-19-19-19-19-45 0-27 19-45 19-19 45-19h224q40 0 68-28 28-28 28-68v-224q0-27 19-45 19-19 45-19z"
                fill="currentColor"
              />
            </svg>
          </Button>
        </div>
      </div>
    </div>
  );
};
