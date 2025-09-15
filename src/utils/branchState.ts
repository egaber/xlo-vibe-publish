/**
 * Branch State Management Utility
 * Handles serialization, storage, and URL generation for shareable Excel states
 */

import { CellData } from "@/types/cellTypes";

// Types for branch state management
export interface BranchState {
  id: string;
  timestamp: number;
  name: string;
  sheets: Array<{
    id: string;
    name: string;
    isProtected: boolean;
    isVisible: boolean;
  }>;
  sheetDataMap: Record<string, {
    cellData: Record<string, CellData>;
    selectedCell: string;
    selectedCellValue: string;
  }>;
  activeSheetId: string;
  columnWidths: number[];
  metadata: {
    createdBy: string;
    description: string;
    version: string;
  };
}

export interface ShareableLink {
  url: string;
  branchId: string;
  shortCode?: string;
}

/**
 * Generate a unique branch ID
 */
export function generateBranchId(): string {
  const timestamp = Date.now();
  const random = Math.random().toString(36).substring(2, 15);
  return `branch_${timestamp}_${random}`;
}

/**
 * Generate a short code for easier sharing
 */
export function generateShortCode(): string {
  return Math.random().toString(36).substring(2, 8).toUpperCase();
}

/**
 * Serialize the current application state into a BranchState
 */
export function serializeCurrentState(
  sheets: Array<{id: string; name: string; isProtected: boolean; isVisible: boolean}>,
  sheetDataMap: Record<string, {
    cellData: Record<string, CellData>;
    selectedCell: string;
    selectedCellValue: string;
  }>,
  activeSheetId: string,
  columnWidths: number[],
  description?: string
): BranchState {
  const branchId = generateBranchId();
  
  return {
    id: branchId,
    timestamp: Date.now(),
    name: `Branch ${new Date().toLocaleString()}`,
    sheets: sheets.map(sheet => ({
      id: sheet.id,
      name: sheet.name,
      isProtected: sheet.isProtected,
      isVisible: sheet.isVisible
    })),
    sheetDataMap,
    activeSheetId,
    columnWidths: [...columnWidths],
    metadata: {
      createdBy: 'User',
      description: description || 'Shared Excel state',
      version: '1.0.0'
    }
  };
}

/**
 * Compress state data using JSON + base64
 */
export function compressState(state: BranchState): string {
  try {
    const jsonString = JSON.stringify(state);
    const base64 = btoa(unescape(encodeURIComponent(jsonString)));
    return base64;
  } catch (error) {
    console.error('Failed to compress state:', error);
    throw new Error('State compression failed');
  }
}

/**
 * Decompress state data
 */
export function decompressState(compressedData: string): BranchState {
  try {
    const jsonString = decodeURIComponent(escape(atob(compressedData)));
    const state = JSON.parse(jsonString) as BranchState;
    return state;
  } catch (error) {
    console.error('Failed to decompress state:', error);
    throw new Error('State decompression failed');
  }
}

/**
 * Store branch state in localStorage
 */
export function storeBranchState(state: BranchState): void {
  try {
    const key = `excel_branch_${state.id}`;
    const compressed = compressState(state);
    localStorage.setItem(key, compressed);
    
    // Also store in an index for management
    const indexKey = 'excel_branch_index';
    const existingIndex = localStorage.getItem(indexKey);
    const index = existingIndex ? JSON.parse(existingIndex) : [];
    
    index.push({
      id: state.id,
      name: state.name,
      timestamp: state.timestamp,
      description: state.metadata.description
    });
    
    // Keep only the last 50 branches
    const sortedIndex = index.sort((a: {timestamp: number}, b: {timestamp: number}) => b.timestamp - a.timestamp).slice(0, 50);
    localStorage.setItem(indexKey, JSON.stringify(sortedIndex));
  } catch (error) {
    console.error('Failed to store branch state:', error);
    throw new Error('Branch state storage failed');
  }
}

/**
 * Retrieve branch state from localStorage
 */
export function retrieveBranchState(branchId: string): BranchState | null {
  try {
    const key = `excel_branch_${branchId}`;
    const compressed = localStorage.getItem(key);
    
    if (!compressed) {
      return null;
    }
    
    return decompressState(compressed);
  } catch (error) {
    console.error('Failed to retrieve branch state:', error);
    return null;
  }
}

/**
 * Generate a shareable URL for the current state
 */
export function generateShareableLink(state: BranchState): ShareableLink {
  const shortCode = generateShortCode();
  
  // Store the state first
  storeBranchState(state);
  
  // Create URL with branch ID and short code
  const baseUrl = window.location.origin + window.location.pathname;
  const url = `${baseUrl}?branch=${state.id}&code=${shortCode}`;
  
  return {
    url,
    branchId: state.id,
    shortCode
  };
}

/**
 * Check if URL contains branch parameters
 */
export function getBranchFromUrl(): {branchId: string; code: string} | null {
  const urlParams = new URLSearchParams(window.location.search);
  const branchId = urlParams.get('branch');
  const code = urlParams.get('code');
  
  if (branchId && code) {
    return { branchId, code };
  }
  
  return null;
}

/**
 * Clean up old branch states (keep last 30 days)
 */
export function cleanupOldBranches(): void {
  try {
    const indexKey = 'excel_branch_index';
    const existingIndex = localStorage.getItem(indexKey);
    
    if (!existingIndex) return;
    
    const index = JSON.parse(existingIndex);
    const thirtyDaysAgo = Date.now() - (30 * 24 * 60 * 60 * 1000);
    
    const validBranches = [];
    
    for (const branch of index) {
      if (branch.timestamp > thirtyDaysAgo) {
        validBranches.push(branch);
      } else {
        // Remove old branch data
        localStorage.removeItem(`excel_branch_${branch.id}`);
      }
    }
    
    localStorage.setItem(indexKey, JSON.stringify(validBranches));
  } catch (error) {
    console.error('Failed to cleanup old branches:', error);
  }
}

/**
 * Get list of all stored branches
 */
export function getBranchList(): Array<{id: string; name: string; timestamp: number; description: string}> {
  try {
    const indexKey = 'excel_branch_index';
    const existingIndex = localStorage.getItem(indexKey);
    
    if (!existingIndex) return [];
    
    return JSON.parse(existingIndex);
  } catch (error) {
    console.error('Failed to get branch list:', error);
    return [];
  }
}