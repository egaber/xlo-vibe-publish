import { useEffect, useCallback, RefObject, DependencyList } from 'react';
import { convertSvgImagesInContainer } from '@/utils/svgIconUtils';

/**
 * Custom hook to automatically convert SVG images to inline SVGs for colorization
 * Use this hook in components that contain SVG icons that need colorization
 */
export const useSvgIconConversion = (containerRef?: RefObject<HTMLElement>, dependencies: DependencyList = []) => {
  const convertSvgs = useCallback(() => {
    if (containerRef?.current) {
      // Convert SVGs within a specific container
      convertSvgImagesInContainer(containerRef.current);
    } else {
      // Convert all SVGs in the document
      import('@/utils/svgIconUtils').then(({ convertSvgImages }) => {
        convertSvgImages();
      });
    }
  }, [containerRef]);

  useEffect(() => {
    // Convert SVGs after component updates
    const timer = setTimeout(convertSvgs, 50);
    return () => clearTimeout(timer);
  }, [convertSvgs, dependencies]);

  return convertSvgs;
};
