// SVG Icon Utilities for Excel Ribbon
// Converts img tags with SVG sources to inline SVG for CSS color support

export function convertSvgImages() {
  const svgImages = document.querySelectorAll('img[src$=".svg"]');

  for (let i = 0; i < svgImages.length; i++) {
    const img = svgImages[i] as HTMLImageElement;

    fetch(img.src)
      .then(response => response.text())
      .then(svgText => {
        const parser = new DOMParser();
        const svgDoc = parser.parseFromString(svgText, "image/svg+xml");
        const svgElement = svgDoc.querySelector("svg");

        if (svgElement) {
          // Copy classes from img to svg
          svgElement.className.baseVal = img.className;
          if (img.alt) svgElement.setAttribute("title", img.alt);
          
          // Copy any data attributes
          Array.from(img.attributes).forEach(attr => {
            if (attr.name.startsWith('data-')) {
              svgElement.setAttribute(attr.name, attr.value);
            }
          });

          // Replace img with inline svg
          img.parentNode?.replaceChild(svgElement, img);
        }
      })
      .catch(err => {
        console.log("Failed to load SVG:", err);
      });
  }
}

// Initialize SVG conversion when DOM is ready
export function initializeSvgIcons() {
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', convertSvgImages);
  } else {
    convertSvgImages();
  }
}

// For React components, call this after component updates
export function convertSvgImagesInContainer(container: HTMLElement) {
  const svgImages = container.querySelectorAll('img[src$=".svg"]');

  for (let i = 0; i < svgImages.length; i++) {
    const img = svgImages[i] as HTMLImageElement;

    fetch(img.src)
      .then(response => response.text())
      .then(svgText => {
        const parser = new DOMParser();
        const svgDoc = parser.parseFromString(svgText, "image/svg+xml");
        const svgElement = svgDoc.querySelector("svg");

        if (svgElement) {
          svgElement.className.baseVal = img.className;
          if (img.alt) svgElement.setAttribute("title", img.alt);
          
          Array.from(img.attributes).forEach(attr => {
            if (attr.name.startsWith('data-')) {
              svgElement.setAttribute(attr.name, attr.value);
            }
          });

          img.parentNode?.replaceChild(svgElement, img);
        }
      })
      .catch(err => {
        console.log("Failed to load SVG:", err);
      });
  }
}
