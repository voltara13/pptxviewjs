/**
 * SVG Renderer Module
 * Handles rendering of SVG vector graphics to Canvas
 * Supports inline SVG data and external SVG files for PptxGenJS compatibility
 */

// import { Logger } from '../utils/utils.js';

/**
 * SVG Renderer - Main class for rendering SVG to Canvas
 */
class SVGRenderer {
    constructor(graphics) {
        this.graphics = graphics;
        this.logger = new Logger('SVGRenderer');
        this.parser = new DOMParser();
    }

    /**
     * Render SVG content to Canvas
     * @param {string|Element} svgContent - SVG content as string or DOM element
     * @param {number} x - X position
     * @param {number} y - Y position
     * @param {number} width - Target width
     * @param {number} height - Target height
     * @param {Object} options - Rendering options
     * @return {Promise<void>} Promise that resolves when rendering is complete
     */
    async renderSVG(svgContent, x, y, width, height, options = {}) {
        try {
            let svgElement = null;

            // Parse SVG content
            if (typeof svgContent === 'string') {
                // Trim whitespace to handle content that starts with newlines/spaces
                const trimmedContent = svgContent.trim();
                
                if (trimmedContent.startsWith('<svg')) {
                    // Inline SVG
                    const svgDoc = this.parser.parseFromString(trimmedContent, 'image/svg+xml');
                    svgElement = svgDoc.documentElement;
                } else if (trimmedContent.startsWith('data:image/svg+xml')) {
                    // Base64 encoded SVG
                    const base64Data = trimmedContent.split(',')[1];
                    const decodedSVG = atob(base64Data);
                    const svgDoc = this.parser.parseFromString(decodedSVG, 'image/svg+xml');
                    svgElement = svgDoc.documentElement;
                } else if (trimmedContent.includes('<svg')) {
                    // Try parsing as XML that contains SVG
                    const svgDoc = this.parser.parseFromString(trimmedContent, 'image/svg+xml');
                    svgElement = svgDoc.documentElement;
                } else {
                    // Assume it's an SVG URL - render placeholder
                    this.renderSVGPlaceholder(x, y, width, height, 'External SVG');
                    return;
                }
            } else if (svgContent instanceof Element) {
                svgElement = svgContent;
            }

            if (!svgElement || svgElement.tagName !== 'svg') {
                this.renderSVGPlaceholder(x, y, width, height, 'Invalid SVG');
                return;
            }

            // Get SVG dimensions
            const svgViewBox = this.getSVGViewBox(svgElement);
            const svgDimensions = this.getSVGDimensions(svgElement, svgViewBox);

            // Calculate scaling
            const scaleX = width / svgDimensions.width;
            const scaleY = height / svgDimensions.height;
            const scale = options.preserveAspectRatio !== false ? 
                         Math.min(scaleX, scaleY) : Math.max(scaleX, scaleY);

            // Render SVG using different strategies
            if (this.canUseImageMethod()) {
                await this.renderSVGViaImage(svgElement, x, y, width, height, scale, options);
            } else {
                this.renderSVGViaParsing(svgElement, x, y, scale, options);
            }

        } catch (error) {
            this.logger.logError(this.constructor.name, 'Error rendering SVG:', error);
            this.renderSVGPlaceholder(x, y, width, height, 'SVG Error');
        }
    }

    /**
     * Get SVG viewBox information
     * @param {Element} svgElement - SVG element
     * @return {Object} ViewBox information
     */
    getSVGViewBox(svgElement) {
        const viewBox = svgElement.getAttribute('viewBox');
        if (viewBox) {
            const values = viewBox.split(/\s+/).map(v => parseFloat(v));
            if (values.length === 4) {
                return {
                    x: values[0],
                    y: values[1],
                    width: values[2],
                    height: values[3]
                };
            }
        }
        
        return null;
    }

    /**
     * Get SVG dimensions
     * @param {Element} svgElement - SVG element
     * @param {Object} viewBox - ViewBox information
     * @return {Object} Dimensions
     */
    getSVGDimensions(svgElement, viewBox) {
        let width = parseFloat(svgElement.getAttribute('width')) || 0;
        let height = parseFloat(svgElement.getAttribute('height')) || 0;

        // Use viewBox if no explicit dimensions
        if (!width && !height && viewBox) {
            width = viewBox.width;
            height = viewBox.height;
        }

        // Default dimensions if nothing found
        if (!width) {width = 100;}
        if (!height) {height = 100;}

        return { width, height };
    }

    /**
     * Check if Image-based rendering is available
     * @return {boolean} True if Image method is available
     */
    canUseImageMethod() {
        // Check if we can use the Image + drawImage method
        const hasImage = typeof Image !== 'undefined';
        const hasContext = this.graphics._context || this.graphics.m_oContext;
        const hasDrawImage = hasContext && typeof hasContext.drawImage === 'function';
        
        return hasImage && hasContext && hasDrawImage;
    }

    /**
     * Enhance SVG string for better rendering
     * @param {string} svgString - Original SVG string
     * @param {number} width - Target width
     * @param {number} height - Target height
     * @return {string} Enhanced SVG string
     */
    enhanceSVGForRendering(svgString, width, height) {
        try {
            // Parse the SVG to potentially modify it
            const doc = this.parser.parseFromString(svgString, 'image/svg+xml');
            const svgElement = doc.documentElement;
            
            // Ensure SVG has proper viewport
            if (!svgElement.hasAttribute('viewBox')) {
                const svgWidth = svgElement.getAttribute('width') || width;
                const svgHeight = svgElement.getAttribute('height') || height;
                svgElement.setAttribute('viewBox', `0 0 ${svgWidth} ${svgHeight}`);
            }
            
            // Ensure dimensions are set
            svgElement.setAttribute('width', width.toString());
            svgElement.setAttribute('height', height.toString());
            
            // Add namespace if missing
            if (!svgElement.hasAttribute('xmlns')) {
                svgElement.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
            }
            
            return new XMLSerializer().serializeToString(svgElement);
        } catch (error) {
            this.logger.log("warn", this.constructor.name, 'Could not enhance SVG, using original:', error);
            return svgString;
        }
    }

    /**
     * Render SVG via Image object (most compatible method)
     * @param {Element} svgElement - SVG element
     * @param {number} x - X position
     * @param {number} y - Y position
     * @param {number} width - Target width
     * @param {number} height - Target height
     * @param {number} scale - Scale factor
     * @param {Object} options - Rendering options
     * @return {Promise<void>} Promise that resolves when image is loaded and rendered
     */
    async renderSVGViaImage(svgElement, x, y, width, height, scale, options) {
        return new Promise((resolve, reject) => {
            try {
                // Convert SVG to data URL
                const serializer = new XMLSerializer();
                const svgString = serializer.serializeToString(svgElement);
                
                // Ensure SVG has proper dimensions for rendering
                const enhancedSvgString = this.enhanceSVGForRendering(svgString, width, height);
                const svgDataUrl = 'data:image/svg+xml;base64,' + btoa(enhancedSvgString);

                // Create image and draw when loaded
                const img = new Image();
                img.onload = () => {
                    try {
                        if (this.graphics._context || this.graphics.m_oContext) {
                            const ctx = this.graphics._context || this.graphics.m_oContext;
                            ctx.save();
                            // Preserve SVG's intrinsic aspect ratio (xMidYMid meet)
                            let drawW = width, drawH = height, drawX = x, drawY = y;
                            if (img.naturalWidth > 0 && img.naturalHeight > 0) {
                                const imgAspect = img.naturalWidth / img.naturalHeight;
                                const boxAspect = width / height;
                                if (imgAspect > boxAspect) {
                                    drawW = width;
                                    drawH = width / imgAspect;
                                    drawY = y + (height - drawH) / 2;
                                } else {
                                    drawH = height;
                                    drawW = height * imgAspect;
                                    drawX = x + (width - drawW) / 2;
                                }
                            }
                            ctx.drawImage(img, drawX, drawY, drawW, drawH);
                            ctx.restore();
                        } else {
                            this.logger.logError(this.constructor.name, 'No canvas context available for SVG drawing');
                        }
                        resolve();
                    } catch (drawError) {
                        this.logger.logError(this.constructor.name, 'Error drawing SVG image:', drawError);
                        this.renderSVGPlaceholder(x, y, width, height, 'SVG Draw Error');
                        reject(drawError);
                    }
                };
                
                img.onerror = (error) => {
                    this.logger.logError(this.constructor.name, 'SVG image load error:', error);
                    this.renderSVGPlaceholder(x, y, width, height, 'SVG Load Error');
                    reject(error);
                };

                // Set crossOrigin to handle data URIs properly
                img.crossOrigin = 'anonymous';
                img.src = svgDataUrl;
                
                // Timeout after 5 seconds
                setTimeout(() => {
                    if (img.complete === false) {
                        this.logger.log("warn", this.constructor.name, 'SVG loading timed out');
                        this.renderSVGPlaceholder(x, y, width, height, 'SVG Timeout');
                        reject(new Error('SVG loading timeout'));
                    }
                }, 5000);

            } catch (error) {
                this.logger.logError(this.constructor.name, 'Error in SVG Image rendering:', error);
                this.renderSVGPlaceholder(x, y, width, height, 'SVG Render Error');
                reject(error);
            }
        });
    }

    /**
     * Render SVG by parsing and drawing elements (fallback method)
     * @param {Element} svgElement - SVG element
     * @param {number} x - X position
     * @param {number} y - Y position
     * @param {number} scale - Scale factor
     * @param {Object} options - Rendering options
     */
    renderSVGViaParsing(svgElement, x, y, scale, options) {
        try {
            if (!this.graphics._context) {return;}

            this.graphics._context.save();
            this.graphics._context.translate(x, y);
            this.graphics._context.scale(scale, scale);

            // Process SVG child elements
            this.processSVGElements(svgElement.children);

            this.graphics._context.restore();

        } catch (error) {
            this.logger.logError(this.constructor.name, 'Error in SVG parsing rendering:', error);
            this.renderSVGPlaceholder(x, y, 100, 100, 'SVG Parse Error');
        }
    }

    /**
     * Process SVG child elements
     * @param {HTMLCollection} elements - SVG child elements
     */
    processSVGElements(elements) {
        for (const element of elements) {
            this.processSVGElement(element);
        }
    }

    /**
     * Process individual SVG element
     * @param {Element} element - SVG element
     */
    processSVGElement(element) {
        const tagName = element.tagName.toLowerCase();

        switch (tagName) {
            case 'rect':
                this.renderSVGRect(element);
                break;
            case 'circle':
                this.renderSVGCircle(element);
                break;
            case 'ellipse':
                this.renderSVGEllipse(element);
                break;
            case 'line':
                this.renderSVGLine(element);
                break;
            case 'polyline':
            case 'polygon':
                this.renderSVGPolygon(element);
                break;
            case 'path':
                this.renderSVGPath(element);
                break;
            case 'text':
                this.renderSVGText(element);
                break;
            case 'g':
                // Group - process children
                this.graphics._context.save();
                this.applySVGTransform(element);
                this.processSVGElements(element.children);
                this.graphics._context.restore();
                break;
            default:
                // Unknown element - skip
                break;
        }
    }

    /**
     * Render SVG rectangle
     * @param {Element} element - Rect element
     */
    renderSVGRect(element) {
        const x = parseFloat(element.getAttribute('x')) || 0;
        const y = parseFloat(element.getAttribute('y')) || 0;
        const width = parseFloat(element.getAttribute('width')) || 0;
        const height = parseFloat(element.getAttribute('height')) || 0;

        this.applySVGStyles(element);
        
        const ctx = this.graphics._context;
        ctx.fillRect(x, y, width, height);
        ctx.strokeRect(x, y, width, height);
    }

    /**
     * Render SVG circle
     * @param {Element} element - Circle element
     */
    renderSVGCircle(element) {
        const cx = parseFloat(element.getAttribute('cx')) || 0;
        const cy = parseFloat(element.getAttribute('cy')) || 0;
        const r = parseFloat(element.getAttribute('r')) || 0;

        this.applySVGStyles(element);

        const ctx = this.graphics._context;
        ctx.beginPath();
        ctx.arc(cx, cy, r, 0, 2 * Math.PI);
        ctx.fill();
        ctx.stroke();
    }

    /**
     * Render SVG line
     * @param {Element} element - Line element
     */
    renderSVGLine(element) {
        const x1 = parseFloat(element.getAttribute('x1')) || 0;
        const y1 = parseFloat(element.getAttribute('y1')) || 0;
        const x2 = parseFloat(element.getAttribute('x2')) || 0;
        const y2 = parseFloat(element.getAttribute('y2')) || 0;

        this.applySVGStyles(element);

        const ctx = this.graphics._context;
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);
        ctx.stroke();
    }

    /**
     * Render SVG text
     * @param {Element} element - Text element
     */
    renderSVGText(element) {
        const x = parseFloat(element.getAttribute('x')) || 0;
        const y = parseFloat(element.getAttribute('y')) || 0;
        const text = element.textContent || '';

        this.applySVGStyles(element);

        const ctx = this.graphics._context;
        ctx.fillText(text, x, y);
    }

    /**
     * Apply SVG styles to Canvas context
     * @param {Element} element - SVG element
     */
    applySVGStyles(element) {
        const ctx = this.graphics._context;
        
        // Fill
        const fill = element.getAttribute('fill');
        if (fill && fill !== 'none') {
            ctx.fillStyle = fill;
        } else {
            ctx.fillStyle = 'transparent';
        }

        // Stroke
        const stroke = element.getAttribute('stroke');
        if (stroke && stroke !== 'none') {
            ctx.strokeStyle = stroke;
        } else {
            ctx.strokeStyle = 'transparent';
        }

        // Stroke width
        const strokeWidth = element.getAttribute('stroke-width');
        if (strokeWidth) {
            ctx.lineWidth = parseFloat(strokeWidth);
        }

        // Opacity
        const opacity = element.getAttribute('opacity');
        if (opacity) {
            ctx.globalAlpha = parseFloat(opacity);
        }
    }

    /**
     * Apply SVG transform to Canvas context
     * @param {Element} element - SVG element
     */
    applySVGTransform(element) {
        const transform = element.getAttribute('transform');
        if (transform) {
            // Basic transform parsing - could be expanded
            const translateMatch = transform.match(/translate\(([^)]+)\)/);
            if (translateMatch) {
                const values = translateMatch[1].split(',').map(v => parseFloat(v.trim()));
                if (values.length >= 2) {
                    this.graphics._context.translate(values[0], values[1]);
                }
            }

            const scaleMatch = transform.match(/scale\(([^)]+)\)/);
            if (scaleMatch) {
                const values = scaleMatch[1].split(',').map(v => parseFloat(v.trim()));
                if (values.length >= 1) {
                    const scaleX = values[0];
                    const scaleY = values.length > 1 ? values[1] : scaleX;
                    this.graphics._context.scale(scaleX, scaleY);
                }
            }

            const rotateMatch = transform.match(/rotate\(([^)]+)\)/);
            if (rotateMatch) {
                const angle = parseFloat(rotateMatch[1]);
                this.graphics._context.rotate(angle * Math.PI / 180);
            }
        }
    }

    /**
     * Render SVG placeholder
     * @param {number} x - X position
     * @param {number} y - Y position
     * @param {number} width - Width
     * @param {number} height - Height
     * @param {string} message - Placeholder message
     */
    renderSVGPlaceholder(x, y, width, height, _message = 'SVG') {
        // Draw placeholder rectangle only (no text)
        this.graphics.fillRect(x, y, width, height, { r: 240, g: 240, b: 240 });
        this.graphics.strokeRect(x, y, width, height, { r: 200, g: 200, b: 200 }, 1);
    }

    /**
     * Render SVG ellipse
     * @param {Element} element - Ellipse element
     */
    renderSVGEllipse(element) {
        const cx = parseFloat(element.getAttribute('cx')) || 0;
        const cy = parseFloat(element.getAttribute('cy')) || 0;
        const rx = parseFloat(element.getAttribute('rx')) || 0;
        const ry = parseFloat(element.getAttribute('ry')) || 0;

        this.applySVGStyles(element);

        const ctx = this.graphics._context;
        ctx.save();
        ctx.translate(cx, cy);
        ctx.scale(rx, ry);
        ctx.beginPath();
        ctx.arc(0, 0, 1, 0, 2 * Math.PI);
        ctx.restore();
        ctx.fill();
        ctx.stroke();
    }
    
    /**
     * Render SVG polygon/polyline
     * @param {Element} element - Polygon/polyline element
     */
    renderSVGPolygon(element) {
        const points = element.getAttribute('points') || '';
        const isPolygon = element.tagName.toLowerCase() === 'polygon';
        
        if (!points.trim()) {return;}

        this.applySVGStyles(element);

        const ctx = this.graphics._context;
        const pointPairs = points.trim().split(/[\s,]+/);
        
        if (pointPairs.length < 4) {return;} // Need at least 2 points (x,y pairs)

        ctx.beginPath();
        
        // Parse first point
        let x = parseFloat(pointPairs[0]);
        let y = parseFloat(pointPairs[1]);
        ctx.moveTo(x, y);
        
        // Parse remaining points
        for (let i = 2; i < pointPairs.length; i += 2) {
            if (i + 1 < pointPairs.length) {
                x = parseFloat(pointPairs[i]);
                y = parseFloat(pointPairs[i + 1]);
                ctx.lineTo(x, y);
            }
        }
        
        // Close path for polygon
        if (isPolygon) {
            ctx.closePath();
        }
        
        ctx.fill();
        ctx.stroke();
    }
    
    /**
     * Render SVG path
     * @param {Element} element - Path element
     */
    renderSVGPath(element) {
        const d = element.getAttribute('d') || '';
        
        if (!d.trim()) {return;}

        this.applySVGStyles(element);

        try {
            const ctx = this.graphics._context;
            const path = this.parseSVGPathData(d);
            
            ctx.beginPath();
            this.executeSVGPathCommands(path, ctx);
            ctx.fill();
            ctx.stroke();
            
        } catch (error) {
            this.logger.log("warn", this.constructor.name, 'Error parsing SVG path:', error);
            // Fallback to simple rectangle
            const bbox = this.getElementBoundingBox(element);
            if (bbox) {
                this.applySVGStyles(element);
                const ctx = this.graphics._context;
                ctx.fillRect(bbox.x, bbox.y, bbox.width, bbox.height);
                ctx.strokeRect(bbox.x, bbox.y, bbox.width, bbox.height);
            }
        }
    }

    /**
     * Parse SVG path data into command array
     * @param {string} pathData - SVG path data string
     * @return {Array} Array of path commands
     */
    parseSVGPathData(pathData) {
        const commands = [];
        const commandRegex = /([MmLlHhVvCcSsQqTtAaZz])((?:\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?\s*,?\s*)*)/g;
        
        let match;
        while ((match = commandRegex.exec(pathData)) !== null) {
            const command = match[1];
            const paramsStr = match[2].trim();
            const params = paramsStr ? paramsStr.split(/[\s,]+/).map(p => parseFloat(p)).filter(p => !isNaN(p)) : [];
            
            commands.push({ command, params });
        }
        
        return commands;
    }

    /**
     * Execute SVG path commands on canvas context
     * @param {Array} commands - Array of path commands
     * @param {CanvasRenderingContext2D} ctx - Canvas context
     */
    executeSVGPathCommands(commands, ctx) {
        let currentX = 0, currentY = 0;
        let lastControlX = 0, lastControlY = 0;
        
        for (const { command, params } of commands) {
            switch (command.toLowerCase()) {
                case 'm': // Move to
                    if (params.length >= 2) {
                        if (command === 'M') {
                            currentX = params[0];
                            currentY = params[1];
                        } else {
                            currentX += params[0];
                            currentY += params[1];
                        }
                        ctx.moveTo(currentX, currentY);
                        
                        // Additional coordinate pairs are treated as lineTo
                        for (let i = 2; i < params.length; i += 2) {
                            if (i + 1 < params.length) {
                                if (command === 'M') {
                                    currentX = params[i];
                                    currentY = params[i + 1];
                                } else {
                                    currentX += params[i];
                                    currentY += params[i + 1];
                                }
                                ctx.lineTo(currentX, currentY);
                            }
                        }
                    }
                    break;
                    
                case 'l': // Line to
                    for (let i = 0; i < params.length; i += 2) {
                        if (i + 1 < params.length) {
                            if (command === 'L') {
                                currentX = params[i];
                                currentY = params[i + 1];
                            } else {
                                currentX += params[i];
                                currentY += params[i + 1];
                            }
                            ctx.lineTo(currentX, currentY);
                        }
                    }
                    break;
                    
                case 'h': // Horizontal line to
                    for (let i = 0; i < params.length; i++) {
                        if (command === 'H') {
                            currentX = params[i];
                        } else {
                            currentX += params[i];
                        }
                        ctx.lineTo(currentX, currentY);
                    }
                    break;
                    
                case 'v': // Vertical line to
                    for (let i = 0; i < params.length; i++) {
                        if (command === 'V') {
                            currentY = params[i];
                        } else {
                            currentY += params[i];
                        }
                        ctx.lineTo(currentX, currentY);
                    }
                    break;
                    
                case 'c': // Cubic Bézier curve to
                    for (let i = 0; i < params.length; i += 6) {
                        if (i + 5 < params.length) {
                            let cp1x, cp1y, cp2x, cp2y, x, y;
                            
                            if (command === 'C') {
                                cp1x = params[i];
                                cp1y = params[i + 1];
                                cp2x = params[i + 2];
                                cp2y = params[i + 3];
                                x = params[i + 4];
                                y = params[i + 5];
                            } else {
                                cp1x = currentX + params[i];
                                cp1y = currentY + params[i + 1];
                                cp2x = currentX + params[i + 2];
                                cp2y = currentY + params[i + 3];
                                x = currentX + params[i + 4];
                                y = currentY + params[i + 5];
                            }
                            
                            ctx.bezierCurveTo(cp1x, cp1y, cp2x, cp2y, x, y);
                            lastControlX = cp2x;
                            lastControlY = cp2y;
                            currentX = x;
                            currentY = y;
                        }
                    }
                    break;
                    
                case 'q': // Quadratic Bézier curve to
                    for (let i = 0; i < params.length; i += 4) {
                        if (i + 3 < params.length) {
                            let cpx, cpy, x, y;
                            
                            if (command === 'Q') {
                                cpx = params[i];
                                cpy = params[i + 1];
                                x = params[i + 2];
                                y = params[i + 3];
                            } else {
                                cpx = currentX + params[i];
                                cpy = currentY + params[i + 1];
                                x = currentX + params[i + 2];
                                y = currentY + params[i + 3];
                            }
                            
                            ctx.quadraticCurveTo(cpx, cpy, x, y);
                            lastControlX = cpx;
                            lastControlY = cpy;
                            currentX = x;
                            currentY = y;
                        }
                    }
                    break;
                    
                case 'a': // Elliptical arc
                    for (let i = 0; i < params.length; i += 7) {
                        if (i + 6 < params.length) {
                            const rx = params[i];
                            const ry = params[i + 1];
                            const rotation = params[i + 2] * Math.PI / 180;
                            const largeArcFlag = params[i + 3];
                            const sweepFlag = params[i + 4];
                            let x, y;
                            
                            if (command === 'A') {
                                x = params[i + 5];
                                y = params[i + 6];
                            } else {
                                x = currentX + params[i + 5];
                                y = currentY + params[i + 6];
                            }
                            
                            // Simplified arc implementation
                            this.drawEllipticalArc(ctx, currentX, currentY, rx, ry, rotation, largeArcFlag, sweepFlag, x, y);
                            currentX = x;
                            currentY = y;
                        }
                    }
                    break;
                    
                case 'z': // Close path
                    ctx.closePath();
                    break;
            }
        }
    }

    /**
     * Draw elliptical arc (simplified implementation)
     * @param {CanvasRenderingContext2D} ctx - Canvas context
     * @param {number} x1 - Start X
     * @param {number} y1 - Start Y
     * @param {number} rx - X radius
     * @param {number} ry - Y radius
     * @param {number} rotation - Rotation angle
     * @param {number} largeArcFlag - Large arc flag
     * @param {number} sweepFlag - Sweep direction flag
     * @param {number} x2 - End X
     * @param {number} y2 - End Y
     */
    drawEllipticalArc(ctx, x1, y1, rx, ry, rotation, largeArcFlag, sweepFlag, x2, y2) {
        // Simplified implementation - draws a line for now
        // Full elliptical arc implementation would require complex math
        ctx.lineTo(x2, y2);
    }

    /**
     * Get bounding box of an element (fallback method)
     * @param {Element} element - SVG element
     * @return {Object|null} Bounding box or null
     */
    getElementBoundingBox(element) {
        // Try to get common attributes for basic shapes
        const x = parseFloat(element.getAttribute('x')) || 0;
        const y = parseFloat(element.getAttribute('y')) || 0;
        const width = parseFloat(element.getAttribute('width')) || 10;
        const height = parseFloat(element.getAttribute('height')) || 10;
        
        if (width > 0 && height > 0) {
            return { x, y, width, height };
        }
        
        return null;
    }
}

// Export class

// Export classes (maintain backward compatibility)
if (typeof window !== 'undefined') {
    window.SVGRenderer = SVGRenderer;
}

// Intentionally no ES module exports to support classic <script> usage in root demo
