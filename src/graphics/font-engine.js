/**
 * Font Engine Module
 * Font processing system
 * Adapted from modern font engine implementation
 */

// Import dependencies
// import { Logger } from '../utils/utils.js';

class FontEngine {
    constructor() {
        this.logger = new Logger('FontEngine');

        // Initialize font system
        this.initializeFontSystem();

        // Simplified list of "available" fonts, mimicking AllFonts.js
        this.availableFonts = new Set([
            'Arial', 'Calibri', 'Times New Roman', 'Courier New', 'Verdana',
            'Georgia', 'Tahoma', 'Helvetica', 'Garamond', 'Impact',
            'Segoe UI', 'Comic Sans MS', 'Trebuchet MS', 'Lucida Sans Unicode',
            'Palatino Linotype', 'Book Antiqua', 'Bookman Old Style', 'Century Gothic'
        ]);

        this.documentFonts = new Set();
        this.missingFonts = new Set();
        this.fontCache = new Map();
        this.fontManager = null;
    }

    /**
     * Initialize the font system
     */
    initializeFontSystem() {
        try {
            // Initialize AscFonts if available
            if (typeof window.AscFonts !== 'undefined') {
                this.fontManager = new window.AscFonts.CFontManagerEngine();
                this.initializeFonts();
            } else {
                this.fontManager = new SimplifiedFontManager();
            }
        } catch (error) {
            this.fontManager = new SimplifiedFontManager();
        }
    }

    /**
     * Initialize font loading
     */
    initializeFonts() {
        try {
            // Check if system fonts API is available
            if (typeof self.queryLocalFonts === 'function') {
                this.loadSystemFonts();
            } else {
                this.loadFallbackFonts();
            }
        } catch (_error) {
            this.loadFallbackFonts();
        }
    }

    /**
     * Load system fonts using the Font Access API
     */
    async loadSystemFonts() {
        try {
            const fonts = await self.queryLocalFonts();

            // Add system fonts to available fonts
            fonts.forEach(font => {
                this.availableFonts.add(font.family);
            });

        } catch (error) {
            this.loadFallbackFonts();
        }
    }

    /**
     * Load fallback fonts when system fonts are not available
     */
    loadFallbackFonts() {
        // Add common web fonts
        const webFonts = [
            'Arial', 'Helvetica', 'Times New Roman', 'Times', 'Courier New', 'Courier',
            'Verdana', 'Georgia', 'Palatino', 'Garamond', 'Bookman', 'Comic Sans MS',
            'Trebuchet MS', 'Arial Black', 'Impact', 'Lucida Sans Unicode',
            'Tahoma', 'Geneva', 'Lucida Grande', 'Segoe UI', 'Calibri', 'Cambria',
            'Candara', 'Consolas', 'Constantia', 'Corbel', 'Franklin Gothic Medium',
            'Gill Sans', 'Lucida Console', 'Lucida Sans Typewriter', 'MS Gothic',
            'MS Mincho', 'MS PGothic', 'MS PMincho', 'MS Reference Sans Serif',
            'MS Reference Specialty', 'MS Sans Serif', 'MS Serif', 'Myriad Pro',
            'Optima', 'Perpetua', 'Rockwell', 'Rockwell Extra Bold', 'Segoe Print',
            'Segoe Script', 'Segoe UI Light', 'Segoe UI Semibold', 'Segoe UI Symbol',
            'Tw Cen MT', 'Tw Cen MT Condensed', 'Tw Cen MT Condensed Extra Bold'
        ];

        webFonts.forEach(font => {
            this.availableFonts.add(font);
        });
    }

    /**
     * Analyze the presentation DOM to find all used fonts
     * @param {CPresentation} presentation - The main presentation object
     */
    analyzeDocument(presentation) {
        this.documentFonts.clear();

        // Analyze default text style
        if (presentation.defaultTextStyle) {
            this.addFontFromStyle(presentation.defaultTextStyle);
        }

        // Analyze slides
        presentation.slides.forEach(slide => {
            if (slide.commonSlideData && slide.commonSlideData.shapeTree) {
                slide.commonSlideData.shapeTree.forEach(shape => {
                    this.analyzeShape(shape);
                });
            }
        });

        return Array.from(this.documentFonts);
    }

    /**
     * Add font from a style object
     * @param {CTextStyle} style - The style to analyze
     */
    addFontFromStyle(style) {
        if (!style) {return;}
        // This is a simplified stub. A real implementation would check
        // the default, level1, level2, etc. paragraph properties.
        if (style.font) {
            this.documentFonts.add(style.font);
        }
    }

    /**
     * Recursively analyze a shape and its children for fonts
     * @param {CShape} shape - The shape to analyze
     */
    analyzeShape(shape) {
        if (!shape) {return;}

        // Check text body for fonts
        if (shape.textBody && shape.textBody.paragraphs) {
            shape.textBody.paragraphs.forEach(para => {
                if (para.runs) {
                    para.runs.forEach(run => {
                        if (run.properties && run.properties.font) {
                            this.documentFonts.add(run.properties.font);
                        }
                    });
                }
            });
        }

        // Recursively check grouped shapes
        if (shape.shapeTree) { // For CGroupShape
            shape.shapeTree.forEach(childShape => {
                this.analyzeShape(childShape);
            });
        }
    }

    /**
     * "Load" the required fonts
     * This simulates checking availability and preparing for loading
     */
    loadFonts() {
        this.missingFonts.clear();

        this.documentFonts.forEach(font => {
            if (!this.availableFonts.has(font)) {
                this.missingFonts.add(font);
            }
        });

    }

    /**
     * Get the results of the font processing
     */
    getResults() {
        return {
            documentFonts: Array.from(this.documentFonts),
            availableFonts: Array.from(this.availableFonts),
            missingFonts: Array.from(this.missingFonts)
        };
    }

    /**
     * Get font metrics for text rendering
     * @param {string} fontName - Font family name
     * @param {number} fontSize - Font size in points
     * @param {string} text - Text to measure
     * @returns {Object} Font metrics
     */
    getFontMetrics(fontName, fontSize, text) {
        const cacheKey = `${fontName}_${fontSize}_${text}`;

        if (this.fontCache.has(cacheKey)) {
            return this.fontCache.get(cacheKey);
        }

        // Simplified font metrics calculation
        const metrics = {
            width: text.length * fontSize * 0.6, // Approximate character width
            height: fontSize * 1.2, // Approximate line height
            ascent: fontSize * 0.8,
            descent: fontSize * 0.2,
            baseline: fontSize * 0.8
        };

        this.fontCache.set(cacheKey, metrics);
        return metrics;
    }

    /**
     * Check if a font is available
     * @param {string} fontName - Font family name
     * @returns {boolean} True if font is available
     */
    isFontAvailable(fontName) {
        return this.availableFonts.has(fontName) ||
               this.documentFonts.has(fontName);
    }

    /**
     * Get fallback font if the requested font is not available
     * @param {string} fontName - Requested font name
     * @returns {string} Fallback font name
     */
    getFallbackFont(fontName) {
        if (this.isFontAvailable(fontName)) {
            return fontName;
        }

        // Return emoji-capable fallback fonts for Unicode character support
        return 'Arial, "Apple Color Emoji", "Segoe UI Emoji", "Noto Color Emoji", "Segoe UI Symbol", sans-serif';
    }

    /**
     * Render text using font system patterns
     * Following the approach from sdkjs/common/libfont/textmeasurer.js and sdkjs/word/Drawing/Graphics.js
     * @param {CanvasRenderingContext2D} ctx - Canvas context
     * @param {string} text - Text to render
     * @param {number} x - X coordinate
     * @param {number} y - Y coordinate
     * @param {Object} style - Text style object
     */
    renderText(ctx, text, x, y, style) {
        // Get font name with fallback following standard patterns
        const fontName = this.getFallbackFont(style.fontFamily || 'Arial');
        const fontSize = style.fontSize || 12;
        const fontWeight = style.fontWeight || (style.bold ? 'bold' : 'normal');
        const fontStyle = style.fontStyle || (style.italic ? 'italic' : 'normal');

        // Set canvas font following standard font setup patterns
        ctx.font = `${fontStyle} ${fontWeight} ${fontSize}px "${fontName}"`;
        
        // Convert color to CSS format if needed
        let fillColor = '#000000';
        if (style.color) {
            if (typeof style.color === 'object' && style.color.r !== undefined) {
                // RGBA object format: {r: 255, g: 0, b: 0, a: 255} or {r: 255, g: 0, b: 0}
                const alpha = style.color.a !== undefined ? style.color.a / 255 : 1;
                fillColor = `rgba(${style.color.r}, ${style.color.g}, ${style.color.b}, ${alpha})`;
            } else if (typeof style.color === 'string') {
                fillColor = style.color;
            } else {
                // Unsupported color format - use default
                fillColor = '#000000';
            }
        } else {
            // No color specified - use default
            fillColor = '#000000';
        }
        
        ctx.fillStyle = fillColor;
        ctx.textBaseline = 'alphabetic';

        // Apply text alignment if specified
        if (style.textAlign) {
            ctx.textAlign = style.textAlign;
        }

        // Render text character by character for proper glyph handling
        // This follows the pattern from sdkjs/word/Drawing/Graphics.js FillText method
        if (text.length === 1) {
            // Single character - direct rendering
            ctx.fillText(text, x, y);
        } else {
            // Multi-character - render character by character for consistency
            let currentX = x;
            for (let i = 0; i < text.length; i++) {
                const char = text.charAt(i);
                ctx.fillText(char, currentX, y);

                // Advance position using proper character measurement
                const metrics = ctx.measureText(char);
                currentX += metrics.width;

                // Apply letter spacing if specified
                if (style.letterSpacing) {
                    currentX += style.letterSpacing;
                }
            }
        }
    }

    /**
     * Measure text following standard measurement patterns
     * Based on sdkjs/common/libfont/textmeasurer.js Measure method
     * @param {CanvasRenderingContext2D} ctx - Canvas context
     * @param {string} text - Text to measure
     * @param {Object} style - Text style object
     * @returns {Object} Text metrics with width and height
     */
    measureText(ctx, text, style) {
        // Get font name with fallback
        const fontName = this.getFallbackFont(style.fontFamily || 'Arial');
        const fontSize = style.fontSize || 12;
        const fontWeight = style.fontWeight || (style.bold ? 'bold' : 'normal');
        const fontStyle = style.fontStyle || (style.italic ? 'italic' : 'normal');

        // Set up font for measurement
        ctx.save();
        ctx.font = `${fontStyle} ${fontWeight} ${fontSize}px "${fontName}"`;

        // Measure text character by character for accuracy
        let totalWidth = 0;
        let maxHeight = fontSize; // Base height on font size

        for (let i = 0; i < text.length; i++) {
            const char = text.charAt(i);
            const metrics = ctx.measureText(char);

            // Accumulate width
            totalWidth += metrics.width;

            // Apply letter spacing if specified
            if (style.letterSpacing) {
                totalWidth += style.letterSpacing;
            }

            // Update height based on actual bounding box if available
            if (metrics.actualBoundingBoxAscent !== undefined && metrics.actualBoundingBoxDescent !== undefined) {
                const charHeight = metrics.actualBoundingBoxAscent + metrics.actualBoundingBoxDescent;
                maxHeight = Math.max(maxHeight, charHeight);
            }
        }

        ctx.restore();

        return {
            width: totalWidth,
            height: maxHeight
        };
    }
}

/**
 * Simplified Font Manager
 * Fallback implementation when font system is not available
 */
class SimplifiedFontManager {
    constructor() {
        this.library = null;
        this.manager = null;
    }

    openFont(_stream, _faceindex) {
        // Simplified implementation - just return a mock font object
        return {
            SetFace: function(face, manager) {
                this.face = face;
                this.manager = manager;
            },
            IsSuccess: function() {
                return true;
            }
        };
    }

    setHintsProps(bIsHinting, bIsSubpixHinting) {
        // Simplified implementation
        this.hinting = bIsHinting;
        this.subpixHinting = bIsSubpixHinting;
    }

    destroy() {
        // Cleanup
        this.library = null;
        this.manager = null;
    }
}

// Export classes (maintain backward compatibility)
if (typeof window !== 'undefined') {
    window.FontEngine = FontEngine;
    window.SimplifiedFontManager = SimplifiedFontManager;
}

// Intentionally no ES module exports to support classic <script> usage in root demo
