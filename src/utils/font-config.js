/**
 * Font Configuration System
 * Centralized font handling for PPTX chart rendering with proper PPTX integration
 */

/**
 * Chart Font Configuration Class
 * Manages font settings across different chart elements with PPTX override support
 */
class ChartFontConfig {
    constructor(options = {}) {
        this.options = this._validateOptions(options);
        
        // Initialize logger with error handling
        try {
            this.logger = typeof Logger !== 'undefined' ? new Logger('ChartFontConfig') : this._createFallbackLogger();
        } catch (error) {
            this.logger = this._createFallbackLogger();
        }
        
        // Initialize default font configurations for different chart elements
        this.initializeDefaults();
        
        // Initialize responsive scaling settings
        this.initializeScaling();
    }

    /**
     * Initialize default font configurations for different chart elements
     * @private
     */
    initializeDefaults() {
        this.defaults = {
            // Global chart defaults
            global: {
                fontFamily: 'Calibri',
                fontSize: 11, // Use original DOM-based sizes
                bold: false,
                italic: false,
                color: { r: 0, g: 0, b: 0 }
            },

            // Chart title fonts
            title: {
                fontFamily: 'Calibri',
                fontSize: 18, // Use original DOM-based sizes
                bold: false, // Don't default to bold - use PPTX font data
                italic: false,
                color: { r: 0, g: 0, b: 0 }
            },

            // Chart subtitle fonts
            subtitle: {
                fontFamily: 'Calibri',
                fontSize: 14, // Use original DOM-based sizes
                bold: false,
                italic: false,
                color: { r: 100, g: 100, b: 100 }
            },

            // Axis title fonts
            axisTitle: {
                fontFamily: 'Calibri',
                fontSize: 12, // Use original DOM-based sizes
                bold: true,
                italic: false,
                color: { r: 68, g: 68, b: 68 }
            },

            // Axis label fonts
            axisLabels: {
                fontFamily: 'Calibri',
                fontSize: 11, // Use original DOM-based sizes
                bold: false,
                italic: false,
                color: { r: 68, g: 68, b: 68 }
            },

            // Legend fonts
            legend: {
                fontFamily: 'Calibri',
                fontSize: 11, // Use original DOM-based sizes
                bold: false,
                italic: false,
                color: { r: 60, g: 60, b: 60 }
            },

            // Data label fonts
            dataLabels: {
                fontFamily: 'Calibri',
                fontSize: 9, // Use original DOM-based sizes
                bold: true,
                italic: false,
                color: { r: 50, g: 50, b: 50 }
            },

            // Tick mark fonts
            tickMarks: {
                fontFamily: 'Calibri',
                fontSize: 10, // Use original DOM-based sizes
                bold: false,
                italic: false,
                color: { r: 68, g: 68, b: 68 }
            }
        };
    }

    /**
     * Initialize responsive scaling settings
     * @private
     */
    initializeScaling() {
        this.scaling = {
            enabled: this.options.enableResponsiveScaling !== false,
            baseDimensions: {
                width: 400,  // Base chart width
                height: 300  // Base chart height
            },
            scaleFactors: {
                min: 0.7,    // Minimum scale factor (70%) - original values
                max: 1.5     // Maximum scale factor (150%) - original values
            },
            dpiSupport: {
                enabled: this.options.enableDpiScaling !== false,
                baseDpi: 96, // Standard screen DPI
                currentDpi: this._getCurrentDpi()
            }
        };
    }

    /**
     * Get font configuration for a specific chart element
     * @param {string} elementType - Type of chart element
     * @param {Object} pptxOverrides - PPTX-specific font overrides
     * @param {Object} customOverrides - Custom font overrides
     * @param {Object} scalingContext - Scaling context (chart dimensions, DPI)
     * @return {Object} Resolved font configuration
     */
    getFont(elementType, pptxOverrides = {}, customOverrides = {}, scalingContext = {}) {
        try {
            // Get base font configuration
            const baseFont = this.defaults[elementType] || this.defaults.global;
            
            // Apply PPTX overrides (highest priority)
            const pptxFont = this._applyPptxOverrides(baseFont, pptxOverrides);
            
            // Apply custom overrides
            const customFont = this._applyCustomOverrides(pptxFont, customOverrides);
            
            // Apply responsive scaling
            const scaledFont = this._applyScaling(customFont, scalingContext);
            
            // Normalize and validate final font configuration
            const finalFont = this._normalizeFontConfig(scaledFont);
            
            this.logger.log("debug", this.constructor.name, `Font resolved for ${elementType}:`, {
                base: baseFont,
                pptx: pptxOverrides,
                custom: customOverrides,
                final: finalFont
            });
            
            return finalFont;
            
        } catch (error) {
            this.logger.logError(this.constructor.name, `Error resolving font for ${elementType}:`, error);
            return this._getFallbackFont();
        }
    }

    /**
     * Apply PPTX font overrides (highest priority)
     * @param {Object} baseFont - Base font configuration
     * @param {Object} pptxOverrides - PPTX font overrides
     * @return {Object} Font with PPTX overrides applied
     * @private
     */
    _applyPptxOverrides(baseFont, pptxOverrides) {
        const font = { ...baseFont };
        
        if (!pptxOverrides || typeof pptxOverrides !== 'object') {
            return font;
        }

        // Handle various PPTX font property formats
        if (pptxOverrides.fontFamily) {
            font.fontFamily = pptxOverrides.fontFamily;
        } else if (pptxOverrides.family) {
            font.fontFamily = pptxOverrides.family;
        } else if (pptxOverrides.typeface) {
            font.fontFamily = pptxOverrides.typeface;
        }

        // Handle font size with PPTX units conversion
        if (pptxOverrides.fontSize !== undefined) {
            font.fontSize = this._convertPptxFontSize(pptxOverrides.fontSize);
        } else if (pptxOverrides.size !== undefined) {
            font.fontSize = this._convertPptxFontSize(pptxOverrides.size);
        } else if (pptxOverrides.sz !== undefined) {
            // PPTX often uses 'sz' attribute with size in hundreds of points
            font.fontSize = this._convertPptxFontSize(pptxOverrides.sz, 'hundredths');
        }

        // Handle font style attributes
        if (pptxOverrides.bold !== undefined) {
            font.bold = Boolean(pptxOverrides.bold);
        } else if (pptxOverrides.b !== undefined) {
            font.bold = pptxOverrides.b !== '0' && pptxOverrides.b !== false;
        }

        if (pptxOverrides.italic !== undefined) {
            font.italic = Boolean(pptxOverrides.italic);
        } else if (pptxOverrides.i !== undefined) {
            font.italic = pptxOverrides.i !== '0' && pptxOverrides.i !== false;
        }

        // Handle color
        if (pptxOverrides.color) {
            font.color = this._normalizeColor(pptxOverrides.color);
        }

        return font;
    }

    /**
     * Apply custom font overrides
     * @param {Object} baseFont - Base font configuration
     * @param {Object} customOverrides - Custom font overrides
     * @return {Object} Font with custom overrides applied
     * @private
     */
    _applyCustomOverrides(baseFont, customOverrides) {
        const font = { ...baseFont };
        
        if (!customOverrides || typeof customOverrides !== 'object') {
            return font;
        }

        // Apply custom overrides with property normalization
        if (customOverrides.fontFamily) {font.fontFamily = customOverrides.fontFamily;}
        if (customOverrides.family) {font.fontFamily = customOverrides.family;}
        
        if (customOverrides.fontSize !== undefined) {
            font.fontSize = Number(customOverrides.fontSize);
        } else if (customOverrides.size !== undefined) {
            font.fontSize = Number(customOverrides.size);
        }
        
        if (customOverrides.bold !== undefined) {font.bold = Boolean(customOverrides.bold);}
        if (customOverrides.italic !== undefined) {font.italic = Boolean(customOverrides.italic);}
        if (customOverrides.color) {font.color = this._normalizeColor(customOverrides.color);}

        return font;
    }

    /**
     * Apply responsive scaling to font size
     * @param {Object} font - Font configuration
     * @param {Object} scalingContext - Scaling context
     * @return {Object} Font with scaling applied
     * @private
     */
    _applyScaling(font, scalingContext) {
        const scaledFont = { ...font };
        
        if (!this.scaling.enabled || !scalingContext.chartWidth || !scalingContext.chartHeight) {
            return scaledFont;
        }

        try {
            // Calculate scale factor based on chart dimensions
            const widthScale = scalingContext.chartWidth / this.scaling.baseDimensions.width;
            const heightScale = scalingContext.chartHeight / this.scaling.baseDimensions.height;
            const averageScale = (widthScale + heightScale) / 2;
            
            // Apply scale factor constraints
            const constrainedScale = Math.max(
                this.scaling.scaleFactors.min,
                Math.min(this.scaling.scaleFactors.max, averageScale)
            );
            
            // Apply DPI scaling if enabled
            let dpiScale = 1;
            if (this.scaling.dpiSupport.enabled && scalingContext.dpi) {
                dpiScale = scalingContext.dpi / this.scaling.dpiSupport.baseDpi;
            } else if (this.scaling.dpiSupport.enabled) {
                dpiScale = this.scaling.dpiSupport.currentDpi / this.scaling.dpiSupport.baseDpi;
            }
            
            // Calculate final font size
            const baseFontSize = font.fontSize || 11;
            const scaledFontSize = Math.round(baseFontSize * constrainedScale * dpiScale);
            
            // Ensure minimum readable font size
            scaledFont.fontSize = Math.max(8, scaledFontSize); // Back to original minimum
            
            this.logger.log("debug", this.constructor.name, 'Font scaling applied:', {
                original: baseFontSize,
                dimensionScale: constrainedScale,
                dpiScale: dpiScale,
                final: scaledFont.fontSize
            });
            
        } catch (error) {
            this.logger.log("warn", this.constructor.name, 'Error applying font scaling:', error);
        }
        
        return scaledFont;
    }

    /**
     * Convert PPTX font size to standard points
     * @param {number|string} pptxSize - PPTX font size
     * @param {string} unit - Unit type ('points', 'hundredths', 'emu')
     * @return {number} Font size in points
     * @private
     */
    _convertPptxFontSize(pptxSize, unit = 'points') {
        const size = Number(pptxSize);
        if (isNaN(size)) {return 11;} // Default fallback
        
        switch (unit) {
            case 'hundredths':
                // PPTX often stores font size as points * 100
                return Math.max(6, Math.round(size / 100));
            case 'emu':
                // EMU (English Metric Units) to points conversion
                return Math.max(6, Math.round(size / 12700));
            case 'points':
            default:
                return Math.max(6, Math.round(size));
        }
    }

    /**
     * Normalize color object
     * @param {Object|string} color - Color in various formats
     * @return {Object} Normalized color object {r, g, b}
     * @private
     */
    _normalizeColor(color) {
        if (!color) {return { r: 0, g: 0, b: 0 };}
        
        // Already normalized
        if (color.r !== undefined && color.g !== undefined && color.b !== undefined) {
            return {
                r: Math.max(0, Math.min(255, Number(color.r))),
                g: Math.max(0, Math.min(255, Number(color.g))),
                b: Math.max(0, Math.min(255, Number(color.b)))
            };
        }
        
        // Hex color string
        if (typeof color === 'string' && color.startsWith('#')) {
            const hex = color.slice(1);
            if (hex.length === 6) {
                return {
                    r: parseInt(hex.substr(0, 2), 16),
                    g: parseInt(hex.substr(2, 2), 16),
                    b: parseInt(hex.substr(4, 2), 16)
                };
            }
        }
        
        // RGB string
        if (typeof color === 'string' && color.startsWith('rgb')) {
            const matches = color.match(/\d+/g);
            if (matches && matches.length >= 3) {
                return {
                    r: Number(matches[0]),
                    g: Number(matches[1]),
                    b: Number(matches[2])
                };
            }
        }
        
        return { r: 0, g: 0, b: 0 }; // Fallback
    }

    /**
     * Normalize font configuration to ensure consistency
     * @param {Object} font - Font configuration to normalize
     * @return {Object} Normalized font configuration
     * @private
     */
    _normalizeFontConfig(font) {
        return {
            fontFamily: font.fontFamily || font.family || 'Calibri',
            fontSize: Math.max(6, Math.min(120, Number(font.fontSize || font.size || 11))),
            bold: Boolean(font.bold),
            italic: Boolean(font.italic),
            color: this._normalizeColor(font.color)
        };
    }

    /**
     * Get fallback font configuration
     * @return {Object} Safe fallback font
     * @private
     */
    _getFallbackFont() {
        return {
            fontFamily: 'Arial, sans-serif',
            fontSize: 11,
            bold: false,
            italic: false,
            color: { r: 0, g: 0, b: 0 }
        };
    }

    /**
     * Get current DPI for scaling calculations
     * @return {number} Current DPI
     * @private
     */
    _getCurrentDpi() {
        try {
            // Try to get device pixel ratio
            const devicePixelRatio = (typeof window !== 'undefined' && window.devicePixelRatio) || 1;
            return 96 * devicePixelRatio;
        } catch (error) {
            return 96; // Standard DPI fallback
        }
    }

    /**
     * Create fallback logger when Logger class is unavailable
     * @return {Object} Fallback logger
     * @private
     */
    _createFallbackLogger() {
        return {
            debug: (...args) => {},
            info: (...args) => {},
            warn: (...args) => {},
            error: (...args) => {}
        };
    }

    /**
     * Validate configuration options
     * @param {Object} options - Configuration options
     * @return {Object} Validated options
     * @private
     */
    _validateOptions(options) {
        return {
            enableResponsiveScaling: options.enableResponsiveScaling !== false,
            enableDpiScaling: options.enableDpiScaling !== false,
            enablePptxOverrides: options.enablePptxOverrides !== false,
            logLevel: options.logLevel || 'warn',
            ...options
        };
    }

    /**
     * Update default font configuration for a specific element type
     * @param {string} elementType - Chart element type
     * @param {Object} fontConfig - New font configuration
     */
    setElementDefault(elementType, fontConfig) {
        if (this.defaults[elementType]) {
            this.defaults[elementType] = {
                ...this.defaults[elementType],
                ...this._normalizeFontConfig(fontConfig)
            };
            this.logger.log("info", this.constructor.name, `Updated default font for ${elementType}`);
        } else {
            this.logger.log("warn", this.constructor.name, `Unknown element type: ${elementType}`);
        }
    }

    /**
     * Get all available element types
     * @return {Array} Array of element type names
     */
    getElementTypes() {
        return Object.keys(this.defaults);
    }

    /**
     * Create font CSS string for web rendering
     * @param {Object} font - Font configuration
     * @return {string} CSS font string
     */
    toCssString(font) {
        const normalizedFont = this._normalizeFontConfig(font);
        const style = normalizedFont.italic ? 'italic ' : '';
        const weight = normalizedFont.bold ? 'bold ' : '';
        
        return `${style}${weight}${normalizedFont.fontSize}px ${normalizedFont.fontFamily}`;
    }

    /**
     * Create font color CSS string
     * @param {Object} font - Font configuration
     * @return {string} CSS color string
     */
    toColorString(font) {
        const color = this._normalizeColor(font.color);
        return `rgb(${color.r}, ${color.g}, ${color.b})`;
    }
}

// Export classes and utilities
if (typeof window !== 'undefined') {
    window.ChartFontConfig = ChartFontConfig;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        ChartFontConfig
    };
}