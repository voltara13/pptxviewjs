/**
 * Theme Processor Module
 * Handles PPTX themes, master slides, and layout processing
 * Ensures complete PptxGenJS compatibility with theme and layout support
 */

/**
 * Theme Information Structure
 */
class ThemeInfo {
    constructor() {
        this.name = '';
        this.colorScheme = {};        // Theme color scheme
        this.fontScheme = {};         // Theme font scheme
        this.formatScheme = {};       // Theme format scheme
        this.backgroundStyles = [];   // Background fill styles
        this.effects = [];           // Theme effects
        this.themeElements = {};     // Raw theme elements
    }
}

/**
 * Color Scheme Structure
 */
class ColorScheme {
    constructor() {
        this.dk1 = null;    // Dark 1 (text/background)
        this.lt1 = null;    // Light 1 (text/background)
        this.dk2 = null;    // Dark 2 (text/background)
        this.lt2 = null;    // Light 2 (text/background)
        this.accent1 = null; // Accent 1
        this.accent2 = null; // Accent 2
        this.accent3 = null; // Accent 3
        this.accent4 = null; // Accent 4
        this.accent5 = null; // Accent 5
        this.accent6 = null; // Accent 6
        this.hlink = null;   // Hyperlink
        this.folHlink = null; // Followed hyperlink
    }
}

/**
 * Font Scheme Structure
 */
class FontScheme {
    constructor() {
        this.majorFont = {};  // Heading fonts
        this.minorFont = {};  // Body fonts
    }
}

/**
 * Master Slide Structure
 */
class MasterSlide {
    constructor() {
        this.id = '';
        this.name = '';
        this.preserve = false;
        this.cSld = null;           // Common slide data
        this.clrMap = null;         // Color map
        this.txStyles = null;       // Text styles
        this.sldLayoutLst = [];     // Associated layouts
        this.themeId = '';          // Theme reference
        this.hf = null;             // Header/footer
        this.timing = null;         // Timing information
        this.transition = null;     // Transition effects
    }
}

/**
 * Layout Slide Structure
 */
class LayoutSlide {
    constructor() {
        this.id = '';
        this.masterId = '';
        this.name = '';
        this.type = 'custom';       // Layout type
        this.preserve = false;
        this.showMasterSp = true;   // Show master shapes
        this.showMasterPhAnim = false; // Show master placeholder animations
        this.userDrawn = false;     // User drawn layout
        this.cSld = null;           // Common slide data
        this.clrMapOvr = null;      // Color map override
        this.hf = null;             // Header/footer
        this.timing = null;         // Timing information
        this.transition = null;     // Transition effects
    }
}

/**
 * Theme Processor - Main class for theme and master slide processing
 */
class ThemeProcessor {
    constructor(context) {
        this.context = context;
        this.logger = new Logger('ThemeProcessor');
        this.themes = new Map();
        this.masters = new Map();
        this.layouts = new Map();
    }

    /**
     * Process theme from XML
     * @param {string} themeXml - Theme XML content
     * @param {string} themeId - Theme ID
     * @return {ThemeInfo} Processed theme information
     */
    processTheme(themeXml, themeId) {
        try {
            const parser = new DOMParser();
            const themeDoc = parser.parseFromString(themeXml, 'text/xml');
            
            const themeElement = themeDoc.querySelector('theme, a\\:theme');
            if (!themeElement) {
                this.logger.log("warn", this.constructor.name, 'No theme element found in XML');
                return this.createDefaultTheme(themeId);
            }

            const theme = new ThemeInfo();
            theme.name = themeElement.getAttribute('name') || `Theme${themeId}`;

            // Process theme elements
            this.processThemeElements(themeElement, theme);
            this.processColorScheme(themeElement, theme);
            this.processFontScheme(themeElement, theme);
            this.processFormatScheme(themeElement, theme);

            // Store theme
            this.themes.set(themeId, theme);
            
            return theme;

        } catch (error) {
            this.logger.logError(this.constructor.name, 'Error processing theme:', error);
            return this.createDefaultTheme(themeId);
        }
    }

    /**
     * Process master slide from XML
     * @param {string} masterXml - Master slide XML content
     * @param {string} masterId - Master slide ID
     * @return {MasterSlide} Processed master slide
     */
    processMasterSlide(masterXml, masterId) {
        try {
            const parser = new DOMParser();
            const masterDoc = parser.parseFromString(masterXml, 'text/xml');
            
            const masterElement = masterDoc.querySelector('sldMaster, p\\:sldMaster');
            if (!masterElement) {
                this.logger.log("warn", this.constructor.name, 'No slide master element found in XML');
                return this.createDefaultMaster(masterId);
            }

            const master = new MasterSlide();
            master.id = masterId;
            master.name = masterElement.getAttribute('name') || `Master${masterId}`;
            master.preserve = masterElement.getAttribute('preserve') === 'true';

            // Process master elements
            this.processMasterElements(masterElement, master);

            // Store master
            this.masters.set(masterId, master);
            
            return master;

        } catch (error) {
            this.logger.logError(this.constructor.name, 'Error processing master slide:', error);
            return this.createDefaultMaster(masterId);
        }
    }

    /**
     * Process layout slide from XML
     * @param {string} layoutXml - Layout slide XML content
     * @param {string} layoutId - Layout slide ID
     * @param {string} masterId - Associated master slide ID
     * @return {LayoutSlide} Processed layout slide
     */
    processLayoutSlide(layoutXml, layoutId, masterId) {
        try {
            const parser = new DOMParser();
            const layoutDoc = parser.parseFromString(layoutXml, 'text/xml');
            
            const layoutElement = layoutDoc.querySelector('sldLayout, p\\:sldLayout');
            if (!layoutElement) {
                this.logger.log("warn", this.constructor.name, 'No slide layout element found in XML');
                return this.createDefaultLayout(layoutId, masterId);
            }

            const layout = new LayoutSlide();
            layout.id = layoutId;
            layout.masterId = masterId;
            layout.name = layoutElement.getAttribute('name') || `Layout${layoutId}`;
            layout.type = layoutElement.getAttribute('type') || 'custom';
            layout.preserve = layoutElement.getAttribute('preserve') === 'true';
            layout.showMasterSp = layoutElement.getAttribute('showMasterSp') !== 'false';
            layout.showMasterPhAnim = layoutElement.getAttribute('showMasterPhAnim') === 'true';
            layout.userDrawn = layoutElement.getAttribute('userDrawn') === 'true';

            // Process layout elements
            this.processLayoutElements(layoutElement, layout);

            // Store layout
            this.layouts.set(layoutId, layout);
            
            // Associate with master
            const master = this.masters.get(masterId);
            if (master) {
                master.sldLayoutLst.push(layoutId);
            }
            
            return layout;

        } catch (error) {
            this.logger.logError(this.constructor.name, 'Error processing layout slide:', error);
            return this.createDefaultLayout(layoutId, masterId);
        }
    }

    /**
     * Process theme elements
     * @param {Element} themeElement - Theme element
     * @param {ThemeInfo} theme - Theme to populate
     */
    processThemeElements(themeElement, theme) {
        const themeElements = themeElement.querySelector('themeElements, a\\:themeElements');
        if (themeElements) {
            theme.themeElements = {
                element: themeElements,
                parsed: false // Lazy parsing flag
            };
        }
    }

    /**
     * Process color scheme
     * @param {Element} themeElement - Theme element
     * @param {ThemeInfo} theme - Theme to populate
     */
    processColorScheme(themeElement, theme) {
        const clrScheme = themeElement.querySelector('clrScheme, a\\:clrScheme');
        if (!clrScheme) {return;}

        const colorScheme = new ColorScheme();

        // Process standard colors
        const colorMappings = {
            'dk1': 'dk1',
            'lt1': 'lt1', 
            'dk2': 'dk2',
            'lt2': 'lt2',
            'accent1': 'accent1',
            'accent2': 'accent2',
            'accent3': 'accent3',
            'accent4': 'accent4',
            'accent5': 'accent5',
            'accent6': 'accent6',
            'hlink': 'hlink',
            'folHlink': 'folHlink'
        };

        for (const [xmlName, propName] of Object.entries(colorMappings)) {
            const colorElement = clrScheme.querySelector(`${xmlName}, a\\:${xmlName}`);
            if (colorElement) {
                colorScheme[propName] = this.parseColor(colorElement);
            }
        }

        theme.colorScheme = colorScheme;
    }

    /**
     * Process font scheme
     * @param {Element} themeElement - Theme element
     * @param {ThemeInfo} theme - Theme to populate
     */
    processFontScheme(themeElement, theme) {
        const fontScheme = themeElement.querySelector('fontScheme, a\\:fontScheme');
        if (!fontScheme) {return;}

        const fonts = new FontScheme();

        // Process major font (headings)
        const majorFont = fontScheme.querySelector('majorFont, a\\:majorFont');
        if (majorFont) {
            fonts.majorFont = this.parseFontCollection(majorFont);
        }

        // Process minor font (body)
        const minorFont = fontScheme.querySelector('minorFont, a\\:minorFont');
        if (minorFont) {
            fonts.minorFont = this.parseFontCollection(minorFont);
        }

        theme.fontScheme = fonts;
    }

    /**
     * Process format scheme
     * @param {Element} themeElement - Theme element
     * @param {ThemeInfo} theme - Theme to populate
     */
    processFormatScheme(themeElement, theme) {
        const fmtScheme = themeElement.querySelector('fmtScheme, a\\:fmtScheme');
        if (fmtScheme) {
            // Process fill styles, line styles, effect styles
            theme.formatScheme = {
                fillStyleLst: this.processFillStyles(fmtScheme),
                lnStyleLst: this.processLineStyles(fmtScheme),
                effectStyleLst: this.processEffectStyles(fmtScheme),
                bgFillStyleLst: this.processBackgroundFillStyles(fmtScheme)
            };
        }
    }

    /**
     * Process master slide elements
     * @param {Element} masterElement - Master slide element
     * @param {MasterSlide} master - Master to populate
     */
    processMasterElements(masterElement, master) {
        // Process common slide data
        const cSld = masterElement.querySelector('cSld, p\\:cSld');
        if (cSld) {
            master.cSld = this.parseCommonSlideData(cSld);
        }

        // Process color map
        const clrMap = masterElement.querySelector('clrMap, p\\:clrMap');
        if (clrMap) {
            master.clrMap = this.parseColorMap(clrMap);
        }

        // Process text styles
        const txStyles = masterElement.querySelector('txStyles, p\\:txStyles');
        if (txStyles) {
            master.txStyles = this.parseTextStyles(txStyles);
        }

        // Process other elements
        master.hf = this.parseHeaderFooter(masterElement);
        master.timing = this.parseTiming(masterElement);
        master.transition = this.parseTransition(masterElement);
    }

    /**
     * Process layout slide elements
     * @param {Element} layoutElement - Layout slide element
     * @param {LayoutSlide} layout - Layout to populate
     */
    processLayoutElements(layoutElement, layout) {
        // Process common slide data
        const cSld = layoutElement.querySelector('cSld, p\\:cSld');
        if (cSld) {
            layout.cSld = this.parseCommonSlideData(cSld);
        }

        // Process color map override
        const clrMapOvr = layoutElement.querySelector('clrMapOvr, p\\:clrMapOvr');
        if (clrMapOvr) {
            layout.clrMapOvr = this.parseColorMapOverride(clrMapOvr);
        }

        // Process other elements
        layout.hf = this.parseHeaderFooter(layoutElement);
        layout.timing = this.parseTiming(layoutElement);
        layout.transition = this.parseTransition(layoutElement);
    }

    /**
     * Get theme by ID
     * @param {string} themeId - Theme ID
     * @return {ThemeInfo|null} Theme information
     */
    getTheme(themeId) {
        return this.themes.get(themeId) || null;
    }

    /**
     * Get master slide by ID
     * @param {string} masterId - Master slide ID
     * @return {MasterSlide|null} Master slide
     */
    getMaster(masterId) {
        return this.masters.get(masterId) || null;
    }

    /**
     * Get layout slide by ID
     * @param {string} layoutId - Layout slide ID
     * @return {LayoutSlide|null} Layout slide
     */
    getLayout(layoutId) {
        return this.layouts.get(layoutId) || null;
    }

    /**
     * Resolve color reference using theme and color map
     * @param {string} colorRef - Color reference (e.g., 'accent1', 'dk1')
     * @param {string} themeId - Theme ID
     * @param {Object} colorMap - Color map overrides
     * @return {Object|null} Resolved color
     */
    resolveColor(colorRef, themeId, colorMap = null) {
        const theme = this.getTheme(themeId);
        if (!theme || !theme.colorScheme) {return null;}

        // Apply color map if provided
        const mappedRef = colorMap && colorMap[colorRef] ? colorMap[colorRef] : colorRef;
        
        return theme.colorScheme[mappedRef] || null;
    }

    /**
     * Resolve font reference using theme
     * @param {string} fontRef - Font reference ('major' or 'minor')
     * @param {string} themeId - Theme ID
     * @param {string} script - Script/language (optional)
     * @return {string|null} Font family name
     */
    resolveFont(fontRef, themeId, script = 'latin') {
        const theme = this.getTheme(themeId);
        if (!theme || !theme.fontScheme) {return null;}

        const fontCollection = fontRef === 'major' ? theme.fontScheme.majorFont : theme.fontScheme.minorFont;
        
        return fontCollection[script] || fontCollection.latin || null;
    }

    // Helper methods (simplified implementations)
    parseColor(colorElement) {
        // Simplified color parsing - could be expanded
        const sysClr = colorElement.querySelector('sysClr, a\\:sysClr');
        const srgbClr = colorElement.querySelector('srgbClr, a\\:srgbClr');
        const schemeClr = colorElement.querySelector('schemeClr, a\\:schemeClr');

        if (srgbClr) {
            const val = srgbClr.getAttribute('val');
            if (val) {
                const r = parseInt(val.substr(0, 2), 16);
                const g = parseInt(val.substr(2, 2), 16);
                const b = parseInt(val.substr(4, 2), 16);
                return { r, g, b };
            }
        }

        if (sysClr) {
            const val = sysClr.getAttribute('val');
            // Map system colors to RGB values (simplified)
            const sysColors = {
                'windowText': { r: 0, g: 0, b: 0 },
                'window': { r: 255, g: 255, b: 255 }
            };
            return sysColors[val] || { r: 0, g: 0, b: 0 };
        }

        return { r: 0, g: 0, b: 0 }; // Default black
    }

    parseFontCollection(fontElement) {
        const fonts = {};
        
        // Parse Latin font
        const latin = fontElement.querySelector('latin, a\\:latin');
        if (latin) {
            fonts.latin = latin.getAttribute('typeface');
        }

        // Parse other script fonts (simplified)
        const cs = fontElement.querySelector('cs, a\\:cs');
        if (cs) {
            fonts.cs = cs.getAttribute('typeface');
        }

        return fonts;
    }

    parseCommonSlideData(cSldElement) {
        return {
            name: cSldElement.getAttribute('name') || '',
            bg: this.parseBackground(cSldElement),
            spTree: this.parseShapeTree(cSldElement)
        };
    }

    parseColorMap(clrMapElement) {
        const colorMap = {};
        const attributes = clrMapElement.attributes;
        
        for (let i = 0; i < attributes.length; i++) {
            const attr = attributes[i];
            colorMap[attr.name] = attr.value;
        }
        
        return colorMap;
    }

    parseColorMapOverride(clrMapOvrElement) {
        // Simplified color map override parsing
        return {};
    }

    parseTextStyles(txStylesElement) {
        return {
            titleStyle: {},
            bodyStyle: {},
            otherStyle: {}
        };
    }

    parseShapeTree(cSldElement) {
        // Would integrate with existing shape parsing
        return [];
    }

    parseBackground(cSldElement) {
        // Would integrate with existing background parsing
        return null;
    }

    parseHeaderFooter(element) { return null; }
    parseTiming(element) { return null; }
    parseTransition(element) { return null; }
    processFillStyles(fmtElement) { return []; }
    processLineStyles(fmtElement) { return []; }
    processEffectStyles(fmtElement) { return []; }
    processBackgroundFillStyles(fmtElement) { return []; }

    // Default creation methods
    createDefaultTheme(themeId) {
        const theme = new ThemeInfo();
        theme.name = `Default Theme ${themeId}`;
        theme.colorScheme = new ColorScheme();
        theme.fontScheme = new FontScheme();
        return theme;
    }

    createDefaultMaster(masterId) {
        const master = new MasterSlide();
        master.id = masterId;
        master.name = `Default Master ${masterId}`;
        return master;
    }

    createDefaultLayout(layoutId, masterId) {
        const layout = new LayoutSlide();
        layout.id = layoutId;
        layout.masterId = masterId;
        layout.name = `Default Layout ${layoutId}`;
        layout.type = 'blank';
        return layout;
    }
}

// Export classes

// Export classes (maintain backward compatibility)
if (typeof window !== 'undefined') {
    window.ThemeInfo = ThemeInfo;
    window.ColorScheme = ColorScheme;
    window.FontScheme = FontScheme;
    window.MasterSlide = MasterSlide;
    window.LayoutSlide = LayoutSlide;
    window.ThemeProcessor = ThemeProcessor;
}

// ES Module exports (disabled for script-tag compatibility)
// export { ThemeInfo,ColorScheme,FontScheme,MasterSlide,LayoutSlide,ThemeProcessor };
