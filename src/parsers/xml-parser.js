/**
 * XML Parser Module
 * Enhanced version with standard StaxParser and XML parsing functionality
 * Supports slide masters and layout extraction
 */

// Import dependencies
// import { Logger } from '../utils/utils.js';

/**
 * XML Parser Context - Enhanced with standard patterns
 * Based on standard XmlParserContext
 */
class XmlParserContext {
    constructor() {
        this.zip = null;
        this.imageMap = {};
        this.layoutsMap = {};
        this.mastersMap = {};
        this.tablesMap = {};
        this.tableStylesMap = {};
        this.smartarts = [];
        this.connectorsPr = [];

        // Enhanced master/layout tracking
        this.slideLayoutRels = new Map();  // Layout ID -> Layout object
        this.slideMasterRels = new Map();  // Master ID -> Master object
        this.masterLayoutMap = new Map();  // Master ID -> Layout IDs[]

        // Theme and color map support
        this.themeMap = {};
        this.colorMaps = {};

        this.logger = new Logger('XmlParserContext');
    }

    /**
     * Register slide master following standard patterns
     */
    registerSlideMaster(masterId, masterData) {
        this.slideMasterRels.set(masterId, masterData);
        this.mastersMap[masterId] = masterData;

        // Initialize layout array for this master
        if (!this.masterLayoutMap.has(masterId)) {
            this.masterLayoutMap.set(masterId, []);
        }

    }

    /**
     * Register slide layout following standard patterns
     */
    registerSlideLayout(layoutId, layoutData, masterId) {
        this.slideLayoutRels.set(layoutId, layoutData);
        this.layoutsMap[layoutId] = layoutData;

        // Link layout to master
        if (masterId) {
            layoutData.masterId = masterId;

            if (!this.masterLayoutMap.has(masterId)) {
                this.masterLayoutMap.set(masterId, []);
            }
            this.masterLayoutMap.get(masterId).push(layoutId);
        }

    }

    /**
     * Get slide master by ID
     */
    getSlideMaster(masterId) {
        return this.slideMasterRels.get(masterId);
    }

    /**
     * Get slide layout by ID
     */
    getSlideLayout(layoutId) {
        return this.slideLayoutRels.get(layoutId);
    }

    /**
     * Get layouts for a specific master
     */
    getLayoutsForMaster(masterId) {
        const layoutIds = this.masterLayoutMap.get(masterId) || [];
        return layoutIds.map(id => this.getSlideLayout(id)).filter(Boolean);
    }

    /**
     * Load data links (images, media) - Enhanced
     */
    loadDataLinks() {
        return this.imageMap;
    }

    /**
     * Generate SmartArts - Enhanced
     */
    generateSmartArts() {
        return this.smartarts;
    }
}

/**
 * Enhanced StAX-style XML Parser with standard patterns
 * Based on standard StaxParser with master/layout support
 */
class StaxParser {
    constructor(xml, part, context) {
        this.xml = xml;
        this.part = part;
        this.context = context;
        this.parser = new DOMParser();
        this.doc = null;
        this.currentElement = null;
        this.depth = 0;
        this.namespaces = {
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        };

        this.logger = new Logger('StaxParser');
    }

    /**
     * Parse XML document with error handling
     */
    parse() {
        try {
            this.doc = this.parser.parseFromString(this.xml, 'text/xml');

            // Check for parser errors
            const parserError = this.doc.querySelector('parsererror');
            if (parserError) {
                throw new Error('XML parsing failed: ' + parserError.textContent);
            }

            return this.doc;
        } catch (error) {
            return null;
        }
    }

    /**
     * Parse slide master from XML - standard style
     */
    parseSlideMaster(masterId) {
        if (!this.doc) {
            this.parse();
        }

        const masterElement = this.doc.documentElement;
        if (!masterElement || masterElement.localName !== 'sldMaster') {
            return null;
        }

        const master = {
            id: masterId,
            type: 'slideMaster',
            name: this.getAttributeValue(masterElement, 'name') || `Master ${masterId}`,
            preserve: this.getAttributeValue(masterElement, 'preserve') === 'true',

            // Common slide data
            cSld: this.parseCommonSlideData(masterElement),

            // Color map
            clrMap: this.parseColorMap(masterElement),

            // Text styles
            txStyles: this.parseTextStyles(masterElement),

            // Slide layouts (will be populated later)
            sldLayoutLst: [],

            // Theme reference
            themeId: this.extractThemeId(masterElement),

            // Header/Footer
            hf: this.parseHeaderFooter(masterElement),

            // Timing and transition
            timing: this.parseTiming(masterElement),
            transition: this.parseTransition(masterElement)
        };

        return master;
    }

    /**
     * Parse slide layout from XML - standard style
     */
    parseSlideLayout(layoutId, masterId) {
        if (!this.doc) {
            this.parse();
        }

        const layoutElement = this.doc.documentElement;
        if (!layoutElement || layoutElement.localName !== 'sldLayout') {
            return null;
        }

        const layout = {
            id: layoutId,
            masterId: masterId,
            type: 'slideLayout',
            name: this.getAttributeValue(layoutElement, 'name') || `Layout ${layoutId}`,
            matchingName: this.getAttributeValue(layoutElement, 'matchingName') || '',
            preserve: this.getAttributeValue(layoutElement, 'preserve') === 'true',
            showMasterSp: this.getAttributeValue(layoutElement, 'showMasterSp') !== 'false',
            showMasterPhAnim: this.getAttributeValue(layoutElement, 'showMasterPhAnim') === 'true',
            userDrawn: this.getAttributeValue(layoutElement, 'userDrawn') !== 'false',

            // Layout type
            layoutType: this.getLayoutType(layoutElement),

            // Common slide data
            cSld: this.parseCommonSlideData(layoutElement),

            // Color map override
            clrMapOvr: this.parseColorMapOverride(layoutElement),

            // Header/Footer
            hf: this.parseHeaderFooter(layoutElement),

            // Timing and transition
            timing: this.parseTiming(layoutElement),
            transition: this.parseTransition(layoutElement)
        };

        return layout;
    }

    /**
     * Parse common slide data (cSld) - standard pattern
     */
    parseCommonSlideData(parentElement) {
        const cSldElement = this.getChildElement(parentElement, 'cSld');
        if (!cSldElement) {
            return { name: '', bg: null, spTree: [] };
        }

        return {
            name: this.getAttributeValue(cSldElement, 'name') || '',
            bg: this.parseBackground(cSldElement),
            spTree: this.parseShapeTree(cSldElement)
        };
    }

    /**
     * Parse shape tree from cSld
     */
    parseShapeTree(cSldElement) {
        const spTreeElement = this.getChildElement(cSldElement, 'spTree');
        if (!spTreeElement) {
            return [];
        }

        const shapes = [];
        const children = spTreeElement.children;

        for (let i = 0; i < children.length; i++) {
            const child = children[i];
            const shape = this.parseShape(child);
            if (shape) {
                shapes.push(shape);
            }
        }

        return shapes;
    }

    /**
     * Parse individual shape element
     */
    parseShape(element) {
        const localName = element.localName;

        switch (localName) {
            case 'sp':
                return this.parseRegularShape(element);
            case 'pic':
                return this.parsePictureShape(element);
            case 'grpSp':
                return this.parseGroupShape(element);
            case 'cxnSp':
                return this.parseConnectorShape(element);
            case 'graphicFrame':
                return this.parseGraphicFrame(element);
            default:
                return null;
        }
    }

    /**
     * Parse regular shape (sp)
     */
    parseRegularShape(element) {
        const placeholder = this.parsePlaceholder(element);
        return {
            type: 'sp',
            id: this.getAttributeValue(element, 'id'),
            name: this.getShapeName(element),
            properties: this.parseShapeProperties(element),
            textBody: this.parseTextBody(element),
            style: this.parseShapeStyle(element),
            placeholder,
            isPlaceholder: !!placeholder
        };
    }

    /**
     * Parse layout type from layout element
     */
    getLayoutType(layoutElement) {
        const typeAttr = this.getAttributeValue(layoutElement, 'type');
        if (typeAttr) {
            return typeAttr;
        }

        // Try to infer from placeholders
        const cSld = this.getChildElement(layoutElement, 'cSld');
        if (cSld) {
            const spTree = this.getChildElement(cSld, 'spTree');
            if (spTree) {
                const placeholders = this.getPlaceholderTypes(spTree);
                return this.inferLayoutType(placeholders);
            }
        }

        return 'blank';
    }

    /**
     * Get placeholder types from shape tree
     */
    getPlaceholderTypes(spTreeElement) {
        const placeholders = [];
        const shapes = spTreeElement.children;

        for (let i = 0; i < shapes.length; i++) {
            const shape = shapes[i];
            const placeholder = this.parsePlaceholder(shape);
            if (placeholder) {
                placeholders.push(placeholder.type);
            }
        }

        return placeholders;
    }

    /**
     * Infer layout type from placeholders
     */
    inferLayoutType(placeholders) {
        if (placeholders.includes('title') && placeholders.includes('body')) {
            return 'titleAndContent';
        } else if (placeholders.includes('title')) {
            return 'titleOnly';
        } else if (placeholders.includes('body')) {
            return 'contentOnly';
        } else if (placeholders.length === 0) {
            return 'blank';
        }
        return 'custom';
    }

    /**
     * Parse placeholder information
     */
    parsePlaceholder(element) {
        const nvSpPr = this.getChildElement(element, 'nvSpPr') ||
                       this.getChildElement(element, 'nvPicPr') ||
                       this.getChildElement(element, 'nvGrpSpPr');

        if (!nvSpPr) {return null;}

        const nvPr = this.getChildElement(nvSpPr, 'nvPr');
        if (!nvPr) {return null;}

        const ph = this.getChildElement(nvPr, 'ph');
        if (!ph) {return null;}

        // Normalize idx to number when present; leave undefined when missing
        const idxAttr = this.getAttributeValue(ph, 'idx');
        const normalizedIdx = (idxAttr === null || idxAttr === undefined) ? undefined : (isNaN(parseInt(idxAttr, 10)) ? undefined : parseInt(idxAttr, 10));

        return {
            type: this.getAttributeValue(ph, 'type') || 'obj',
            idx: normalizedIdx,
            orient: this.getAttributeValue(ph, 'orient'),
            sz: this.getAttributeValue(ph, 'sz')
        };
    }

    /**
     * Parse text styles for master
     */
    parseTextStyles(masterElement) {
        const txStylesElement = this.getChildElement(masterElement, 'txStyles');
        if (!txStylesElement) {return null;}

        return {
            titleStyle: this.parseTextStyle(this.getChildElement(txStylesElement, 'titleStyle')),
            bodyStyle: this.parseTextStyle(this.getChildElement(txStylesElement, 'bodyStyle')),
            otherStyle: this.parseTextStyle(this.getChildElement(txStylesElement, 'otherStyle'))
        };
    }

    /**
     * Parse individual text style
     */
    parseTextStyle(styleElement) {
        if (!styleElement) {return null;}

        // This would parse detailed text style information
        // For now, return a simplified structure
        return {
            element: styleElement,
            parsed: false  // Flag for lazy parsing
        };
    }

    /**
     * Parse color map
     */
    parseColorMap(element) {
        const clrMapElement = this.getChildElement(element, 'clrMap');
        if (!clrMapElement) {return null;}

        const colorMap = {};
        const attributes = clrMapElement.attributes;

        for (let i = 0; i < attributes.length; i++) {
            const attr = attributes[i];
            colorMap[attr.name] = attr.value;
        }

        return colorMap;
    }

    /**
     * Parse color map override
     */
    parseColorMapOverride(element) {
        const clrMapOvrElement = this.getChildElement(element, 'clrMapOvr');
        if (!clrMapOvrElement) {return null;}

        const masterClrMapping = this.getChildElement(clrMapOvrElement, 'masterClrMapping');
        if (masterClrMapping) {
            return this.parseColorMap(masterClrMapping);
        }

        const overrideClrMapping = this.getChildElement(clrMapOvrElement, 'overrideClrMapping');
        if (overrideClrMapping) {
            return this.parseColorMap(overrideClrMapping);
        }

        return null;
    }

    /**
     * Parse background
     */
    parseBackground(element) {
        const bgElement = this.getChildElement(element, 'bg');
        if (!bgElement) {return null;}

        const bgPr = this.getChildElement(bgElement, 'bgPr');
        if (bgPr) {
            return {
                type: 'bgPr',
                fill: this.parseFill(bgPr)
            };
        }

        const bgRef = this.getChildElement(bgElement, 'bgRef');
        if (bgRef) {
            return {
                type: 'bgRef',
                idx: this.getAttributeValue(bgRef, 'idx')
            };
        }

        return null;
    }

    /**
     * Utility methods
     */
    getChildElement(parent, tagName) {
        if (!parent) {return null;}

        // Try with namespace
        for (const ns of Object.keys(this.namespaces)) {
            const nsElement = parent.querySelector(`${ns}\\:${tagName}, ${tagName}`);
            if (nsElement) {return nsElement;}
        }

        // Try without namespace
        return parent.querySelector(tagName);
    }

    getAttributeValue(element, attributeName) {
        return element ? element.getAttribute(attributeName) : null;
    }

    getTextContent(element) {
        return element ? element.textContent : '';
    }

    getElementsByTagName(tagName) {
        if (!this.doc) {
            this.parse();
        }
        return this.doc ? this.doc.getElementsByTagName(tagName) : [];
    }

    getElementsByTagNameNS(namespace, localName) {
        if (!this.doc) {
            this.parse();
        }
        return this.doc ? this.doc.getElementsByTagNameNS(namespace, localName) : [];
    }

    // Placeholder methods for complex parsing (to be implemented)
    parseShapeProperties(element) { return {}; }
    parseTextBody(element) { return null; }
    parseShapeStyle(element) { return null; }
    getShapeName(element) { 
        // Look for cNvPr element which contains the name (handle namespaces)
        const cNvPr = element.querySelector('cNvPr') || 
                      element.querySelector('p\\:cNvPr') ||
                      element.querySelector('[*|localName="cNvPr"]');
        if (cNvPr) {
            return cNvPr.getAttribute('name') || '';
        }
        return ''; 
    }
    parsePictureShape(element) { 
        const ph = this.parsePlaceholder(element);
        const shape = {
            type: 'pic',
            id: this.getAttributeValue(element, 'id'),
            name: this.getShapeName(element),
            properties: this.parseShapeProperties(element),
            style: this.parseShapeStyle(element),
            placeholder: ph,
            isPlaceholder: !!ph
        };

        // Parse image relationship ID from blipFill/blip element
        const blipFillElement = element.querySelector('blipFill, p\\:blipFill');
        if (blipFillElement) {
            const blipElement = blipFillElement.querySelector('blip, a\\:blip');
            if (blipElement) {
                // First check for SVG alternative in extension list (priority over PNG fallback)
                const svgRelId = this.extractSVGRelationshipId(blipElement);
                
                if (svgRelId) {
                    // Use SVG version if available
                    shape.imageRelId = svgRelId;
                    shape.hasSVGAlternative = true; // Flag to indicate SVG preference
                } else {
                    // Fallback to PNG/standard relationship ID
                    const embedId = blipElement.getAttribute('r:embed') || blipElement.getAttribute('r:id');
                    if (embedId) {
                        shape.imageRelId = embedId;
                    }
                }

                // Parse image effects if present
                const effectLstElement = blipElement.querySelector('effectLst, a\\:effectLst');
                if (effectLstElement) {
                    shape.imageEffects = this.parseImageEffects(effectLstElement);
                }

                // Parse image fill mode
                const fillModeAttr = blipFillElement.getAttribute('dpi') || blipFillElement.getAttribute('rotWithShape');
                if (fillModeAttr) {
                    shape.imageFillMode = fillModeAttr;
                }
            }
        }

        // Parse transform (position and size)
        shape.transform = this.parseTransform(element);
        
        return shape;
    }

    /**
     * Extract SVG relationship ID from extension list
     * Looks for asvg:svgBlip elements in a:extLst extensions
     * @param {Element} blipElement - The blip element containing potential SVG alternatives
     * @returns {string|null} SVG relationship ID if found, null otherwise
     */
    extractSVGRelationshipId(blipElement) {
        try {
            // Look for extension list (a:extLst) within the blip element
            const extLstElement = blipElement.querySelector('extLst, a\\:extLst, [*|localName="extLst"]');
            if (!extLstElement) {
                return null;
            }

            // Look for extensions with the SVG URI (more flexible matching)
            const allExtensions = extLstElement.querySelectorAll('ext, a\\:ext, [*|localName="ext"]');
            
            for (const ext of allExtensions) {
                const uri = ext.getAttribute('uri');
                
                if (uri === '{96DAC541-7B7A-43D3-8B79-37D633B846F1}') {
                    
                    // Look for asvg:svgBlip within the extension (flexible namespace matching)
                    const svgBlipElement = ext.querySelector('svgBlip, asvg\\:svgBlip, [*|localName="svgBlip"]');
                    
                    if (svgBlipElement) {
                        // Extract the SVG relationship ID
                        const svgRelId = svgBlipElement.getAttribute('r:embed') || svgBlipElement.getAttribute('r:id');
                        if (svgRelId) {
                            return svgRelId;
                        } else {
                        }
                    } else {
                    }
                }
            }

            return null;
        } catch (error) {
            return null;
        }
    }

    parseGroupShape(element) { return { type: 'grpSp' }; }
    parseConnectorShape(element) { return { type: 'cxnSp' }; }
    parseGraphicFrame(element) { 
        const shape = { 
            type: 'graphicFrame',
            name: this.getShapeName(element),
            position: this.parseTransform(element)
        };
        
        // Look for graphic data to identify content type (handle namespaces)
        const graphic = element.querySelector('graphic') || 
                       element.querySelector('a\\:graphic') || 
                       element.querySelector('[*|localName="graphic"]');
        if (graphic) {
            const graphicData = graphic.querySelector('graphicData') || 
                               graphic.querySelector('a\\:graphicData') || 
                               graphic.querySelector('[*|localName="graphicData"]');
            if (graphicData) {
                const uri = graphicData.getAttribute('uri');
                shape.graphicData = { uri: uri };
                
                // If this is a table, extract the table XML
                if (uri === 'http://schemas.openxmlformats.org/drawingml/2006/table') {
                    const tableElement = graphicData.querySelector('tbl');
                    if (tableElement) {
                        // Serialize the table XML for later parsing
                        shape.graphicData.tableXml = new XMLSerializer().serializeToString(tableElement);
                    }
                }
                // If this is a chart, parse chart data immediately
                else if (uri === 'http://schemas.openxmlformats.org/drawingml/2006/chart') {
                    
                    // Store the graphicData element for later processing
                    shape.graphicData.element = graphicData;
                    
                    // Extract chart reference information
                    const chartElement = graphicData.querySelector('chart, c\\:chart');
                    if (chartElement) {
                        const rId = chartElement.getAttribute('r:id');
                        if (rId) {
                            shape.graphicData.chartRef = rId;
                        }
                    }
                    
                    // Try to parse chart data if ChartProcessor is available
                    if (window.ChartProcessor) {
                        try {
                            const chartProcessor = new ChartProcessor(this.context);
                            
                            // Try async parsing (for relationship-based charts)
                            chartProcessor.parseChartFromGraphicFrame(element).then(chartData => {
                                if (chartData) {
                                    shape.chartData = chartData;
                                    
                                    // Trigger re-render if shape is already processed
                                    if (window.currentProcessor && window.currentProcessor.reRenderShape) {
                                        window.currentProcessor.reRenderShape(shape);
                                    }
                                } else {
                                    
                                    // Fallback: try embedded chart parsing
                                    const embeddedData = chartProcessor.parseEmbeddedChartData(graphicData);
                                    if (embeddedData) {
                                        shape.chartData = embeddedData;
                                    }
                                }
                            }).catch(error => {
                                
                                // Fallback: try embedded chart parsing
                                try {
                                    const embeddedData = chartProcessor.parseEmbeddedChartData(graphicData);
                                    if (embeddedData) {
                                        shape.chartData = embeddedData;
                                    }
                                } catch (embeddedError) {
                                }
                            });
                            
                        } catch (error) {
                        }
                    } else {
                        
                        // Store basic chart information for later processing
                        shape.graphicData.isChart = true;
                        shape.graphicData.chartXml = new XMLSerializer().serializeToString(graphicData);
                    }
                }
            }
        }
        
        return shape;
    }
    parseHeaderFooter(element) { return null; }
    parseTiming(element) { return null; }
    parseTransition(element) { return null; }
    extractThemeId(element) { return null; }
    parseFill(element) { return null; }
    parseImageEffects(element) { return null; }
    
    /**
     * Parse transform (xfrm) element to extract position and size
     */
    parseTransform(element) {
        const xfrm = element.querySelector('xfrm') || 
                     element.querySelector('p\\:xfrm') || 
                     element.querySelector('[*|localName="xfrm"]');
        if (!xfrm) {
            return { x: 0, y: 0, width: 0, height: 0 };
        }
        
        // Parse offset (position) - handle namespaces
        const off = xfrm.querySelector('off') || 
                    xfrm.querySelector('a\\:off') || 
                    xfrm.querySelector('[*|localName="off"]');
        let x = 0, y = 0;
        if (off) {
            x = parseInt(off.getAttribute('x') || '0', 10);
            y = parseInt(off.getAttribute('y') || '0', 10);
        }
        
        // Parse extent (size) - handle namespaces
        const ext = xfrm.querySelector('ext') || 
                    xfrm.querySelector('a\\:ext') || 
                    xfrm.querySelector('[*|localName="ext"]');
        let width = 0, height = 0;
        if (ext) {
            width = parseInt(ext.getAttribute('cx') || '0', 10);
            height = parseInt(ext.getAttribute('cy') || '0', 10);
        }
        
        return { x, y, width, height };
    }
    
    /**
     * Parse slide from slide element
     * This method is expected by chart integration tests
     * @param {Element} slideElement - The slide element
     * @returns {Object} Parsed slide with shapes array
     */
    parseSlide(slideElement) {
        try {
            const slide = {
                shapes: [],
                background: null,
                name: '',
                id: null
            };
            
            // Get slide ID if available
            const slideId = slideElement.getAttribute('id');
            if (slideId) {
                slide.id = slideId;
            }
            
            // Find common slide data
            const cSldElement = this.getChildElement(slideElement, 'cSld');
            if (cSldElement) {
                // Get slide name
                const nameAttr = cSldElement.getAttribute('name');
                if (nameAttr) {
                    slide.name = nameAttr;
                }
                
                // Parse background
                slide.background = this.parseBackground(cSldElement);
                
                // Parse shape tree
                slide.shapes = this.parseShapeTree(cSldElement);
            }
            
            return slide;
        } catch (error) {
            return {
                shapes: [],
                background: null,
                name: '',
                id: null
            };
        }
    }
}

// Export classes (maintain backward compatibility)
// Expose to both window and globalThis for bundler/UMD environments
if (typeof window !== 'undefined') {
    window.XmlParserContext = XmlParserContext;
    window.StaxParser = StaxParser;
}
if (typeof globalThis !== 'undefined') {
    globalThis.XmlParserContext = XmlParserContext;
    globalThis.StaxParser = StaxParser;
}

// Intentionally no ES module exports to support classic <script> usage in root demo
