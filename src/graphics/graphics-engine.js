/**
 * Enhanced Graphics Engine Module
 * Based on Standard CGraphics and shape rendering system
 * Supports comprehensive shape types, group shapes, and advanced rendering
 * 
 * Copyright 2025 gptsci.com
 * 
 * Licensed under the MIT License
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *     http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Standard Formula Types - Based on Standard SDKJS implementation
 */
const FORMULA_TYPE_MULT_DIV = 0;     // */
const FORMULA_TYPE_PLUS_MINUS = 1;   // +-
const FORMULA_TYPE_PLUS_DIV = 2;     // +/
const FORMULA_TYPE_IF_ELSE = 3;      // ?:
const FORMULA_TYPE_ABS = 4;          // abs
const FORMULA_TYPE_AT2 = 5;          // at2 (atan2)
const FORMULA_TYPE_CAT2 = 6;         // cat2
const FORMULA_TYPE_COS = 7;          // cos
const FORMULA_TYPE_MAX = 8;          // max
const FORMULA_TYPE_MIN = 9;          // min
const FORMULA_TYPE_MOD = 10;         // mod (sqrt(x²+y²+z²))
const FORMULA_TYPE_PIN = 11;         // pin
const FORMULA_TYPE_SAT2 = 12;        // sat2
const FORMULA_TYPE_SIN = 13;         // sin
const FORMULA_TYPE_SQRT = 14;        // sqrt
const FORMULA_TYPE_TAN = 15;         // tan
const FORMULA_TYPE_VALUE = 16;       // val

/**
 * Standard geometry constants (based on Standard)
 */
const GEOMETRY_CONSTANTS = {
    // Basic dimensions
    w: 'w',      // shape width
    h: 'h',      // shape height
    ss: 'ss',    // min(w, h)
    ls: 'ls',    // max(w, h)
    
    // Centers
    hc: 'hc',    // horizontal center (w/2)
    vc: 'vc',    // vertical center (h/2)
    
    // Edges
    l: 'l',      // left (0)
    t: 't',      // top (0)
    r: 'r',      // right (w)
    b: 'b',      // bottom (h)
    
    // Quarters
    wd2: 'wd2',  // w/2
    hd2: 'hd2',  // h/2
    wd4: 'wd4',  // w/4
    hd4: 'hd4',  // h/4
    wd6: 'wd6',  // w/6
    hd6: 'hd6',  // h/6
    wd8: 'wd8',  // w/8
    hd8: 'hd8',  // h/8
    
    // Angular constants (in 60000ths of a degree)
    cd2: 'cd2',     // 10800000 (180°)
    cd4: 'cd4',     // 5400000 (90°)
    cd8: 'cd8',     // 2700000 (45°)
    _3cd4: '_3cd4', // 16200000 (270°)
    _3cd8: '_3cd8', // 8100000 (135°)
    _5cd8: '_5cd8', // 13500000 (225°)
    _7cd8: '_7cd8'  // 18900000 (315°)
};

/**
 * Standard Shape Type Mapping
 * Maps PowerPoint preset geometry types to Standard shape types
 */
const StandardShapeTypes = {
    // Basic shapes
    'rect': 1,
    'roundRect': 2,
    'ellipse': 3,
    'diamond': 4,
    'triangle': 5,
    'rtTriangle': 6,
    'parallelogram': 7,
    'trapezoid': 8,
    'hexagon': 9,
    'octagon': 10,
    'plus': 11,
    'star5': 12,
    'rightArrow': 13,
    'leftArrow': 66,
    'upArrow': 68,
    'downArrow': 67,
    'leftRightArrow': 69,
    'upDownArrow': 70,
    'quadArrow': 76,
    'heart': 74,
    'lightningBolt': 73,
    'can': 22,
    'cube': 16,
    'donut': 23,
    'arc': 19,
    'line': 20,
    'plaque': 21,
    'bevel': 1000, // Custom
    'frame': 75,
    'pentagon': 56,
    'star4': 1000,
    'star6': 1000,
    'star8': 58,
    'star12': 1000,
    'star16': 59,
    'star24': 1000,
    'star32': 60,
    'teardrop': 1000, // Custom
    'flowChartProcess': 1000,
    'flowChartDecision': 4, // diamond
    'flowChartTerminator': 2, // roundRect
    'flowChartInputOutput': 7, // parallelogram
    'flowChartDocument': 1000,
    'flowChartMultidocument': 1000,
    'flowChartPreparation': 1000,
    'flowChartManualInput': 1000,
    'flowChartManualOperation': 1000,
    'flowChartConnector': 3, // ellipse
    'flowChartOffpageConnector': 1000,
    'flowChartPunchedCard': 1000,
    'flowChartPunchedTape': 1000,
    'flowChartSummingJunction': 1000,
    'flowChartOr': 1000,
    'flowChartCollate': 1000,
    'flowChartSort': 1000,
    'flowChartExtract': 1000,
    'flowChartMerge': 1000,
    'flowChartStoredData': 1000,
    'flowChartDelay': 1000,
    'flowChartSequentialAccessStorage': 1000,
    'flowChartMagneticDisk': 1000,
    'flowChartDirectAccessStorage': 1000,
    'flowChartDisplay': 1000,
    'flowChartInternalStorage': 1000,
    'flowChartAlternateProcess': 1000,
    'flowChartMagneticDrum': 1000,
    'flowChartMagneticTape': 1000,
    'callout1': 1000,
    'callout2': 1000,
    'callout3': 1000,
    'accentCallout1': 1000,
    'accentCallout2': 1000,
    'accentCallout3': 1000,
    'borderCallout1': 1000,
    'borderCallout2': 1000,
    'borderCallout3': 1000,
    'accentBorderCallout1': 1000,
    'accentBorderCallout2': 1000,
    'accentBorderCallout3': 1000,
    'ribbon': 1000,
    'ribbon2': 1000,
    'ellipseRibbon': 1000,
    'ellipseRibbon2': 1000,
    'leftRightRibbon': 1000,
    'verticalScroll': 1000,
    'horizontalScroll': 1000,
    'wave': 1000,
    'doubleWave': 1000,
    'bentArrow': 1000,
    'uturnArrow': 1000,
    'circularArrow': 1000,
    'leftUpArrow': 1000,
    'bentUpArrow': 1000,
    'curvedRightArrow': 1000,
    'curvedLeftArrow': 1000,
    'curvedUpArrow': 1000,
    'curvedDownArrow': 1000,
    'stripedRightArrow': 1000,
    'notchedRightArrow': 1000,
    'blockArc': 1000,
    'swooshArrow': 1000,
    'leftBracket': 1000,
    'rightBracket': 1000,
    'leftBrace': 1000,
    'rightBrace': 1000,
    'bracketPair': 1000,
    'bracePair': 1000,
    'straightConnector1': 20, // line
    'bentConnector2': 1000,
    'bentConnector3': 1000,
    'bentConnector4': 1000,
    'bentConnector5': 1000,
    'curvedConnector2': 1000,
    'curvedConnector3': 1000,
    'curvedConnector4': 1000,
    'curvedConnector5': 1000,
    'sun': 1000,
    'moon': 1000,
    'cloud': 1000,
    'irregularSeal1': 1000,
    'irregularSeal2': 1000,
    'foldedCorner': 1000,
    'smileyFace': 1000,
    'noSmoking': 1000,
    'mathPlus': 11,
    'mathMinus': 1000,
    'mathMultiply': 1000,
    'mathDivide': 1000,
    'mathEqual': 1000,
    'mathNotEqual': 1000,
    'actionButtonBlank': 1000,
    'actionButtonHome': 1000,
    'actionButtonHelp': 1000,
    'actionButtonInformation': 1000,
    'actionButtonForwardNext': 1000,
    'actionButtonBackPrevious': 1000,
    'actionButtonEnd': 1000,
    'actionButtonBeginning': 1000,
    'actionButtonReturn': 1000,
    'actionButtonDocument': 1000,
    'actionButtonSound': 1000,
    'actionButtonMovie': 1000
};

/**
 * OpenXML Shape Types - Enhanced with more PowerPoint innate shapes
 */
const OpenXmlShapeTypes = {
    // Basic shapes
    RECTANGLE: 'rect',
    ELLIPSE: 'ellipse',
    LINE: 'line',
    ROUND_RECTANGLE: 'roundRect',

    // Arrow shapes
    RIGHT_ARROW: 'rightArrow',
    LEFT_ARROW: 'leftArrow',
    UP_ARROW: 'upArrow',
    DOWN_ARROW: 'downArrow',
    LEFT_RIGHT_ARROW: 'leftRightArrow',
    UP_DOWN_ARROW: 'upDownArrow',
    QUAD_ARROW: 'quadArrow',
    LEFT_RIGHT_UP_ARROW: 'leftRightUpArrow',
    BENT_ARROW: 'bentArrow',
    UTURN_ARROW: 'uturnArrow',
    LEFT_UP_ARROW: 'leftUpArrow',
    BENT_UP_ARROW: 'bentUpArrow',
    CURVED_RIGHT_ARROW: 'curvedRightArrow',
    CURVED_LEFT_ARROW: 'curvedLeftArrow',
    CURVED_UP_ARROW: 'curvedUpArrow',
    CURVED_DOWN_ARROW: 'curvedDownArrow',
    STRIPED_RIGHT_ARROW: 'stripedRightArrow',
    NOTCHED_RIGHT_ARROW: 'notchedRightArrow',
    BLOCK_ARC: 'blockArc',
    SWOOSH_ARROW: 'swooshArrow',
    CIRCULAR_ARROW: 'circularArrow',

    // Flowchart shapes
    FLOWCHART_PROCESS: 'flowChartProcess',
    FLOWCHART_DECISION: 'flowChartDecision',
    FLOWCHART_START_END: 'flowChartTerminator',
    FLOWCHART_INPUT_OUTPUT: 'flowChartInputOutput',
    FLOWCHART_PREDEFINED_PROCESS: 'flowChartPredefinedProcess',
    FLOWCHART_INTERNAL_STORAGE: 'flowChartInternalStorage',
    FLOWCHART_DOCUMENT: 'flowChartDocument',
    FLOWCHART_MULTIDOCUMENT: 'flowChartMultidocument',
    FLOWCHART_PREPARATION: 'flowChartPreparation',
    FLOWCHART_MANUAL_INPUT: 'flowChartManualInput',
    FLOWCHART_MANUAL_OPERATION: 'flowChartManualOperation',
    FLOWCHART_CONNECTOR: 'flowChartConnector',
    FLOWCHART_OFFPAGE_CONNECTOR: 'flowChartOffpageConnector',
    FLOWCHART_PUNCHED_CARD: 'flowChartPunchedCard',
    FLOWCHART_PUNCHED_TAPE: 'flowChartPunchedTape',
    FLOWCHART_SUMMING_JUNCTION: 'flowChartSummingJunction',
    FLOWCHART_OR: 'flowChartOr',
    FLOWCHART_COLLATE: 'flowChartCollate',
    FLOWCHART_SORT: 'flowChartSort',
    FLOWCHART_EXTRACT: 'flowChartExtract',
    FLOWCHART_MERGE: 'flowChartMerge',
    FLOWCHART_STORED_DATA: 'flowChartStoredData',
    FLOWCHART_DELAY: 'flowChartDelay',
    FLOWCHART_SEQUENTIAL_ACCESS_STORAGE: 'flowChartSequentialAccessStorage',
    FLOWCHART_MAGNETIC_DISK: 'flowChartMagneticDisk',
    FLOWCHART_DIRECT_ACCESS_STORAGE: 'flowChartDirectAccessStorage',
    FLOWCHART_DISPLAY: 'flowChartDisplay',
    FLOWCHART_ALTERNATE_PROCESS: 'flowChartAlternateProcess',
    FLOWCHART_MAGNETIC_DRUM: 'flowChartMagneticDrum',
    FLOWCHART_MAGNETIC_TAPE: 'flowChartMagneticTape',

    // Callout shapes
    CALLOUT_1: 'callout1',
    CALLOUT_2: 'callout2',
    CALLOUT_3: 'callout3',
    ACCENT_CALLOUT_1: 'accentCallout1',
    ACCENT_CALLOUT_2: 'accentCallout2',
    ACCENT_CALLOUT_3: 'accentCallout3',
    BORDER_CALLOUT_1: 'borderCallout1',
    BORDER_CALLOUT_2: 'borderCallout2',
    BORDER_CALLOUT_3: 'borderCallout3',
    ACCENT_BORDER_CALLOUT_1: 'accentBorderCallout1',
    ACCENT_BORDER_CALLOUT_2: 'accentBorderCallout2',
    ACCENT_BORDER_CALLOUT_3: 'accentBorderCallout3',

    // Star shapes
    STAR_4: 'star4',
    STAR_5: 'star5',
    STAR_6: 'star6',
    STAR_8: 'star8',
    STAR_12: 'star12',
    STAR_16: 'star16',
    STAR_24: 'star24',
    STAR_32: 'star32',

    // Basic geometric shapes
    TRIANGLE: 'triangle',
    RIGHT_TRIANGLE: 'rtTriangle',
    DIAMOND: 'diamond',
    PENTAGON: 'pentagon',
    HEXAGON: 'hexagon',
    HEPTAGON: 'heptagon',
    OCTAGON: 'octagon',
    DECAGON: 'decagon',
    DODECAGON: 'dodecagon',
    PARALLELOGRAM: 'parallelogram',
    TRAPEZOID: 'trapezoid',
    PLUS: 'plus',
    CROSS: 'cross',

    // Connector shapes
    STRAIGHT_CONNECTOR_1: 'straightConnector1',
    BENT_CONNECTOR_2: 'bentConnector2',
    BENT_CONNECTOR_3: 'bentConnector3',
    BENT_CONNECTOR_4: 'bentConnector4',
    BENT_CONNECTOR_5: 'bentConnector5',
    CURVED_CONNECTOR_2: 'curvedConnector2',
    CURVED_CONNECTOR_3: 'curvedConnector3',
    CURVED_CONNECTOR_4: 'curvedConnector4',
    CURVED_CONNECTOR_5: 'curvedConnector5',

    // Ribbon and banner shapes
    RIBBON: 'ribbon',
    RIBBON_2: 'ribbon2',
    ELLIPSE_RIBBON: 'ellipseRibbon',
    ELLIPSE_RIBBON_2: 'ellipseRibbon2',
    LEFT_RIGHT_RIBBON: 'leftRightRibbon',
    VERTICAL_SCROLL: 'verticalScroll',
    HORIZONTAL_SCROLL: 'horizontalScroll',
    WAVE: 'wave',
    DOUBLE_WAVE: 'doubleWave',

    // Bracket shapes
    LEFT_BRACKET: 'leftBracket',
    RIGHT_BRACKET: 'rightBracket',
    LEFT_BRACE: 'leftBrace',
    RIGHT_BRACE: 'rightBrace',
    BRACKET_PAIR: 'bracketPair',
    BRACE_PAIR: 'bracePair',

    // Miscellaneous shapes
    FRAME: 'frame',
    L_SHAPE: 'lShape',
    DIAGONAL_STRIPE: 'diagStripe',
    CHORD: 'chord',
    ARC: 'arc',

    // 3D shapes
    CUBE: 'cube',
    CAN: 'can',
    CONE: 'cone',
    PYRAMID: 'pyramid',

    // Symbols and icons
    HEART: 'heart',
    LIGHTNING_BOLT: 'lightningBolt',
    SUN: 'sun',
    MOON: 'moon',
    CLOUD: 'cloud',
    SMILEY_FACE: 'smileyFace',
    IRREGULAR_SEAL_1: 'irregularSeal1',
    IRREGULAR_SEAL_2: 'irregularSeal2',
    FOLDED_CORNER: 'foldedCorner',
    BEVEL: 'bevel',
    DONUT: 'donut',
    NO_SMOKING: 'noSmoking',
    BLOCK_ARC_2: 'blockArc2',

    // Mathematical symbols
    PLUS_MATH: 'mathPlus',
    MINUS: 'mathMinus',
    MULTIPLY: 'mathMultiply',
    DIVIDE: 'mathDivide',
    EQUAL: 'mathEqual',
    NOT_EQUAL: 'mathNotEqual',

    // Action button shapes
    ACTION_BUTTON_BLANK: 'actionButtonBlank',
    ACTION_BUTTON_HOME: 'actionButtonHome',
    ACTION_BUTTON_HELP: 'actionButtonHelp',
    ACTION_BUTTON_INFORMATION: 'actionButtonInformation',
    ACTION_BUTTON_FORWARD_NEXT: 'actionButtonForwardNext',
    ACTION_BUTTON_BACK_PREVIOUS: 'actionButtonBackPrevious',
    ACTION_BUTTON_END: 'actionButtonEnd',
    ACTION_BUTTON_BEGINNING: 'actionButtonBeginning',
    ACTION_BUTTON_RETURN: 'actionButtonReturn',
    ACTION_BUTTON_DOCUMENT: 'actionButtonDocument',
    ACTION_BUTTON_SOUND: 'actionButtonSound',
    ACTION_BUTTON_MOVIE: 'actionButtonMovie',
    TEARDROP: 'teardrop'
};

/**
 * Enhanced CGraphics class with Standard integration
 * Provides comprehensive shape rendering using Standard's native shape processing
 */
class CGraphics {
    constructor(canvas) {
        // Set up high-resolution canvas support
        this.setupHighResolutionCanvas(canvas);
        
        this.m_oContext = canvas.getContext('2d');
        this.m_dDpiX = 96.0;
        this.m_dDpiY = 96.0;
        this.m_dWidthMM = 210;
        this.m_dHeightMM = 297;
        // Use logical display dimensions, not scaled canvas dimensions
        // This ensures coordinate transformations work correctly with high-DPI scaling
        if (canvas.style.width && canvas.style.height) {
            this.m_dWidth = parseFloat(canvas.style.width);
            this.m_dHeight = parseFloat(canvas.style.height);
        } else {
            this.m_dWidth = canvas.width / (window.devicePixelRatio || 1);
            this.m_dHeight = canvas.height / (window.devicePixelRatio || 1);
        }
        this.m_oTransform = new CMatrix();
        
        // Store device pixel ratio for image rendering
        this.devicePixelRatio = window.devicePixelRatio || 1;
        this.m_oCoordTransform = new CMatrix();
        this.m_oFullTransform = new CMatrix();
        this.m_oTransformStack = [];
        this.m_oClipStack = [];
        this.m_oPen = new CPen();
        this.m_oBrush = new CBrush();
        this.m_oFont = new CFont();
        this.m_bIntegerGrid = false;
        this.m_bGlobalAlpha = false;
        this.m_dGlobalAlpha = 1.0;
        this.m_bIsClipping = false;
        this.m_oClipRect = null;
        this.m_oPath = null;
        this.m_bIsDrawing = false;
        
        // Enhanced geometry processing
        this.m_oEnhancedGeometryProcessor = new EnhancedGeometryProcessor();
        this.m_oCustomGeometryProcessor = new GeometryProcessor();

        // Initialize coordinate transformations
        this.calculateCoordTransform();
        this.calculateFullTransform();

        // Set default properties
        this.m_oContext.lineCap = 'round';
        this.m_oContext.lineJoin = 'round';
        this.m_oContext.textBaseline = 'alphabetic';
        this.m_oContext.textAlign = 'left';
    }

    /**
     * Set up high-resolution canvas for crisp rendering on high-DPI displays
     */
    setupHighResolutionCanvas(canvas) {
        const ctx = canvas.getContext('2d');
        const devicePixelRatio = window.devicePixelRatio || 1;
        
        // Get the display size (canvas.style size in CSS pixels)
        const displayWidth = canvas.clientWidth || canvas.width;
        const displayHeight = canvas.clientHeight || canvas.height;
        
        // Set the actual canvas size in memory (scaled up for high-DPI)
        canvas.width = displayWidth * devicePixelRatio;
        canvas.height = displayHeight * devicePixelRatio;
        
        // Scale the canvas back down using CSS for proper display
        canvas.style.width = displayWidth + 'px';
        canvas.style.height = displayHeight + 'px';
        
        // Scale the drawing context to match device pixel ratio
        ctx.scale(devicePixelRatio, devicePixelRatio);
        
        // Enable high-quality image rendering
        ctx.imageSmoothingEnabled = true;
        ctx.imageSmoothingQuality = 'high';
    }

    /**
     * Draw image with high-resolution support for crisp rendering
     */
    drawImageHighRes(image, x, y, width, height, sourceX = 0, sourceY = 0, sourceWidth = null, sourceHeight = null) {
        if (!this.m_oContext || !image) {
            return;
        }

        // Save current context state
        this.m_oContext.save();
        
        // Enable high-quality image scaling
        this.m_oContext.imageSmoothingEnabled = true;
        this.m_oContext.imageSmoothingQuality = 'high';
        
        try {
            if (sourceWidth !== null && sourceHeight !== null) {
                // Draw with source rectangle (cropping)
                this.m_oContext.drawImage(
                    image, 
                    sourceX, sourceY, sourceWidth, sourceHeight,
                    x, y, width, height
                );
            } else {
                // Draw entire image scaled to fit
                this.m_oContext.drawImage(image, x, y, width, height);
            }
        } catch (error) {
        } finally {
            // Restore context state
            this.m_oContext.restore();
        }
    }

    /**
     * Legacy drawImage method for compatibility - now uses high-res rendering
     */
    drawImage(image, x, y, width, height) {
        this.drawImageHighRes(image, x, y, width, height);
    }

    /**
     * Enhanced preset geometry drawing using Standard patterns
     * This method properly integrates with Standard's shape processing system
     */
    drawPresetGeometry(shapeType, x, y, width, height, fillColor, strokeColor, lineWidth = 1, strokeInfo = null, adjustments = {}) {
        if (!this.m_oContext) {return;}
        this.SaveGrState();

        try {
            // Get Standard shape type mapping
            const onlyOfficeType = StandardShapeTypes[shapeType] || 1000; // Default to custom

            if (onlyOfficeType === 1000) {
                // Custom shape - use Standard custom geometry processor
                this.drawCustomShape(shapeType, x, y, width, height, fillColor, strokeColor, lineWidth, strokeInfo, adjustments);
            } else {
                // Standard Standard preset shape
                this.drawStandardPresetShape(onlyOfficeType, shapeType, x, y, width, height, fillColor, strokeColor, lineWidth, strokeInfo, adjustments);
            }

        } catch (error) {
            // Fallback to rectangle
            this.drawRectangle(x, y, width, height, fillColor, strokeColor, lineWidth, strokeInfo);
        } finally {
            this.RestoreGrState();
        }
    }

    /**
     * Draw standard Standard preset shapes using enhanced geometry processor
     */
    drawStandardPresetShape(onlyOfficeType, shapeType, x, y, width, height, fillColor, strokeColor, lineWidth, strokeInfo, adjustments) {
        // Convert dimensions to shape units (21600 = 100%)
        const shapeWidth = 21600;
        const shapeHeight = 21600;
        // Process geometry using enhanced processor
        const geometry = this.m_oEnhancedGeometryProcessor.processPresetGeometry(
            shapeType, 
            shapeWidth, 
            shapeHeight, 
            adjustments
        );

        if (geometry) {
            this.renderEnhancedGeometry(geometry, x, y, width, height, fillColor, strokeColor, lineWidth, strokeInfo);
        } else {
            // Fallback to basic shape rendering
            this.drawBasicShape(shapeType, x, y, width, height, fillColor, strokeColor, lineWidth, strokeInfo);
        }
    }

    /**
     * Render enhanced geometry to canvas
     */
    renderEnhancedGeometry(geometry, x, y, width, height, fillColor, strokeColor, lineWidth, strokeInfo) {
        const ctx = this.m_oContext;
        const pathList = geometry.getPathList();
        
        // Calculate scaling factors
        const scaleX = width / geometry.width;
        const scaleY = height / geometry.height;
        
        ctx.save();
        ctx.translate(x, y);
        ctx.scale(scaleX, scaleY);
        
        // Render each path
        pathList.forEach(path => {
            this.renderEnhancedPath(path, geometry, fillColor, strokeColor, lineWidth, strokeInfo, { x, y, w: width, h: height });
        });
        
        ctx.restore();
    }

    /**
     * Render individual enhanced path
     */
    renderEnhancedPath(path, geometry, fillColor, strokeColor, lineWidth, strokeInfo, bounds = null) {
        const ctx = this.m_oContext;
        ctx.beginPath();
        
        // Process path commands
        path.commands.forEach(cmd => {
            switch (cmd.type) {
                case 'M': // Move to
                    ctx.moveTo(cmd.args[0], cmd.args[1]);
                    break;
                case 'L': // Line to
                    ctx.lineTo(cmd.args[0], cmd.args[1]);
                    break;
                case 'C': // Cubic Bezier curve
                    ctx.bezierCurveTo(cmd.args[0], cmd.args[1], cmd.args[2], cmd.args[3], cmd.args[4], cmd.args[5]);
                    break;
                case 'Q': // Quadratic Bezier curve
                    ctx.quadraticCurveTo(cmd.args[0], cmd.args[1], cmd.args[2], cmd.args[3]);
                    break;
                case 'A': // Arc
                    this.renderArcCommand(ctx, cmd.args);
                    break;
                case 'Z': // Close path
                    ctx.closePath();
                    break;
            }
        });
        
        // Apply fill and stroke with Standard styling
        if (path.fill !== 'none' && fillColor) {
            // Check if fillColor is a gradient object
            if (fillColor && typeof fillColor === 'object' && fillColor.type && fillColor.stops) {
                // Use the passed bounds or fallback
                const currentBounds = bounds || { x: 0, y: 0, w: 200, h: 50 };
                const gradient = this.createGradient(currentBounds.x, currentBounds.y, currentBounds.w, currentBounds.h, fillColor);
                if (gradient) {
                    ctx.fillStyle = gradient;
                } else {
                    ctx.fillStyle = this.colorToRgb(fillColor.stops[0]?.color || { r: 0, g: 0, b: 0 });
                }
            } else {
                ctx.fillStyle = this.colorToRgb(fillColor);
            }
            ctx.fill();
        }

        if (path.stroke && strokeColor) {
            ctx.strokeStyle = this.colorToRgb(strokeColor);
            // lineWidth is in mm; the context has been scaled (scaleX = bounds.w / geometry.width).
            // To get correct screen-pixel stroke width, divide by the scale factor.
            // Use the min scale so all edges are at least lineWidthPx thick.
            const lineWidthPx = CoordinateTransform.mmToPixels(lineWidth || 1);
            const geomW = geometry ? geometry.width : 21600;
            const geomH = geometry ? geometry.height : 21600;
            const scaleX = (bounds && bounds.w > 0) ? bounds.w / geomW : 1;
            const scaleY = (bounds && bounds.h > 0) ? bounds.h / geomH : 1;
            const minScale = Math.min(scaleX, scaleY) || 1;
            ctx.lineWidth = lineWidthPx / minScale;

            // Apply stroke information if available
            if (strokeInfo) {
                this.applyStrokeInfo(strokeInfo);
            }

            ctx.stroke();
        }
    }

    /**
     * Render arc command (simplified - for full Standard compatibility, 
     * this would need complete arc-to-bezier conversion)
     */
    renderArcCommand(ctx, args) {
        const [rx, ry, xAxisRotation, largeArcFlag, sweepFlag, x, y] = args;
        
        // Get current point
        const currentPoint = this.getCurrentPoint(ctx);
        if (!currentPoint) {return;}
        
        // For now, use a simple arc approximation
        // In full Standard implementation, this would convert to bezier curves
        try {
            const centerX = (currentPoint.x + x) / 2;
            const centerY = (currentPoint.y + y) / 2;
            const radius = Math.min(rx, ry);
            
            const startAngle = Math.atan2(currentPoint.y - centerY, currentPoint.x - centerX);
            const endAngle = Math.atan2(y - centerY, x - centerX);
            
            ctx.arc(centerX, centerY, radius, startAngle, endAngle, !sweepFlag);
        } catch (e) {
            // Fallback to line
            ctx.lineTo(x, y);
        }
    }

    /**
     * Get current point from canvas context (approximation)
     */
    getCurrentPoint(ctx) {
        // This is a simplification - real implementation would track current point
        return { x: 0, y: 0 };
    }

    /**
     * Draw custom shapes using Standard custom geometry processor
     */
    drawCustomShape(shapeType, x, y, width, height, fillColor, strokeColor, lineWidth, strokeInfo, adjustments) {
        // Use custom geometry processor for complex shapes
        const customGeometry = this.m_oCustomGeometryProcessor.processCustomShape(shapeType, {
            x: x,
            y: y,
            width: width,
            height: height,
            adjustments: adjustments
        });

        if (customGeometry && customGeometry.pathLst) {
            this.drawCustomGeometry(customGeometry, x, y, width, height, fillColor, strokeColor, lineWidth);
        } else {
            // Fallback to basic shape rendering
            this.drawBasicShape(shapeType, x, y, width, height, fillColor, strokeColor, lineWidth, strokeInfo);
        }
    }

    /**
     * Apply stroke information (dash patterns, line caps, etc.)
     */
    applyStrokeInfo(strokeInfo) {
        if (!strokeInfo || !this.m_oContext) {return;}

        const ctx = this.m_oContext;

        // Apply dash pattern
        if (strokeInfo.dashArray && strokeInfo.dashArray.length > 0) {
            // Scale dash array by line width for proper visibility
            const lineWidth = ctx.lineWidth || 1;
            const scaledDashArray = strokeInfo.dashArray.map(dash => dash * lineWidth);
            ctx.setLineDash(scaledDashArray);
        }
        
        // Apply line cap (support aliases)
        const cap = strokeInfo.lineCap || strokeInfo.cap;
        if (cap) {
            switch (cap) {
                case 'rnd':
                case 'round':
                    ctx.lineCap = 'round';
                    break;
                case 'sq':
                case 'square':
                    ctx.lineCap = 'square';
                    break;
                case 'flat':
                case 'butt':
                default:
                    ctx.lineCap = 'butt';
                    break;
            }
        }
        
        // Apply line join (support aliases)
        const join = strokeInfo.lineJoin || strokeInfo.join;
        if (join) {
            switch (join) {
                case 'round':
                    ctx.lineJoin = 'round';
                    break;
                case 'bevel':
                    ctx.lineJoin = 'bevel';
                    break;
                case 'miter':
                default:
                    ctx.lineJoin = 'miter';
                    break;
            }
        }
        
        // Apply miter limit
        if (strokeInfo.miterLimit) {
            ctx.miterLimit = strokeInfo.miterLimit;
        }
    }

    /**
     * Draw basic shapes for fallback cases
     */
    drawBasicShape(shapeType, x, y, width, height, fillColor, strokeColor, lineWidth, strokeInfo) {
        const ctx = this.m_oContext;
        
        ctx.save();
        ctx.beginPath();

        switch (shapeType) {
            case 'rect':
            case 'rectangle':
                ctx.rect(x, y, width, height);
                break;
            case 'ellipse':
            case 'oval':
                ctx.ellipse(x + width/2, y + height/2, width/2, height/2, 0, 0, 2 * Math.PI);
                break;
            case 'roundRect':
                this.drawRoundRectPath(ctx, x, y, width, height, Math.min(width, height) * 0.16667);
                break;
            case 'triangle':
                this.drawTrianglePath(ctx, x, y, width, height);
                break;
            case 'diamond':
                this.drawDiamondPath(ctx, x, y, width, height);
                break;
            case 'pentagon':
                this.drawPentagonPath(ctx, x, y, width, height);
                break;
            case 'hexagon':
                this.drawHexagonPath(ctx, x, y, width, height);
                break;
            case 'octagon':
                this.drawOctagonPath(ctx, x, y, width, height);
                break;
            case 'star':
            case 'star5':
                this.drawStarPath(ctx, x, y, width, height);
                break;
            case 'rightArrow':
                this.drawRightArrowPath(ctx, x, y, width, height);
                break;
            case 'leftArrow':
                this.drawLeftArrowPath(ctx, x, y, width, height);
                break;
            case 'upArrow':
                this.drawUpArrowPath(ctx, x, y, width, height);
                break;
            case 'downArrow':
                this.drawDownArrowPath(ctx, x, y, width, height);
                break;
            default:
                // Default to rectangle
                ctx.rect(x, y, width, height);
                break;
        }

        // Apply fill and stroke
        if (fillColor) {
            ctx.fillStyle = this.colorToRgb(fillColor);
            ctx.fill();
        }

        if (strokeColor) {
            ctx.strokeStyle = this.colorToRgb(strokeColor);
            ctx.lineWidth = CoordinateTransform.mmToPixels(lineWidth || 1);
            
            if (strokeInfo) {
                this.applyStrokeInfo(strokeInfo);
            }
            
            ctx.stroke();
        }

        ctx.restore();
    }

    /**
     * Draw custom geometry using Standard path system
     */
    drawCustomGeometry(geometry, x, y, width, height, fillColor, strokeColor, lineWidth) {
        try {

            if (!geometry || !geometry.pathLst || geometry.pathLst.length === 0) {
                return;
            }


            this.m_oContext.save();

            // Apply coordinate transformation
            this.m_oContext.translate(x, y);
            
            // Check for coordSize, use defaults if missing
            const coordWidth = geometry.coordSize?.width || 1000;
            const coordHeight = geometry.coordSize?.height || 1000;
            
            this.m_oContext.scale(width / coordWidth, height / coordHeight);

            // Process each path in the path list
            geometry.pathLst.forEach((path, index) => {
                this.drawGeometryPath(path, fillColor, strokeColor, lineWidth);
            });

            this.m_oContext.restore();
        } catch (error) {
            this.m_oContext.restore(); // Ensure context is restored even on error
        }
    }

    /**
     * Draw a single geometry path
     */
    drawGeometryPath(path, fillColor, strokeColor, lineWidth) {
        if (!path || !path.commands || path.commands.length === 0) {
            return;
        }

        this.m_oContext.beginPath();

        // Process path commands
        path.commands.forEach(command => {
            switch (command.type) {
                case 'moveTo':
                    this.m_oContext.moveTo(command.x, command.y);
                    break;
                case 'lineTo':
                    this.m_oContext.lineTo(command.x, command.y);
                    break;
                case 'curveTo':
                    this.m_oContext.bezierCurveTo(command.x1, command.y1, command.x2, command.y2, command.x, command.y);
                    break;
                case 'close':
                    // Only close path if we're doing fill, not stroke-only
                    if (fillColor && path.fill !== false && path.fill !== 'none') {
                        this.m_oContext.closePath();
                    }
                    break;
                case 'arcTo':
                    this.m_oContext.arcTo(command.x1, command.y1, command.x2, command.y2, command.radius);
                    break;
                case 'quadraticCurveTo':
                    this.m_oContext.quadraticCurveTo(command.x1, command.y1, command.x, command.y);
                    break;
            }
        });

        // Apply stroke first (for paths with both fill and stroke, stroke should be visible)
        if (strokeColor && path.stroke !== false) {
            this.m_oContext.strokeStyle = this.colorToRgb(strokeColor);
            this.m_oContext.lineWidth = CoordinateTransform.mmToPixels(lineWidth || 1);
            this.m_oContext.stroke();
        }

        // Apply fill only if explicitly specified and not 'none'
        if (fillColor && path.fill !== false && path.fill !== 'none') {
            this.m_oContext.fillStyle = this.colorToRgb(fillColor);
            this.m_oContext.fill();
        }
    }

    /**
     * Draw shape path using Standard path data
     */
    drawShapePath(pathData, fillColor, strokeColor, lineWidth, strokeInfo) {
        if (!pathData || !pathData.commands) {
            return;
        }

        this.m_oContext.beginPath();

        // Process Standard path commands
        pathData.commands.forEach(command => {
            switch (command.type) {
                case 'M': // Move to
                    this.m_oContext.moveTo(command.x, command.y);
                    break;
                case 'L': // Line to
                    this.m_oContext.lineTo(command.x, command.y);
                    break;
                case 'C': // Cubic Bezier curve
                    this.m_oContext.bezierCurveTo(command.x1, command.y1, command.x2, command.y2, command.x, command.y);
                    break;
                case 'Q': // Quadratic Bezier curve
                    this.m_oContext.quadraticCurveTo(command.x1, command.y1, command.x, command.y);
                    break;
                case 'A': // Arc
                    this.drawArcCommand(command);
                    break;
                case 'Z': // Close path
                    // Only close path if we're doing fill, not stroke-only
                    if (fillColor && path.fill !== false && path.fill !== 'none') {
                        this.m_oContext.closePath();
                    }
                    break;
            }
        });

        // Apply fill and stroke with Standard styling
        if (fillColor) {
            this.m_oContext.fillStyle = this.colorToRgb(fillColor);
            this.m_oContext.fill();
        }

        if (strokeColor) {
            this.m_oContext.strokeStyle = this.colorToRgb(strokeColor);
            this.m_oContext.lineWidth = lineWidth || 1;

            // Apply stroke information if available
            if (strokeInfo) {
                this.applyStrokeInfo(strokeInfo);
            }

            this.m_oContext.stroke();
        }
    }

    /**
     * Apply Standard stroke information
     */
    applyStrokeInfoAlternative(strokeInfo) {
        if (!strokeInfo) {return;}

        // Apply line dash pattern
        if (strokeInfo.dashArray && strokeInfo.dashArray.length > 0) {
            this.m_oContext.setLineDash(strokeInfo.dashArray);
        }

        // Apply line cap style
        if (strokeInfo.lineCap) {
            this.m_oContext.lineCap = strokeInfo.lineCap;
        }

        // Apply line join style
        if (strokeInfo.lineJoin) {
            this.m_oContext.lineJoin = strokeInfo.lineJoin;
        }

        // Apply miter limit
        if (strokeInfo.miterLimit) {
            this.m_oContext.miterLimit = strokeInfo.miterLimit;
        }
    }

    /**
     * Draw arc command for Standard path system
     */
    drawArcCommand(command) {
        // Convert Standard arc parameters to Canvas arc
        const centerX = command.cx || command.x;
        const centerY = command.cy || command.y;
        const radius = command.r || command.radius;
        const startAngle = command.startAngle || 0;
        const endAngle = command.endAngle || Math.PI * 2;
        const counterclockwise = command.counterclockwise || false;

        this.m_oContext.arc(centerX, centerY, radius, startAngle, endAngle, counterclockwise);
    }

    // ... existing methods remain the same ...

    drawRectangle(x, y, width, height, fillColor, strokeColor, lineWidth, strokeInfo) {
        if (!this.m_oContext) {return;}

        this.m_oContext.save();

        if (fillColor) {
            if (typeof fillColor === 'object' && fillColor.type && fillColor.stops) {
                const gradient = this.createGradient(x, y, width, height, fillColor);
                this.m_oContext.fillStyle = gradient || this.colorToRgb(fillColor.stops[0]?.color || { r: 0, g: 0, b: 0 });
            } else {
                this.m_oContext.fillStyle = this.colorToRgb(fillColor);
            }
            this.m_oContext.fillRect(x, y, width, height);
        }

        if (strokeColor) {
            this.m_oContext.strokeStyle = this.colorToRgb(strokeColor);
            this.m_oContext.lineWidth = lineWidth || 1;

            if (strokeInfo) {
                this.applyStrokeInfo(strokeInfo);
            }

            this.m_oContext.strokeRect(x, y, width, height);
        }

        this.m_oContext.restore();
    }

    drawEllipse(x, y, width, height, fillColor, strokeColor, lineWidth) {
        if (!this.m_oContext) {return;}

        this.m_oContext.save();
        this.m_oContext.beginPath();

        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const radiusX = width / 2;
        const radiusY = height / 2;

        this.m_oContext.ellipse(centerX, centerY, radiusX, radiusY, 0, 0, 2 * Math.PI);

        if (fillColor) {
            this.m_oContext.fillStyle = this.colorToRgb(fillColor);
            this.m_oContext.fill();
        }

        if (strokeColor) {
            this.m_oContext.strokeStyle = this.colorToRgb(strokeColor);
            this.m_oContext.lineWidth = lineWidth || 1;
            this.m_oContext.stroke();
        }

        this.m_oContext.restore();
    }

    drawRoundedRectangle(x, y, width, height, radius, fillColor, strokeColor, lineWidth) {
        if (!this.m_oContext) {return;}

        this.m_oContext.save();
        this.m_oContext.beginPath();

        // Ensure radius doesn't exceed half the width or height
        const maxRadius = Math.min(width, height) / 2;
        radius = Math.min(radius, maxRadius);

        this.m_oContext.moveTo(x + radius, y);
        this.m_oContext.lineTo(x + width - radius, y);
        this.m_oContext.quadraticCurveTo(x + width, y, x + width, y + radius);
        this.m_oContext.lineTo(x + width, y + height - radius);
        this.m_oContext.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
        this.m_oContext.lineTo(x + radius, y + height);
        this.m_oContext.quadraticCurveTo(x, y + height, x, y + height - radius);
        this.m_oContext.lineTo(x, y + radius);
        this.m_oContext.quadraticCurveTo(x, y, x + radius, y);
        this.m_oContext.closePath();

        if (fillColor) {
            this.m_oContext.fillStyle = this.colorToRgb(fillColor);
            this.m_oContext.fill();
        }

        if (strokeColor) {
            this.m_oContext.strokeStyle = this.colorToRgb(strokeColor);
            this.m_oContext.lineWidth = lineWidth || 1;
            this.m_oContext.stroke();
        }

        this.m_oContext.restore();
    }

    drawTriangle(x, y, width, height, fillColor, strokeColor, lineWidth) {
        if (!this.m_oContext) {return;}

        this.m_oContext.save();
        this.m_oContext.beginPath();

        // Isosceles triangle
        this.m_oContext.moveTo(x + width / 2, y);
        this.m_oContext.lineTo(x + width, y + height);
        this.m_oContext.lineTo(x, y + height);
        this.m_oContext.closePath();

        if (fillColor) {
            this.m_oContext.fillStyle = this.colorToRgb(fillColor);
            this.m_oContext.fill();
        }

        if (strokeColor) {
            this.m_oContext.strokeStyle = this.colorToRgb(strokeColor);
            this.m_oContext.lineWidth = lineWidth || 1;
            this.m_oContext.stroke();
        }

        this.m_oContext.restore();
    }

    drawDiamond(x, y, width, height, fillColor, strokeColor, lineWidth) {
        if (!this.m_oContext) {return;}

        this.m_oContext.save();
        this.m_oContext.beginPath();

        const centerX = x + width / 2;
        const centerY = y + height / 2;

        this.m_oContext.moveTo(centerX, y);
        this.m_oContext.lineTo(x + width, centerY);
        this.m_oContext.lineTo(centerX, y + height);
        this.m_oContext.lineTo(x, centerY);
        this.m_oContext.closePath();

        if (fillColor) {
            this.m_oContext.fillStyle = this.colorToRgb(fillColor);
            this.m_oContext.fill();
        }

        if (strokeColor) {
            this.m_oContext.strokeStyle = this.colorToRgb(strokeColor);
            this.m_oContext.lineWidth = lineWidth || 1;
            this.m_oContext.stroke();
        }

        this.m_oContext.restore();
    }

    drawRegularPolygon(x, y, width, height, sides, fillColor, strokeColor, lineWidth) {
        if (!this.m_oContext || sides < 3) {return;}

        this.m_oContext.save();
        this.m_oContext.beginPath();

        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const radius = Math.min(width, height) / 2;

        for (let i = 0; i < sides; i++) {
            const angle = (i * 2 * Math.PI) / sides - Math.PI / 2;
            const px = centerX + radius * Math.cos(angle);
            const py = centerY + radius * Math.sin(angle);

            if (i === 0) {
                this.m_oContext.moveTo(px, py);
            } else {
                this.m_oContext.lineTo(px, py);
            }
        }

        this.m_oContext.closePath();

        if (fillColor) {
            this.m_oContext.fillStyle = this.colorToRgb(fillColor);
            this.m_oContext.fill();
        }

        if (strokeColor) {
            this.m_oContext.strokeStyle = this.colorToRgb(strokeColor);
            this.m_oContext.lineWidth = lineWidth || 1;
            this.m_oContext.stroke();
        }

        this.m_oContext.restore();
    }

    drawStar(x, y, width, height, points, fillColor, strokeColor, lineWidth) {
        if (!this.m_oContext || points < 3) {return;}

        this.m_oContext.save();
        this.m_oContext.beginPath();

        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const outerRadius = Math.min(width, height) / 2;
        const innerRadius = outerRadius * 0.4;

        for (let i = 0; i < points * 2; i++) {
            const angle = (i * Math.PI) / points - Math.PI / 2;
            const radius = i % 2 === 0 ? outerRadius : innerRadius;
            const px = centerX + radius * Math.cos(angle);
            const py = centerY + radius * Math.sin(angle);

            if (i === 0) {
                this.m_oContext.moveTo(px, py);
            } else {
                this.m_oContext.lineTo(px, py);
            }
        }

        this.m_oContext.closePath();

        if (fillColor) {
            this.m_oContext.fillStyle = this.colorToRgb(fillColor);
            this.m_oContext.fill();
        }

        if (strokeColor) {
            this.m_oContext.strokeStyle = this.colorToRgb(strokeColor);
            this.m_oContext.lineWidth = lineWidth || 1;
            this.m_oContext.stroke();
        }

        this.m_oContext.restore();
    }

    drawRightArrow(x, y, width, height, fillColor, strokeColor, lineWidth) {
        if (!this.m_oContext) {return;}

        this.m_oContext.save();
        this.m_oContext.beginPath();

        const arrowWidth = width * 0.7;
        const arrowHeight = height * 0.4;
        const headWidth = width * 0.3;
        const headHeight = height * 0.6;

        // Arrow shaft
        this.m_oContext.moveTo(x, y + (height - arrowHeight) / 2);
        this.m_oContext.lineTo(x + arrowWidth, y + (height - arrowHeight) / 2);

        // Arrow head top
        this.m_oContext.lineTo(x + arrowWidth, y + (height - headHeight) / 2);
        this.m_oContext.lineTo(x + width, y + height / 2);

        // Arrow head bottom
        this.m_oContext.lineTo(x + arrowWidth, y + (height + headHeight) / 2);
        this.m_oContext.lineTo(x + arrowWidth, y + (height + arrowHeight) / 2);

        // Arrow shaft bottom
        this.m_oContext.lineTo(x, y + (height + arrowHeight) / 2);
        this.m_oContext.closePath();

        if (fillColor) {
            this.m_oContext.fillStyle = this.colorToRgb(fillColor);
            this.m_oContext.fill();
        }

        if (strokeColor) {
            this.m_oContext.strokeStyle = this.colorToRgb(strokeColor);
            this.m_oContext.lineWidth = lineWidth || 1;
            this.m_oContext.stroke();
        }

        this.m_oContext.restore();
    }

    drawLeftArrow(x, y, width, height, fillColor, strokeColor, lineWidth) {
        if (!this.m_oContext) {return;}

        this.m_oContext.save();
        this.m_oContext.beginPath();

        const arrowWidth = width * 0.7;
        const arrowHeight = height * 0.4;
        const headWidth = width * 0.3;
        const headHeight = height * 0.6;

        // Arrow head top
        this.m_oContext.moveTo(x, y + height / 2);
        this.m_oContext.lineTo(x + headWidth, y + (height - headHeight) / 2);
        this.m_oContext.lineTo(x + headWidth, y + (height - arrowHeight) / 2);

        // Arrow shaft
        this.m_oContext.lineTo(x + width, y + (height - arrowHeight) / 2);
        this.m_oContext.lineTo(x + width, y + (height + arrowHeight) / 2);

        // Arrow shaft bottom
        this.m_oContext.lineTo(x + headWidth, y + (height + arrowHeight) / 2);
        this.m_oContext.lineTo(x + headWidth, y + (height + headHeight) / 2);
        this.m_oContext.closePath();

        if (fillColor) {
            this.m_oContext.fillStyle = this.colorToRgb(fillColor);
            this.m_oContext.fill();
        }

        if (strokeColor) {
            this.m_oContext.strokeStyle = this.colorToRgb(strokeColor);
            this.m_oContext.lineWidth = lineWidth || 1;
            this.m_oContext.stroke();
        }

        this.m_oContext.restore();
    }

    drawHeart(x, y, width, height, fillColor, strokeColor, lineWidth) {
        if (!this.m_oContext) {return;}

        this.m_oContext.save();
        this.m_oContext.beginPath();

        const centerX = x + width / 2;
        const topY = y + height * 0.3;

        // Left curve
        this.m_oContext.moveTo(centerX, topY);
        this.m_oContext.bezierCurveTo(centerX - width * 0.3, y, x, y + height * 0.3, centerX, y + height * 0.8);

        // Right curve
        this.m_oContext.bezierCurveTo(x + width, y + height * 0.3, centerX + width * 0.3, y, centerX, topY);

        this.m_oContext.closePath();

        if (fillColor) {
            this.m_oContext.fillStyle = this.colorToRgb(fillColor);
            this.m_oContext.fill();
        }

        if (strokeColor) {
            this.m_oContext.strokeStyle = this.colorToRgb(strokeColor);
            this.m_oContext.lineWidth = lineWidth || 1;
            this.m_oContext.stroke();
        }

        this.m_oContext.restore();
    }

    drawTeardrop(x, y, width, height, fillColor, strokeColor, lineWidth) {
        if (!this.m_oContext) {return;}

        this.m_oContext.save();
        this.m_oContext.beginPath();

        const centerX = x + width / 2;
        const centerY = y + height * 0.7;
        const radius = Math.min(width, height) * 0.3;

        // Draw the circular part
        this.m_oContext.arc(centerX, centerY, radius, 0, Math.PI * 2);

        // Draw the teardrop point
        this.m_oContext.moveTo(centerX, centerY - radius);
        this.m_oContext.quadraticCurveTo(centerX - radius * 0.5, y + height * 0.2, centerX, y);
        this.m_oContext.quadraticCurveTo(centerX + radius * 0.5, y + height * 0.2, centerX, centerY - radius);

        if (fillColor) {
            this.m_oContext.fillStyle = this.colorToRgb(fillColor);
            this.m_oContext.fill();
        }

        if (strokeColor) {
            this.m_oContext.strokeStyle = this.colorToRgb(strokeColor);
            this.m_oContext.lineWidth = lineWidth || 1;
            this.m_oContext.stroke();
        }

        this.m_oContext.restore();
    }

    drawCylinder(x, y, width, height, fillColor, strokeColor, lineWidth) {
        if (!this.m_oContext) {return;}

        this.m_oContext.save();

        const ellipseHeight = height * 0.15;
        const radiusX = width / 2;
        const radiusY = ellipseHeight / 2;
        const centerX = x + width / 2;

        // Draw the cylinder body
        if (fillColor) {
            this.m_oContext.fillStyle = this.colorToRgb(fillColor);
            this.m_oContext.fillRect(x, y + radiusY, width, height - ellipseHeight);
        }

        // Draw top ellipse
        this.m_oContext.beginPath();
        this.m_oContext.ellipse(centerX, y + radiusY, radiusX, radiusY, 0, 0, 2 * Math.PI);

        if (fillColor) {
            this.m_oContext.fillStyle = this.colorToRgb(fillColor);
            this.m_oContext.fill();
        }

        // Draw bottom ellipse
        this.m_oContext.beginPath();
        this.m_oContext.ellipse(centerX, y + height - radiusY, radiusX, radiusY, 0, 0, 2 * Math.PI);

        if (fillColor) {
            this.m_oContext.fillStyle = this.colorToRgb(fillColor);
            this.m_oContext.fill();
        }

        // Draw strokes
        if (strokeColor) {
            this.m_oContext.strokeStyle = this.colorToRgb(strokeColor);
            this.m_oContext.lineWidth = lineWidth || 1;

            // Side lines
            this.m_oContext.beginPath();
            this.m_oContext.moveTo(x, y + radiusY);
            this.m_oContext.lineTo(x, y + height - radiusY);
            this.m_oContext.moveTo(x + width, y + radiusY);
            this.m_oContext.lineTo(x + width, y + height - radiusY);
            this.m_oContext.stroke();

            // Top ellipse
            this.m_oContext.beginPath();
            this.m_oContext.ellipse(centerX, y + radiusY, radiusX, radiusY, 0, 0, 2 * Math.PI);
            this.m_oContext.stroke();

            // Bottom ellipse
            this.m_oContext.beginPath();
            this.m_oContext.ellipse(centerX, y + height - radiusY, radiusX, radiusY, 0, 0, 2 * Math.PI);
            this.m_oContext.stroke();
        }

        this.m_oContext.restore();
    }

    drawCube(x, y, width, height, fillColor, strokeColor, lineWidth) {
        if (!this.m_oContext) {return;}

        this.m_oContext.save();

        const depth = Math.min(width, height) * 0.3;

        // Front face
        if (fillColor) {
            this.m_oContext.fillStyle = this.colorToRgb(fillColor);
            this.m_oContext.fillRect(x, y + depth, width - depth, height - depth);
        }

        // Top face
        this.m_oContext.beginPath();
        this.m_oContext.moveTo(x, y + depth);
        this.m_oContext.lineTo(x + depth, y);
        this.m_oContext.lineTo(x + width, y);
        this.m_oContext.lineTo(x + width - depth, y + depth);
        this.m_oContext.closePath();

        if (fillColor) {
            const darkerFill = this.darkenColor(fillColor, 0.8);
            this.m_oContext.fillStyle = this.colorToRgb(darkerFill);
            this.m_oContext.fill();
        }

        // Right face
        this.m_oContext.beginPath();
        this.m_oContext.moveTo(x + width - depth, y + depth);
        this.m_oContext.lineTo(x + width, y);
        this.m_oContext.lineTo(x + width, y + height - depth);
        this.m_oContext.lineTo(x + width - depth, y + height);
        this.m_oContext.closePath();

        if (fillColor) {
            const darkerFill = this.darkenColor(fillColor, 0.6);
            this.m_oContext.fillStyle = this.colorToRgb(darkerFill);
            this.m_oContext.fill();
        }

        // Draw strokes
        if (strokeColor) {
            this.m_oContext.strokeStyle = this.colorToRgb(strokeColor);
            this.m_oContext.lineWidth = lineWidth || 1;

            // Front face
            this.m_oContext.strokeRect(x, y + depth, width - depth, height - depth);

            // Top face
            this.m_oContext.beginPath();
            this.m_oContext.moveTo(x, y + depth);
            this.m_oContext.lineTo(x + depth, y);
            this.m_oContext.lineTo(x + width, y);
            this.m_oContext.lineTo(x + width - depth, y + depth);
            this.m_oContext.closePath();
            this.m_oContext.stroke();

            // Right face
            this.m_oContext.beginPath();
            this.m_oContext.moveTo(x + width - depth, y + depth);
            this.m_oContext.lineTo(x + width, y);
            this.m_oContext.lineTo(x + width, y + height - depth);
            this.m_oContext.lineTo(x + width - depth, y + height);
            this.m_oContext.closePath();
            this.m_oContext.stroke();

            // Connection lines
            this.m_oContext.beginPath();
            this.m_oContext.moveTo(x, y + depth);
            this.m_oContext.lineTo(x + depth, y);
            this.m_oContext.moveTo(x + width - depth, y + depth);
            this.m_oContext.lineTo(x + width, y);
            this.m_oContext.moveTo(x + width - depth, y + height);
            this.m_oContext.lineTo(x + width, y + height - depth);
            this.m_oContext.stroke();
        }

        this.m_oContext.restore();
    }

    drawPlus(x, y, width, height, fillColor, strokeColor, lineWidth) {
        if (!this.m_oContext) {return;}

        this.m_oContext.save();
        this.m_oContext.beginPath();

        const thickness = Math.min(width, height) * 0.3;
        const centerX = x + width / 2;
        const centerY = y + height / 2;

        // Horizontal bar
        this.m_oContext.moveTo(x, centerY - thickness / 2);
        this.m_oContext.lineTo(x + width, centerY - thickness / 2);
        this.m_oContext.lineTo(x + width, centerY + thickness / 2);
        this.m_oContext.lineTo(x, centerY + thickness / 2);
        this.m_oContext.closePath();

        // Vertical bar
        this.m_oContext.moveTo(centerX - thickness / 2, y);
        this.m_oContext.lineTo(centerX + thickness / 2, y);
        this.m_oContext.lineTo(centerX + thickness / 2, y + height);
        this.m_oContext.lineTo(centerX - thickness / 2, y + height);
        this.m_oContext.closePath();

        if (fillColor) {
            this.m_oContext.fillStyle = this.colorToRgb(fillColor);
            this.m_oContext.fill();
        }

        if (strokeColor) {
            this.m_oContext.strokeStyle = this.colorToRgb(strokeColor);
            this.m_oContext.lineWidth = lineWidth || 1;
            this.m_oContext.stroke();
        }

        this.m_oContext.restore();
    }

    drawLine(x1, y1, x2, y2, strokeColor, lineWidth) {
        if (!this.m_oContext) {return;}

        this.m_oContext.save();
        this.m_oContext.beginPath();

        this.m_oContext.moveTo(x1, y1);
        this.m_oContext.lineTo(x2, y2);

        // Default to black if no stroke color specified
        const color = this.colorToRgb(strokeColor) || 'rgba(0,0,0,1)';
        this.m_oContext.strokeStyle = color;
        // Convert mm to px for consistency
        const pxWidth = CoordinateTransform.mmToPixels(lineWidth || 1);
        this.m_oContext.lineWidth = pxWidth;
        this.m_oContext.stroke();

        this.m_oContext.restore();
    }

    drawLineWithArrows(x1, y1, x2, y2, strokeColor, lineWidth = 1, strokeInfo = null) {
        if (!this.m_oContext) {return;}

        const ctx = this.m_oContext;
        ctx.save();

        // Draw main line
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);

        // Default to black if no stroke color specified
        const color = this.colorToRgb(strokeColor) || 'rgba(0,0,0,1)';
        ctx.strokeStyle = color;
        // Convert mm to px (ensure visible min width)
        const pxWidth = Math.max(CoordinateTransform.mmToPixels(lineWidth || 1), 1);
        ctx.lineWidth = pxWidth;
        if (strokeInfo) { this.applyStrokeInfo(strokeInfo); }
        ctx.stroke();

        // Map PPTX arrow size attributes to scale factors
        // PPTX uses: sm (small), med (medium), lg (large) for both w (width) and len (length)
        const mapSize = (s) => {
            const v = (s || '').toString().toLowerCase();
            // PPTX specification size mappings
            // These values are based on PowerPoint's rendering behavior
            if (v === 'lg' || v === 'large') {return 3.0;}  // Large arrows are ~3x base size
            if (v === 'med' || v === 'medium') {return 2.0;} // Medium arrows are ~2x base size
            if (v === 'sm' || v === 'small') {return 1.0;}   // Small arrows are 1x base size
            return 2.0; // Default to medium if not specified
        };
        // Per request: headEnd applies at start (x1,y1), tailEnd applies at end (x2,y2)
        let startDef = strokeInfo && strokeInfo.headEnd ? {
            type: strokeInfo.headEnd.type || strokeInfo.headEnd.val || 'none',
            lengthScale: mapSize(strokeInfo.headEnd.len),
            widthScale: mapSize(strokeInfo.headEnd.w)
        } : null;
        let endDef = strokeInfo && strokeInfo.tailEnd ? {
            type: strokeInfo.tailEnd.type || strokeInfo.tailEnd.val || 'none',
            lengthScale: mapSize(strokeInfo.tailEnd.len),
            widthScale: mapSize(strokeInfo.tailEnd.w)
        } : null;

        // Skip drawing when type is 'none'
        if (startDef && (startDef.type === 'none' || startDef.type === null || startDef.type === undefined)) {
            startDef = null;
        }
        if (endDef && (endDef.type === 'none' || endDef.type === null || endDef.type === undefined)) {
            endDef = null;
        }

        // Draw arrowheads if requested
        if (startDef || endDef) {
            const angle = Math.atan2(y2 - y1, x2 - x1);
            // Get current transformation scale to make arrow heads proportional to display size
            const transform = ctx.getTransform();
            const currentScale = Math.sqrt(transform.a * transform.a + transform.b * transform.b);
            
            // Calculate arrow head size from line width
            // PPTX arrow heads scale with line width: length ~= 3*lineWidth, width ~= 2.5*lineWidth
            // These multipliers match PowerPoint's actual rendering behavior more closely
            const headLen = ctx.lineWidth * 3.0;  // Length proportional to line width
            const headWid = ctx.lineWidth * 2.5;  // Width proportional to line width

            

            if (startDef) {
                const scaleL = (startDef.lengthScale || 1);
                const scaleW = (startDef.widthScale || 1);
                this._drawArrowHead(ctx, x1, y1, angle + Math.PI, headLen * scaleL, headWid * scaleW, color, ctx.lineWidth, startDef.type);
            }
            if (endDef) {
                const scaleL = (endDef.lengthScale || 1);
                const scaleW = (endDef.widthScale || 1);
                this._drawArrowHead(ctx, x2, y2, angle, headLen * scaleL, headWid * scaleW, color, ctx.lineWidth, endDef.type);
            }
        } else {
            
        }

        ctx.restore();
    }

    _drawArrowHead(ctx, x, y, angle, length, width, strokeColor, lineWidth, type = 'triangle') {
        ctx.save();
        ctx.translate(x, y);
        ctx.rotate(angle);

        const halfW = width / 2;
        const normalizedType = (type || 'triangle').toString().toLowerCase();

        ctx.beginPath();
        // Draw different arrow types based on PPTX type attribute
        // Types: triangle, stealth, diamond, oval, arrow, open
        switch (normalizedType) {
            case 'open':
                // Open arrow - two lines forming a V shape
                ctx.moveTo(0, 0);
                ctx.lineTo(-length, -halfW);
                ctx.moveTo(0, 0);
                ctx.lineTo(-length, halfW);
                break;
            case 'stealth':
                // Stealth arrow - pointed/streamlined triangle with concave back
                ctx.moveTo(0, 0);
                ctx.lineTo(-length, -(halfW * 0.5));
                ctx.lineTo(-length * 0.85, 0); // Concave back for stealth look
                ctx.lineTo(-length, (halfW * 0.5));
                ctx.closePath();
                break;
            case 'diamond':
                // Diamond arrow - rhombus shape
                ctx.moveTo(0, 0);
                ctx.lineTo(-length / 2, -halfW);
                ctx.lineTo(-length, 0);
                ctx.lineTo(-length / 2, halfW);
                ctx.closePath();
                break;
            case 'oval':
                // Oval arrow - ellipse/circle
                ctx.ellipse(-length / 2, 0, length / 2, halfW, 0, 0, Math.PI * 2);
                break;
            case 'arrow': // Map to triangle (PPTX synonym)
            case 'triangle':
            case 'tri':
            default:
                // Default: triangle - standard filled triangle
                ctx.moveTo(0, 0);
                ctx.lineTo(-length, -halfW);
                ctx.lineTo(-length, halfW);
                ctx.closePath();
                break;
        }

        const color = typeof strokeColor === 'string' ? strokeColor : this.colorToRgb(strokeColor) || 'rgba(0,0,0,1)';
        if (normalizedType === 'open') {
            // Open arrows are stroked only
            ctx.strokeStyle = color;
            ctx.lineWidth = lineWidth || 1;
            ctx.stroke();
        } else {
            // All other arrow types are filled with a thin outline
            ctx.fillStyle = color;
            ctx.fill();
            ctx.strokeStyle = color;
            ctx.lineWidth = Math.max(0.5, (lineWidth || 1) * 0.3);
            ctx.stroke();
        }

        ctx.restore();
    }

    // ... continue with existing utility methods ...

    /**
     * Utility method to darken a color
     */
    darkenColor(color, factor) {
        if (typeof color === 'string') {
            // Convert hex to RGB
            const rgb = this.hexToRgb(color);
            return {
                r: Math.floor(rgb.r * factor),
                g: Math.floor(rgb.g * factor),
                b: Math.floor(rgb.b * factor),
                a: rgb.a
            };
        } else if (color.r !== undefined) {
            return {
                r: Math.floor(color.r * factor),
                g: Math.floor(color.g * factor),
                b: Math.floor(color.b * factor),
                a: color.a
            };
        }
        return color;
    }

    /**
     * Convert color to RGB string
     */
    colorToRgb(color) {
        if (typeof color === 'string') {
            return color;
        } else if (color && color.r !== undefined) {
            const a = color.a !== undefined ? color.a / 255 : 1;
            return `rgba(${color.r}, ${color.g}, ${color.b}, ${a})`;
        }
        return null;
    }

    /**
     * Create a gradient canvas gradient object for use as fillStyle
     */
    createGradient(x, y, w, h, gradient) {
        if (!this.m_oContext || !gradient || !gradient.stops || gradient.stops.length === 0) {
            return null;
        }

        let grad;
        if (gradient.type === 'linear') {
            // Linear gradient
            const angle = gradient.angle || 0;
            const radians = (angle * Math.PI) / 180;

            const x1 = x + w / 2 - (Math.cos(radians) * w) / 2;
            const y1 = y + h / 2 - (Math.sin(radians) * h) / 2;
            const x2 = x + w / 2 + (Math.cos(radians) * w) / 2;
            const y2 = y + h / 2 + (Math.sin(radians) * h) / 2;

            grad = this.m_oContext.createLinearGradient(x1, y1, x2, y2);
        } else {
            // Radial gradient
            const centerX = x + w / 2;
            const centerY = y + h / 2;
            const radius = Math.max(w, h) / 2;

            grad = this.m_oContext.createRadialGradient(centerX, centerY, 0, centerX, centerY, radius);
        }

        // Add color stops
        for (const stop of gradient.stops) {
            grad.addColorStop(stop.position, this.colorToRgb(stop.color));
        }

        return grad;
    }

    /**
     * Fill rectangle with color or gradient
     */
    fillRect(x, y, width, height, fillColor) {
        if (!this.m_oContext || !fillColor) {
            return;
        }

        this.m_oContext.save();

        // Check if fillColor is a gradient object
        if (fillColor && typeof fillColor === 'object' && fillColor.type === 'linear' && fillColor.stops) {
            const gradient = this.createGradient(x, y, width, height, fillColor);
            if (gradient) {
                this.m_oContext.fillStyle = gradient;
            } else {
                this.m_oContext.fillStyle = this.colorToRgb(fillColor.stops[0]?.color || { r: 0, g: 0, b: 0 });
            }
        } else {
            this.m_oContext.fillStyle = this.colorToRgb(fillColor);
        }

        this.m_oContext.fillRect(x, y, width, height);
        this.m_oContext.restore();
    }

    /**
     * Convert hex color to RGB object
     */
    hexToRgb(hex) {
        const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
        return result ? {
            r: parseInt(result[1], 16),
            g: parseInt(result[2], 16),
            b: parseInt(result[3], 16),
            a: 255
        } : { r: 0, g: 0, b: 0, a: 255 };
    }

    /**
     * Save graphics state
     */
    SaveGrState() {
        this.m_oContext.save();
        this.m_oTransformStack.push(this.m_oTransform.CreateDublicate());
    }

    /**
     * Restore graphics state
     */
    RestoreGrState() {
        this.m_oContext.restore();
        if (this.m_oTransformStack.length > 0) {
            this.m_oTransform = this.m_oTransformStack.pop();
        }
    }

    /**
     * Calculate coordinate transformation
     */
    calculateCoordTransform() {
        this.m_oCoordTransform.Reset();
        this.m_oCoordTransform.sx = this.m_dWidth / this.m_dWidthMM;
        this.m_oCoordTransform.sy = this.m_dHeight / this.m_dHeightMM;
    }

    /**
     * Calculate full transformation matrix
     */
    calculateFullTransform() {
        this.m_oFullTransform = this.m_oCoordTransform.CreateDublicate();
        this.m_oFullTransform.Multiply(this.m_oTransform, 1);
    }

    // ... rest of the existing methods remain the same ...

    /**
     * Helper methods for drawing basic shape paths
     */
    drawRoundRectPath(ctx, x, y, width, height, radius) {
        ctx.beginPath();
        ctx.moveTo(x + radius, y);
        ctx.lineTo(x + width - radius, y);
        ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
        ctx.lineTo(x + width, y + height - radius);
        ctx.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
        ctx.lineTo(x + radius, y + height);
        ctx.quadraticCurveTo(x, y + height, x, y + height - radius);
        ctx.lineTo(x, y + radius);
        ctx.quadraticCurveTo(x, y, x + radius, y);
    }

    drawTrianglePath(ctx, x, y, width, height) {
        ctx.moveTo(x + width / 2, y);
        ctx.lineTo(x + width, y + height);
        ctx.lineTo(x, y + height);
        ctx.closePath();
    }

    drawDiamondPath(ctx, x, y, width, height) {
        ctx.moveTo(x + width / 2, y);
        ctx.lineTo(x + width, y + height / 2);
        ctx.lineTo(x + width / 2, y + height);
        ctx.lineTo(x, y + height / 2);
        ctx.closePath();
    }

    drawPentagonPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const radius = Math.min(width, height) / 2;
        
        ctx.moveTo(centerX, y);
        for (let i = 1; i < 5; i++) {
            const angle = (i * 2 * Math.PI) / 5 - Math.PI / 2;
            const px = centerX + radius * Math.cos(angle);
            const py = centerY + radius * Math.sin(angle);
            ctx.lineTo(px, py);
        }
        ctx.closePath();
    }

    drawHexagonPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const radius = Math.min(width, height) / 2;
        
        ctx.moveTo(centerX + radius, centerY);
        for (let i = 1; i < 6; i++) {
            const angle = (i * Math.PI) / 3;
            const px = centerX + radius * Math.cos(angle);
            const py = centerY + radius * Math.sin(angle);
            ctx.lineTo(px, py);
        }
        ctx.closePath();
    }

    drawOctagonPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const radius = Math.min(width, height) / 2;
        
        ctx.moveTo(centerX + radius, centerY);
        for (let i = 1; i < 8; i++) {
            const angle = (i * Math.PI) / 4;
            const px = centerX + radius * Math.cos(angle);
            const py = centerY + radius * Math.sin(angle);
            ctx.lineTo(px, py);
        }
        ctx.closePath();
    }

    drawStarPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const outerRadius = Math.min(width, height) / 2;
        const innerRadius = outerRadius * 0.4;
        
        ctx.moveTo(centerX, y);
        for (let i = 0; i < 10; i++) {
            const angle = (i * Math.PI) / 5 - Math.PI / 2;
            const radius = i % 2 === 0 ? outerRadius : innerRadius;
            const px = centerX + radius * Math.cos(angle);
            const py = centerY + radius * Math.sin(angle);
            ctx.lineTo(px, py);
        }
        ctx.closePath();
    }

    drawRightArrowPath(ctx, x, y, width, height) {
        const headWidth = width * 0.6;
        const headHeight = height;
        const tailHeight = height * 0.6;
        const tailY = y + (height - tailHeight) / 2;
        
        ctx.moveTo(x, tailY);
        ctx.lineTo(x + headWidth, tailY);
        ctx.lineTo(x + headWidth, y);
        ctx.lineTo(x + width, y + height / 2);
        ctx.lineTo(x + headWidth, y + height);
        ctx.lineTo(x + headWidth, tailY + tailHeight);
        ctx.lineTo(x, tailY + tailHeight);
        ctx.closePath();
    }

    drawLeftArrowPath(ctx, x, y, width, height) {
        const headWidth = width * 0.4;
        const tailHeight = height * 0.6;
        const tailY = y + (height - tailHeight) / 2;
        
        ctx.moveTo(x, y + height / 2);
        ctx.lineTo(x + headWidth, y);
        ctx.lineTo(x + headWidth, tailY);
        ctx.lineTo(x + width, tailY);
        ctx.lineTo(x + width, tailY + tailHeight);
        ctx.lineTo(x + headWidth, tailY + tailHeight);
        ctx.lineTo(x + headWidth, y + height);
        ctx.closePath();
    }

    drawUpArrowPath(ctx, x, y, width, height) {
        const headHeight = height * 0.4;
        const tailWidth = width * 0.6;
        const tailX = x + (width - tailWidth) / 2;
        
        ctx.moveTo(x + width / 2, y);
        ctx.lineTo(x + width, y + headHeight);
        ctx.lineTo(tailX + tailWidth, y + headHeight);
        ctx.lineTo(tailX + tailWidth, y + height);
        ctx.lineTo(tailX, y + height);
        ctx.lineTo(tailX, y + headHeight);
        ctx.lineTo(x, y + headHeight);
        ctx.closePath();
    }

    drawDownArrowPath(ctx, x, y, width, height) {
        const headHeight = height * 0.4;
        const tailWidth = width * 0.6;
        const tailX = x + (width - tailWidth) / 2;
        const headY = y + height - headHeight;
        
        ctx.moveTo(tailX, y);
        ctx.lineTo(tailX + tailWidth, y);
        ctx.lineTo(tailX + tailWidth, headY);
        ctx.lineTo(x + width, headY);
        ctx.lineTo(x + width / 2, y + height);
        ctx.lineTo(x, headY);
        ctx.lineTo(tailX, headY);
        ctx.closePath();
    }
}

/**
 * Standard Shape Processor
 * Handles standard Standard preset shapes
 */
class StandardShapeProcessor {
    constructor() {
        this.shapeCache = new Map();
        this.pathCache = new Map();
    }

    /**
     * Process Standard preset shape
     */
    processPresetShape(onlyOfficeType, shapeType, params) {
        const cacheKey = `${onlyOfficeType}_${shapeType}_${params.width}_${params.height}`;

        if (this.shapeCache.has(cacheKey)) {
            return this.shapeCache.get(cacheKey);
        }

        let shapeData = null;

        // Process based on Standard shape type
        switch (onlyOfficeType) {
            case 1: // Rectangle
                shapeData = this.createRectangleShape(params);
                break;
            case 2: // RoundRect
                shapeData = this.createRoundRectShape(params);
                break;
            case 3: // Ellipse
                shapeData = this.createEllipseShape(params);
                break;
            case 4: // Diamond
                shapeData = this.createDiamondShape(params);
                break;
            case 5: // Triangle
                shapeData = this.createTriangleShape(params);
                break;
            case 6: // Right Triangle
                shapeData = this.createRightTriangleShape(params);
                break;
            case 7: // Parallelogram
                shapeData = this.createParallelogramShape(params);
                break;
            case 8: // Trapezoid
                shapeData = this.createTrapezoidShape(params);
                break;
            case 9: // Hexagon
                shapeData = this.createHexagonShape(params);
                break;
            case 10: // Octagon
                shapeData = this.createOctagonShape(params);
                break;
            case 11: // Plus
                shapeData = this.createPlusShape(params);
                break;
            case 12: // Star5
                shapeData = this.createStar5Shape(params);
                break;
            case 13: // Right Arrow
                shapeData = this.createRightArrowShape(params);
                break;
            case 20: // Line
                shapeData = this.createLineShape(params);
                break;
            case 22: // Can (Cylinder)
                shapeData = this.createCylinderShape(params);
                break;
            case 56: // Pentagon
                shapeData = this.createPentagonShape(params);
                break;
            case 66: // Left Arrow
                shapeData = this.createLeftArrowShape(params);
                break;
            case 67: // Down Arrow
                shapeData = this.createDownArrowShape(params);
                break;
            case 68: // Up Arrow
                shapeData = this.createUpArrowShape(params);
                break;
            case 69: // Left Right Arrow
                shapeData = this.createLeftRightArrowShape(params);
                break;
            case 70: // Up Down Arrow
                shapeData = this.createUpDownArrowShape(params);
                break;
            case 74: // Heart
                shapeData = this.createHeartShape(params);
                break;
            case 76: // Quad Arrow
                shapeData = this.createQuadArrowShape(params);
                break;
            default:
                // Return null for unknown shapes - will fall back to custom processing
                return null;
        }

        if (shapeData) {
            this.shapeCache.set(cacheKey, shapeData);
        }

        return shapeData;
    }

    /**
     * Create rectangle shape data
     */
    createRectangleShape(params) {
        return {
            type: 'preset',
            onlyOfficeType: 1,
            path: {
                commands: [
                    { type: 'M', x: params.x, y: params.y },
                    { type: 'L', x: params.x + params.width, y: params.y },
                    { type: 'L', x: params.x + params.width, y: params.y + params.height },
                    { type: 'L', x: params.x, y: params.y + params.height },
                    { type: 'Z' }
                ]
            }
        };
    }

    /**
     * Create ellipse shape data
     */
    createEllipseShape(params) {
        const centerX = params.x + params.width / 2;
        const centerY = params.y + params.height / 2;
        const radiusX = params.width / 2;
        const radiusY = params.height / 2;

        return {
            type: 'preset',
            onlyOfficeType: 3,
            path: {
                commands: [
                    { type: 'M', x: centerX + radiusX, y: centerY },
                    { type: 'A', rx: radiusX, ry: radiusY, x: centerX - radiusX, y: centerY, rotation: 0, largeArc: 0, sweep: 1 },
                    { type: 'A', rx: radiusX, ry: radiusY, x: centerX + radiusX, y: centerY, rotation: 0, largeArc: 0, sweep: 1 },
                    { type: 'Z' }
                ]
            }
        };
    }

    /**
     * Create diamond shape data
     */
    createDiamondShape(params) {
        const centerX = params.x + params.width / 2;
        const centerY = params.y + params.height / 2;

        return {
            type: 'preset',
            onlyOfficeType: 4,
            path: {
                commands: [
                    { type: 'M', x: centerX, y: params.y },
                    { type: 'L', x: params.x + params.width, y: centerY },
                    { type: 'L', x: centerX, y: params.y + params.height },
                    { type: 'L', x: params.x, y: centerY },
                    { type: 'Z' }
                ]
            }
        };
    }

    /**
     * Create triangle shape data
     */
    createTriangleShape(params) {
        return {
            type: 'preset',
            onlyOfficeType: 5,
            path: {
                commands: [
                    { type: 'M', x: params.x + params.width / 2, y: params.y },
                    { type: 'L', x: params.x + params.width, y: params.y + params.height },
                    { type: 'L', x: params.x, y: params.y + params.height },
                    { type: 'Z' }
                ]
            }
        };
    }

    /**
     * Create right arrow shape data
     */
    createRightArrowShape(params) {
        const arrowWidth = params.width * 0.7;
        const arrowHeight = params.height * 0.4;
        const headWidth = params.width * 0.3;
        const headHeight = params.height * 0.6;

        return {
            type: 'preset',
            onlyOfficeType: 13,
            path: {
                commands: [
                    { type: 'M', x: params.x, y: params.y + (params.height - arrowHeight) / 2 },
                    { type: 'L', x: params.x + arrowWidth, y: params.y + (params.height - arrowHeight) / 2 },
                    { type: 'L', x: params.x + arrowWidth, y: params.y + (params.height - headHeight) / 2 },
                    { type: 'L', x: params.x + params.width, y: params.y + params.height / 2 },
                    { type: 'L', x: params.x + arrowWidth, y: params.y + (params.height + headHeight) / 2 },
                    { type: 'L', x: params.x + arrowWidth, y: params.y + (params.height + arrowHeight) / 2 },
                    { type: 'L', x: params.x, y: params.y + (params.height + arrowHeight) / 2 },
                    { type: 'Z' }
                ]
            }
        };
    }

    /**
     * Create heart shape data
     */
    createHeartShape(params) {
        const centerX = params.x + params.width / 2;
        const topY = params.y + params.height * 0.3;

        return {
            type: 'preset',
            onlyOfficeType: 74,
            path: {
                commands: [
                    { type: 'M', x: centerX, y: topY },
                    { type: 'C', x1: centerX - params.width * 0.3, y1: params.y, x2: params.x, y2: params.y + params.height * 0.3, x: centerX, y: params.y + params.height * 0.8 },
                    { type: 'C', x1: params.x + params.width, y1: params.y + params.height * 0.3, x2: centerX + params.width * 0.3, y2: params.y, x: centerX, y: topY },
                    { type: 'Z' }
                ]
            }
        };
    }

    // Add more shape creation methods as needed...

    /**
     * Create left arrow shape data
     */
    createLeftArrowShape(params) {
        const arrowWidth = params.width * 0.7;
        const arrowHeight = params.height * 0.4;
        const headWidth = params.width * 0.3;
        const headHeight = params.height * 0.6;

        return {
            type: 'preset',
            onlyOfficeType: 66,
            path: {
                commands: [
                    { type: 'M', x: params.x, y: params.y + params.height / 2 },
                    { type: 'L', x: params.x + headWidth, y: params.y + (params.height - headHeight) / 2 },
                    { type: 'L', x: params.x + headWidth, y: params.y + (params.height - arrowHeight) / 2 },
                    { type: 'L', x: params.x + params.width, y: params.y + (params.height - arrowHeight) / 2 },
                    { type: 'L', x: params.x + params.width, y: params.y + (params.height + arrowHeight) / 2 },
                    { type: 'L', x: params.x + headWidth, y: params.y + (params.height + arrowHeight) / 2 },
                    { type: 'L', x: params.x + headWidth, y: params.y + (params.height + headHeight) / 2 },
                    { type: 'Z' }
                ]
            }
        };
    }

    /**
     * Create plus shape data
     */
    createPlusShape(params) {
        const thickness = Math.min(params.width, params.height) * 0.3;
        const centerX = params.x + params.width / 2;
        const centerY = params.y + params.height / 2;

        return {
            type: 'preset',
            onlyOfficeType: 11,
            path: {
                commands: [
                    // Horizontal bar
                    { type: 'M', x: params.x, y: centerY - thickness / 2 },
                    { type: 'L', x: params.x + params.width, y: centerY - thickness / 2 },
                    { type: 'L', x: params.x + params.width, y: centerY + thickness / 2 },
                    { type: 'L', x: params.x, y: centerY + thickness / 2 },
                    { type: 'Z' },
                    // Vertical bar
                    { type: 'M', x: centerX - thickness / 2, y: params.y },
                    { type: 'L', x: centerX + thickness / 2, y: params.y },
                    { type: 'L', x: centerX + thickness / 2, y: params.y + params.height },
                    { type: 'L', x: centerX - thickness / 2, y: params.y + params.height },
                    { type: 'Z' }
                ]
            }
        };
    }

    // Additional shape creation methods can be added here...
}


// ... continue with existing coordinate transformation classes and utility functions ...

/**
 * Matrix transformation class
 */
class CMatrix {
    constructor() {
        this.sx = 1.0;
        this.shy = 0.0;
        this.shx = 0.0;
        this.sy = 1.0;
        this.tx = 0.0;
        this.ty = 0.0;
    }

    Reset() {
        this.sx = 1.0;
        this.shy = 0.0;
        this.shx = 0.0;
        this.sy = 1.0;
        this.tx = 0.0;
        this.ty = 0.0;
    }

    CreateDublicate() {
        const matrix = new CMatrix();
        matrix.sx = this.sx;
        matrix.shy = this.shy;
        matrix.shx = this.shx;
        matrix.sy = this.sy;
        matrix.tx = this.tx;
        matrix.ty = this.ty;
        return matrix;
    }

    Multiply(matrix, order) {
        if (order === 1) {
            // Post-multiply
            const sx = this.sx * matrix.sx + this.shy * matrix.shx;
            const shy = this.sx * matrix.shy + this.shy * matrix.sy;
            const shx = this.shx * matrix.sx + this.sy * matrix.shx;
            const sy = this.shx * matrix.shy + this.sy * matrix.sy;
            const tx = this.tx * matrix.sx + this.ty * matrix.shx + matrix.tx;
            const ty = this.tx * matrix.shy + this.ty * matrix.sy + matrix.ty;

            this.sx = sx;
            this.shy = shy;
            this.shx = shx;
            this.sy = sy;
            this.tx = tx;
            this.ty = ty;
        } else {
            // Pre-multiply
            const sx = matrix.sx * this.sx + matrix.shy * this.shx;
            const shy = matrix.sx * this.shy + matrix.shy * this.sy;
            const shx = matrix.shx * this.sx + matrix.sy * this.shx;
            const sy = matrix.shx * this.shy + matrix.sy * this.sy;
            const tx = matrix.tx * this.sx + matrix.ty * this.shx + this.tx;
            const ty = matrix.tx * this.shy + matrix.ty * this.sy + this.ty;

            this.sx = sx;
            this.shy = shy;
            this.shx = shx;
            this.sy = sy;
            this.tx = tx;
            this.ty = ty;
        }
    }
}

/**
 * Pen class for stroke properties
 */
class CPen {
    constructor() {
        this.Color = { r: 0, g: 0, b: 0, a: 255 };
        this.Alpha = 255;
        this.Size = 1.0;
        this.DashStyle = 0;
        this.LineJoin = 0;
        this.LineCap = 0;
        this.MiterLimit = 10.0;
    }
}

/**
 * Brush class for fill properties
 */
class CBrush {
    constructor() {
        this.Type = 0;
        this.Color1 = { r: 0, g: 0, b: 0, a: 255 };
        this.Color2 = { r: 255, g: 255, b: 255, a: 255 };
        this.Alpha1 = 255;
        this.Alpha2 = 255;
        this.TexturePath = '';
        this.TextureMode = 0;
        this.TextureAlpha = 255;
        this.LinearAngle = 0.0;
        this.Rectable = { x: 0, y: 0, w: 0, h: 0 };
    }
}

/**
 * Font class for text properties
 */
class CFont {
    constructor() {
        this.Name = 'Arial';
        this.Size = 12.0;
        this.Bold = false;
        this.Italic = false;
        this.Underline = false;
        this.Strikeout = false;
        this.Path = '';
        this.FaceIndex = 0;
        this.CharSpace = 0.0;
    }
}

// Export the enhanced graphics engine
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        CGraphics,
        StandardShapeProcessor,
        StandardShapeTypes,
        OpenXmlShapeTypes,
        CMatrix,
        CPen,
        CBrush,
        CFont
    };
}

/**
 * Mathematical functions used in Standard formulas
 */
const GeometryMath = {
    /**
     * Convert degrees to radians
     */
    toRadians: function(degrees) {
        return degrees * Math.PI / 180;
    },
    
    /**
     * Convert radians to degrees
     */
    toDegrees: function(radians) {
        return radians * 180 / Math.PI;
    },
    
    /**
     * Standard Sin function (handles 60000ths of degree input)
     */
    Sin: function(angle60000) {
        return Math.sin(angle60000 * Math.PI / 10800000);
    },
    
    /**
     * Standard Cos function (handles 60000ths of degree input)
     */
    Cos: function(angle60000) {
        return Math.cos(angle60000 * Math.PI / 10800000);
    },
    
    /**
     * Standard Tan function (handles 60000ths of degree input)
     */
    Tan: function(angle60000) {
        return Math.tan(angle60000 * Math.PI / 10800000);
    },
    
    /**
     * Standard ATan2 function (returns 60000ths of degree)
     */
    ATan2: function(y, x) {
        return Math.atan2(y, x) * 10800000 / Math.PI;
    },
    
    /**
     * Standard CAt2 function
     */
    CAt2: function(x, y, z) {
        return x * Math.cos(Math.atan2(z, y));
    },
    
    /**
     * Standard SAt2 function  
     */
    SAt2: function(x, y, z) {
        return x * Math.sin(Math.atan2(z, y));
    }
};

/**
 * Enhanced Geometry Class - Standard Compatible
 * Handles preset and custom geometries with full adjustment value support
 */
class EnhancedGeometry {
    constructor() {
        this.preset = null;
        this.pathLst = [];
        this.gdLst = {};          // Guide list (calculated values)
        this.gdLstInfo = [];      // Guide calculation info
        this.avLst = {};          // Adjustment value availability
        this.adjLst = {};         // Adjustment values
        this.ahXYLst = [];        // XY adjustment handles
        this.ahPolarLst = [];     // Polar adjustment handles
        this.cnxLst = [];         // Connection points
        this.rect = null;         // Text rectangle
        this.width = 21600;       // Default width in shape units
        this.height = 21600;      // Default height in shape units
        this.isCalculated = false;
    }
    
    /**
     * Set preset geometry type
     */
    setPreset(preset) {
        this.preset = preset;
        this.isCalculated = false;
    }
    
    /**
     * Add adjustment value
     */
    addAdjustment(name, value, min, max) {
        this.adjLst[name] = {
            value: value,
            min: min || null,
            max: max || null
        };
        this.avLst[name] = true;
        this.isCalculated = false;
    }
    
    /**
     * Add guide formula
     */
    addGuide(name, formula, x, y, z) {
        this.gdLstInfo.push({
            name: name,
            formula: formula,
            x: x,
            y: y,
            z: z
        });
        this.isCalculated = false;
    }
    
    /**
     * Calculate guide value using Standard formula system
     */
    calculateGuideValue(name, formula, x, y, z) {
        const xt = this.getValue(x);
        const yt = this.getValue(y);
        const zt = this.getValue(z);
        
        let result = 0;
        
        switch (formula) {
            case FORMULA_TYPE_MULT_DIV:  // */
                result = (zt !== 0) ? (xt * yt) / zt : 0;
                break;
            case FORMULA_TYPE_PLUS_MINUS: // +-
                result = xt + yt - zt;
                break;
            case FORMULA_TYPE_PLUS_DIV:  // +/
                result = (zt !== 0) ? (xt + yt) / zt : 0;
                break;
            case FORMULA_TYPE_IF_ELSE:   // ?:
                result = (xt > 0) ? yt : zt;
                break;
            case FORMULA_TYPE_ABS:       // abs
                result = Math.abs(xt);
                break;
            case FORMULA_TYPE_AT2:       // at2 (atan2)
                result = GeometryMath.ATan2(yt, xt);
                break;
            case FORMULA_TYPE_CAT2:      // cat2
                result = GeometryMath.CAt2(xt, yt, zt);
                break;
            case FORMULA_TYPE_COS:       // cos
                result = xt * GeometryMath.Cos(yt);
                break;
            case FORMULA_TYPE_MAX:       // max
                result = Math.max(xt, yt);
                break;
            case FORMULA_TYPE_MIN:       // min
                result = Math.min(xt, yt);
                break;
            case FORMULA_TYPE_MOD:       // mod
                result = Math.sqrt(xt*xt + yt*yt + zt*zt);
                break;
            case FORMULA_TYPE_PIN:       // pin
                if (yt < xt) {result = xt;}
                else if (yt > zt) {result = zt;}
                else {result = yt;}
                break;
            case FORMULA_TYPE_SAT2:      // sat2
                result = GeometryMath.SAt2(xt, yt, zt);
                break;
            case FORMULA_TYPE_SIN:       // sin
                result = xt * GeometryMath.Sin(yt);
                break;
            case FORMULA_TYPE_SQRT:      // sqrt
                result = Math.sqrt(Math.max(0, xt));
                break;
            case FORMULA_TYPE_TAN:       // tan
                result = xt * GeometryMath.Tan(yt);
                break;
            case FORMULA_TYPE_VALUE:     // val
                result = xt;
                break;
            default:
                result = 0;
                break;
        }
        
        if (isNaN(result)) {
            result = 0;
        }
        
        this.gdLst[name] = result;
        return result;
    }
    
    /**
     * Get value by name (from constants, guides, or adjustments)
     */
    getValue(name) {
        if (typeof name === 'number') {
            return name;
        }
        
        if (typeof name === 'string') {
            // Check for numeric string
            const numValue = parseFloat(name);
            if (!isNaN(numValue)) {
                return numValue;
            }
            
            // Check geometry constants
            switch (name) {
                case 'w': return this.width;
                case 'h': return this.height;
                case 'ss': return Math.min(this.width, this.height);
                case 'ls': return Math.max(this.width, this.height);
                case 'hc': return this.width / 2;
                case 'vc': return this.height / 2;
                case 'l': return 0;
                case 't': return 0;
                case 'r': return this.width;
                case 'b': return this.height;
                case 'wd2': return this.width / 2;
                case 'hd2': return this.height / 2;
                case 'wd4': return this.width / 4;
                case 'hd4': return this.height / 4;
                case 'wd6': return this.width / 6;
                case 'hd6': return this.height / 6;
                case 'wd8': return this.width / 8;
                case 'hd8': return this.height / 8;
                case 'cd2': return 10800000;
                case 'cd4': return 5400000;
                case 'cd8': return 2700000;
                case '_3cd4': return 16200000;
                case '_3cd8': return 8100000;
                case '_5cd8': return 13500000;
                case '_7cd8': return 18900000;
            }
            
            // Check calculated guides
            if (this.gdLst.hasOwnProperty(name)) {
                return this.gdLst[name];
            }
            
            // Check adjustment values
            if (this.adjLst.hasOwnProperty(name)) {
                return this.adjLst[name].value;
            }
        }
        
        return 0;
    }
    
    /**
     * Calculate all guides in proper dependency order
     */
    calculateGuides() {
        // Clear previous calculations
        this.gdLst = {};
        
        // Calculate guides in order (dependencies should be calculated first)
        for (let i = 0; i < this.gdLstInfo.length; i++) {
            const guide = this.gdLstInfo[i];
            this.calculateGuideValue(guide.name, guide.formula, guide.x, guide.y, guide.z);
        }
        
        this.isCalculated = true;
    }
    
    /**
     * Recalculate geometry with given width and height
     */
    recalculate(width, height) {
        this.width = width || 21600;
        this.height = height || 21600;
        this.calculateGuides();
    }
    
    /**
     * Add path to the geometry
     */
    addPath(path) {
        this.pathLst.push(path);
    }
    
    /**
     * Get path list for rendering
     */
    getPathList() {
        if (!this.isCalculated) {
            this.calculateGuides();
        }
        return this.pathLst;
    }
}

/**
 * Enhanced Path Class - Standard Compatible
 */
class EnhancedPath {
    constructor() {
        this.commands = [];
        this.fill = 'norm';          // Fill mode: 'norm', 'none'
        this.stroke = true;          // Whether to stroke
        this.w = undefined;          // Path width
        this.h = undefined;          // Path height
        this.extrusionOk = false;    // 3D extrusion allowed
    }
    
    /**
     * Add path command
     */
    addCommand(type, ...args) {
        this.commands.push({
            type: type,
            args: args
        });
    }
    
    /**
     * Move to point
     */
    moveTo(x, y) {
        this.addCommand('M', x, y);
    }
    
    /**
     * Line to point
     */
    lineTo(x, y) {
        this.addCommand('L', x, y);
    }
    
    /**
     * Cubic bezier curve
     */
    curveTo(x1, y1, x2, y2, x, y) {
        this.addCommand('C', x1, y1, x2, y2, x, y);
    }
    
    /**
     * Quadratic bezier curve
     */
    quadTo(x1, y1, x, y) {
        this.addCommand('Q', x1, y1, x, y);
    }
    
    /**
     * Arc
     */
    arcTo(rx, ry, xAxisRotation, largeArcFlag, sweepFlag, x, y) {
        this.addCommand('A', rx, ry, xAxisRotation, largeArcFlag, sweepFlag, x, y);
    }
    
    /**
     * Close path
     */
    close() {
        this.addCommand('Z');
    }
    
    /**
     * Set fill mode
     */
    setFill(fill) {
        this.fill = fill;
    }
    
    /**
     * Set stroke
     */
    setStroke(stroke) {
        this.stroke = stroke;
    }
}

/**
 * Enhanced Geometry Processor - Standard Compatible
 * Processes preset geometries using actual Standard geometry definitions
 */
class EnhancedGeometryProcessor {
    constructor() {
        this.shapeCache = new Map();
        this.presetDefinitions = new Map();
        this.initializePresetDefinitions();
    }
    
    /**
     * Initialize preset geometry definitions
     * Based on Standard CreateGeometry.js
     */
    initializePresetDefinitions() {
        // Rectangle
        this.presetDefinitions.set('rect', {
            guides: [],
            paths: [
                {
                    commands: [
                        { type: 'M', args: ['l', 't'] },
                        { type: 'L', args: ['r', 't'] },
                        { type: 'L', args: ['r', 'b'] },
                        { type: 'L', args: ['l', 'b'] },
                        { type: 'Z', args: [] }
                    ],
                    fill: 'norm',
                    stroke: true
                }
            ],
            connections: [
                { ang: 'cd4', x: 'hc', y: 't' },
                { ang: '0', x: 'r', y: 'vc' },
                { ang: '_3cd4', x: 'hc', y: 'b' },
                { ang: 'cd2', x: 'l', y: 'vc' }
            ],
            textRect: { l: 'l', t: 't', r: 'r', b: 'b' }
        });
        
        // Rounded Rectangle
        this.presetDefinitions.set('roundRect', {
            adjustments: [
                { name: 'adj', formula: FORMULA_TYPE_VALUE, value: 16667 }
            ],
            guides: [
                { name: 'a', formula: FORMULA_TYPE_PIN, x: '0', y: 'adj', z: '50000' },
                { name: 'x1', formula: FORMULA_TYPE_MULT_DIV, x: 'ss', y: 'a', z: '100000' },
                { name: 'y1', formula: FORMULA_TYPE_MULT_DIV, x: 'ss', y: 'a', z: '100000' },
                { name: 'x2', formula: FORMULA_TYPE_PLUS_MINUS, x: 'r', y: '0', z: 'x1' },
                { name: 'y2', formula: FORMULA_TYPE_PLUS_MINUS, x: 'b', y: '0', z: 'y1' },
                { name: 'il', formula: FORMULA_TYPE_MULT_DIV, x: 'x1', y: '29289', z: '100000' },
                { name: 'ir', formula: FORMULA_TYPE_PLUS_MINUS, x: 'r', y: '0', z: 'il' },
                { name: 'it', formula: FORMULA_TYPE_MULT_DIV, x: 'y1', y: '29289', z: '100000' },
                { name: 'ib', formula: FORMULA_TYPE_PLUS_MINUS, x: 'b', y: '0', z: 'it' }
            ],
            paths: [
                {
                    commands: [
                        { type: 'M', args: ['l', 'y1'] },
                        { type: 'Q', args: ['l', 't', 'x1', 't'] },
                        { type: 'L', args: ['x2', 't'] },
                        { type: 'Q', args: ['r', 't', 'r', 'y1'] },
                        { type: 'L', args: ['r', 'y2'] },
                        { type: 'Q', args: ['r', 'b', 'x2', 'b'] },
                        { type: 'L', args: ['x1', 'b'] },
                        { type: 'Q', args: ['l', 'b', 'l', 'y2'] },
                        { type: 'Z', args: [] }
                    ],
                    fill: 'norm',
                    stroke: true
                }
            ],
            handles: [
                { type: 'xy', gdRefX: 'adj', minX: '0', maxX: '50000', x: 'x1', y: 't' }
            ],
            connections: [
                { ang: 'cd4', x: 'hc', y: 't' },
                { ang: '0', x: 'r', y: 'vc' },
                { ang: '_3cd4', x: 'hc', y: 'b' },
                { ang: 'cd2', x: 'l', y: 'vc' }
            ],
            textRect: { l: 'il', t: 'it', r: 'ir', b: 'ib' }
        });
        
        // Ellipse
        this.presetDefinitions.set('ellipse', {
            guides: [
                { name: 'idx', formula: FORMULA_TYPE_MULT_DIV, x: 'wd2', y: '2700000', z: '21600000' },
                { name: 'idy', formula: FORMULA_TYPE_MULT_DIV, x: 'hd2', y: '2700000', z: '21600000' },
                { name: 'il', formula: FORMULA_TYPE_PLUS_MINUS, x: 'hc', y: '0', z: 'idx' },
                { name: 'ir', formula: FORMULA_TYPE_PLUS_MINUS, x: 'hc', y: 'idx', z: '0' },
                { name: 'it', formula: FORMULA_TYPE_PLUS_MINUS, x: 'vc', y: '0', z: 'idy' },
                { name: 'ib', formula: FORMULA_TYPE_PLUS_MINUS, x: 'vc', y: 'idy', z: '0' }
            ],
            paths: [
                {
                    commands: [
                        { type: 'M', args: ['l', 'vc'] },
                        { type: 'A', args: ['wd2', 'hd2', '0', '1', '1', 'r', 'vc'] },
                        { type: 'A', args: ['wd2', 'hd2', '0', '1', '1', 'l', 'vc'] },
                        { type: 'Z', args: [] }
                    ],
                    fill: 'norm',
                    stroke: true
                }
            ],
            connections: [
                { ang: 'cd4', x: 'hc', y: 't' },
                { ang: '0', x: 'r', y: 'vc' },
                { ang: '_3cd4', x: 'hc', y: 'b' },
                { ang: 'cd2', x: 'l', y: 'vc' }
            ],
            textRect: { l: 'il', t: 'it', r: 'ir', b: 'ib' }
        });
        
        // Add more preset definitions as needed...
        // This is a foundation that can be extended with more complex shapes
    }
    
    /**
     * Process preset geometry
     */
    processPresetGeometry(preset, width, height, adjustments = {}) {
        const cacheKey = `${preset}_${width}_${height}_${JSON.stringify(adjustments)}`;
        
        if (this.shapeCache.has(cacheKey)) {
            return this.shapeCache.get(cacheKey);
        }
        
        const definition = this.presetDefinitions.get(preset);
        if (!definition) {
            return this.createDefaultGeometry(width, height);
        }
        
        const geometry = new EnhancedGeometry();
        geometry.setPreset(preset);
        geometry.recalculate(width, height);
        
        // Set adjustments
        if (definition.adjustments) {
            definition.adjustments.forEach(adj => {
                const value = adjustments[adj.name] || adj.value;
                geometry.addAdjustment(adj.name, value, adj.min, adj.max);
            });
        }
        
        // Add guides
        if (definition.guides) {
            definition.guides.forEach(guide => {
                geometry.addGuide(guide.name, guide.formula, guide.x, guide.y, guide.z);
            });
        }
        
        // Create paths
        if (definition.paths) {
            definition.paths.forEach(pathDef => {
                const path = new EnhancedPath();
                path.setFill(pathDef.fill);
                path.setStroke(pathDef.stroke);
                
                pathDef.commands.forEach(cmd => {
                    const resolvedArgs = cmd.args.map(arg => {
                        return typeof arg === 'string' ? geometry.getValue(arg) : arg;
                    });
                    path.addCommand(cmd.type, ...resolvedArgs);
                });
                
                geometry.addPath(path);
            });
        }
        
        // Calculate final geometry
        geometry.calculateGuides();
        
        this.shapeCache.set(cacheKey, geometry);
        return geometry;
    }
    
    /**
     * Create default geometry (rectangle)
     */
    createDefaultGeometry(width, height) {
        return this.processPresetGeometry('rect', width, height);
    }

    /**
     * Fill text with style properties
     * Expected interface: fillText(text, x, y, styleProperties)
     * @param {string} text - Text to render
     * @param {number} x - X coordinate
     * @param {number} y - Y coordinate
     * @param {Object} styleProperties - Style properties including fontSize, color, fontFamily, etc.
     */
    fillText(text, x, y, styleProperties = {}) {
        if (!this.m_oContext || !text) {
            return;
        }

        this.m_oContext.save();

        try {
            // Extract style properties with defaults and apply scaling
            const baseFontSize = styleProperties.fontSize || 12;
            // Get scaling from graphics adapter if available, otherwise use 1.0
            const scaleFactor = this.m_oGraphicsAdapter && this.m_oGraphicsAdapter.getTextScaleFactor ? 
                                this.m_oGraphicsAdapter.getTextScaleFactor() : 1.0;
            const fontSize = baseFontSize * scaleFactor;
            const fontFamily = styleProperties.fontFamily || 'Arial';
            const fontWeight = styleProperties.fontWeight || (styleProperties.bold ? 'bold' : 'normal');
            const fontStyle = styleProperties.fontStyle || (styleProperties.italic ? 'italic' : 'normal');
            
            // Set font with emoji support
            const fontStack = `"${fontFamily}", "Apple Color Emoji", "Segoe UI Emoji", "Noto Color Emoji", "Segoe UI Symbol", Arial, sans-serif`;
            this.m_oContext.font = `${fontStyle} ${fontWeight} ${fontSize}px ${fontStack}`;
            
            // Set color
            if (styleProperties.color) {
                if (typeof styleProperties.color === 'object' && styleProperties.color.r !== undefined) {
                    // RGBA object format: {r: 255, g: 0, b: 0, a: 255} or {r: 255, g: 0, b: 0}
                    const alpha = styleProperties.color.a !== undefined ? styleProperties.color.a / 255 : 1;
                    const colorString = `rgba(${styleProperties.color.r}, ${styleProperties.color.g}, ${styleProperties.color.b}, ${alpha})`;
                    this.m_oContext.fillStyle = colorString;
                } else if (typeof styleProperties.color === 'string') {
                    // String format: "#FF0000" or "red"
                    this.m_oContext.fillStyle = styleProperties.color;
                } else {
                    // Fallback to black if color format is unexpected
                    this.m_oContext.fillStyle = '#000000';
                }
            } else {
                // Fallback to black if no color specified
                this.m_oContext.fillStyle = '#000000';
            }

            // Set text baseline and alignment
            this.m_oContext.textBaseline = styleProperties.textBaseline || 'alphabetic';
            this.m_oContext.textAlign = styleProperties.textAlign || 'left';

            // Apply transformations if needed
            if (this.m_oFullTransform && !this.m_oFullTransform.IsIdentity()) {
                const transform = this.m_oFullTransform;
                this.m_oContext.setTransform(
                    transform.sx, transform.shy,
                    transform.shx, transform.sy,
                    transform.tx, transform.ty
                );
            }

            // Draw highlight background behind text if present
            if (styleProperties.highlight) {
                const hl = styleProperties.highlight;
                let hlColor;
                if (typeof hl === 'object' && hl.r !== undefined) {
                    const hlAlpha = hl.a !== undefined ? hl.a / 255 : 1;
                    hlColor = `rgba(${hl.r}, ${hl.g}, ${hl.b}, ${hlAlpha})`;
                } else if (typeof hl === 'string') {
                    hlColor = hl;
                }
                if (hlColor) {
                    const metrics = this.m_oContext.measureText(text);
                    const hlHeight = fontSize * 1.2;
                    const savedFill = this.m_oContext.fillStyle;
                    this.m_oContext.fillStyle = hlColor;
                    this.m_oContext.fillRect(x, y - fontSize, metrics.width, hlHeight);
                    this.m_oContext.fillStyle = savedFill;
                }
            }

            // Apply text effects before rendering
            this.applyTextEffects(styleProperties.effects);

            // Render the text
            this.m_oContext.fillText(text, x, y);
            
            // Reset shadow after rendering
            this.resetTextEffects();
            
        } catch (error) {
        } finally {
            this.m_oContext.restore();
        }
    }

    /**
     * Apply text effects to canvas context
     */
    applyTextEffects(effects) {
        if (!effects || !this.m_oContext) {
            return;
        }

        // Apply outer shadow
        if (effects.outerShadow) {
            this.applyOuterShadow(effects.outerShadow);
        }

        // Apply glow effect
        if (effects.glow) {
            this.applyGlowEffect(effects.glow);
        }

        // Apply inner shadow
        if (effects.innerShadow) {
            this.applyInnerShadow(effects.innerShadow);
        }
    }

    /**
     * Apply outer shadow effect
     */
    applyOuterShadow(shadow) {
        if (!shadow || !this.m_oContext) {
            return;
        }

        // Convert EMU values to pixels
        const blurRadius = (shadow.blurRadius || 0) / 9525; // EMU to pixels (approx)
        const distance = (shadow.distance || 0) / 9525;
        const direction = (shadow.direction || 0) / 60000; // Convert to degrees

        // Calculate offset from distance and direction
        const angleRad = (direction * Math.PI) / 180;
        const offsetX = Math.cos(angleRad) * distance;
        const offsetY = Math.sin(angleRad) * distance;

        // Set shadow properties
        this.m_oContext.shadowOffsetX = offsetX;
        this.m_oContext.shadowOffsetY = offsetY;
        this.m_oContext.shadowBlur = blurRadius;

        // Set shadow color
        if (shadow.color) {
            const colorStr = this.convertColorToString(shadow.color);
            this.m_oContext.shadowColor = colorStr;
        } else {
            this.m_oContext.shadowColor = 'rgba(0, 0, 0, 0.5)';
        }
    }

    /**
     * Apply glow effect (simulated with shadow)
     */
    applyGlowEffect(glow) {
        if (!glow || !this.m_oContext) {
            return;
        }

        const radius = (glow.radius || 0) / 9525; // EMU to pixels

        // Set glow properties (no offset for glow)
        this.m_oContext.shadowOffsetX = 0;
        this.m_oContext.shadowOffsetY = 0;
        this.m_oContext.shadowBlur = radius;

        // Set glow color
        if (glow.color) {
            const colorStr = this.convertColorToString(glow.color);
            this.m_oContext.shadowColor = colorStr;
        } else {
            this.m_oContext.shadowColor = 'rgba(255, 255, 255, 0.8)';
        }
    }

    /**
     * Apply inner shadow effect (limited canvas support)
     */
    applyInnerShadow(shadow) {
        // Inner shadow is not directly supported by canvas
        // Could be implemented with composite operations in future
    }

    /**
     * Reset text effects
     */
    resetTextEffects() {
        if (!this.m_oContext) {
            return;
        }

        this.m_oContext.shadowOffsetX = 0;
        this.m_oContext.shadowOffsetY = 0;
        this.m_oContext.shadowBlur = 0;
        this.m_oContext.shadowColor = 'transparent';
    }

    /**
     * Convert parsed color object to CSS color string
     */
    convertColorToString(colorObj) {
        if (!colorObj) {
            return 'black';
        }

        if (colorObj.type === 'srgb') {
            const hex = colorObj.value;
            const r = parseInt(hex.substr(0, 2), 16);
            const g = parseInt(hex.substr(2, 2), 16);
            const b = parseInt(hex.substr(4, 2), 16);
            const alpha = colorObj.alpha || 1;
            return `rgba(${r}, ${g}, ${b}, ${alpha})`;
        } else if (colorObj.type === 'scheme') {
            // For scheme colors, use approximations
            const schemeColors = {
                'dk1': '#000000',
                'lt1': '#FFFFFF',
                'dk2': '#44546A',
                'lt2': '#E7E6E6'
            };
            const color = schemeColors[colorObj.value] || '#000000';
            const alpha = colorObj.alpha || 1;
            
            // Convert hex to rgba
            const r = parseInt(color.substr(1, 2), 16);
            const g = parseInt(color.substr(3, 2), 16);
            const b = parseInt(color.substr(5, 2), 16);
            return `rgba(${r}, ${g}, ${b}, ${alpha})`;
        }

        return 'black';
    }

    /**
     * Measure text with style properties
     * Expected interface: measureText(text, styleProperties)
     * @param {string} text - Text to measure
     * @param {Object} styleProperties - Style properties including fontSize, fontFamily, etc.
     * @returns {Object} Text metrics with width and height properties
     */
    measureText(text, styleProperties = {}) {
        if (!this.m_oContext || !text) {
            return { width: 0, height: 0 };
        }

        this.m_oContext.save();

        try {
            // Extract style properties with defaults and apply scaling
            const baseFontSize = styleProperties.fontSize || 12;
            // Get scaling from graphics adapter if available, otherwise use 1.0
            const scaleFactor = this.m_oGraphicsAdapter && this.m_oGraphicsAdapter.getTextScaleFactor ? 
                                this.m_oGraphicsAdapter.getTextScaleFactor() : 1.0;
            const fontSize = baseFontSize * scaleFactor;
            const fontFamily = styleProperties.fontFamily || 'Arial';
            const fontWeight = styleProperties.fontWeight || (styleProperties.bold ? 'bold' : 'normal');
            const fontStyle = styleProperties.fontStyle || (styleProperties.italic ? 'italic' : 'normal');
            
            // Set font for measurement with emoji support
            const fontStack = `"${fontFamily}", "Apple Color Emoji", "Segoe UI Emoji", "Noto Color Emoji", "Segoe UI Symbol", Arial, sans-serif`;
            this.m_oContext.font = `${fontStyle} ${fontWeight} ${fontSize}px ${fontStack}`;
            
            // Measure the text
            const metrics = this.m_oContext.measureText(text);
            
            // Calculate height based on font size (approximation)
            let height = fontSize;
            
            // Use actualBoundingBox if available for more accurate height
            if (metrics.actualBoundingBoxAscent !== undefined && metrics.actualBoundingBoxDescent !== undefined) {
                height = metrics.actualBoundingBoxAscent + metrics.actualBoundingBoxDescent;
            }
            
            return {
                width: metrics.width,
                height: height
            };
            
        } catch (error) {
            return { width: 0, height: 0 };
        } finally {
            this.m_oContext.restore();
        }
    }
}

// Export classes (maintain backward compatibility)
if (typeof window !== 'undefined') {
    window.CGraphics = CGraphics;
    window.StandardShapeProcessor = StandardShapeProcessor;
}

// Intentionally no ES module exports to support classic <script> usage in root demo
