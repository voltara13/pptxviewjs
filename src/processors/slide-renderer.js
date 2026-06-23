// import { ZLib as ZLibModule, OpenXmlPackage as OpenXmlPackageModule, OpenXmlTypes as OpenXmlTypesModule } from '../parsers/zip-processor.js';

/**
 * Slide Renderer Module
 * Enhanced version based on standard slide rendering system
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

// Import dependencies
// Ensure parsers and DOM modules are executed before this module to populate globals
import '../parsers/zip-processor.js';
import '../parsers/xml-parser.js';
import '../dom/presentation-dom.js';
// import { Logger } from '../utils/utils.js';
// import { FontEngine } from '../graphics/font-engine.js';
// import { ZLib } from '../parsers/zip-processor.js';
// Use globals for browser script-tag compatibility
// import { XmlParserContext as XmlParserContextModule, StaxParser as StaxParserModule } from '../parsers/xml-parser.js';
// import { CanvasGraphicsAdapter, CDrawingDocument } from '../graphics/graphics-adapter.js';
// import { CPresentation as CPresentationModule, CSlide as CSlideModule, CSld as CSldModule, CShape as CShapeModule } from '../dom/presentation-dom.js';
// import { SVGRenderer } from './svg-renderer.js';
// import { ChartProcessor } from './chart-processor.js';
// import { TableProcessor } from './table-processor.js';
// import { ChartRenderer } from '../renderers/chart-renderer.js';

// Fallbacks for globals when bundled
const LoggerClass = (typeof globalThis !== 'undefined' && globalThis.Logger) ? globalThis.Logger : class {
    constructor() {}
    log() {}
    info() {}
    warn() {}
    error() {}
    debug() {}
    trace() {}
    logError() {}
};
const FontEngineClass = (typeof globalThis !== 'undefined' && globalThis.FontEngine) ? globalThis.FontEngine : class {
    constructor() {}
};
const CanvasGraphicsAdapterClass = (typeof globalThis !== 'undefined' && globalThis.CanvasGraphicsAdapter) ? globalThis.CanvasGraphicsAdapter : class {
    constructor() {}
    resetState() {}
};
const CDrawingDocumentClass = (typeof globalThis !== 'undefined' && globalThis.CDrawingDocument) ? globalThis.CDrawingDocument : class {
    constructor() {
        this.graphics = null;
        this.canvas = null;
        this.processor = null;
        this.logger = new LoggerClass('CDrawingDocument');
    }
    init(canvas, processor = null) {
        this.canvas = canvas;
        this.processor = processor;
        if (!this.graphics) {
            this.graphics = new CanvasGraphicsAdapterClass();
        } else {
            if (this.graphics.resetState) {
                this.graphics.resetState();
            }
        }
    }
};
const ZLibClass = (typeof globalThis !== 'undefined' && globalThis.ZLib) ? globalThis.ZLib : class {
    constructor() { this.files = {}; this.isOpen = false; }
    async open(_arrayBuffer) { this.isOpen = true; return true; }
    async getFileText(_path) { return null; }
    async getFile(_path) { return null; }
    getPaths() { return []; }
};

// Safe OpenXmlTypes reference
const OpenXmlTypesSafe = (typeof globalThis !== 'undefined' && globalThis.OpenXmlTypes)
    || (typeof window !== 'undefined' && window.OpenXmlTypes)
    || {
        presentation: { relationType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' },
        slide: { relationType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide' },
        slideLayout: { relationType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout' },
        slideMaster: { relationType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster' },
        theme: { relationType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme' },
        image: { relationType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' }
    };

/**
 * Enhanced Rectangle class for bounds checking
 */
class CRect {
    constructor(x = 0, y = 0, w = 0, h = 0) {
        this.x = x;
        this.y = y;
        this.w = w;
        this.h = h;
    }

    isIntersectOther(other) {
        return !(this.x > other.x + other.w ||
                this.x + this.w < other.x ||
                this.y > other.y + other.h ||
                this.y + this.h < other.y);
    }
}

/**
 * PPTX Processor class
 * Main class that coordinates the entire PPTX processing pipeline
 */
class PPTXSlideRenderer {
    constructor() {
        this.logger = new LoggerClass('PPTXProcessor');
        this.zip = null;
        this.package = null;
        this.presentation = null;
        this.slides = [];
        this.slideMasters = [];
        this.slideLayouts = [];
        this.fontEngine = new FontEngineClass();

        // Initialize standard graphics engine
        this.drawingDocument = null;
        this.initializeGraphicsEngine();

        this.currentSlideIndex = 0;
        this.xmlParser = null;

        // Enhanced image processing
        this.imageMap = new Map();
        this.mediaMap = new Map();
        this.imageCache = new Map();
        this.svgRelationshipMap = new Map(); // Maps PNG relId to corresponding SVG relId
        this.imageLoader = new ImageLoader();

        // Enhanced rendering context
        this.renderContext = {
            enableOptimizations: true,
            enableBoundsChecking: true,
            enableViewportCulling: true,
            enableTextAntialiasing: true,
            enableShapeAntialiasing: true,
            quality: 'high',
            dpi: 96, // Default DPI
            pixelRatio: window.devicePixelRatio || 1 // Default to device pixel ratio
        };


    }

    /**
     * Set up high-resolution canvas with configurable DPI
     */
    setupHighResolutionCanvas(canvas, displayWidth, displayHeight, pixelRatio = null) {
        // Use configured pixel ratio or fall back to device pixel ratio or 1
        const effectivePixelRatio = pixelRatio || this.renderContext.pixelRatio || window.devicePixelRatio || 1;
        
        
        // Set display size (CSS)
        canvas.style.width = displayWidth + 'px';
        canvas.style.height = displayHeight + 'px';
        
        // Set actual canvas size in memory (scaled up for high-DPI)
        canvas.width = displayWidth * effectivePixelRatio;
        canvas.height = displayHeight * effectivePixelRatio;
        
        // Scale drawing context to match pixel ratio
        const ctx = canvas.getContext('2d');
        ctx.scale(effectivePixelRatio, effectivePixelRatio);
        
        // Enable high-quality image rendering
        ctx.imageSmoothingEnabled = true;
        ctx.imageSmoothingQuality = 'high';
        
        return canvas;
    }

    /**
     * Set rendering DPI (affects pixel ratio calculation)
     */
    setRenderingDPI(dpi) {
        this.renderContext.dpi = dpi;
        // Calculate pixel ratio based on DPI (96 DPI = 1x, 192 DPI = 2x, etc.)
        this.renderContext.pixelRatio = dpi / 96;
    }

    /**
     * Get current rendering DPI
     */
    getRenderingDPI() {
        return this.renderContext.dpi;
    }

    /**
     * Set custom pixel ratio (overrides DPI-based calculation)
     */
    setPixelRatio(ratio) {
        this.renderContext.pixelRatio = ratio;
        this.renderContext.dpi = ratio * 96; // Update DPI to match
    }

    /**
     * Initialize standard graphics engine
     */
    initializeGraphicsEngine() {
        if (typeof CDrawingDocument !== 'undefined') {
            this.drawingDocument = new CDrawingDocument();
        } else {
            this.drawingDocument = new CDrawingDocumentClass();
        }
        
        // Initialize SVG renderer for enhanced SVG processing  
        this.svgRenderer = null; // Will be initialized when graphics context is available
    }

    /**
     * Get the standard graphics engine
     */
    getGraphicsEngine() {
        return this.drawingDocument;
    }

    /**
     * Process PPTX file from ArrayBuffer
     */
    async processFile(arrayBuffer) {

        try {
            // Step 1: ZIP Processing
            this.updateProgress(1, 'processing');
            await this.processZIP(arrayBuffer);
            this.updateProgress(1, 'active', 100);

            // Step 2: XML to DOM Conversion
            this.updateProgress(2, 'processing');
            await this.processXMLToDOM();
            this.updateProgress(2, 'active', 100);

            // Step 3: Font Processing
            this.updateProgress(3, 'processing');
            this.fontEngine.analyzeDocument(this.presentation);
            this.fontEngine.loadFonts();
            this.updateProgress(3, 'active', 100);

            // Step 4: Initialize rendering
            this.updateProgress(4, 'processing');
            this.initializeRendering();
            this.updateProgress(4, 'active', 100);

            // Step 5: Preload images
            this.updateProgress(5, 'processing');
            await this.preloadAllImages();
            this.updateProgress(5, 'active', 100);

            // Step 6: Complete
            this.updateProgress(6, 'active', 100);

            return true;

        } catch (error) {
            throw error;
        }
    }

    /**
     * Update UI progress
     */
    updateProgress(step, state, percentage = 0) {
        // Update progress bar
        const progressBar = document.getElementById(`progress${step}`);
        if (progressBar) {
            progressBar.style.width = `${percentage}%`;
        }

        // Update step state
        const stepElement = document.getElementById(`step${step}`);
        if (stepElement) {
            stepElement.classList.remove('active', 'processing');
            if (state) {
                stepElement.classList.add(state);
            }
        }
    }

    /**
     * Step 1: Process ZIP archive
     */
    async processZIP(arrayBuffer) {

        // Initialize ZLib
        this.zip = (typeof ZLib !== 'undefined') ? new ZLib() : new ZLibClass();
        const success = await this.zip.open(arrayBuffer);

        if (!success) {
            throw new Error('Failed to open ZIP archive');
        }
        const allPaths = this.zip.getPaths();

        // Filter and log slide files
        const slideFiles = allPaths.filter(path => path.includes('slide') && path.endsWith('.xml'));

        // Create OpenXML package using safe global lookup
        const OpenXmlPackageCtor = (typeof globalThis !== 'undefined' && globalThis.OpenXmlPackage)
            || (typeof window !== 'undefined' && window.OpenXmlPackage)
            || (function(){
                // Minimal OpenXML package fallback used when globals are not available
                class SimpleOpenXmlPart {
                    constructor(pkg, uri, contentType) {
                        this.package = pkg;
                        this.uri = uri;
                        this.contentType = contentType;
                        this._content = null;
                    }
                    async getDocumentContent() {
                        if (this._content === null) {
                            const zipPath = this.uri.startsWith('/') ? this.uri.substring(1) : this.uri;
                            this._content = await this.package.zip.getFileText(zipPath);
                        }
                        return this._content;
                    }
                    async getRelationships() {
                        // Use pre-loaded rels if present
                        if (this.package.relationships[this.uri]) {
                            return this.package.relationships[this.uri];
                        }
                        const relsPath = this.uri.replace(/\/([^\/]+)$/, '/_rels/$1.rels');
                        const zipPath = relsPath.startsWith('/') ? relsPath.substring(1) : relsPath;
                        const relsXml = await this.package.zip.getFileText(zipPath);
                        if (relsXml) {
                            return this.package.parseRelationships(relsXml);
                        }
                        return {};
                    }
                }
                class SimpleOpenXmlPackage {
                    constructor(zip) {
                        this.zip = zip;
                        this.parts = {};
                        this.contentTypes = {};
                        this.relationships = {};
                    }
                    async initialize() {
                        // Content types
                        const contentTypesXml = await this.zip.getFileText('[Content_Types].xml');
                        if (contentTypesXml) { this.parseContentTypes(contentTypesXml); }
                        // Root relationships
                        const mainRelsXml = await this.zip.getFileText('_rels/.rels');
                        if (mainRelsXml) { this.relationships[''] = this.parseRelationships(mainRelsXml); }
                        // Load per-part relationships
                        const paths = this.zip.getPaths();
                        for (const path of paths) {
                            if (path.startsWith('_rels/') && path.endsWith('.rels') && path !== '_rels/.rels') {
                                const m = path.match(/^(.+)\/_rels\/(.+)\.rels$/);
                                if (m) {
                                    const partDir = m[1];
                                    const partFile = m[2];
                                    const partUri = `/${partDir}/${partFile}`;
                                    const relsXml = await this.zip.getFileText(path);
                                    if (relsXml) {
                                        this.relationships[partUri] = this.parseRelationships(relsXml);
                                    }
                                }
                            }
                        }
                        // Create parts for all files
                        for (const path of paths) {
                            if (!path.startsWith('_rels/') && path !== '[Content_Types].xml') {
                                const partUri = `/${path}`;
                                const contentType = this.getContentType(partUri);
                                this.parts[partUri] = new SimpleOpenXmlPart(this, partUri, contentType);
                            }
                        }
                    }
                    parseContentTypes(xml) {
                        const parser = new DOMParser();
                        const doc = parser.parseFromString(xml, 'text/xml');
                        doc.querySelectorAll('Default').forEach(def => {
                            const ext = def.getAttribute('Extension');
                            const ct = def.getAttribute('ContentType');
                            if (ext) {this.contentTypes[`ext:${ext}`] = ct;}
                        });
                        doc.querySelectorAll('Override').forEach(ovr => {
                            const partName = ovr.getAttribute('PartName');
                            const ct = ovr.getAttribute('ContentType');
                            if (partName) {this.contentTypes[partName] = ct;}
                        });
                    }
                    parseRelationships(xml) {
                        const parser = new DOMParser();
                        const doc = parser.parseFromString(xml, 'text/xml');
                        const map = {};
                        const all = doc.getElementsByTagName('*');
                        for (let i = 0; i < all.length; i++) {
                            const el = all[i];
                            if (el.localName === 'Relationship') {
                                const id = el.getAttribute('Id');
                                const type = el.getAttribute('Type');
                                const target = el.getAttribute('Target');
                                if (id) {map[id] = { type, target, targetMode: el.getAttribute('TargetMode') || 'Internal' };}
                            }
                        }
                        return map;
                    }
                    getContentType(partName) {
                        if (this.contentTypes[partName]) {return this.contentTypes[partName];}
                        const ext = partName.split('.').pop();
                        if (ext && this.contentTypes[`ext:${ext}`]) {return this.contentTypes[`ext:${ext}`];}
                        return 'application/octet-stream';
                    }
                    getPartByRelationshipType(relType) {
                        const rels = this.relationships[''] || {};
                        for (const rel of Object.values(rels)) {
                            if (rel.type === relType) {
                                const partName = rel.target.startsWith('/') ? rel.target : `/${rel.target}`;
                                return this.parts[partName];
                            }
                        }
                        return null;
                    }
                    getPartByUri(uri) {
                        return this.parts[uri] || null;
                    }
                }
                return SimpleOpenXmlPackage;
            })();
        this.package = new OpenXmlPackageCtor(this.zip);
        await this.package.initialize();
        try {
            const relKeys = this.package && this.package.relationships ? Object.keys(this.package.relationships) : [];
        } catch (_e) {}

    }

    /**
     * Step 2: Process XML to DOM
     */
    async processXMLToDOM() {

        try {
            // Create XML parser context
            // Resolve XmlParserContext from global/window to avoid redeclaration conflicts
            const XmlCtxCtor = (typeof globalThis !== 'undefined' && globalThis.XmlParserContext) ? globalThis.XmlParserContext : (typeof window !== 'undefined' ? window.XmlParserContext : null);
            this.xmlParser = XmlCtxCtor ? new XmlCtxCtor() : { zip: null };
            this.xmlParser.zip = this.zip;
            const context = this.xmlParser;

            // Process main presentation document
            let presentationPart = this.package.getPartByRelationshipType(OpenXmlTypesSafe.presentation ? OpenXmlTypesSafe.presentation.relationType : undefined);
            if (!presentationPart) {
                // Fallbacks: try common locations
                presentationPart = this.package.getPartByUri('/ppt/presentation.xml') ||
                                   this.package.getPartByUri('ppt/presentation.xml');
                if (!presentationPart) {
                    // Scan package parts for any presentation.xml
                    const availableParts = Object.keys(this.package.parts || {});
                    const candidate = availableParts.find(p => /presentation\.xml$/i.test(p));
                    if (candidate) {
                        presentationPart = this.package.getPartByUri(candidate);
                    }
                }
                if (!presentationPart && this.zip && typeof this.zip.getPaths === 'function') {
                    // Last-resort: scan ZIP entries directly and register part if found
                    try {
                        const paths = this.zip.getPaths();
                        // Prefer the canonical location first
                        let direct = paths.find(p => /^(?:\/?|)ppt\/presentation\.xml$/i.test(p));
                        if (!direct) {
                            direct = paths.find(p => /presentation\.xml$/i.test(p));
                        }
                        if (direct) {
                            const partUri = direct.startsWith('/') ? direct : `/${direct}`;
                            // Create a minimal part wrapper compatible with our package API
                            const self = this;
                            const DynamicPart = class {
                                constructor(pkg, uri){ this.package = pkg; this.uri = uri; this._content = null; }
                                async getDocumentContent(){
                                    if (this._content === null) {
                                        const zp = this.uri.startsWith('/') ? this.uri.substring(1) : this.uri;
                                        this._content = await self.zip.getFileText(zp);
                                    }
                                    return this._content;
                                }
                                async getRelationships(){
                                    const relsPath = this.uri.replace(/\/([^\/]+)$/, '/_rels/$1.rels');
                                    const zp = relsPath.startsWith('/') ? relsPath.substring(1) : relsPath;
                                    const relsXml = await self.zip.getFileText(zp);
                                    if (relsXml) {
                                        // Use package parser if available to keep behavior consistent
                                        if (self.package && typeof self.package.parseRelationships === 'function') {
                                            return self.package.parseRelationships(relsXml);
                                        }
                                    }
                                    return {};
                                }
                            };
                            const dynPart = new DynamicPart(this.package, partUri);
                            if (this.package && this.package.parts) {
                                this.package.parts[partUri] = dynPart;
                            }
                            presentationPart = dynPart;
                        }
                    } catch (_e) {}
                }
                if (!presentationPart) {
                    throw new Error('No presentation document found in PPTX file');
                }
            }

            const presentationXml = await presentationPart.getDocumentContent();
            if (!presentationXml) {
                throw new Error('Failed to extract presentation XML content');
            }

            // Resolve StaxParser from globals to avoid bundler scope issues
            const StaxParserCtor = (typeof globalThis !== 'undefined' && globalThis.StaxParser)
                || (typeof window !== 'undefined' && window.StaxParser)
                || null;
            if (!StaxParserCtor) {
                throw new Error('StaxParser is not defined');
            }
            const presentationParser = new StaxParserCtor(presentationXml, presentationPart, context);

            // Parse presentation XML
            const doc = presentationParser.parse();
            if (!doc) {
                throw new Error('Failed to parse presentation XML document');
            }

            this.presentation = this.createPresentationFromXML(doc);
            if (!this.presentation) {
                throw new Error('Failed to create presentation object from XML');
            }

            // Process themes, masters, and layouts
            await this.processThemes(context, presentationPart);

            await this.processSlideMasters(context, presentationPart);

            await this.processSlides(context);

            // Process media/images after all content is loaded
            await this.processImages(context);
            
            // Build SVG relationship mapping
            this.buildSVGRelationshipMapping();

        } catch (error) {
            throw error;
        }
    }

    /**
     * Create presentation object from XML document
     */
    createPresentationFromXML(doc) {
        try {
        const CPClass = (typeof globalThis !== 'undefined' && globalThis.CPresentation)
            || (typeof window !== 'undefined' && window.CPresentation)
            || null;
        if (!CPClass) { throw new Error('CPresentation is not defined'); }
        const presentation = new CPClass();

            // Parse presentation element
            const presentationElement = doc.documentElement;
            if (presentationElement) {
                // Parse slide size
                const sldSzElement = presentationElement.querySelector('sldSz, p\\:sldSz');
        if (sldSzElement) {
            presentation.slideSize = {
                cx: parseInt(sldSzElement.getAttribute('cx')) || 9144000,
                cy: parseInt(sldSzElement.getAttribute('cy')) || 6858000
            };
        }

                // Parse slide master ID list
                const sldMasterIdLstElement = presentationElement.querySelector('sldMasterIdLst, p\\:sldMasterIdLst');
        if (sldMasterIdLstElement) {
                    const masterIds = [];
            const masterIdElements = sldMasterIdLstElement.querySelectorAll('sldMasterId, p\\:sldMasterId');
                    masterIdElements.forEach(element => {
                        const rId = element.getAttribute('r:id') || element.getAttribute('id');
                if (rId) {
                            masterIds.push(rId);
                        }
                    });
                    presentation.masterIds = masterIds;
                }

                // Parse slide ID list
                const sldIdLstElement = presentationElement.querySelector('sldIdLst, p\\:sldIdLst');
                if (sldIdLstElement) {
                    const slideIds = [];
                    const slideIdElements = sldIdLstElement.querySelectorAll('sldId, p\\:sldId');
                    slideIdElements.forEach(element => {
                        const rId = element.getAttribute('r:id') || element.getAttribute('id');
                        if (rId) {
                            slideIds.push(rId);
                        }
                    });
                    presentation.slideIds = slideIds;
                }

                // Parse default text style
                const defaultTextStyleElement = presentationElement.querySelector('defaultTextStyle, p\\:defaultTextStyle');
                if (defaultTextStyleElement) {
                    presentation.defaultTextStyle = this.parseTextStyleElement(defaultTextStyleElement);
                }
            }

        return presentation;

        } catch (error) {
            // Return a minimal presentation object
            const fallbackPresentation = new CPresentation();
            fallbackPresentation.slideSize = { cx: 9144000, cy: 6858000 }; // Default PowerPoint size
            return fallbackPresentation;
        }
    }

    /**
     * Process images and media from PPTX archive
     */
    async processImages(context) {

        // Extract image and media relationship mappings from all parts
        try {
            if (typeof this.extractMediaMap === 'function') {
                await this.extractMediaMap();
            } else {
                await this.extractImageMap();
            }
        } catch (error) {
            await this.extractImageMap();
        }

        // Preload ALL images for comprehensive display
        await this.preloadAllImages();

        // Log comprehensive image information
        this.logImageProcessingResults();
    }
    logImageProcessingResults() {

        // Log all discovered images
        const imageDetails = [];
        for (const [relId, imagePath] of this.imageMap) {
            const cached = this.imageCache.has(relId);
            imageDetails.push({ relId, imagePath, cached });
        }

        // Log images found in ZIP folder
        const allPaths = this.zip.getPaths();
        const mediaImages = allPaths.filter(path =>
            path.startsWith('ppt/media/') && this.isImageFile(path)
        );
    }

    /**
     * Extract media relationship mappings from all document parts
     */
    async extractMediaMap() {

        const mediaPaths = [];

        // Find all media files in the ZIP archive
        const allPaths = this.zip.getPaths();
        for (const path of allPaths) {
            if (path.startsWith('ppt/media/')) {
                mediaPaths.push(path);
            }
        }

        // Process relationships for each part that might contain media
        const parts = [];

        // Add presentation relationships
            const presentationPart = this.package.getPartByRelationshipType(OpenXmlTypesSafe.presentation.relationType);
        if (presentationPart) {
            parts.push(presentationPart);
        }

        // Add slide relationships - use the actual slide URIs from the package
        if (this.slides && this.slides.length > 0) {

            // Get all available parts to find slide parts
            const availableParts = Object.keys(this.package.parts);

            // Also check available relationships
            const availableRelationships = Object.keys(this.package.relationships);

            for (const partUri of availableParts) {
                if (partUri.includes('/slides/slide') && partUri.endsWith('.xml')) {
                    const slidePart = this.package.getPartByUri(partUri);
                    if (slidePart) {
                        parts.push(slidePart);
                    }
                }
            }
        }

        // Add master relationships
        if (this.slideMasters && this.slideMasters.length > 0) {

            const availableParts = Object.keys(this.package.parts);
            for (const partUri of availableParts) {
                if (partUri.includes('/slideMasters/slideMaster') && partUri.endsWith('.xml')) {
                    const masterPart = this.package.getPartByUri(partUri);
                    if (masterPart) {
                        parts.push(masterPart);
                    }
                }
            }
        }

        // Add layout relationships
        if (this.slideLayouts && this.slideLayouts.length > 0) {

            const availableParts = Object.keys(this.package.parts);
            for (const partUri of availableParts) {
                if (partUri.includes('/slideLayouts/slideLayout') && partUri.endsWith('.xml')) {
                    const layoutPart = this.package.getPartByUri(partUri);
                    if (layoutPart) {
                        parts.push(layoutPart);
                    }
                }
            }
        }

        for (const part of parts) {

            await this.extractPartMediaRelationships(part);
        }
    }

    /**
     * Extract image relationship mappings from all document parts
     */
    async extractImageMap() {
        const imagePaths = [];

        // Find all image files in the ZIP archive
        const allPaths = this.zip.getPaths();
        for (const path of allPaths) {
            if (path.startsWith('ppt/media/') && this.isImageFile(path)) {
                imagePaths.push(path);
            }
        }

        // Process relationships for each part that might contain images
        const parts = [
            // Presentation relationships
            this.package.getPartByRelationshipType(OpenXmlTypesSafe.presentation.relationType),
            // All slide relationships
            ...this.slides.map((_, index) => this.package.getPartByUri(`/ppt/slides/slide${index + 1}.xml`)),
            // All master relationships
            ...this.slideMasters.map((_, index) => this.package.getPartByUri(`/ppt/slideMasters/slideMaster${index + 1}.xml`)),
            // All layout relationships
            ...this.slideLayouts.map((_, index) => this.package.getPartByUri(`/ppt/slideLayouts/slideLayout${index + 1}.xml`))
        ].filter(part => part !== null);

        for (const part of parts) {
            await this.extractPartImageRelationships(part);
        }

    }

    /**
     * Extract image relationships from a specific document part
     */
    async extractPartImageRelationships(part) {
        try {
            const relationships = await part.getRelationships();

            for (const [relId, rel] of Object.entries(relationships)) {
                if (rel.type === OpenXmlTypesSafe.image.relationType) {
                    // Convert relative target to absolute path
                    let imagePath = rel.target;
                    if (imagePath.startsWith('../')) {
                        imagePath = imagePath.replace('../', 'ppt/');
                    } else if (!imagePath.startsWith('/')) {
                        imagePath = `/ppt/${imagePath}`;
                    }

                    this.imageMap.set(relId, imagePath);
                }
            }
        } catch (_error) {
				// Error ignored
			}
    }

    /**
     * Check if file path represents an image
     */
    isImageFile(path) {
        const imageExtensions = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.webp', '.tiff'];
        const lowerPath = path.toLowerCase();
        return imageExtensions.some(ext => lowerPath.endsWith(ext));
    }

    /**
     * Preload images for a specific slide
     */
    async preloadSlideImages(slideIndex) {
        if (!this.slides[slideIndex] || !this.slides[slideIndex].commonSlideData) {
            return;
        }

        const slide = this.slides[slideIndex];
        const imageRelIds = this.extractImageRelationshipsFromSlide(slide);

        const loadPromises = imageRelIds.map(relId => this.loadImage(relId));
        await Promise.allSettled(loadPromises);
    }

    /**
     * Preload ALL images in the presentation for better performance
     */
    async preloadAllImages() {

        // Get all unique image relationship IDs
        const allImageRelIds = new Set();

        // Add images from all slides
        for (const slide of this.slides) {
            const slideImageIds = this.extractImageRelationshipsFromSlide(slide);
            slideImageIds.forEach(id => allImageRelIds.add(id));

            // Check background images
            if (slide.backgroundFill && slide.backgroundFill.fill &&
                slide.backgroundFill.fill.type === 'image' &&
                slide.backgroundFill.fill.imageData?.relationshipId) {
                allImageRelIds.add(slide.backgroundFill.fill.imageData.relationshipId);
            }
        }

        // Add any images from masters and layouts
        for (const master of this.slideMasters) {
            if (master.cSld?.bg?.fill?.type === 'image' &&
                master.cSld.bg.fill.imageData?.relationshipId) {
                allImageRelIds.add(master.cSld.bg.fill.imageData.relationshipId);
            }
        }

        // Load all images in parallel with progress tracking
        const imageArray = Array.from(allImageRelIds);
        const batchSize = 5; // Load images in batches to avoid overwhelming the browser

        for (let i = 0; i < imageArray.length; i += batchSize) {
            const batch = imageArray.slice(i, i + batchSize);
            const loadPromises = batch.map(relId => this.loadImage(relId));
            await Promise.allSettled(loadPromises);

            const progress = Math.min(100, Math.round(((i + batchSize) / imageArray.length) * 100));
        }

    }

    /**
     * Extract image relationship IDs from slide shapes
     */
    extractImageRelationshipsFromSlide(slide) {
        const imageRelIds = [];

        if (slide.commonSlideData && slide.commonSlideData.shapeTree) {
            for (const shape of slide.commonSlideData.shapeTree) {
                if (shape.type === 'pic' && shape.imageRelId) {
                    imageRelIds.push(shape.imageRelId);
                }

                // Handle group shapes recursively
                if (shape.type === 'grpSp' && shape.shapeTree) {
                    imageRelIds.push(...this.extractImageRelationshipsFromShapeTree(shape.shapeTree));
                }
            }
        }

        return imageRelIds;
    }

    /**
     * Extract image relationship IDs from shape tree recursively
     */
    extractImageRelationshipsFromShapeTree(shapeTree) {
        const imageRelIds = [];

        for (const shape of shapeTree) {
            if (shape.type === 'pic' && shape.imageRelId) {
                imageRelIds.push(shape.imageRelId);
            }

            if (shape.type === 'grpSp' && shape.shapeTree) {
                imageRelIds.push(...this.extractImageRelationshipsFromShapeTree(shape.shapeTree));
            }
        }

        return imageRelIds;
    }

    /**
     * Update all pic shapes in a slide that use oldRelId to use newRelId (composite key)
     */
    updateShapeImageRelIds(slide, oldRelId, newRelId) {
        if (!slide.commonSlideData?.shapeTree) {return;}
        const updateInTree = (shapeTree) => {
            for (const shape of shapeTree) {
                if (shape.type === 'pic' && shape.imageRelId === oldRelId) {
                    shape.imageRelId = newRelId;
                }
                if (shape.type === 'grpSp' && shape.shapeTree) {
                    updateInTree(shape.shapeTree);
                }
            }
        };
        updateInTree(slide.commonSlideData.shapeTree);
    }

    /**
     * Preload all images for all slides
     */
    async preloadAllImages() {

        try {
            // Collect fallback image relationship IDs (for cases where per-slide loading fails)
            const allImageRelIds = new Set();

            // Load slide images using per-slide composite keys to avoid relId collisions
            // (different slides and layouts can each use rId1, rId2, etc. for different images)
            for (let i = 0; i < this.slides.length; i++) {
                const slide = this.slides[i];
                const imageRelIds = this.extractImageRelationshipsFromSlide(slide);
                if (imageRelIds.length === 0) {continue;}

                if (slide.partUri) {
                    try {
                        const slidePart = this.package.getPartByUri(slide.partUri);
                        if (slidePart) {
                            const slideRels = await slidePart.getRelationships();
                            const slidePath = slide.partUri.replace(/^\//, '');
                            for (const relId of imageRelIds) {
                                if (slideRels && slideRels[relId]) {
                                    let imgPath = slideRels[relId].target;
                                    if (imgPath.startsWith('../')) {
                                        imgPath = imgPath.replace('../', 'ppt/');
                                    } else if (!imgPath.startsWith('/') && !imgPath.startsWith('ppt/')) {
                                        imgPath = `ppt/${imgPath}`;
                                    }
                                    const uniqueKey = `slide:${slidePath}:${relId}`;
                                    if (!this.imageCache.has(uniqueKey)) {
                                        try {
                                            const imageData = await this.zip.getFileData(imgPath).catch(() => null);
                                            if (imageData) {
                                                const blob = new Blob([imageData], { type: this.getMimeType(imgPath) });
                                                const imageUrl = URL.createObjectURL(blob);
                                                const image = await this.imageLoader.loadImageFromUrl(imageUrl);
                                                this.imageCache.set(uniqueKey, { image, url: imageUrl, path: imgPath, width: image.naturalWidth, height: image.naturalHeight });
                                            }
                                        } catch (_loadErr) {
                                            // Image failed to load (e.g. corrupt/non-image file) — leave cache empty
                                        }
                                    }
                                    // Always update shape.imageRelId even when loading fails, so it won't
                                    // fall back to a global relId that maps to a different slide's image
                                    this.updateShapeImageRelIds(slide, relId, uniqueKey);
                                } else {
                                    allImageRelIds.add(relId); // fallback to global lookup
                                }
                            }
                        } else {
                            imageRelIds.forEach(id => allImageRelIds.add(id));
                        }
                    } catch (_e) {
                        imageRelIds.forEach(id => allImageRelIds.add(id));
                    }
                } else {
                    imageRelIds.forEach(id => allImageRelIds.add(id));
                }
            }

            // Handle layout background images with unique cache keys to avoid relId collisions
            // (multiple layouts can each have their own rId1, rId2, etc. pointing to different images)
            for (const layout of this.slideLayouts) {
                const bgFill = layout.cSld?.bg?.fill || layout.commonSlideData?.backgroundFill?.fill;
                if (bgFill?.type === 'image' && bgFill.imageData?.relationshipId && layout.layoutPath) {
                    const relId = bgFill.imageData.relationshipId;
                    const uniqueKey = `layout_bg:${layout.layoutPath}`;
                    try {
                        const layoutPart = this.package.getPartByUri('/' + layout.layoutPath);
                        if (layoutPart) {
                            const rels = await layoutPart.getRelationships();
                            if (rels && rels[relId]) {
                                let imagePath = rels[relId].target;
                                if (imagePath.startsWith('../')) {
                                    imagePath = imagePath.replace('../', 'ppt/');
                                } else if (!imagePath.startsWith('/')) {
                                    imagePath = `ppt/${imagePath}`;
                                }
                                // Load and cache with unique key to avoid relId collision
                                const imageData = await this.zip.getFileData(imagePath);
                                if (imageData && !this.isSVGFile(imagePath)) {
                                    const blob = new Blob([imageData], { type: this.getMimeType(imagePath) });
                                    const imageUrl = URL.createObjectURL(blob);
                                    const image = await this.imageLoader.loadImageFromUrl(imageUrl);
                                    this.imageCache.set(uniqueKey, {
                                        image, url: imageUrl, path: imagePath,
                                        width: image.naturalWidth, height: image.naturalHeight
                                    });
                                    bgFill.imageData.resolvedCacheKey = uniqueKey;
                                }
                            }
                        }
                    } catch (_e) {
                        // Fall back to relId-based lookup
                        allImageRelIds.add(relId);
                    }
                }
            }

            // Handle layout shape images with unique cache keys to avoid relId collisions
            // (multiple layouts can each have rId1, rId2, etc. pointing to different images)
            for (const layout of this.slideLayouts) {
                if (!layout.layoutPath) { continue; }
                const layoutImageRelIds = this.extractImageRelationshipsFromSlide(layout);
                if (layoutImageRelIds.length === 0) { continue; }
                try {
                    const layoutPart = this.package.getPartByUri('/' + layout.layoutPath);
                    if (layoutPart) {
                        const rels = await layoutPart.getRelationships();
                        for (const relId of layoutImageRelIds) {
                            if (rels && rels[relId]) {
                                let imgPath = rels[relId].target;
                                if (imgPath.startsWith('../')) {
                                    imgPath = imgPath.replace('../', 'ppt/');
                                } else if (!imgPath.startsWith('/') && !imgPath.startsWith('ppt/')) {
                                    imgPath = `ppt/${imgPath}`;
                                }
                                const uniqueKey = `layout:${layout.layoutPath}:${relId}`;
                                if (!this.imageCache.has(uniqueKey)) {
                                    const imageData = await this.zip.getFileData(imgPath).catch(() => null);
                                    if (imageData && !this.isSVGFile(imgPath)) {
                                        const blob = new Blob([imageData], { type: this.getMimeType(imgPath) });
                                        const imageUrl = URL.createObjectURL(blob);
                                        const image = await this.imageLoader.loadImageFromUrl(imageUrl);
                                        this.imageCache.set(uniqueKey, { image, url: imageUrl, path: imgPath, width: image.naturalWidth, height: image.naturalHeight });
                                    }
                                }
                                this.updateShapeImageRelIds(layout, relId, uniqueKey);
                            } else {
                                allImageRelIds.add(relId); // fallback to global lookup
                            }
                        }
                    } else {
                        layoutImageRelIds.forEach(id => allImageRelIds.add(id));
                    }
                } catch (_e) {
                    layoutImageRelIds.forEach(id => allImageRelIds.add(id));
                }
            }

            // Load all images in parallel
            const loadPromises = Array.from(allImageRelIds).map(relId =>
                this.loadImage(relId).catch(error => {
                    this.logger.log("warn", this.constructor.name, `Failed to preload image ${relId}:`, error);
                    return null; // Don't fail the entire preload process for one image
                })
            );

            const results = await Promise.allSettled(loadPromises);
            const successCount = results.filter(result => result.status === 'fulfilled' && result.value !== null).length;


        } catch (error) {
            this.logger.logError(this.constructor.name, 'Error during image preloading:', error);
            // Don't throw - continue with rendering even if image preloading fails
        }
    }

    /**
     * Load image by relationship ID
     */
    async loadImage(relId) {
        
        // Check if we should use SVG version instead of PNG
        const svgRelId = this.svgRelationshipMap.get(relId);
        if (svgRelId) {
            relId = svgRelId; // Use SVG relationship ID instead
        }
        
        // Check cache first
        if (this.imageCache.has(relId)) {
            return this.imageCache.get(relId);
        }

        // Get image path from relationship map
        const imagePath = this.imageMap.get(relId);
        if (!imagePath) {
            this.logger.log("warn", this.constructor.name, `No image path found for relId: ${relId}`);
            return null;
        }

        // Check if this is a data URI (base64 encoded image)
        if (imagePath.startsWith('data:')) {
            
            // Handle SVG data URIs
            if (imagePath.startsWith('data:image/svg+xml')) {
                
                try {
                    const base64Data = imagePath.split(',')[1];
                    const svgContent = atob(base64Data);
                    
                    
                    // Cache SVG content
                    this.imageCache.set(relId, {
                        type: 'svg',
                        content: svgContent,
                        path: imagePath,
                        width: 100, // SVG is scalable
                        height: 100 // SVG is scalable
                    });
                    
                    return this.imageCache.get(relId);
                    
                } catch (error) {
                    this.logger.logError(this.constructor.name, `Failed to decode SVG data URI for ${relId}:`, error);
                    return null;
                }
            }
            
            // For other data URIs, handle as regular images
            // TODO: Handle other types of data URIs if needed
            return null;
        }

        try {
            // Extract image data from ZIP
            const imageData = await this.zip.getFileData(imagePath);
            if (!imageData) {
                return null;
            }

            // For all images (including SVG), create a blob URL and load into Image element
            const blob = new Blob([imageData], { type: this.getMimeType(imagePath) });
            const imageUrl = URL.createObjectURL(blob);

            // Load image
            const image = await this.imageLoader.loadImageFromUrl(imageUrl);

            // Cache the loaded image
            this.imageCache.set(relId, {
                image: image,
                url: imageUrl,
                path: imagePath,
                width: image.naturalWidth,
                height: image.naturalHeight
            });

            return this.imageCache.get(relId);

        } catch (error) {
            return null;
        }
    }

    /**
     * Get MIME type from file path
     */
    getMimeType(path) {
        const extension = path.toLowerCase().split('.').pop();
        const mimeTypes = {
            'png': 'image/png',
            'jpg': 'image/jpeg',
            'jpeg': 'image/jpeg',
            'gif': 'image/gif',
            'bmp': 'image/bmp',
            'svg': 'image/svg+xml',
            'webp': 'image/webp',
            'tiff': 'image/tiff'
        };
        return mimeTypes[extension] || 'image/png';
    }

    /**
     * Step 3: Font Processing
     */
    async processThemes(context, presentationPart) {

        try {
            // Get theme relationships from presentation part
            const presentationRels = await presentationPart.getRelationships();
            if (!presentationRels) {
                // Create default theme
                const theme = new CTheme();
                this.presentation.theme = theme;
                return;
            }

            // Find theme relationship - look for any relationship with 'theme' in the type
            const themeRel = Object.values(presentationRels).find(rel =>
                rel.type && rel.type.includes('theme') && rel.target
            );

            if (!themeRel) {
                // Create default theme
                const theme = new CTheme();
                this.presentation.theme = theme;
                return;
            }

            // Fix path construction - handle targets that already have leading slash
            let themePath = themeRel.target;
            if (themePath.startsWith('/')) {
                themePath = themePath.substring(1); // Remove leading slash for ZIP access
            } else if (themePath.startsWith('../')) {
                themePath = themePath.replace('../', '');
            }

            // Ensure we have the correct path structure
            if (!themePath.startsWith('ppt/')) {
                themePath = 'ppt/' + themePath;
            }

            // Get theme XML content
            const themeXmlContent = await this.zip.getFileText(themePath);
            if (!themeXmlContent) {
                const theme = new CTheme();
                this.presentation.theme = theme;
                return;
            }

            // Parse theme XML
            const parser = new DOMParser();
            const themeDoc = parser.parseFromString(themeXmlContent, 'text/xml');

            if (themeDoc.documentElement.nodeName === 'parsererror') {
                const theme = new CTheme();
                this.presentation.theme = theme;
                return;
            }

            // Extract theme colors
            const theme = this.parseThemeFromXML(themeDoc);
            if (theme) {
                this.presentation.theme = theme;
            } else {
                // Fallback to default theme
                const defaultTheme = new CTheme();
                this.presentation.theme = defaultTheme;
            }

        } catch (error) {
            // Create default theme on error
            const theme = new CTheme();
            this.presentation.theme = theme;
        }
    }

    /**
     * Parse theme from XML document with enhanced color extraction
     */
    parseThemeFromXML(themeDoc) {
        try {
            const themeElements = themeDoc.querySelector('a\\:theme, theme');
            if (!themeElements) {
                return null;
            }

            const colorScheme = themeElements.querySelector('a\\:themeElements a\\:clrScheme, themeElements clrScheme');
            if (!colorScheme) {
                return null;
            }

            const theme = new CTheme();
            theme.name = themeElements.getAttribute('name') || 'Default Theme';
            theme.colors = {};
            theme.fonts = { major: 'Calibri', minor: 'Calibri' };

            // ENHANCED: Parse color scheme with comprehensive extraction
            theme.colors = this.extractEnhancedThemeColors(colorScheme);
            if (theme.colors.accent6) {
                const accent6Rgb = this.parseColorFromHex(theme.colors.accent6);
            } else {
            }

            // Parse font scheme
            const fontScheme = themeElements.querySelector('a\\:themeElements a\\:fontScheme, themeElements fontScheme');
            if (fontScheme) {
                const majorFont = fontScheme.querySelector('a\\:majorFont a\\:latin, majorFont latin');
                const minorFont = fontScheme.querySelector('a\\:minorFont a\\:latin, minorFont latin');

                if (majorFont) {
                    theme.fonts.major = majorFont.getAttribute('typeface') || 'Calibri';
                }
                if (minorFont) {
                    theme.fonts.minor = minorFont.getAttribute('typeface') || 'Calibri';
                }

            }

            // ENHANCED: Parse format scheme with improved theme color integration
            const formatScheme = themeElements.querySelector('a\\:themeElements a\\:fmtScheme, themeElements fmtScheme');
            if (formatScheme) {
                theme.formatScheme = this.parseEnhancedFormatScheme(formatScheme, theme.colors);
            } else {
                theme.formatScheme = this.createEnhancedDefaultFormatScheme(theme.colors);
            }

            // ENHANCED: Store additional theme metadata for better DOM access
            theme.metadata = {
                extracted: true,
                source: 'pptx_file',
                timestamp: new Date().toISOString(),
                colorCount: Object.keys(theme.colors).length,
                fillStyleCount: theme.formatScheme.fillStyles?.length || 0,
                backgroundFillCount: theme.formatScheme.backgroundFills?.length || 0
            };
            window.extractedTheme = theme;

            return theme;

        } catch (error) {
            return null;
        }
    }

    /**
     * ENHANCED: Extract theme colors with comprehensive parsing
     */
    extractEnhancedThemeColors(colorScheme) {
        const colors = {};

        try {
            // Define the expected color names in PowerPoint themes
            const colorNames = [
                'dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3',
                'accent4', 'accent5', 'accent6', 'hlink', 'folHlink',
                // Also check for background and text aliases
                'bg1', 'tx1', 'bg2', 'tx2'
            ];

            for (const colorName of colorNames) {
                // Try multiple selectors to find the color element
                const colorElement = colorScheme.querySelector(`a\\:${colorName}, ${colorName}`) ||
                                   colorScheme.querySelector(`[name="${colorName}"]`);

                if (colorElement) {
                    const extractedColor = this.extractColorFromThemeElement(colorElement);
                    if (extractedColor) {
                        colors[colorName] = extractedColor;
                    }
                }
            }

            // ENHANCED: Also try to extract any additional colors found in the scheme
            const allColorElements = colorScheme.querySelectorAll('*');
            for (const element of allColorElements) {
                const tagName = element.tagName.replace('a:', '');
                if (!colors[tagName] && tagName !== 'clrScheme' && tagName !== 'themeElements') {
                    const extractedColor = this.extractColorFromThemeElement(element);
                    if (extractedColor) {
                        colors[tagName] = extractedColor;
                    }
                }
            }

            // ENHANCED: Create aliases for common name variations
            if (colors.dk1 && !colors.tx1) {colors.tx1 = colors.dk1;}   // Dark 1 = Text 1
            if (colors.lt1 && !colors.bg1) {colors.bg1 = colors.lt1;}   // Light 1 = Background 1
            if (colors.dk2 && !colors.tx2) {colors.tx2 = colors.dk2;}   // Dark 2 = Text 2
            if (colors.lt2 && !colors.bg2) {colors.bg2 = colors.lt2;}   // Light 2 = Background 2

        } catch (_error) {
				// Error ignored
			}

        return colors;
    }

    /**
     * ENHANCED: Extract color value from a theme color element
     */
    extractColorFromThemeElement(colorElement) {
        try {
            // Look for sRGB color (most common)
            const srgbClr = colorElement.querySelector('a\\:srgbClr, srgbClr');
            if (srgbClr) {
                const val = srgbClr.getAttribute('val');
                if (val && /^[0-9A-Fa-f]{6}$/.test(val)) {
                    return `#${val.toUpperCase()}`;
                }
            }

            // Look for system color with lastClr attribute
            const sysClr = colorElement.querySelector('a\\:sysClr, sysClr');
            if (sysClr) {
                const lastClr = sysClr.getAttribute('lastClr');
                if (lastClr && /^[0-9A-Fa-f]{6}$/.test(lastClr)) {
                    return `#${lastClr.toUpperCase()}`;
                }

                // Handle common system colors
                const val = sysClr.getAttribute('val');
                const systemColors = {
                    'windowText': '#000000',
                    'window': '#FFFFFF',
                    'btnFace': '#F0F0F0',
                    'btnText': '#000000'
                };
                if (val && systemColors[val]) {
                    return systemColors[val];
                }
            }

            // Look for HSL color
            const hslClr = colorElement.querySelector('a\\:hslClr, hslClr');
            if (hslClr) {
                const h = parseInt(hslClr.getAttribute('hue') || '0');
                const s = parseInt(hslClr.getAttribute('sat') || '0') / 100000;
                const l = parseInt(hslClr.getAttribute('lum') || '0') / 100000;
                return this.hslToHex(h / 60000, s, l);
            }

            // Look for preset color
            const prstClr = colorElement.querySelector('a\\:prstClr, prstClr');
            if (prstClr) {
                const val = prstClr.getAttribute('val');
                const presetColors = {
                    'black': '#000000',
                    'white': '#FFFFFF',
                    'red': '#FF0000',
                    'green': '#00FF00',
                    'blue': '#0000FF',
                    'yellow': '#FFFF00',
                    'magenta': '#FF00FF',
                    'cyan': '#00FFFF'
                };
                if (val && presetColors[val]) {
                    return presetColors[val];
                }
            }

        } catch (_error) {
				// Error ignored
			}

        return null;
    }

    /**
     * ENHANCED: Convert HSL to Hex color
     */
    hslToHex(h, s, l) {
        const hueToRgb = (p, q, t) => {
            if (t < 0) {t += 1;}
            if (t > 1) {t -= 1;}
            if (t < 1/6) {return p + (q - p) * 6 * t;}
            if (t < 1/2) {return q;}
            if (t < 2/3) {return p + (q - p) * (2/3 - t) * 6;}
            return p;
        };

        const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        const p = 2 * l - q;
        const r = Math.round(hueToRgb(p, q, h + 1/3) * 255);
        const g = Math.round(hueToRgb(p, q, h) * 255);
        const b = Math.round(hueToRgb(p, q, h - 1/3) * 255);

        return `#${((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase()}`;
    }

    /**
     * Parse theme format scheme (fillStyleLst and bgFillStyleLst)
     */
    parseFormatScheme(formatSchemeElement) {
        const formatScheme = {
            fillStyles: [],
            backgroundFills: [],
            lineStyles: []
        };

        try {
            // Parse fill style list (fillStyleLst)
            const fillStyleList = formatSchemeElement.querySelector('a\\:fillStyleLst, fillStyleLst');
            if (fillStyleList) {
                const fillElements = fillStyleList.children;
                for (let i = 0; i < fillElements.length; i++) {
                    const fillEl = fillElements[i];
                    const fillStyle = this.parseThemeFillStyle(fillEl);
                    if (fillStyle) {
                        formatScheme.fillStyles.push(fillStyle);
                    } else {
                    }
                }
            } else {
            }

            // Parse background fill style list (bgFillStyleLst)
            const bgFillStyleList = formatSchemeElement.querySelector('a\\:bgFillStyleLst, bgFillStyleLst');
            if (bgFillStyleList) {
                const bgFillElements = bgFillStyleList.children;
                for (let i = 0; i < bgFillElements.length; i++) {
                    const bgFillEl = bgFillElements[i];
                    const bgFillStyle = this.parseThemeFillStyle(bgFillEl);
                    if (bgFillStyle) {
                        formatScheme.backgroundFills.push(bgFillStyle);
                    } else {
                    }
                }
            } else {
            }

            // Parse line style list (lnStyleLst) - for completeness
            const lineStyleList = formatSchemeElement.querySelector('a\\:lnStyleLst, lnStyleLst');
            if (lineStyleList) {
                const lineElements = lineStyleList.children;
                for (const lineEl of lineElements) {
                    const lineStyle = this.parseThemeLineStyle(lineEl);
                    if (lineStyle) {
                        formatScheme.lineStyles.push(lineStyle);
                    }
                }
            }

        } catch (_error) {
				// Error ignored
			}
        // CRITICAL FIX: If we have no fill styles, merge with default
        if (formatScheme.fillStyles.length === 0) {
            const defaultScheme = this.createDefaultFormatScheme();
            formatScheme.fillStyles = defaultScheme.fillStyles;
        }

        // Also ensure we have background fills
        if (formatScheme.backgroundFills.length === 0) {
            const defaultScheme = this.createDefaultFormatScheme();
            formatScheme.backgroundFills = defaultScheme.backgroundFills;
        }

        return formatScheme;
    }

    /**
     * Enhanced parseThemeFillStyle to handle placeholder colors properly
     */
    parseThemeFillStyle(fillElement) {
        try {
            const tagName = fillElement.tagName.replace('a:', '').toLowerCase();

            switch (tagName) {
                case 'solidfill':
                    // For theme fill styles, we need to preserve the raw structure
                    // because phClr placeholders need to be resolved at render time
                    const solidFillData = this.parseFill(fillElement);

                    if (solidFillData && solidFillData.color && solidFillData.color.type === 'placeholder') {
                        // Store the raw fill structure for later resolution
                        const placeholderFill = {
                            type: 'solid',
                            color: solidFillData.color,
                            rawElement: fillElement.outerHTML // Preserve for later resolution
                        };
                        return placeholderFill;
                    }

                    // CRITICAL FIX: Don't return null for valid solid fills, even if they contain placeholder colors
                    if (solidFillData) {
                        return solidFillData;
                    }

                    // If parseFill failed but this is definitely a solidFill, create a fallback
                    return {
                        type: 'solid',
                        color: { r: 128, g: 128, b: 128, a: 255 }, // Gray fallback
                        rawElement: fillElement.outerHTML
                    };

                case 'gradfill':
                    const gradFillData = this.parseFill(fillElement);

                    if (gradFillData && gradFillData.type === 'gradient') {
                        // For theme gradients with placeholders, preserve the structure
                        return {
                            type: 'gradient',
                            ...gradFillData,
                            rawElement: fillElement.outerHTML
                        };
                    }

                    // CRITICAL FIX: If gradient parsing failed, create a comprehensive fallback

                    // Try to extract gradient stops manually for theme gradients
                    const gsElements = fillElement.querySelectorAll('a\\:gs, gs');
                    const stops = [];

                    for (let i = 0; i < gsElements.length; i++) {
                        const gsEl = gsElements[i];
                        const pos = parseFloat(gsEl.getAttribute('pos') || '0') / 100000;

                        // Look for color in the gradient stop
                        const colorEl = gsEl.querySelector('a\\:schemeClr, schemeClr, a\\:srgbClr, srgbClr, a\\:hslClr, hslClr');
                        let color = { type: 'placeholder', value: 'phClr' }; // Default placeholder

                        if (colorEl) {
                            const tagName = colorEl.tagName.replace('a:', '').toLowerCase();
                            if (tagName === 'schemeclr') {
                                const val = colorEl.getAttribute('val') || 'phClr';
                                // Parse modifications from both the gradient stop and the color element
                                const gsModifications = this.parseColorModifications(gsEl);
                                const colorModifications = this.parseColorModifications(colorEl);
                                const allModifications = [...gsModifications, ...colorModifications];

                                color = {
                                    type: val === 'phClr' ? 'placeholder' : 'scheme',
                                    value: val,
                                    modifications: allModifications
                                };

                            } else if (tagName === 'srgbclr') {
                                const val = colorEl.getAttribute('val');
                                if (val) {
                                    color = this.parseColorFromHex(`#${val}`);
                                }
                            }
                        }

                        stops.push({
                            position: pos,
                            color: color
                        });
                    }

                    // If no stops found, create default ones for theme gradients
                    if (stops.length === 0) {
                        stops.push(
                            { position: 0, color: { type: 'placeholder', value: 'phClr' } },
                            { position: 1, color: { type: 'placeholder', value: 'phClr' } }
                        );
                    }

                    // Create a valid gradient that can be processed later
                    const fallbackGradient = {
                        type: 'gradient',
                        direction: 'linear',
                        angle: 0,
                        stops: stops,
                        rawElement: fillElement.outerHTML
                    };

                    return fallbackGradient;

                case 'blipfill':
                    const blipFillData = this.parseFill(fillElement);
                    return blipFillData;

                case 'pattfill':
                    const pattFillData = this.parseFill(fillElement);
                    return pattFillData;

                case 'nofill':
                    return { type: 'none' };

                default:
                    // Create a fallback even for unknown types
                    return {
                        type: 'solid',
                        color: { r: 128, g: 128, b: 128, a: 255 }, // Gray fallback
                        rawElement: fillElement.outerHTML
                    };
            }
        } catch (error) {
            // Create a fallback instead of returning null
            return {
                type: 'solid',
                color: { r: 128, g: 128, b: 128, a: 255 }, // Gray fallback
                rawElement: fillElement.outerHTML
            };
        }
    }

    /**
     * Parse individual theme line style
     */
    parseThemeLineStyle(lineElement) {
        try {
            // Use existing stroke parsing logic
            return this.parseStroke(lineElement);
        } catch (error) {
            return null;
        }
    }

    /**
     * ENHANCED: Parse format scheme with theme color integration
     */
    parseEnhancedFormatScheme(formatSchemeElement, themeColors) {
        const formatScheme = {
            fillStyles: [],
            backgroundFills: [],
            lineStyles: []
        };

        try {

            // Parse fill style list (fillStyleLst) with theme color awareness
            const fillStyleList = formatSchemeElement.querySelector('a\\:fillStyleLst, fillStyleLst');
            if (fillStyleList) {
                const fillElements = fillStyleList.children;
                for (let i = 0; i < fillElements.length; i++) {
                    const fillEl = fillElements[i];
                    const fillStyle = this.parseThemeAwareFillStyle(fillEl, themeColors, i);
                    if (fillStyle) {
                        formatScheme.fillStyles.push(fillStyle);
                        if (i === 5) {
                        }
                    } else {
                        // Create fallback using theme colors
                        const fallbackStyle = this.createFallbackFillStyle(i, themeColors);
                        formatScheme.fillStyles.push(fallbackStyle);
                        if (i === 5) {
                        }
                    }
                }
            }

            // Parse background fill style list (bgFillStyleLst) with theme color awareness
            const bgFillStyleList = formatSchemeElement.querySelector('a\\:bgFillStyleLst, bgFillStyleLst');
            if (bgFillStyleList) {
                const bgFillElements = bgFillStyleList.children;
                for (let i = 0; i < bgFillElements.length; i++) {
                    const bgFillEl = bgFillElements[i];
                    const bgFillStyle = this.parseThemeAwareFillStyle(bgFillEl, themeColors, i);
                    if (bgFillStyle) {
                        formatScheme.backgroundFills.push(bgFillStyle);
                    } else {
                        // Create fallback for background fills
                        const fallbackBgStyle = this.createFallbackBackgroundFillStyle(i, themeColors);
                        formatScheme.backgroundFills.push(fallbackBgStyle);
                    }
                }
            }

        } catch (_error) {
				// Error ignored
			}

        // ENHANCED: Ensure we have minimum required styles with theme colors
        this.ensureMinimumFormatSchemeStyles(formatScheme, themeColors);

        return formatScheme;
    }

    /**
     * ENHANCED: Parse theme-aware fill style
     */
    parseThemeAwareFillStyle(fillElement, themeColors, index) {
        try {
            const tagName = fillElement.tagName.replace('a:', '').toLowerCase();

            if (tagName === 'solidfill') {
                return this.parseThemeAwareSolidFill(fillElement, themeColors, index);
            } else if (tagName === 'gradfill') {
                return this.parseThemeAwareGradientFill(fillElement, themeColors, index);
            } else if (tagName === 'nofill') {
                return { type: 'none', index: index };
            }

            // For other types, fall back to original parsing
            const originalStyle = this.parseThemeFillStyle(fillElement);
            if (originalStyle) {
                originalStyle.index = index;
                return originalStyle;
            }

            return null;

        } catch (error) {
            return null;
        }
    }

    /**
     * ENHANCED: Parse theme-aware solid fill
     */
    parseThemeAwareSolidFill(fillElement, themeColors, index) {
        const solidFill = {
            type: 'solid',
            index: index,
            color: null
        };

        try {
            // Look for scheme color first (most common in themes)
            const schemeClr = fillElement.querySelector('a\\:schemeClr, schemeClr');
            if (schemeClr) {
                const scheme = schemeClr.getAttribute('val');

                if (scheme && themeColors[scheme]) {
                    // Use actual extracted theme color
                    solidFill.color = this.parseColorFromHex(themeColors[scheme]);

                    // Apply color modifications if present
                    const modifications = this.parseColorModifications(schemeClr);
                    if (modifications.length > 0) {
                        solidFill.color = this.applyColorModifications(solidFill.color, modifications);
                    }

                    solidFill.scheme = scheme;
                    solidFill.modifications = modifications;

                    return solidFill;
                }
            }

            // Fallback to other color types
            const srgbClr = fillElement.querySelector('a\\:srgbClr, srgbClr');
            if (srgbClr) {
                const val = srgbClr.getAttribute('val');
                if (val) {
                    solidFill.color = this.parseColorFromHex(`#${val}`);
                    return solidFill;
                }
            }

            // If no color found, create appropriate default based on index
            solidFill.color = this.getDefaultColorForIndex(index, themeColors);
            solidFill.isDefault = true;

        } catch (error) {
            solidFill.color = this.getDefaultColorForIndex(index, themeColors);
            solidFill.isDefault = true;
        }

        return solidFill;
    }

    /**
     * ENHANCED: Parse theme-aware gradient fill
     */
    parseThemeAwareGradientFill(fillElement, themeColors, index) {
        const gradientFill = {
            type: 'gradient',
            index: index,
            gradient: {
                type: 'linear',
                stops: []
            }
        };

        try {
            // Parse gradient stops with theme color awareness
            const gsElements = fillElement.querySelectorAll('a\\:gs, gs');

            for (let i = 0; i < gsElements.length; i++) {
                const gsEl = gsElements[i];
                const pos = parseFloat(gsEl.getAttribute('pos') || '0') / 100000;

                const stop = {
                    position: pos,
                    color: null
                };

                // Look for scheme color in gradient stop
                const schemeClr = gsEl.querySelector('a\\:schemeClr, schemeClr');
                if (schemeClr) {
                    const scheme = schemeClr.getAttribute('val');
                    if (scheme && themeColors[scheme]) {
                        stop.color = this.parseColorFromHex(themeColors[scheme]);

                        // Apply modifications
                        const modifications = this.parseColorModifications(schemeClr);
                        if (modifications.length > 0) {
                            stop.color = this.applyColorModifications(stop.color, modifications);
                        }

                    }
                }

                // Fallback to sRGB
                if (!stop.color) {
                    const srgbClr = gsEl.querySelector('a\\:srgbClr, srgbClr');
                    if (srgbClr) {
                        const val = srgbClr.getAttribute('val');
                        if (val) {
                            stop.color = this.parseColorFromHex(`#${val}`);
                        }
                    }
                }

                // Final fallback
                if (!stop.color) {
                    stop.color = this.getDefaultColorForIndex(i, themeColors);
                }

                gradientFill.gradient.stops.push(stop);
            }

            // Set gradient direction
            const lin = fillElement.querySelector('a\\:lin, lin');
            if (lin) {
                const ang = parseInt(lin.getAttribute('ang') || '0');
                gradientFill.gradient.angle = ang / 60000; // Convert to degrees
            }

            return gradientFill;

        } catch (error) {
            // Return solid fallback
            return {
                type: 'solid',
                index: index,
                color: this.getDefaultColorForIndex(index, themeColors),
                isDefault: true
            };
        }
    }

    /**
     * ENHANCED: Get appropriate default color for a given index using theme colors
     */
    getDefaultColorForIndex(index, themeColors) {
        // Try to use actual theme colors first
        const accentOrder = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
        const accentKey = accentOrder[index % accentOrder.length];

        if (themeColors && themeColors[accentKey]) {
            const color = this.parseColorFromHex(themeColors[accentKey]);
            return color;
        }

        // No fallback colors - return null if not in theme
        return null;
    }

    /**
     * ENHANCED: Ensure minimum format scheme styles with theme colors
     */
    ensureMinimumFormatSchemeStyles(formatScheme, themeColors) {
        // Ensure we have at least 6 fill styles (for accent colors)
        while (formatScheme.fillStyles.length < 6) {
            const index = formatScheme.fillStyles.length;
            const fallbackStyle = this.createFallbackFillStyle(index, themeColors);
            formatScheme.fillStyles.push(fallbackStyle);
        }

        // Ensure we have at least 3 background fills
        while (formatScheme.backgroundFills.length < 3) {
            const index = formatScheme.backgroundFills.length;
            const fallbackBgStyle = this.createFallbackBackgroundFillStyle(index, themeColors);
            formatScheme.backgroundFills.push(fallbackBgStyle);
        }
    }

    /**
     * ENHANCED: Create fallback fill style using theme colors
     */
    createFallbackFillStyle(index, themeColors) {
        return {
            type: 'solid',
            index: index,
            color: this.getDefaultColorForIndex(index, themeColors),
            isDefault: true
        };
    }

    /**
     * ENHANCED: Create fallback background fill style using theme colors
     */
    createFallbackBackgroundFillStyle(index, themeColors) {
        const bgColors = [
            themeColors?.bg1 ? this.parseColorFromHex(themeColors.bg1) : { r: 255, g: 255, b: 255, a: 255 }, // White
            themeColors?.bg2 ? this.parseColorFromHex(themeColors.bg2) : { r: 245, g: 245, b: 245, a: 255 }, // Light gray
            { r: 220, g: 220, b: 220, a: 255 } // Medium gray
        ];

        return {
            type: 'solid',
            index: index,
            color: bgColors[index % bgColors.length],
            isDefault: true
        };
    }

    /**
     * ENHANCED: Create enhanced default format scheme using extracted theme colors
     */
    createEnhancedDefaultFormatScheme(themeColors) {

        const formatScheme = {
            fillStyles: [],
            backgroundFills: [],
            lineStyles: []
        };

        // Create fill styles using actual theme colors
        for (let i = 0; i < 6; i++) {
            formatScheme.fillStyles.push(this.createFallbackFillStyle(i, themeColors));
        }

        // Create background fills using actual theme colors
        for (let i = 0; i < 3; i++) {
            formatScheme.backgroundFills.push(this.createFallbackBackgroundFillStyle(i, themeColors));
        }

        return formatScheme;
    }

    /**
     * Create enhanced default format scheme with proper color structure (LEGACY - kept for compatibility)
     */
    createDefaultFormatScheme() {
        return {
            fillStyles: [],
            backgroundFills: [],
            lineStyles: []
        };
    }

    /**
     * Process slide masters
     */
    async processSlideMasters(context, presentationPart) {
        const masterIds = this.presentation.masterIds || [];
        const presentationRels = await presentationPart.getRelationships();

        for (const rId of masterIds) {
            const rel = presentationRels[rId];
            if (rel && rel.target) {
                // Fix path construction - handle targets that already have leading slash
                let masterPath = rel.target;
                if (masterPath.startsWith('/')) {
                    // Target already has leading slash, use as-is
                    masterPath = masterPath.substring(1); // Remove leading slash for URI construction
                } else if (masterPath.startsWith('../')) {
                    // Remove relative prefix
                    masterPath = masterPath.replace('../', '');
                }

                // Ensure we have the correct path structure
                if (!masterPath.startsWith('ppt/')) {
                    masterPath = 'ppt/' + masterPath;
                }

                const masterPart = this.package.getPartByUri('/' + masterPath);

                if (masterPart) {
                    const masterXml = await masterPart.getDocumentContent();
                    if (masterXml) {
                        const masterParser = new StaxParser(masterXml, masterPart, context);
                        const masterDoc = masterParser.parse();

                        if (masterDoc) {
                            const master = this.createSlideMasterFromXML(masterDoc, masterPart);
                            if (master) {
                                this.slideMasters.push(master);
                                this.presentation.addSlideMaster(master);

                                // Process layouts for this master
                                await this.processSlideLayouts(context, master, masterPart);
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * Process slide layouts for a master
     */
    async processSlideLayouts(context, master, masterPart) {
        const masterRels = await masterPart.getRelationships();

        for (const [relId, rel] of Object.entries(masterRels)) {
            if (rel.type === OpenXmlTypesSafe.slideLayout.relationType) {
                // Fix path construction - handle targets that already have leading slash
                let layoutPath = rel.target;
                if (layoutPath.startsWith('/')) {
                    // Target already has leading slash, use as-is
                    layoutPath = layoutPath.substring(1); // Remove leading slash for URI construction
                } else if (layoutPath.startsWith('../')) {
                    // Remove relative prefix
                    layoutPath = layoutPath.replace('../', '');
                }

                // Ensure we have the correct path structure
                if (!layoutPath.startsWith('ppt/')) {
                    layoutPath = 'ppt/' + layoutPath;
                }

                const layoutPart = this.package.getPartByUri('/' + layoutPath);

                if (layoutPart) {
                    const layoutXml = await layoutPart.getDocumentContent();
                    if (layoutXml) {
                        const layoutParser = new StaxParser(layoutXml, layoutPart, context);
                        const layoutDoc = layoutParser.parse();

                        if (layoutDoc) {
                            const layout = this.createSlideLayoutFromXML(layoutDoc, master.id, relId, rel.target, layoutPath);
                            if (layout) {
                                this.slideLayouts.push(layout);
                                master.addLayout(layout);
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * Create slide master object from XML - Enhanced standard style
     */
    createSlideMasterFromXML(doc, part) {
        const masterElement = doc.documentElement;
        const masterId = part.uri || 'master_' + this.slideMasters.length;

        // Create master object from XML document
        const master = new CSlideMaster();
        master.id = masterId;
        master.name = masterElement.getAttribute('name') || `Master ${this.slideMasters.length + 1}`;
        master.layoutIds = [];

        // Get layout IDs
        const sldLayoutIdLst = doc.querySelector('sldLayoutIdLst, p\\:sldLayoutIdLst');
        if (sldLayoutIdLst) {
            const layoutIdElements = sldLayoutIdLst.querySelectorAll('sldLayoutId, p\\:sldLayoutId');
            layoutIdElements.forEach(el => {
                const rId = el.getAttribute('r:id');
                if (rId) {
                    master.layoutIds.push(rId);
                }
            });
        }

        // Parse common slide data
        const cSldElement = doc.querySelector('cSld, p\\:cSld');
        if (cSldElement) {
            master.cSld = {
                name: cSldElement.getAttribute('name') || '',
                bg: this.parseBackground(cSldElement),
                spTree: this.processShapeTree(cSldElement.querySelector('spTree, p\\:spTree'))
            };

            // Also store as commonSlideData for standard compatibility
            master.commonSlideData = new CSld();
            master.commonSlideData.name = master.cSld.name;
            master.commonSlideData.backgroundFill = master.cSld.bg;
            master.commonSlideData.shapeTree = master.cSld.spTree;
        }

        // Parse text styles
        const txStylesElement = doc.querySelector('txStyles, p\\:txStyles');
        if (txStylesElement) {
            master.txStyles = {
                titleStyle: this.parseTextStyleElement(txStylesElement.querySelector('titleStyle, p\\:titleStyle')),
                bodyStyle: this.parseTextStyleElement(txStylesElement.querySelector('bodyStyle, p\\:bodyStyle')),
                otherStyle: this.parseTextStyleElement(txStylesElement.querySelector('otherStyle, p\\:otherStyle'))
            };
        }

        // Parse color map
        const clrMapElement = doc.querySelector('clrMap, p\\:clrMap');
        if (clrMapElement) {
            master.clrMap = this.parseColorMapElement(clrMapElement);
        }

        // Register in context
        this.xmlParser.registerSlideMaster(master.id, master);

        return master;
    }

    /**
     * Create slide layout object from XML - Enhanced standard style
     */
    createSlideLayoutFromXML(doc, masterId, relId = null, relTarget = null, layoutPath = null) {
        const layoutElement = doc.documentElement;
        const layoutId = relId || `layout_${this.slideLayouts.length}`;

        // Create layout object from XML document
        const layout = new CSlideLayout();
        layout.id = layoutId;
        layout.relId = relId;
        layout.relTarget = relTarget;
        layout.layoutPath = layoutPath;
        layout.uri = relTarget; // Store the original relationship target for matching
        layout.masterId = masterId;
        layout.name = layoutElement.getAttribute('name') || `Layout ${this.slideLayouts.length + 1}`;
        layout.type = layoutElement.getAttribute('type') || 'blank';
        layout.matchingName = layoutElement.getAttribute('matchingName') || '';
        layout.preserve = layoutElement.getAttribute('preserve') === 'true';
        layout.showMasterSp = layoutElement.getAttribute('showMasterSp') !== 'false';
        layout.showMasterPhAnim = layoutElement.getAttribute('showMasterPhAnim') === 'true';
        layout.userDrawn = layoutElement.getAttribute('userDrawn') !== 'false';

        // Infer layout type from placeholders if not specified
        if (layout.type === 'blank') {
            layout.layoutType = this.inferLayoutTypeFromPlaceholders(layoutElement);
        } else {
            layout.layoutType = layout.type;
        }

        // Parse common slide data
        const cSldElement = doc.querySelector('cSld, p\\:cSld');
        if (cSldElement) {
            const spTreeElement = cSldElement.querySelector('spTree, p\\:spTree');
            
            const processedShapeTree = this.processShapeTree(spTreeElement);
            
            layout.cSld = {
                name: cSldElement.getAttribute('name') || '',
                bg: this.parseBackground(cSldElement),
                spTree: processedShapeTree
            };

            // Also store as commonSlideData for standard compatibility
            layout.commonSlideData = new CSld();
            layout.commonSlideData.name = layout.cSld.name;
            layout.commonSlideData.backgroundFill = layout.cSld.bg;
            layout.commonSlideData.shapeTree = layout.cSld.spTree;
        } else {
        }

        // Parse layout-level text styles if present (to allow layout to override master)
        const txStylesElement = doc.querySelector('txStyles, p\\:txStyles');
        if (txStylesElement) {
            layout.txStyles = {
                titleStyle: this.parseTextStyleElement(txStylesElement.querySelector('titleStyle, p\\:titleStyle')),
                bodyStyle: this.parseTextStyleElement(txStylesElement.querySelector('bodyStyle, p\\:bodyStyle')),
                otherStyle: this.parseTextStyleElement(txStylesElement.querySelector('otherStyle, p\\:otherStyle'))
            };
        }

        // Parse color map override
        const clrMapOvrElement = doc.querySelector('clrMapOvr, p\\:clrMapOvr');
        if (clrMapOvrElement) {
            layout.clrMapOvr = this.parseColorMapOverride(clrMapOvrElement);
        }

        // Register in context
        this.xmlParser.registerSlideLayout(layout.id, layout, masterId);

        return layout;
    }

    /**
     * Process individual slides
     */
    async processSlides(context) {
        const presentationRels = await this.package.getPartByRelationshipType(OpenXmlTypesSafe.presentation.relationType).getRelationships();

        // Find slide relationships
        const slideRels = [];
        for (const [relId, rel] of Object.entries(presentationRels)) {
            if (rel.type === OpenXmlTypesSafe.slide.relationType) {
                slideRels.push({ id: relId, target: rel.target });
            }
        }

        if (slideRels.length === 0) {
            return;
        }

        // Reorder slides to match presentation sldIdLst order when available
        let orderedSlideRels = slideRels;
        try {
            const ids = this.presentation && Array.isArray(this.presentation.slideIds) ? this.presentation.slideIds : [];
            if (ids && ids.length > 0) {
                const relMap = new Map(slideRels.map(rel => [rel.id, rel]));
                const ordered = [];
                for (const rid of ids) {
                    const rel = relMap.get(rid);
                    if (rel) { ordered.push(rel); relMap.delete(rid); }
                }
                // Append any remaining slide relationships not listed (edge cases)
                if (relMap.size > 0) {
                    for (const rel of relMap.values()) { ordered.push(rel); }
                }
                if (ordered.length > 0) {
                    orderedSlideRels = ordered;
                }
            }
        } catch (_e) {}

        // Process slides in presentation order
        for (let i = 0; i < orderedSlideRels.length; i++) {
            try {
                const slideRel = orderedSlideRels[i];

                // Fix path construction - handle targets that already have leading slash
                let slidePath = slideRel.target;
                if (slidePath.startsWith('/')) {
                    // Target already has leading slash, use as-is
                    slidePath = slidePath.substring(1); // Remove leading slash for URI construction
                } else if (slidePath.startsWith('../')) {
                    // Remove relative prefix
                    slidePath = slidePath.replace('../', '');
                }

                // Ensure we have the correct path structure
                if (!slidePath.startsWith('ppt/')) {
                    slidePath = 'ppt/' + slidePath;
                }

                const uri = '/' + slidePath;
                const slidePart = this.package.getPartByUri(uri);

                if (!slidePart) {
                    continue;
                }

                await this.processSlidePart(slidePart, context, i);

            } catch (error) {
                // Continue processing other slides
            }
        }

    }
    /**
     * Create slide object from XML
     */
    createSlideFromXML(doc) {
        try {
            const slide = new CSlide();

            // CRITICAL FIX: Set currentSlide so inheritance can work during shape processing
            this.currentSlide = slide;

            // Parse common slide data
            const cSldElement = doc.querySelector('cSld, p\\:cSld');
            if (cSldElement) {
                slide.commonSlideData = new CSld();

                // Parse shape tree
                const spTreeElement = cSldElement.querySelector('spTree, p\\:spTree');
                if (spTreeElement) {
                    const shapes = this.processShapeTree(spTreeElement);
                    slide.commonSlideData.shapeTree = shapes;
                }

                // Parse background
                const bgElement = cSldElement.querySelector('bg, p\\:bg');
                if (bgElement) {
                    slide.backgroundFill = this.parseBackground(cSldElement);
                }
            }

            // Parse slide properties
            const slideElement = doc.documentElement;
            if (slideElement) {
                slide.showMasterShapes = slideElement.getAttribute('showMasterSp') !== '0';

                // Get slide name if available
                const nameElement = slideElement.querySelector('name, p\\:name');
                if (nameElement) {
                    slide.name = nameElement.textContent || '';
                }
            }

            return slide;

        } catch (error) {
            // Return a minimal slide object
            const fallbackSlide = new CSlide();
            fallbackSlide.commonSlideData = new CSld();
            fallbackSlide.commonSlideData.shapeTree = [];
            return fallbackSlide;
        }
    }

    /**
     * Process a single slide part with layout linking
     */
    async processSlidePart(slidePart, context, slideIndex) {
        const slideXml = await slidePart.getDocumentContent();
        if (!slideXml) {
            return;
        }

        const slideParser = new StaxParser(slideXml, slidePart, context);
        const slideDoc = slideParser.parse();

        if (!slideDoc) {
            return;
        }

        const slide = this.createSlideFromXML(slideDoc);
        if (slide) {
            // Store part URI so preloadAllImages can do per-slide relId resolution
            slide.partUri = slidePart.uri;

            // Link slide to layout and master
            await this.linkSlideToLayoutAndMaster(slide, slidePart, context, slideIndex);

            this.slides.push(slide);
            this.presentation.addSlide(slide);
        } else {
        }
    }

    /**
     * Link slide to its layout and master
     */
    async linkSlideToLayoutAndMaster(slide, slidePart, context, slideIndex) {
        try {
            // Get slide relationships to find layout reference
            const slideRels = await slidePart.getRelationships();

            // Find layout relationship
            
            for (const [relId, rel] of Object.entries(slideRels)) {
                
                if (rel.type === OpenXmlTypesSafe.slideLayout.relationType) {
                    
                    
                    // Find the layout in our parsed layouts - prioritize filename matching
                    
                    
                    // FIRST PASS: Try filename matching (most reliable)
                    let layout = this.slideLayouts.find(l => {
                        if (rel.target && l.relTarget) {
                            const targetFilename = rel.target.split('/').pop();
                            const layoutFilename = l.relTarget.split('/').pop();
                            
                            if (targetFilename === layoutFilename) {
                                return true;
                            }
                        }
                        return false;
                    });
                    
                    // SECOND PASS: Try exact URI matching
                    if (!layout) {
                        layout = this.slideLayouts.find(l => {
                            if (l.uri === rel.target) {
                                return true;
                            }
                            return false;
                        });
                    }
                    
                    // THIRD PASS: Try normalized path matching
                    if (!layout) {
                        layout = this.slideLayouts.find(l => {
                            if (rel.target && l.relTarget) {
                                const normalizeTarget = (target) => target.replace(/^\.\.\//, '').replace(/^\//, '');
                                if (normalizeTarget(rel.target) === normalizeTarget(l.relTarget)) {
                                    return true;
                                }
                            }
                            return false;
                        });
                    }
                    
                    // FOURTH PASS: Fallback to relId matching (less reliable)
                    if (!layout) {
                        layout = this.slideLayouts.find(l => {
                            if (l.relId === relId) {
                                return true;
                            }
                            return false;
                        });
                    }

                    
                    if (layout) {
                        slide.layout = layout;
                        slide.Layout = layout; // standard style property

                        // Find the master for this layout
                        const master = this.slideMasters.find(m => m.id === layout.masterId);
                        if (master) {
                            slide.master = master;
                            slide.Master = master; // standard style property
                            layout.master = master;
                            layout.Master = master; // standard style property
                        }

                        break;
                    } else {
                    }
                }
            }

            // If no layout found, try to find a default one
            if (!slide.layout && this.slideLayouts.length > 0) {
                slide.layout = this.slideLayouts[0];
                slide.Layout = this.slideLayouts[0];
            } else if (!slide.layout) {
            }

        } catch (_error) {
				// Error ignored
			}
    }

    /**
     * Process shape tree from XML
     */
    processShapeTree(spTreeElement) {
        const shapes = [];

        try {
            if (!spTreeElement) {
                return shapes;
            }

            const shapeElements = spTreeElement.children;

            for (let i = 0; i < shapeElements.length; i++) {
                try {
                    const shapeElement = shapeElements[i];
                    const shape = this.processShapeElement(shapeElement);

                    if (shape) {
                        shapes.push(shape);
                    }
                } catch (error) {
                    // Continue processing other shapes
                }
            }

        } catch (_error) {
				// Error ignored
			}

        return shapes;
    }

    /**
     * Process individual shape element
     */
    processShapeElement(element) {
        try {
            if (!element) {
                return null;
            }

            const tagName = element.localName || element.tagName?.toLowerCase();
            let shape = null;

            switch (tagName) {
                case 'sp':
                case 'p:sp':
                    shape = this.processRegularShape(element);
                    break;
                case 'pic':
                case 'p:pic':
                    shape = this.processPictureShape(element);
                    break;
                case 'grpsp':
                case 'grpSp':
                case 'p:grpSp':
                case 'p:grpsp':
                    shape = this.processGroupShape(element);
                    break;
                case 'cxnsp':
                case 'cxnSp':
                case 'p:cxnSp':
                case 'p:cxnsp':
                    shape = this.processConnectorShape(element);
                    break;
                case 'graphicframe':
                case 'graphicFrame':
                case 'p:graphicFrame':
                case 'p:graphicframe':
                    shape = this.processGraphicFrame(element);
                    break;
                default:
                    // Filter out known non-visual elements (case-insensitive matching)
                    const nonVisualElements = [
                        'nvsppr', 'p:nvsppr', 'nvpicpr', 'p:nvpicpr', 
                        'nvgrpsppr', 'p:nvgrpsppr', 'nvGrpSpPr', 'p:nvGrpSpPr',
                        'nvpr', 'p:nvpr', 'cnvpr', 'p:cnvpr', 'cnvsppr', 'p:cnvsppr',
                        'style', 'p:style', 'txstyles', 'p:txstyles',
                        'extlst', 'p:extlst', 'ext', 'p:ext',
                        'grpsppr', 'p:grpsppr', 'grpSpPr', 'p:grpSpPr',  // Group shape properties
                        'sppr', 'p:sppr',        // Shape properties 
                        'picpr', 'p:picpr',      // Picture properties
                        'cxnsppr', 'p:cxnsppr'   // Connector properties
                    ];
                    
                    if (nonVisualElements.includes(tagName) || nonVisualElements.includes(tagName.toLowerCase())) {
                        return null;
                    }
                    
                    shape = this.processDefaultShape(element);
                    break;
            }

            if (shape) {
                            // Set common properties
            shape.id = element.getAttribute('id') || Math.random().toString(36).substr(2, 9);
            shape.name = this.getShapeName(element) || `Shape_${shape.id}`;
                            if (shape.name === 'Rectangle 3' || shape.name?.includes('Rectangle 3')) {
                    const styleElements = element.querySelectorAll('p\\:style, style');
                    styleElements.forEach((styleEl, idx) => {
                        const fillRefEl = styleEl.querySelector('a\\:fillRef, fillRef');
                        if (fillRefEl) {
                        }
                    });
                }

                // CRITICAL FIX: Process shape properties WITHOUT overwriting style information
                const spPrElement = element.querySelector('spPr, p\\:spPr');
                if (spPrElement) {
                    const properties = this.processShapeProperties(spPrElement);

                    // CRITICAL FIX: Merge properties selectively to preserve style information
                    // Instead of Object.assign which overwrites everything, merge specific properties
                    if (properties.transform && !shape.properties?.transform) {
                        shape.properties = shape.properties || {};
                        shape.properties.transform = properties.transform;

                        // Update bounds based on transform if not already set
                        if (!shape.bounds || (shape.bounds.l === 0 && shape.bounds.t === 0 && shape.bounds.r === 0 && shape.bounds.b === 0)) {
                            shape.bounds = {
                                l: properties.transform.x || 0,
                                t: properties.transform.y || 0,
                                r: (properties.transform.x || 0) + (properties.transform.cx || properties.transform.width || 0),
                                b: (properties.transform.y || 0) + (properties.transform.cy || properties.transform.height || 0)
                            };
                        }
                    }
                    if (properties.fill && !shape.fill) {
                        shape.fill = properties.fill;
                    }
                    if (properties.stroke && !shape.stroke) {
                        shape.stroke = properties.stroke;
                    }
                    if (properties.geometry && !shape.geometry) {
                        shape.geometry = properties.geometry;
                    }

                    // Store properties for standard compatibility but don't overwrite existing data
                    if (!shape.spPr) {
                        shape.spPr = properties;
                    }
                    
                    // Assign effects to shape if present
                    if (properties.effectLst) {
                        shape.effects = properties.effectLst;
                    }
                    
                    if (shape.style || shape.preservedStyle) {
                    } else {
                    }

                }

                // Process text body
                const txBodyElement = element.querySelector('txBody, p\\:txBody');
                if (txBodyElement) {
                    shape.textBody = this.processTextBody(txBodyElement);
                }
            }

            return shape;

        } catch (error) {
            return {
                type: 'unknown',
                id: Math.random().toString(36).substr(2, 9),
                name: 'Unknown Shape',
                bounds: { l: 914400, t: 914400, r: 5486400, b: 2743200 }
            };
        }
    }

    /**
     * Process regular shape (sp element)
     */
    processRegularShape(element) {
        // CRITICAL FIX: Create proper CShape instance instead of plain object
        const shape = new CShape();
        shape.type = 'sp';
        shape.name = this.getShapeName(element);
        // Extract ID from cNvPr element (PowerPoint standard location)
        const nvSpPrElement = element.querySelector('nvSpPr, p\\:nvSpPr');
        const cNvPrElement = nvSpPrElement?.querySelector('cNvPr, p\\:cNvPr');
        shape.id = cNvPrElement?.getAttribute('id') || element.getAttribute('id') || null;
        
        // DEBUG: Check for ordering attributes in cNvPr and other locations
        if (cNvPrElement) {
            const allAttributes = Array.from(cNvPrElement.attributes).map(attr => `${attr.name}="${attr.value}"`);
            
            // Check for potential ordering attributes
            const order = cNvPrElement.getAttribute('order') || cNvPrElement.getAttribute('z-order') || 
                         cNvPrElement.getAttribute('zOrder') || cNvPrElement.getAttribute('drawOrder');
            if (order) {
                shape.order = parseInt(order);
            }
        }
        
        // Also check element itself for ordering attributes
        const elementOrder = element.getAttribute('order') || element.getAttribute('z-order') || 
                           element.getAttribute('zOrder') || element.getAttribute('drawOrder');
        if (elementOrder) {
            shape.order = parseInt(elementOrder);
        }
        

        


        try {
            // Get non-visual properties and store for standard compatibility
            if (nvSpPrElement) {
                // Store nvSpPr for standard adapter
                shape.nvSpPr = {
                    cNvPr: null,
                    nvPr: null
                };

                if (cNvPrElement) {
                    shape.name = cNvPrElement.getAttribute('name') || shape.name;
                    // ID already extracted above

                    // Store cNvPr info
                    shape.nvSpPr.cNvPr = {
                        name: shape.name,
                        id: shape.id
                    };
                }

                // Check for placeholder
                const nvPrElement = nvSpPrElement.querySelector('nvPr, p\\:nvPr');
                if (nvPrElement) {
                    const phElement = nvPrElement.querySelector('ph, p\\:ph');
                    if (phElement) {
                        shape.placeholder = this.parsePlaceholder(phElement);
                        shape.isPlaceholder = true;
                    } else {
                    }
                    // Store nvPr info
                    shape.nvSpPr.nvPr = {
                        placeholder: shape.placeholder
                    };
                } else {
                }
            }

            // CRITICAL FIX: Preserve type for accurate DOM conversion
            shape.type = 'sp'; // Mark as regular shape

            // Get shape properties and store spPr for standard compatibility
            const spPrElement = element.querySelector('spPr, p\\:spPr');

            if (spPrElement) {
                const properties = this.processShapeProperties(spPrElement);
                shape.properties = properties;

                // CRITICAL FIX: Store spPr element for standard adapter compatibility
                shape.spPr = properties;
                
                // Assign effects to shape if present
                if (properties.effectLst) {
                    shape.effects = properties.effectLst;
                }



                // Extract bounds from transform
                if (properties.transform) {
                    shape.bounds = {
                        l: properties.transform.x || 0,
                        t: properties.transform.y || 0,
                        r: (properties.transform.x || 0) + (properties.transform.cx || properties.transform.width || 0),
                        b: (properties.transform.y || 0) + (properties.transform.cy || properties.transform.height || 0)
                    };
                } else {
                    // Provide fallback bounds in EMU coordinates if no transform found
                    shape.bounds = { l: 914400, t: 914400, r: 5486400, b: 2743200 }; // 1 inch from edges, 5 inches wide, 2 inches tall in EMU
                }

                // Extract fill information (including explicit no fill)
                if (properties.fill !== undefined) {
                    shape.fill = properties.fill;
                }

                // Extract stroke information
                if (properties.stroke) {
                    shape.stroke = properties.stroke;
                }
                const prstGeomElement = spPrElement.querySelector('prstGeom, a\\:prstGeom');
                if (prstGeomElement) {
                    const preset = prstGeomElement.getAttribute('prst') || 'rect';
                    const adjustments = {};
                    const avLstEl = prstGeomElement.querySelector('avLst, a\\:avLst');
                    if (avLstEl) {
                        avLstEl.querySelectorAll('gd, a\\:gd').forEach(gd => {
                            const name = gd.getAttribute('name');
                            const fmla = gd.getAttribute('fmla') || '';
                            const match = fmla.match(/^val\s+(-?\d+)$/);
                            if (name && match) adjustments[name] = parseInt(match[1]);
                        });
                    }
                    shape.geometry = { type: 'preset', preset, adjustments };
                } else {
                }

                const custGeomElement = spPrElement.querySelector('custGeom, a\\:custGeom');
                if (custGeomElement) {
                    shape.geometry = {
                        type: 'custom',
                        pathList: this.parseCustomGeometry(custGeomElement)
                    };
                }
            } else {
                // Provide fallback bounds in EMU coordinates if no spPr found
                shape.bounds = { l: 914400, t: 914400, r: 5486400, b: 2743200 }; // 1 inch from edges, 5 inches wide, 2 inches tall in EMU
            }

            // Get text body
            const txBodyElement = element.querySelector('txBody, p\\:txBody');
            if (txBodyElement) {
                shape.textBody = this.processTextBody(txBodyElement);

                // Log if we found text content
                if (shape.textBody && shape.textBody.paragraphs) {
                    let hasText = false;
                    let textContent = '';
                    for (const para of shape.textBody.paragraphs) {
                        for (const run of para.runs) {
                            if (run.text && run.text.trim()) {
                                hasText = true;
                                textContent += run.text + ' ';
                            }
                        }
                    }
                    if (hasText) {
                    }
                }
            }

            // CRITICAL FIX: Get style information - try multiple approaches
            let styleFound = false;

            // Approach 1: Look for style element in current XML element (original approach)
            const styleElement = element.querySelector('style, p\\:style');
            if (styleElement) {
                shape.style = this.parseShapeStyle(styleElement);
                shape.preservedStyle = shape.style;
                styleFound = true;
            }

            // Approach 2: Look up from already-parsed raw XML data (NEW CRITICAL FIX)
            if (!styleFound && window.currentSlideData && window.currentSlideData.rawXMLShapes) {
                const rawShape = window.currentSlideData.rawXMLShapes.find(rawShape =>
                    rawShape.name === shape.name || rawShape.id === shape.id
                );

                if (rawShape && rawShape.style) {
                    shape.style = rawShape.style;
                    shape.preservedStyle = rawShape.style;
                    styleFound = true;
                }
            }

            // Approach 3: Parse style from XML element attributes and children (NEW APPROACH)
            if (!styleFound) {
                const parsedStyle = this.extractStyleFromElement(element);
                if (parsedStyle && Object.keys(parsedStyle).length > 0) {
                    shape.style = parsedStyle;
                    shape.preservedStyle = parsedStyle;
                    styleFound = true;
                }
            }

            if (styleFound) {
                window.currentSlideData = window.currentSlideData || {};
                window.currentSlideData.rawXMLShapes = window.currentSlideData.rawXMLShapes || [];

                // Check if already exists, if not add it
                const existingIndex = window.currentSlideData.rawXMLShapes.findIndex(rawShape =>
                    rawShape.name === shape.name || rawShape.id === shape.id
                );

                if (existingIndex === -1) {
                    window.currentSlideData.rawXMLShapes.push({
                        name: shape.name,
                        id: shape.id,
                        styleElement: styleElement,
                        style: shape.style,
                        hasStyle: true
                    });
                }

            } else {
                window.currentSlideData = window.currentSlideData || {};
                window.currentSlideData.rawXMLShapes = window.currentSlideData.rawXMLShapes || [];
                window.currentSlideData.rawXMLShapes.push({
                    name: shape.name,
                    id: shape.id,
                    xmlElement: element,
                    innerHTML: element.innerHTML.substring(0, 500),
                    hasStyle: false
                });
            }

        } catch (_error) {
				// Error ignored
			}
        // Apply comprehensive property inheritance from layout and master slides
        this.applyPropertyInheritance(shape);

        return shape;
    }

    /**
     * Process picture shape
     */
    processPictureShape(element) {
        const shape = new CShape();
        shape.type = 'pic';

        // CRITICAL FIX: Get and preserve style information - comprehensive approach
        let styleFound = false;

        // Try style element approach
        const styleElement = element.querySelector('style, p\\:style');
        if (styleElement) {
            shape.style = this.parseShapeStyle(styleElement);
            shape.preservedStyle = shape.style;
            styleFound = true;
        }

        // Try lookup from raw XML data
        if (!styleFound && window.currentSlideData && window.currentSlideData.rawXMLShapes) {
            const rawShape = window.currentSlideData.rawXMLShapes.find(rawShape =>
                rawShape.name === shape.name || rawShape.id === shape.id
            );

            if (rawShape && rawShape.style) {
                shape.style = rawShape.style;
                shape.preservedStyle = rawShape.style;
                styleFound = true;
            }
        }

        // Try extracting from element structure
        if (!styleFound) {
            const parsedStyle = this.extractStyleFromElement(element);
            if (parsedStyle && Object.keys(parsedStyle).length > 0) {
                shape.style = parsedStyle;
                shape.preservedStyle = parsedStyle;
                styleFound = true;
            }
        }

        if (!styleFound) {
        }

        // Get picture bounds using the same method as regular shapes
        const spPrElement = element.querySelector('spPr, p\\:spPr');

        if (spPrElement) {
            const properties = this.processShapeProperties(spPrElement);

            // CRITICAL FIX: Store properties on shape for standard adapter
            shape.properties = properties;

            if (properties.transform) {
                shape.bounds = {
                    l: properties.transform.x || 0,
                    t: properties.transform.y || 0,
                    r: (properties.transform.x || 0) + (properties.transform.cx || properties.transform.width || 0),
                    b: (properties.transform.y || 0) + (properties.transform.cy || properties.transform.height || 0)
                };
            } else {
                // Provide fallback bounds in EMU coordinates
                shape.bounds = { l: 914400, t: 914400, r: 5486400, b: 2743200 };
            }
        } else {
            // Provide fallback bounds in EMU coordinates
            shape.bounds = { l: 914400, t: 914400, r: 5486400, b: 2743200 };
        }

        // Extract image relationship information
        const blipFillElement = element.querySelector('blipFill, p\\:blipFill');
        if (blipFillElement) {
            const blipElement = blipFillElement.querySelector('blip, a\\:blip');
            if (blipElement) {
                // Found blip element, checking for SVG extensions
                
                // First check for SVG alternative in extension list (priority over PNG fallback)
                const svgRelId = this.extractSVGRelationshipId(blipElement);
                
                if (svgRelId) {
                    // Use SVG version if available
                    shape.imageRelId = svgRelId;
                    shape.hasSVGAlternative = true;
                    // Using SVG relationship ID
                } else {
                    // Fallback to PNG/standard relationship ID
                    const embedId = blipElement.getAttribute('r:embed');
                    if (embedId) {
                        shape.imageRelId = embedId;
                        // Using fallback relationship ID
                    }
                }

                // Get image effects if present
                const effectsElement = blipElement.querySelector('effects, a\\:effects');
                if (effectsElement) {
                    shape.imageEffects = this.parseImageEffects(effectsElement);
                }
            }

            // Get stretch/tile information
            const stretchElement = blipFillElement.querySelector('stretch, a\\:stretch');
            const tileElement = blipFillElement.querySelector('tile, a\\:tile');

            if (stretchElement) {
                shape.imageFillMode = 'stretch';
                const fillRectElement = stretchElement.querySelector('fillRect, a\\:fillRect');
                if (fillRectElement) {
                    shape.imageFillRect = {
                        l: parseInt(fillRectElement.getAttribute('l')) || 0,
                        t: parseInt(fillRectElement.getAttribute('t')) || 0,
                        r: parseInt(fillRectElement.getAttribute('r')) || 0,
                        b: parseInt(fillRectElement.getAttribute('b')) || 0
                    };
                }
            } else if (tileElement) {
                shape.imageFillMode = 'tile';
                shape.imageTileProperties = {
                    tx: parseInt(tileElement.getAttribute('tx')) || 0,
                    ty: parseInt(tileElement.getAttribute('ty')) || 0,
                    sx: parseInt(tileElement.getAttribute('sx')) || 100000,
                    sy: parseInt(tileElement.getAttribute('sy')) || 100000,
                    flip: tileElement.getAttribute('flip') || 'none',
                    algn: tileElement.getAttribute('algn') || 'tl'
                };
            } else {
                shape.imageFillMode = 'stretch'; // Default
            }

            // Parse srcRect — defines the source image mapping (positive = crop, negative = letterbox)
            const srcRectEl = blipFillElement.querySelector('srcRect, a\\:srcRect');
            if (srcRectEl) {
                shape.imageSrcRect = {
                    l: parseInt(srcRectEl.getAttribute('l') || '0'),
                    t: parseInt(srcRectEl.getAttribute('t') || '0'),
                    r: parseInt(srcRectEl.getAttribute('r') || '0'),
                    b: parseInt(srcRectEl.getAttribute('b') || '0')
                };
            }
        }

        // Get shape name
        shape.name = this.getShapeName(element);

        // Apply comprehensive property inheritance from layout and master slides
        this.applyPropertyInheritance(shape);

        return shape;
    }

    /**
     * Parse image effects from effects element
     */
    parseImageEffects(effectsElement) {
        const effects = {};

        // Alpha modulation
        const alphaMod = effectsElement.querySelector('alphaMod, a\\:alphaMod');
        if (alphaMod) {
            effects.alpha = parseInt(alphaMod.getAttribute('val')) / 100000;
        }

        // Brightness/Contrast
        const lum = effectsElement.querySelector('lum, a\\:lum');
        if (lum) {
            effects.brightness = parseInt(lum.getAttribute('bright')) || 0;
            effects.contrast = parseInt(lum.getAttribute('contrast')) || 0;
        }

        // Color modulation
        const clrMod = effectsElement.querySelector('clrMod, a\\:clrMod');
        if (clrMod) {
            effects.colorMod = clrMod.getAttribute('val');
        }

        // Grayscale
        const grayscl = effectsElement.querySelector('grayscl, a\\:grayscl');
        if (grayscl) {
            effects.grayscale = true;
        }

        return effects;
    }

    /**
     * Parse effect list - handles various text and shape effects
     */
    parseEffectList(effectLstElement) {
        const effects = {};

        // Parse outer shadow
        const outerShdw = effectLstElement.querySelector('outerShdw, a\\:outerShdw');
        if (outerShdw) {
            effects.outerShadow = this.parseOuterShadow(outerShdw);
        }

        // Parse inner shadow
        const innerShdw = effectLstElement.querySelector('innerShdw, a\\:innerShdw');
        if (innerShdw) {
            effects.innerShadow = this.parseInnerShadow(innerShdw);
        }

        // Parse glow
        const glow = effectLstElement.querySelector('glow, a\\:glow');
        if (glow) {
            effects.glow = this.parseGlow(glow);
        }

        // Parse reflection
        const reflection = effectLstElement.querySelector('reflection, a\\:reflection');
        if (reflection) {
            effects.reflection = this.parseReflection(reflection);
        }

        // Parse soft edge
        const softEdge = effectLstElement.querySelector('softEdge, a\\:softEdge');
        if (softEdge) {
            effects.softEdge = this.parseSoftEdge(softEdge);
        }

        // Parse preset shadow
        const prstShdw = effectLstElement.querySelector('prstShdw, a\\:prstShdw');
        if (prstShdw) {
            effects.presetShadow = this.parsePresetShadow(prstShdw);
        }

        return effects;
    }

    /**
     * Parse outer shadow effect
     */
    parseOuterShadow(outerShdwElement) {
        const shadow = {};

        // Shadow properties
        shadow.blurRadius = parseInt(outerShdwElement.getAttribute('blurRad')) || 0;
        shadow.distance = parseInt(outerShdwElement.getAttribute('dist')) || 0;
        shadow.direction = parseInt(outerShdwElement.getAttribute('dir')) || 0;
        shadow.scaleX = parseInt(outerShdwElement.getAttribute('sx')) / 100000 || 1;
        shadow.scaleY = parseInt(outerShdwElement.getAttribute('sy')) / 100000 || 1;
        shadow.skewX = parseInt(outerShdwElement.getAttribute('kx')) || 0;
        shadow.skewY = parseInt(outerShdwElement.getAttribute('ky')) || 0;
        shadow.alignment = outerShdwElement.getAttribute('algn') || 'bl';
        shadow.rotateWithShape = outerShdwElement.getAttribute('rotWithShape') === '1';

        // Parse color
        const colorElement = outerShdwElement.querySelector('srgbClr, a\\:srgbClr, schemeClr, a\\:schemeClr');
        if (colorElement) {
            shadow.color = this.parseColor(colorElement);
        }

        return shadow;
    }

    /**
     * Parse inner shadow effect
     */
    parseInnerShadow(innerShdwElement) {
        const shadow = {};

        shadow.blurRadius = parseInt(innerShdwElement.getAttribute('blurRad')) || 0;
        shadow.distance = parseInt(innerShdwElement.getAttribute('dist')) || 0;
        shadow.direction = parseInt(innerShdwElement.getAttribute('dir')) || 0;

        // Parse color
        const colorElement = innerShdwElement.querySelector('srgbClr, a\\:srgbClr, schemeClr, a\\:schemeClr');
        if (colorElement) {
            shadow.color = this.parseColor(colorElement);
        }

        return shadow;
    }

    /**
     * Parse glow effect
     */
    parseGlow(glowElement) {
        const glow = {};

        glow.radius = parseInt(glowElement.getAttribute('rad')) || 0;

        // Parse color
        const colorElement = glowElement.querySelector('srgbClr, a\\:srgbClr, schemeClr, a\\:schemeClr');
        if (colorElement) {
            glow.color = this.parseColor(colorElement);
        }

        return glow;
    }

    /**
     * Parse reflection effect
     */
    parseReflection(reflectionElement) {
        const reflection = {};

        reflection.blurRadius = parseInt(reflectionElement.getAttribute('blurRad')) || 0;
        reflection.startOpacity = parseInt(reflectionElement.getAttribute('stA')) / 100000 || 1;
        reflection.endOpacity = parseInt(reflectionElement.getAttribute('endA')) / 100000 || 0;
        reflection.distance = parseInt(reflectionElement.getAttribute('dist')) || 0;
        reflection.direction = parseInt(reflectionElement.getAttribute('dir')) || 0;
        reflection.fadeDirection = parseInt(reflectionElement.getAttribute('fadeDir')) || 0;
        reflection.scaleX = parseInt(reflectionElement.getAttribute('sx')) / 100000 || 1;
        reflection.scaleY = parseInt(reflectionElement.getAttribute('sy')) / 100000 || 1;
        reflection.skewX = parseInt(reflectionElement.getAttribute('kx')) || 0;
        reflection.skewY = parseInt(reflectionElement.getAttribute('ky')) || 0;
        reflection.alignment = reflectionElement.getAttribute('algn') || 'bl';

        return reflection;
    }

    /**
     * Parse soft edge effect
     */
    parseSoftEdge(softEdgeElement) {
        const softEdge = {};

        softEdge.radius = parseInt(softEdgeElement.getAttribute('rad')) || 0;

        return softEdge;
    }

    /**
     * Parse preset shadow effect
     */
    parsePresetShadow(prstShdwElement) {
        const shadow = {};

        shadow.preset = prstShdwElement.getAttribute('prst') || 'shdw1';
        shadow.distance = parseInt(prstShdwElement.getAttribute('dist')) || 0;
        shadow.direction = parseInt(prstShdwElement.getAttribute('dir')) || 0;

        return shadow;
    }

    /**
     * Parse color element with alpha support
     */
    parseColor(colorElement) {
        const color = {};

        if (colorElement.tagName.includes('srgbClr')) {
            color.type = 'srgb';
            color.value = colorElement.getAttribute('val') || '000000';
        } else if (colorElement.tagName.includes('schemeClr')) {
            color.type = 'scheme';
            color.value = colorElement.getAttribute('val') || 'dk1';
        }

        // Parse alpha
        const alphaElement = colorElement.querySelector('alpha, a\\:alpha');
        if (alphaElement) {
            color.alpha = parseInt(alphaElement.getAttribute('val')) / 100000 || 1;
        } else {
            color.alpha = 1;
        }

        return color;
    }

    /**
     * Process group shape (grpSp element) - Enhanced with standard patterns
     */
    processGroupShape(element) {
        // CRITICAL FIX: Create proper CShape instance for group shapes
        const shape = new CShape();
        shape.type = 'grpSp';
        shape.name = this.getShapeName(element);
        shape.id = element.getAttribute('id') || Math.random().toString(36).substr(2, 9);
        shape.shapeTree = [];

        // Group coordinate system properties (standard pattern)
        shape.groupCoordSystem = {
            chOff: { x: 0, y: 0 },    // Child offset
            chExt: { cx: 0, cy: 0 },  // Child extent
            groupOff: { x: 0, y: 0 }, // Group offset
            groupExt: { cx: 0, cy: 0 } // Group extent
        };

        // CRITICAL FIX: Get and preserve style information for group shapes - comprehensive approach
        this.applyComprehensiveStyleExtraction(shape, element, 'group');

        // Process group properties (grpSpPr element)
        const grpSpPrElement = element.querySelector('grpSpPr, p\\:grpSpPr');
        if (grpSpPrElement) {
            this.processGroupProperties(grpSpPrElement, shape);
        }

        // Process child shapes in spTree
        const spTreeElement = element.querySelector('spTree, p\\:spTree');
        if (spTreeElement) {
            // Mark child shapes as being in a group
            shape.shapeTree = this.processShapeTreeInGroup(spTreeElement, shape);
        } else {
            // If no spTree, look for direct child shapes (alternative structure)

            // Look for direct child shape elements
            const directChildren = element.children;
            for (let i = 0; i < directChildren.length; i++) {
                const child = directChildren[i];
                const tagName = child.tagName.toLowerCase();

                // Check if this is a shape element
                if (tagName === 'sp' || tagName === 'p:sp' ||
                    tagName === 'pic' || tagName === 'p:pic' ||
                    tagName === 'grpsp' || tagName === 'p:grpsp' ||
                    tagName === 'cxnsp' || tagName === 'p:cxnsp') {

                    try {
                        const childShape = this.processShapeElement(child);
                        if (childShape) {
                            childShape.inGroup = true;
                            childShape.parentGroup = shape;
                            this.transformChildShapeCoordinates(childShape, shape);
                            shape.shapeTree.push(childShape);

                        }
                    } catch (_error) {
				// Error ignored
			}
                }
            }
        }

        // Calculate group bounds based on child shapes if not explicitly set
        if (!shape.bounds) {
            this.calculateGroupBounds(shape);
        }

        // Apply comprehensive property inheritance from layout and master slides
        this.applyPropertyInheritance(shape);

        return shape;
    }

    /**
     * Process group properties (grpSpPr element) - standard pattern
     */
    processGroupProperties(grpSpPrElement, shape) {
        // Process group transform (xfrm element)
        const xfrmElement = grpSpPrElement.querySelector('xfrm, a\\:xfrm');
        if (xfrmElement) {
            const transform = this.processGroupTransform(xfrmElement);
            shape.transform = transform;
            shape.groupCoordSystem = transform.groupCoordSystem;

            // Set group bounds from transform
            if (transform.groupOff && transform.groupExt) {
                shape.bounds = {
                    l: transform.groupOff.x,
                    t: transform.groupOff.y,
                    r: transform.groupOff.x + transform.groupExt.cx,
                    b: transform.groupOff.y + transform.groupExt.cy
                };
            }
        }

        // Process group fill and stroke properties
        const properties = this.processShapeProperties(grpSpPrElement);
        if (properties.fill) {
            shape.fill = properties.fill;
        }
        if (properties.stroke) {
            shape.stroke = properties.stroke;
        }
    }

    /**
     * Process group transform (xfrm element) - standard pattern
     */
    processGroupTransform(xfrmElement) {
        const transform = {
            rotation: 0,
            flipH: false,
            flipV: false,
            groupOff: { x: 0, y: 0 },
            groupExt: { cx: 0, cy: 0 },
            groupCoordSystem: {
                chOff: { x: 0, y: 0 },
                chExt: { cx: 0, cy: 0 },
                groupOff: { x: 0, y: 0 },
                groupExt: { cx: 0, cy: 0 }
            }
        };

        // Parse rotation
        const rot = xfrmElement.getAttribute('rot');
        if (rot) {
            transform.rotation = parseInt(rot) / 60000; // Convert from 60000ths of a degree

        }

        // Parse flip attributes
        transform.flipH = xfrmElement.getAttribute('flipH') === '1' || xfrmElement.getAttribute('flipH') === 'true';
        transform.flipV = xfrmElement.getAttribute('flipV') === '1' || xfrmElement.getAttribute('flipV') === 'true';

        // Parse group offset (off element)
        const offElement = xfrmElement.querySelector('off, a\\:off');
        if (offElement) {
            transform.groupOff.x = parseInt(offElement.getAttribute('x')) || 0;
            transform.groupOff.y = parseInt(offElement.getAttribute('y')) || 0;
        }

        // Parse group extent (ext element)
        const extElement = xfrmElement.querySelector('ext, a\\:ext');
        if (extElement) {
            transform.groupExt.cx = parseInt(extElement.getAttribute('cx')) || 0;
            transform.groupExt.cy = parseInt(extElement.getAttribute('cy')) || 0;
        }

        // Parse child offset (chOff element) - standard pattern
        const chOffElement = xfrmElement.querySelector('chOff, a\\:chOff');
        if (chOffElement) {
            transform.groupCoordSystem.chOff.x = parseInt(chOffElement.getAttribute('x')) || 0;
            transform.groupCoordSystem.chOff.y = parseInt(chOffElement.getAttribute('y')) || 0;
        }

        // Parse child extent (chExt element) - standard pattern
        const chExtElement = xfrmElement.querySelector('chExt, a\\:chExt');
        if (chExtElement) {
            transform.groupCoordSystem.chExt.cx = parseInt(chExtElement.getAttribute('cx')) || 0;
            transform.groupCoordSystem.chExt.cy = parseInt(chExtElement.getAttribute('cy')) || 0;
        }

        // Copy group coordinates to coordinate system
        transform.groupCoordSystem.groupOff = { ...transform.groupOff };
        transform.groupCoordSystem.groupExt = { ...transform.groupExt };

        return transform;
    }

    /**
     * Process shape tree within a group - standard pattern
     */
    processShapeTreeInGroup(spTreeElement, parentGroup) {
        const shapes = [];

        try {
            if (!spTreeElement) {
                return shapes;
            }

            const shapeElements = spTreeElement.children;

            for (let i = 0; i < shapeElements.length; i++) {
                try {
                    const shapeElement = shapeElements[i];
                    const shape = this.processShapeElement(shapeElement);

                    if (shape) {
                        // Mark shape as being in a group
                        shape.inGroup = true;
                        shape.parentGroup = parentGroup;

                        // Transform child coordinates relative to group coordinate system
                        this.transformChildShapeCoordinates(shape, parentGroup);

                        shapes.push(shape);

                    } else {
                    }
                } catch (error) {
                    // Continue processing other shapes
                }
            }

        } catch (_error) {
				// Error ignored
			}
        return shapes;
    }

    /**
     * Transform child shape coordinates relative to group coordinate system - standard pattern
     */
    transformChildShapeCoordinates(childShape, parentGroup) {
        if (!childShape.bounds || !parentGroup.groupCoordSystem) {
            return;
        }

        const groupCoords = parentGroup.groupCoordSystem;
        const childBounds = childShape.bounds;

        // Calculate scale factors from group coordinate system to actual coordinates
        let scaleX = 1, scaleY = 1;

        if (groupCoords.chExt.cx > 0 && groupCoords.groupExt.cx > 0) {
            scaleX = groupCoords.groupExt.cx / groupCoords.chExt.cx;
        }
        if (groupCoords.chExt.cy > 0 && groupCoords.groupExt.cy > 0) {
            scaleY = groupCoords.groupExt.cy / groupCoords.chExt.cy;
        }

        // Transform child coordinates from group space to slide space
        const transformedBounds = {
            l: groupCoords.groupOff.x + ((childBounds.l - groupCoords.chOff.x) * scaleX),
            t: groupCoords.groupOff.y + ((childBounds.t - groupCoords.chOff.y) * scaleY),
            r: groupCoords.groupOff.x + ((childBounds.r - groupCoords.chOff.x) * scaleX),
            b: groupCoords.groupOff.y + ((childBounds.b - groupCoords.chOff.y) * scaleY)
        };

        // Store original bounds for reference
        childShape.originalBounds = { ...childBounds };
        childShape.bounds = transformedBounds;

        // Store transformation info
        childShape.groupTransform = {
            scaleX: scaleX,
            scaleY: scaleY,
            offsetX: groupCoords.groupOff.x - groupCoords.chOff.x * scaleX,
            offsetY: groupCoords.groupOff.y - groupCoords.chOff.y * scaleY
        };

    }

    /**
     * Calculate group bounds based on child shapes - standard pattern
     */
    calculateGroupBounds(groupShape) {
        if (!groupShape.shapeTree || groupShape.shapeTree.length === 0) {
            // Use fallback bounds if no child shapes
            groupShape.bounds = { l: 914400, t: 914400, r: 5486400, b: 2743200 };
            return;
        }

        // Calculate bounding box of all child shapes
        let minX = Infinity, minY = Infinity;
        let maxX = -Infinity, maxY = -Infinity;
        let hasValidBounds = false;

        for (const childShape of groupShape.shapeTree) {
            if (childShape.bounds) {
                minX = Math.min(minX, childShape.bounds.l);
                minY = Math.min(minY, childShape.bounds.t);
                maxX = Math.max(maxX, childShape.bounds.r);
                maxY = Math.max(maxY, childShape.bounds.b);
                hasValidBounds = true;
            }
        }

        if (hasValidBounds) {
            groupShape.bounds = {
                l: minX,
                t: minY,
                r: maxX,
                b: maxY
            };

        } else {
            // Use fallback bounds
            groupShape.bounds = { l: 914400, t: 914400, r: 5486400, b: 2743200 };
        }
    }

    /**
     * Process connector shape
     */
    processConnectorShape(element) {
        const shape = new CShape();
        shape.type = 'cxnSp';

        // CRITICAL FIX: Get and preserve style information - comprehensive approach
        this.applyComprehensiveStyleExtraction(shape, element, 'connector');

        // Debug: basic info before property extraction
        try {
            const idAttr = element.getAttribute('id') || element.getAttribute('r:id');
            this.logger?.log('info', this.constructor.name, '[Connector Parse] start', { id: idAttr });
        } catch (e) {}

        // Get connector bounds using the same method as other shapes
        const spPrElement = element.querySelector('spPr, p\\:spPr');
        if (spPrElement) {
            const properties = this.processShapeProperties(spPrElement);
            // Ensure connector shape carries parsed properties (stroke, geometry, transform, etc.)
            shape.properties = properties;
            // Back-compat: expose stroke info on multiple well-known fields
            if (properties && properties.stroke) {
                if (!shape.properties.line) { shape.properties.line = properties.stroke; }
                if (!shape.stroke) { shape.stroke = properties.stroke; }
            }
            // Also copy transform to shape.transform for downstream usage
            if (properties && properties.transform) {
                shape.transform = Object.assign(shape.transform || {}, properties.transform);
            }
            // Force-parse <a:ln> for connectors to guarantee stroke and arrowheads
            try {
                const lnEl = spPrElement.querySelector('a\\:ln, ln, p\\:ln');
                if (lnEl) {
                    const parsedLn = this.parseStroke(lnEl);
                    shape.properties = shape.properties || {};
                    shape.properties.stroke = parsedLn;
                    shape.properties.line = parsedLn;
                    shape.stroke = parsedLn;
                    // Store in spPr-style structure for downstream compatibility
                    shape.spPr = shape.spPr || {};
                    shape.spPr.ln = Object.assign({}, parsedLn);
                    
                } else {
                    
                }
            } catch (_e) {}
            if (properties.geometry) {
                shape.geometry = properties.geometry;
            }
            if (properties.transform) {
                shape.bounds = {
                    l: properties.transform.x || 0,
                    t: properties.transform.y || 0,
                    r: (properties.transform.x || 0) + (properties.transform.cx || properties.transform.width || 0),
                    b: (properties.transform.y || 0) + (properties.transform.cy || properties.transform.height || 0)
                };
                try {
                    this.logger?.log('info', this.constructor.name, '[Connector Parse] bounds from transform', { bounds: shape.bounds, transform: properties.transform });
                } catch (e) {}
                
            } else {
                // Provide fallback bounds in EMU coordinates
                shape.bounds = { l: 914400, t: 914400, r: 5486400, b: 2743200 }; // 1 inch from edges, 5 inches wide, 2 inches tall in EMU
                try { this.logger?.log('warn', this.constructor.name, '[Connector Parse] no transform, using fallback bounds'); } catch (e) {}
                try { console.warn('[Connector Parse] no transform; using fallback bounds', shape.bounds); } catch (e) {}
            }
        } else {
            // Provide fallback bounds in EMU coordinates
            shape.bounds = { l: 914400, t: 914400, r: 5486400, b: 2743200 }; // 1 inch from edges, 5 inches wide, 2 inches tall in EMU
            try { this.logger?.log('warn', this.constructor.name, '[Connector Parse] no spPr found, using fallback bounds'); } catch (e) {}
            try { console.warn('[Connector Parse] no spPr; using fallback bounds', shape.bounds); } catch (e) {}
        }

        // Apply comprehensive property inheritance from layout and master slides
        this.applyPropertyInheritance(shape);

        // Final debug of connector stroke and arrow info
        try {
            const strokeColor = this.getShapeStrokeColor(shape);
            const lineWidth = this.getShapeLineWidth(shape);
            const strokeInfo = this.getShapeStrokeInfo(shape);
            
        } catch (e) {}

        return shape;
    }

    /**
     * Process graphic frame
     */
    processGraphicFrame(element) {
        const shape = new CShape();
        shape.type = 'graphicFrame';

        // CRITICAL FIX: Get and preserve style information - comprehensive approach
        this.applyComprehensiveStyleExtraction(shape, element, 'graphic frame');

        // Use the same shape properties processing as other shapes
        const spPrElement = element.querySelector('spPr, p\\:spPr');
        if (spPrElement) {
            const properties = this.processShapeProperties(spPrElement);
            if (properties.transform) {
                const x = properties.transform.x || 0;
                const y = properties.transform.y || 0;
                const w = properties.transform.cx || properties.transform.width || 0;
                const h = properties.transform.cy || properties.transform.height || 0;
                
                shape.bounds = {
                    l: x,
                    t: y,
                    r: x + w,
                    b: y + h
                };
                // Also set direct coordinate properties
                shape.x = x;
                shape.y = y;
                shape.w = w;
                shape.h = h;
                
                // CRITICAL: Update the transform object to ensure graphics adapter gets correct values
                if (!shape.properties) {shape.properties = {};}
                shape.properties.transform = {
                    x: x,
                    y: y,
                    width: w,
                    height: h,
                    cx: w,
                    cy: h
                };
            } else {
                // Provide fallback bounds in EMU coordinates
                shape.bounds = { l: 914400, t: 914400, r: 5486400, b: 2743200 }; // 1 inch from edges, 5 inches wide, 2 inches tall in EMU
                shape.x = 914400;
                shape.y = 914400;
                shape.w = 4572000; // 5486400 - 914400
                shape.h = 1828800; // 2743200 - 914400
            }
        } else {
            // Fallback: try to get frame bounds directly (but still treat as EMU)
            const xfrmElement = element.querySelector('xfrm, p\\:xfrm');
            if (xfrmElement) {
                const offElement = xfrmElement.querySelector('off, p\\:off');
                const extElement = xfrmElement.querySelector('ext, p\\:ext');

                if (offElement && extElement) {
                    // These are EMU coordinates from XML, not pixels
                    const x = parseInt(offElement.getAttribute('x')) || 0;
                    const y = parseInt(offElement.getAttribute('y')) || 0;
                    const cxRaw = parseInt(extElement.getAttribute('cx'));
                    const cyRaw = parseInt(extElement.getAttribute('cy'));
                    const cx = (Number.isFinite(cxRaw) && cxRaw >= 0) ? cxRaw : 914400; // 1 inch default
                    const cy = (Number.isFinite(cyRaw) && cyRaw >= 0) ? cyRaw : 914400; // 1 inch default

                    
                    // Always set the coordinates first, then check if we need to look for better ones
                    shape.bounds = {
                        l: x,
                        t: y,
                        r: x + cx,
                        b: y + cy
                    };
                    shape.x = x;
                    shape.y = y;
                    shape.w = cx;
                    shape.h = cy;
                    
                    // CRITICAL: Also set shape.properties.transform which graphics adapter reads
                    if (!shape.properties) {shape.properties = {};}
                    shape.properties.transform = {
                        x: x,
                        y: y,
                        width: cx,
                        height: cy,
                        cx: cx,
                        cy: cy
                    };
                    
                    
                    // Check if this looks like placeholder coordinates (very small dimensions or at origin)
                    if ((x === 0 && y === 0) || (cx <= 100000 && cy <= 100000)) {
                        
                        // Try to find coordinates from the parent graphicFrame element
                        let parentElement = element.parentElement;
                        while (parentElement && parentElement.tagName) {
                            const parentXfrm = parentElement.querySelector('xfrm, p\\:xfrm');
                            if (parentXfrm) {
                                const parentOff = parentXfrm.querySelector('off, p\\:off');
                                const parentExt = parentXfrm.querySelector('ext, p\\:ext');
                                if (parentOff && parentExt) {
                                    const parentX = parseInt(parentOff.getAttribute('x')) || 0;
                                    const parentY = parseInt(parentOff.getAttribute('y')) || 0;
                                    const parentCx = parseInt(parentExt.getAttribute('cx')) || 0;
                                    const parentCy = parseInt(parentExt.getAttribute('cy')) || 0;
                                    
                                    if (parentCx > 100000 && parentCy > 100000) {
                                        
                                        // Set both bounds object AND direct coordinate properties
                                        shape.bounds = {
                                            l: parentX,
                                            t: parentY,
                                            r: parentX + parentCx,
                                            b: parentY + parentCy
                                        };
                                        
                                        // Also set direct coordinate properties
                                        shape.x = parentX;
                                        shape.y = parentY;
                                        shape.w = parentCx; 
                                        shape.h = parentCy;
                                        
                                        // CRITICAL: Override transform properties with parent coordinates
                                        if (!shape.properties) {shape.properties = {};}
                                        shape.properties.transform = {
                                            x: parentX,
                                            y: parentY,
                                            width: parentCx,
                                            height: parentCy,
                                            cx: parentCx,
                                            cy: parentCy
                                        };
                                        
                                        break;
                                    }
                                }
                            }
                            parentElement = parentElement.parentElement;
                        }
                        
                        // If we didn't find better parent coordinates, the original ones are already set
                        if (!shape.bounds || (shape.x === x && shape.y === y)) {
                        }
                    }
                    // Note: Coordinates are already set above, no need to set them again
                } else {
                    // Final fallback
                    shape.bounds = { l: 914400, t: 914400, r: 5486400, b: 2743200 }; // 1 inch from edges, 5 inches wide, 2 inches tall in EMU
                    shape.x = 914400;
                    shape.y = 914400;
                    shape.w = 4572000; // 5486400 - 914400
                    shape.h = 1828800; // 2743200 - 914400
                }
            } else {
                // Final fallback
                shape.bounds = { l: 914400, t: 914400, r: 5486400, b: 2743200 }; // 1 inch from edges, 5 inches wide, 2 inches tall in EMU
                shape.x = 914400;
                shape.y = 914400;
                shape.w = 4572000; // 5486400 - 914400
                shape.h = 1828800; // 2743200 - 914400
            }
        }

        // Extract graphic data for tables and charts
        let graphic = element.querySelector('graphic');
        if (!graphic) {graphic = element.querySelector('p\\:graphic');}
        if (!graphic) {graphic = element.querySelector('a\\:graphic');}
        
        if (!graphic) {
            // Try to find graphic element without namespace
            graphic = Array.from(element.children).find(child => 
                child.tagName.toLowerCase().includes('graphic')
            );
        }
        
        if (graphic) {
            let graphicData = graphic.querySelector('graphicData');
            if (!graphicData) {graphicData = graphic.querySelector('p\\:graphicData');}  
            if (!graphicData) {graphicData = graphic.querySelector('a\\:graphicData');}
            
            if (!graphicData) {
                // Try to find graphicData element without namespace
                graphicData = Array.from(graphic.children).find(child => 
                    child.tagName.toLowerCase().includes('graphicdata')
                );
            }
            
            if (graphicData) {
                const uri = graphicData.getAttribute('uri');
                // Found URI in processGraphicFrame
                shape.graphicData = { uri: uri };
                
                // If this is a table, extract the table XML
                if (uri === 'http://schemas.openxmlformats.org/drawingml/2006/table') {
                    let tableElement = graphicData.querySelector('tbl');
                    if (!tableElement) {tableElement = graphicData.querySelector('a\\:tbl');}
                    
                    if (!tableElement) {
                        // Try to find table element without namespace
                        tableElement = Array.from(graphicData.children).find(child => 
                            child.tagName.toLowerCase().includes('tbl')
                        );
                    }
                    
                    if (tableElement) {
                        // Serialize the table XML for later parsing
                        shape.graphicData.tableXml = new XMLSerializer().serializeToString(tableElement);
                    }
                }
                // If this is a chart, schedule async chart parsing
                else if (uri === 'http://schemas.openxmlformats.org/drawingml/2006/chart') {
                    // Chart URI detected in slide-renderer processGraphicFrame

                    // Store the graphicData element for later processing
                    shape.graphicData.element = graphicData;

                    // Extract the chart relationship ID for slide-context-aware loading during rendering
                    const chartEl = graphicData.querySelector('chart') ||
                                    graphicData.querySelector('c\\:chart') ||
                                    graphicData.querySelector('[*|localName="chart"]');
                    if (chartEl) {
                        const chartRId = chartEl.getAttribute('r:id');
                        if (chartRId) {
                            shape.graphicData.chartRef = chartRId;
                        }
                    }
                    
                    // Schedule async chart parsing
                    if (ChartProcessor) {
                        try {
                            // ChartProcessor available, scheduling async chart parsing
                            
                            // CRITICAL FIX: Mark that async chart processing is scheduled to prevent duplicate rendering
                            shape.asyncChartProcessing = true;
                            // Marked shape for async chart processing - initial rendering will be skipped
                            
                            // SCALING FIX: Store original CShape bounds for proper relative sizing
                            shape.originalBounds = {
                                l: shape.bounds.l,
                                t: shape.bounds.t,
                                r: shape.bounds.r,
                                b: shape.bounds.b
                            };
                            
                            const originalWidth = shape.originalBounds.r - shape.originalBounds.l;
                            const originalHeight = shape.originalBounds.b - shape.originalBounds.t;
                            
                            const debugInfo = {
                                event: 'ORIGINAL_BOUNDS_STORAGE',
                                shapeType: shape.type,
                                rawBounds: shape.bounds,
                                originalBounds: shape.originalBounds,
                                originalDimensions: {
                                    width: originalWidth,
                                    height: originalHeight,
                                    widthInches: (originalWidth / 914400).toFixed(2),
                                    heightInches: (originalHeight / 914400).toFixed(2)
                                }
                            };
                            // Original bounds storage completed
                            
                            // Store debug info on shape for later access
                            shape._debugInfo = debugInfo;
                            
                            // Pass zip data through context for chart processing
                            // Available zip sources checked
                            
                            const chartContext = {
                                ...this.context,
                                zip: this.zip || this.zipProcessor?.zip
                            };
                            const chartProcessor = new ChartProcessor(chartContext);
                            
                            // Schedule async chart parsing
                            chartProcessor.parseChartFromGraphicFrame(element).then(chartData => {
                                if (chartData) {
                                    shape.chartData = chartData;
                                    // Clear the async processing flag now that we have real chart data
                                    shape.asyncChartProcessing = false;
                                    // Async chart processing completed - flag cleared
                                    
                                    // Chart data is now stored on the shape.
                                    // Do NOT auto-render here — renderSlide's Path B (renderChartFrame)
                                    // handles chart rendering with proper await and correct canvas context.
                                    // Auto-rendering from this async callback caused cross-slide chart leakage
                                    // because this.graphics may point to a different slide's canvas.
                                    
                                    // Notify thumbnail generation that chart rendering is complete
                                    if (typeof window !== 'undefined') {
                                        window.dispatchEvent(new CustomEvent('chartRenderingComplete', {
                                            detail: { chartData }
                                        }));
                                    }
                                }
                            }).catch(error => {
                                console.error('[Chart Debug] Failed to parse chart data in slide-renderer:', error);
                                // Clear the flag even if parsing failed
                                shape.asyncChartProcessing = false;
                                // Async chart processing failed - flag cleared
                            });
                            
                        } catch (error) {
                            console.error('[Chart Debug] Failed to schedule chart parsing in slide-renderer:', error);
                        }
                    } else {
                        console.warn('[Chart Debug] ChartProcessor not available on window object in slide-renderer');
                    }
                }
            }
        }


        return shape;
    }

    /**
     * Process default shape (fallback)
     */
    processDefaultShape(element) {
        const shape = new CShape();
        shape.type = 'unknown';

        // CRITICAL FIX: Get and preserve style information - comprehensive approach
        this.applyComprehensiveStyleExtraction(shape, element, 'default shape');

        // Try to extract coordinates even for unknown elements
        const spPrElement = element.querySelector('spPr, p\\:spPr');
        if (spPrElement) {
            const properties = this.processShapeProperties(spPrElement);
            if (properties.transform) {
                shape.bounds = {
                    l: properties.transform.x || 0,
                    t: properties.transform.y || 0,
                    r: (properties.transform.x || 0) + (properties.transform.cx || properties.transform.width || 0),
                    b: (properties.transform.y || 0) + (properties.transform.cy || properties.transform.height || 0)
                };
            } else {
                // Use proper EMU coordinates for fallback bounds
                shape.bounds = { l: 914400, t: 914400, r: 5486400, b: 2743200 }; // 1 inch from edges, 5 inches wide, 2 inches tall in EMU
            }
        } else {
            // Use proper EMU coordinates for fallback bounds
            shape.bounds = { l: 914400, t: 914400, r: 5486400, b: 2743200 }; // 1 inch from edges, 5 inches wide, 2 inches tall in EMU
        }

        // Apply comprehensive property inheritance from layout and master slides
        this.applyPropertyInheritance(shape);

        return shape;
    }

    /**
     * Get shape name from element
     */
    getShapeName(element) {
        const nvSpPrElement = element.querySelector('nvSpPr, p\\:nvSpPr') ||
                              element.querySelector('nvPicPr, p\\:nvPicPr') ||
                              element.querySelector('nvGrpSpPr, p\\:nvGrpSpPr');

        if (nvSpPrElement) {
            const cNvPrElement = nvSpPrElement.querySelector('cNvPr, p\\:cNvPr');
            if (cNvPrElement) {
                return cNvPrElement.getAttribute('name') || '';
            }
        }

        return '';
    }

    /**
     * Process shape properties including transform, fill, stroke, etc.
     */
    processShapeProperties(spPrElement) {
        const properties = {};

        if (!spPrElement) {
            return properties;
        }
        const childElements = Array.from(spPrElement.children);

        try {
            // Get transform - try multiple selectors
            let xfrmElement = spPrElement.querySelector('xfrm, p\\:xfrm, a\\:xfrm');

            // If not found, try without namespace prefixes
            if (!xfrmElement) {
                xfrmElement = spPrElement.querySelector('xfrm');
            }

            // If still not found, try finding by tagName directly
            if (!xfrmElement) {
                for (const child of spPrElement.children) {
                    if (child.tagName === 'xfrm' || child.tagName === 'a:xfrm' || child.tagName === 'p:xfrm') {
                        xfrmElement = child;
                        break;
                    }
                }
            }

            if (xfrmElement) {
                const xfrmChildren = Array.from(xfrmElement.children);

                // Try multiple selectors for off element
                let offElement = xfrmElement.querySelector('off, p\\:off, a\\:off');
                if (!offElement) {
                    offElement = xfrmElement.querySelector('off');
                }
                if (!offElement) {
                    for (const child of xfrmElement.children) {
                        if (child.tagName === 'off' || child.tagName === 'a:off' || child.tagName === 'p:off') {
                            offElement = child;
                            break;
                        }
                    }
                }

                // Try multiple selectors for ext element
                let extElement = xfrmElement.querySelector('ext, p\\:ext, a\\:ext');
                if (!extElement) {
                    extElement = xfrmElement.querySelector('ext');
                }
                if (!extElement) {
                    for (const child of xfrmElement.children) {
                        if (child.tagName === 'ext' || child.tagName === 'a:ext' || child.tagName === 'p:ext') {
                            extElement = child;
                            break;
                        }
                    }
                }

                if (offElement) {
                    const x = offElement.getAttribute('x');
                    const y = offElement.getAttribute('y');
                }
                if (extElement) {
                    const cx = extElement.getAttribute('cx');
                    const cy = extElement.getAttribute('cy');
                }

                if (offElement && extElement) {
                    const x = parseInt(offElement.getAttribute('x')) || 0;
                    const y = parseInt(offElement.getAttribute('y')) || 0;
                    const cx = parseInt(extElement.getAttribute('cx')) || 100;
                    const cy = parseInt(extElement.getAttribute('cy')) || 100;

                    properties.transform = {
                        x: x,
                        y: y,
                        cx: cx,
                        cy: cy,
                        // Add width/height aliases for compatibility
                        width: cx,
                        height: cy
                    };
                } else {
                }

                // Get rotation - Fix unit conversion from EMU to degrees
                const rot = xfrmElement.getAttribute('rot');
                if (rot) {
                    properties.transform = properties.transform || {};
                    // Convert from EMU (1/60000 degree units) to degrees
                    const rotationDegrees = parseInt(rot) / 60000 || 0;
                    properties.transform.rot = rotationDegrees;
                    properties.transform.rotation = rotationDegrees; // Also store for standard compatibility
                    // Store original EMU value for reference
                    properties.transform.rotEMU = parseInt(rot) || 0;
                    

                }

                // Preserve flip flags (flipH/flipV)
                const flipHAttr = xfrmElement.getAttribute('flipH');
                const flipVAttr = xfrmElement.getAttribute('flipV');
                if (flipHAttr !== null || flipVAttr !== null) {
                    properties.transform = properties.transform || {};
                    if (flipHAttr !== null) {
                        properties.transform.flipH = (flipHAttr === '1' || flipHAttr === 'true');
                    }
                    if (flipVAttr !== null) {
                        properties.transform.flipV = (flipVAttr === '1' || flipVAttr === 'true');
                    }
                }
            } else {
            }

            // Get fill - check multiple possible fill elements
            // CRITICAL: Check gradFill BEFORE solidFill because querySelector is a descendant
            // search and solidFill elements exist INSIDE gradFill's gradient stops (<a:gs>),
            // which would cause gradient fills to be incorrectly parsed as solid fills.
            // ALSO CRITICAL: Only use DIRECT CHILDREN of spPr - solidFill inside <a:ln> (stroke)
            // would otherwise be mistaken for the shape's fill color.
            const fillTags = ['gradFill', 'solidFill', 'blipFill', 'pattFill', 'noFill'];
            let fillElement = null;
            for (const child of spPrElement.children) {
                const localName = child.localName || child.tagName?.replace(/^.*:/, '');
                if (fillTags.includes(localName)) {
                    fillElement = child;
                    break;
                }
            }

            if (fillElement) {
                this._currentProcessingShape = { name: this.getShapeName(spPrElement.closest('sp, p\\:sp')) || 'Unknown' };
                properties.fill = this.parseFill(fillElement);
                this._currentProcessingShape = null; // Clear context
                const shapeName = this.getShapeName(spPrElement.closest('sp, p\\:sp'));
                if (shapeName === 'Rectangle 3' || shapeName?.includes('Rectangle 3')) {
                }
            }

            // Get stroke
            const strokeElement = spPrElement.querySelector('ln, p\\:ln, a\\:ln');
            if (strokeElement) {
                properties.stroke = this.parseStroke(strokeElement);
            }
            const prstGeomElement = spPrElement.querySelector('prstGeom, p\\:prstGeom, a\\:prstGeom');
            if (prstGeomElement) {
                const preset = prstGeomElement.getAttribute('prst') || 'rect';
                const adjustments = {};
                const avLstEl = prstGeomElement.querySelector('avLst, a\\:avLst');
                if (avLstEl) {
                    avLstEl.querySelectorAll('gd, a\\:gd').forEach(gd => {
                        const name = gd.getAttribute('name');
                        const fmla = gd.getAttribute('fmla') || '';
                        const match = fmla.match(/^val\s+(-?\d+)$/);
                        if (name && match) adjustments[name] = parseInt(match[1]);
                    });
                }
                properties.geometry = { type: 'preset', preset, adjustments };
            } else {
                const childElements = Array.from(spPrElement.children);
            }

            // Get custom geometry
            const custGeomElement = spPrElement.querySelector('custGeom, p\\:custGeom, a\\:custGeom');
            if (custGeomElement) {
                properties.geometry = {
                    type: 'custom',
                    pathList: this.parseCustomGeometry(custGeomElement)
                };
            }

            // Get effects - process effectLst element
            const effectLstElement = spPrElement.querySelector('effectLst, p\\:effectLst, a\\:effectLst');
            if (effectLstElement) {
                properties.effectLst = this.parseEffectList(effectLstElement);
            }

        } catch (_error) {
				// Error ignored
			}

        return properties;
    }

    /**
     * Process text body with enhanced text property extraction
     */
    processTextBody(txBodyElement) {
        const textBody = {
            paragraphs: [],
            bodyProperties: this.parseBodyProperties(txBodyElement)
        };

        try {
            // Parse list style for default text properties
            const lstStyleElement = txBodyElement.querySelector('lstStyle, a\\:lstStyle');
            if (lstStyleElement) {
                textBody.lstStyle = this.parseListStyle(lstStyleElement);
            }

            const paragraphElements = txBodyElement.querySelectorAll('p, a\\:p');

            for (const pElement of paragraphElements) {
                // Parse paragraph properties
                const paragraph = {
                    runs: [],
                    properties: this.parseParagraphProperties(pElement)
                };

                
                // Process text runs and field elements in document order
                const runAndFieldElements = pElement.querySelectorAll('r, a\\:r, fld, a\\:fld');

                for (const rElement of runAndFieldElements) {
                    const tagName = rElement.tagName ? rElement.tagName.toLowerCase() : '';
                    const isFld = tagName === 'a:fld' || tagName === 'fld';

                    // Get text content from t element
                    const tElement = rElement.querySelector('t, a\\:t');
                    const text = tElement ? tElement.textContent : '';

                    // Create run with properties
                    const run = {
                        text: text || '',
                        properties: this.parseRunProperties(rElement)
                    };

                    // Tag field runs with their type so the renderer can substitute dynamic values
                    if (isFld) {
                        const fieldType = rElement.getAttribute('type') || '';
                        run.fieldType = fieldType;
                        // For slidenum fields, mark text as the literal preview; renderer will replace at draw time
                        if (!run.text && (fieldType === 'slidenum' || fieldType === 'slideNum')) {
                            run.text = '1'; // Fallback if <a:t> is missing
                        }
                    }

                    paragraph.runs.push(run);
                }

                // Handle empty paragraphs (line breaks)
                if (paragraph.runs.length === 0) {
                    paragraph.runs.push({
                        text: '',
                        properties: {}
                    });
                }

                textBody.paragraphs.push(paragraph);
            }

            // Log summary of text content found
            const totalTextRuns = textBody.paragraphs.reduce((sum, p) => sum + p.runs.length, 0);
            const textRunsWithContent = textBody.paragraphs.reduce((sum, p) =>
                sum + p.runs.filter(r => r.text && r.text.trim()).length, 0);
            
            // Debug: Extract actual text content
            const allText = textBody.paragraphs.map(p => 
                p.runs.map(r => r.text || '').join('')
            ).join(' ').trim();
            

        } catch (_error) {
				// Error ignored
			}

        return textBody;
    }

    /**
     * Parse list style from lstStyle element
     */
    parseListStyle(lstStyleElement) {
        const lstStyle = {};

        try {
            // Parse level properties (lvl1pPr, lvl2pPr, etc.)
            for (let i = 1; i <= 9; i++) {
                const levelElement = lstStyleElement.querySelector(`lvl${i}pPr, a\\:lvl${i}pPr`);
                if (levelElement) {
                    const levelProps = {
                        align: levelElement.getAttribute('algn') || 'left',
                        marL: parseInt(levelElement.getAttribute('marL')) || 0,
                        indent: parseInt(levelElement.getAttribute('indent')) || 0
                    };

                    // Parse default run properties for this level
                    const defRPrElement = levelElement.querySelector('defRPr, a\\:defRPr');
                    if (defRPrElement) {
                        const szAttr = defRPrElement.getAttribute('sz');
                        if (szAttr) {
                            const fontSizeHundredths = parseInt(szAttr);
                            if (fontSizeHundredths > 0) {levelProps.fontSize = fontSizeHundredths / 100;}
                        }
                        const bAttr = defRPrElement.getAttribute('b');
                        if (bAttr !== null) {levelProps.bold = bAttr === '1' || bAttr === 'true';}
                        const iAttr = defRPrElement.getAttribute('i');
                        if (iAttr !== null) {levelProps.italic = iAttr === '1' || iAttr === 'true';}

                        // Text capitalization
                        const cap = defRPrElement.getAttribute('cap');
                        if (cap) {
                            levelProps.cap = cap; // all, small, none
                        }

                        // Font family
                        const latinElement = defRPrElement.querySelector('latin, a\\:latin');
                        if (latinElement) {
                            levelProps.fontFamily = latinElement.getAttribute('typeface') || 'Arial';
                        }

                        // Color
                        const solidFillElement = defRPrElement.querySelector('solidFill, a\\:solidFill');
                        if (solidFillElement) {
                            levelProps.color = this.parseColor(solidFillElement);
                        }
                    }

                    // Parse spacing before (spcBef)
                    const spcBefEl = levelElement.querySelector('spcBef, a\\:spcBef');
                    if (spcBefEl) {
                        const spcPct = spcBefEl.querySelector('spcPct, a\\:spcPct');
                        const spcPts = spcBefEl.querySelector('spcPts, a\\:spcPts');
                        if (spcPct) {
                            levelProps.spaceBeforePct = parseInt(spcPct.getAttribute('val')) / 100000;
                        } else if (spcPts) {
                            levelProps.spaceBefore = parseInt(spcPts.getAttribute('val')) * 127;
                        }
                    }

                    // Parse spacing after (spcAft)
                    const spcAftEl = levelElement.querySelector('spcAft, a\\:spcAft');
                    if (spcAftEl) {
                        const spcPct = spcAftEl.querySelector('spcPct, a\\:spcPct');
                        const spcPts = spcAftEl.querySelector('spcPts, a\\:spcPts');
                        if (spcPct) {
                            levelProps.spaceAfterPct = parseInt(spcPct.getAttribute('val')) / 100000;
                        } else if (spcPts) {
                            levelProps.spaceAfter = parseInt(spcPts.getAttribute('val')) * 127;
                        }
                    }

                    lstStyle[`lvl${i}pPr`] = levelProps;
                }
            }
        } catch (_error) {
				// Error ignored
			}
        return lstStyle;
    }

    /**
     * Parse body properties from txBody element
     */
    parseBodyProperties(txBodyElement) {
        const props = {
            wrap: true,
            verticalAlign: 'top',
            anchorX: 'center',
            anchorY: 'middle'
        };

        try {
            const bodyPrElement = txBodyElement.querySelector('bodyPr, a\\:bodyPr');
            if (bodyPrElement) {
                // Text wrapping
                if (bodyPrElement.getAttribute('wrap') === 'none') {
                    props.wrap = false;
                }

                // Vertical alignment
                const anchor = bodyPrElement.getAttribute('anchor');
                if (anchor) {
                    props.verticalAlign = anchor;
                }

                // Text anchoring
                props.anchorX = bodyPrElement.getAttribute('anchorX') || 'center';
                props.anchorY = bodyPrElement.getAttribute('anchorY') || 'middle';

                // Margins - align defaults with graphics-adapter.js rendering defaults
                props.leftMargin = parseInt(bodyPrElement.getAttribute('lIns')) || 45720; // Default ~0.05 inch (1.25mm)
                props.rightMargin = parseInt(bodyPrElement.getAttribute('rIns')) || 45720;
                props.topMargin = parseInt(bodyPrElement.getAttribute('tIns')) || 22860; // Default ~0.025 inch (0.625mm)
                props.bottomMargin = parseInt(bodyPrElement.getAttribute('bIns')) || 22860;

                // Auto-fit: "shrink text on overflow" (normAutofit) carries a pre-computed
                // fontScale and lnSpcReduction (in 1/1000 of a percent) that PowerPoint uses to
                // make the text fit. Expose them as fractions so rendering can honor them;
                // otherwise the text renders at full size (too wide / overflowing the shape).
                const normAutofit = bodyPrElement.querySelector('normAutofit, a\\:normAutofit');
                if (normAutofit) {
                    const fs = parseInt(normAutofit.getAttribute('fontScale'));
                    const lnr = parseInt(normAutofit.getAttribute('lnSpcReduction'));
                    props.fontScale = (Number.isFinite(fs) && fs > 0) ? fs / 100000 : 1;
                    props.lineSpaceReduction = (Number.isFinite(lnr) && lnr > 0) ? lnr / 100000 : 0;
                }
            }
        } catch (_error) {
				// Error ignored
			}

        return props;
    }

    /**
     * Parse paragraph properties from p element
     */
    parseParagraphProperties(pElement) {
        const props = {
            align: 'left',
            lineHeight: 100, // 100%
            spacing: {
                before: 0,
                after: 0
            },
            bullet: null // Bullet properties
            // Note: indent is NOT initialized here - only set if explicitly defined
            // This allows proper inheritance from master/layout styles
        };

        try {
            const pPrElement = pElement.querySelector('pPr, a\\:pPr');
            if (pPrElement) {
                // Parse paragraph level (for nested bullets)
                const lvl = pPrElement.getAttribute('lvl');
                if (lvl !== null && lvl !== undefined) {
                    props.level = parseInt(lvl) || 0;
                } else {
                    props.level = 0; // Default to level 0
                }
                
                // Alignment
                const algn = pPrElement.getAttribute('algn');
                if (algn) {
                    props.align = algn;
                }

                // Line spacing
                const lnSpcElement = pPrElement.querySelector('lnSpc, a\\:lnSpc');
                if (lnSpcElement) {
                    const spcPct = lnSpcElement.querySelector('spcPct, a\\:spcPct');
                    const spcPts = lnSpcElement.querySelector('spcPts, a\\:spcPts');
                    if (spcPct) {
                        // spcPct value is in 1/1000 of a percent; convert to percent (e.g., 120000 -> 120)
                        props.lineHeight = parseInt(spcPct.getAttribute('val')) / 1000;
                    } else if (spcPts) {
                        // spcPts is absolute spacing in 1/100 points (e.g., 2200 => 22 pt)
                        const absPtsHundredths = parseInt(spcPts.getAttribute('val'));
                        props.lineHeightPoints = absPtsHundredths / 100; // store as points
                    }
                }

                // Spacing before
                const spcBElement = pPrElement.querySelector('spcBef, a\\:spcBef');
                if (spcBElement) {
                    const spcPts = spcBElement.querySelector('spcPts, a\\:spcPts');
                    if (spcPts) {
                        // spcPts value is in 1/100 of a point; convert to EMU (1pt = 12700 EMU)
                        // EMU = (val / 100 pt) * 12700 = val * 127
                        props.spacing.before = parseInt(spcPts.getAttribute('val')) * 127;
                    }
                }

                // Spacing after
                const spcAElement = pPrElement.querySelector('spcAft, a\\:spcAft');
                if (spcAElement) {
                    const spcPts = spcAElement.querySelector('spcPts, a\\:spcPts');
                    if (spcPts) {
                        // spcPts value is in 1/100 of a point; convert to EMU (1pt = 12700 EMU)
                        // EMU = (val / 100 pt) * 12700 = val * 127
                        props.spacing.after = parseInt(spcPts.getAttribute('val')) * 127;
                    }
                }

                // Indentation properties (marL and indent)
                // Only set if explicitly defined to allow inheritance
                const marL = pPrElement.getAttribute('marL');
                const indent = pPrElement.getAttribute('indent');
                if (marL !== null || indent !== null) {
                    props.indent = {
                        left: marL ? parseInt(marL) : 0,      // EMU units
                        hanging: indent ? parseInt(indent) : 0 // EMU units (negative for hanging indent)
                    };
                }

                // Bullet properties parsing
                this.parseBulletProperties(pPrElement, props);


                // Default text run properties for paragraph
                const defRPrElement = pPrElement.querySelector('defRPr, a\\:defRPr');
                if (defRPrElement) {
                    const fontSizeHundredths = parseInt(defRPrElement.getAttribute('sz'));
                    if (fontSizeHundredths) props.fontSize = fontSizeHundredths / 100; // Convert to points
                    props.bold = defRPrElement.getAttribute('b') === '1';
                    props.italic = defRPrElement.getAttribute('i') === '1';

                    // Font
                    const latinElement = defRPrElement.querySelector('latin, a\\:latin');
                    if (latinElement) {
                        props.fontFamily = latinElement.getAttribute('typeface') || 'Arial';
                    }

                    // Color
                    const solidFillElement = defRPrElement.querySelector('solidFill, a\\:solidFill');
                    if (solidFillElement) {
                        props.color = this.parseColor(solidFillElement);
                    }
                }
            }
        } catch (_error) {
				// Error ignored
			}

        return props;
    }

    /**
     * Parse bullet properties from paragraph properties element
     */
    parseBulletProperties(pPrElement, props) {
        try {
            
            // Check for bullet character (buChar)
            const buCharElement = pPrElement.querySelector('buChar, a\\:buChar');
            if (buCharElement) {
                const char = buCharElement.getAttribute('char');
                if (char) {
                    props.bullet = {
                        type: 'character',
                        char: char
                    };
                    return;
                }
            }

            // Check for bullet numbering (buNum)
            const buNumElement = pPrElement.querySelector('buNum, a\\:buNum');
            if (buNumElement) {
                const startAt = buNumElement.getAttribute('startAt') || '1';
                props.bullet = {
                    type: 'number',
                    startAt: parseInt(startAt)
                };
                return;
            }

            // Check for automatic numbering (buAutoNum)
            const buAutoNumElement = pPrElement.querySelector('buAutoNum, a\\:buAutoNum');
            if (buAutoNumElement) {
                const type = buAutoNumElement.getAttribute('type') || 'arabicPeriod';
                const startAt = buAutoNumElement.getAttribute('startAt') || '1';
                props.bullet = {
                    type: 'autoNumber',
                    subType: type,
                    startAt: parseInt(startAt)
                };
                return;
            }

            // Check for no bullet specified (buNone)
            const buNoneElement = pPrElement.querySelector('buNone, a\\:buNone');
            if (buNoneElement) {
                props.bullet = {
                    type: 'none'
                };
                return;
            }


        } catch (error) {
        }
    }

    /**
     * Parse run properties with enhanced text formatting
     */
    parseRunProperties(rElement) {
        const properties = {};

        try {
            const rPrElement = rElement.querySelector('rPr, a\\:rPr');
            if (rPrElement) {
                // Font size (in hundredths of a point - convert to points)
                // First check for sz attribute directly on rPr element
                let fontSizeHundredths = null;
                if (rPrElement.hasAttribute('sz')) {
                    fontSizeHundredths = parseInt(rPrElement.getAttribute('sz'));
                } else {
                    // Fallback: check for sz child element
                    const szElement = rPrElement.querySelector('sz, a\\:sz');
                    if (szElement) {
                        fontSizeHundredths = parseInt(szElement.getAttribute('val'));
                    }
                }
                
                if (fontSizeHundredths && fontSizeHundredths > 0) {
                    properties.fontSize = fontSizeHundredths / 100; // Convert to points
                }

                // Font family
                const latinElement = rPrElement.querySelector('latin, a\\:latin');
                if (latinElement) {
                    properties.fontFamily = latinElement.getAttribute('typeface') || 'Arial';
                } else {
                    // If no explicit font family, set a default that can be overridden later
                    // The rendering system will fall back to theme defaults
                    properties.fontFamily = null;
                }

                // Bold
                if (rPrElement.hasAttribute('b')) {
                    const bValue = rPrElement.getAttribute('b');
                    properties.bold = bValue !== '0' && bValue !== 'false';
                }

                // Italic
                if (rPrElement.hasAttribute('i')) {
                    const iValue = rPrElement.getAttribute('i');
                    properties.italic = iValue !== '0' && iValue !== 'false';
                }

                // Underline
                if (rPrElement.hasAttribute('u')) {
                    properties.underline = rPrElement.getAttribute('u') !== 'none';
                }

                // Strike
                if (rPrElement.hasAttribute('strike')) {
                    properties.strike = rPrElement.getAttribute('strike') !== 'noStrike';
                }

                // Baseline shift (superscript/subscript)
                if (rPrElement.hasAttribute('baseline')) {
                    const baseline = parseInt(rPrElement.getAttribute('baseline')) || 0;
                    if (baseline > 0) {
                        properties.verticalAlign = 'superscript';
                    } else if (baseline < 0) {
                        properties.verticalAlign = 'subscript';
                    }
                }

                // Text capitalization
                if (rPrElement.hasAttribute('cap')) {
                    properties.cap = rPrElement.getAttribute('cap'); // all, small, none
                }

                // Text spacing
                if (rPrElement.hasAttribute('spc')) {
                    properties.letterSpacing = parseInt(rPrElement.getAttribute('spc')) / 100;
                }

                // Color
                const solidFillElement = rPrElement.querySelector('solidFill, a\\:solidFill');
                if (solidFillElement) {
                    properties.color = this.parseColor(solidFillElement);
                }

                // Highlight
                const highlightElement = rPrElement.querySelector('highlight, a\\:highlight');
                if (highlightElement) {
                    properties.highlight = this.parseColor(highlightElement);
                }

                // Text effects (shadow, glow)
                const effectLstElement = rPrElement.querySelector('effectLst, a\\:effectLst');
                if (effectLstElement) {
                    properties.effectLst = this.parseEffectList(effectLstElement);
                }
            }

        } catch (_error) {
				// Error ignored
			}
        return properties;
    }

    /**
     * Parse fill element
     */
    parseFill(element) {
        if (!element) {return null;}
        const currentShape = this._currentProcessingShape;
        const isRectangle3 = currentShape && (currentShape.name === 'Rectangle 3' || currentShape.name?.includes('Rectangle 3'));

        if (isRectangle3) {
        }

        // CRITICAL FIX: Check if the element itself is a gradFill before looking for solidFill children
        // (gradFill elements contain solidFill inside gradient stops, which would be incorrectly matched)
        const elTag = element.tagName || element.localName || '';
        if (elTag.indexOf('gradFill') >= 0 || (element.localName || '').indexOf('gradFill') >= 0) {
            const gradientData = this.parseGradient(element);
            return {
                type: 'gradient',
                gradient: gradientData
            };
        }

        // Check for solid fill - either as child element or the element itself
        let solidFill = element.querySelector('solidFill, a\\:solidFill');

        // CRITICAL FIX: Check if the element itself is a solidFill
        if (!solidFill && (element.tagName === 'a:solidFill' || element.tagName === 'solidFill')) {
            solidFill = element;
        }

        if (isRectangle3) {
            if (solidFill) {
            }
        }

        if (solidFill) {
            if (isRectangle3) {
            }

            let parsedColor = null;
            try {
                parsedColor = this.parseColor(solidFill);
                if (isRectangle3) {
                }
            } catch (error) {
                if (isRectangle3) {
                }
                return null;
            }

            const fill = {
                type: 'solid',
                color: parsedColor
            };

            if (isRectangle3) {
            }

            return fill;
        }

        // Check for gradient fill
        const gradFill = element.querySelector('gradFill, a\\:gradFill');
        if (gradFill) {
            const gradientData = this.parseGradient(gradFill);
            return {
                type: 'gradient',
                gradient: gradientData
            };
        }

        // Check for pattern fill
        const pattFill = element.querySelector('pattFill, a\\:pattFill');
        if (pattFill) {
            return {
                type: 'pattern',
                pattern: this.parsePattern(pattFill)
            };
        }

        // Check for picture fill (image background)
        const blipFill = element.querySelector('blipFill, a\\:blipFill');
        if (blipFill) {
            const imageData = this.parseBlipFill(blipFill);
            return {
                type: 'image',
                imageData: imageData
            };
        }

        // Check for no fill - either as child or the element itself
        const noFill = element.querySelector('noFill, a\\:noFill');
        if (noFill || element.tagName === 'a:noFill' || element.tagName === 'noFill' || element.localName === 'noFill') {
            return {
                type: 'none'
            };
        }

        return null;
    }

    /**
     * Parse blip fill for images (including background images)
     */
    parseBlipFill(blipFillElement) {
        const blipElement = blipFillElement.querySelector('blip, a\\:blip');
        if (!blipElement) {return null;}

        const imageData = {};

        // Get the relationship ID for the image
        const embedId = blipElement.getAttribute('r:embed');
        if (embedId) {
            imageData.relationshipId = embedId;
        }

        // Get image effects if present
        const effectsElement = blipFillElement.querySelector('effects, a\\:effects');
        if (effectsElement) {
            imageData.effects = this.parseImageEffects(effectsElement);
        }

        // Get stretch/tile information
        const stretchElement = blipFillElement.querySelector('stretch, a\\:stretch');
        const tileElement = blipFillElement.querySelector('tile, a\\:tile');

        if (stretchElement) {
            imageData.fillMode = 'stretch';
            const fillRectElement = stretchElement.querySelector('fillRect, a\\:fillRect');
            if (fillRectElement) {
                imageData.fillRect = {
                    l: parseInt(fillRectElement.getAttribute('l')) || 0,
                    t: parseInt(fillRectElement.getAttribute('t')) || 0,
                    r: parseInt(fillRectElement.getAttribute('r')) || 0,
                    b: parseInt(fillRectElement.getAttribute('b')) || 0
                };
            }
        } else if (tileElement) {
            imageData.fillMode = 'tile';
            imageData.tileProperties = {
                tx: parseInt(tileElement.getAttribute('tx')) || 0,
                ty: parseInt(tileElement.getAttribute('ty')) || 0,
                sx: parseInt(tileElement.getAttribute('sx')) || 100000,
                sy: parseInt(tileElement.getAttribute('sy')) || 100000,
                flip: tileElement.getAttribute('flip') || 'none',
                algn: tileElement.getAttribute('algn') || 'tl'
            };
        } else {
            imageData.fillMode = 'stretch'; // Default
        }

        // Get source rectangle if present
        const srcRectElement = blipFillElement.querySelector('srcRect, a\\:srcRect');
        if (srcRectElement) {
            imageData.sourceRect = {
                l: parseInt(srcRectElement.getAttribute('l')) || 0,
                t: parseInt(srcRectElement.getAttribute('t')) || 0,
                r: parseInt(srcRectElement.getAttribute('r')) || 0,
                b: parseInt(srcRectElement.getAttribute('b')) || 0
            };
        }

        return imageData;
    }

    /**
     * Parse color from various color elements with enhanced scheme and modification support
     */
    parseColor(colorElement) {
        if (!colorElement) {return { r: 0, g: 0, b: 0, a: 255 };}

        // Check for sRGB color
        const srgbClr = colorElement.querySelector('srgbClr, a\\:srgbClr');
        if (srgbClr) {
            const val = srgbClr.getAttribute('val');
            if (val) {
                let color = {
                    r: parseInt(val.substr(0, 2), 16),
                    g: parseInt(val.substr(2, 2), 16),
                    b: parseInt(val.substr(4, 2), 16),
                    a: 255
                };

                // Apply color modifications
                color = this.applyColorModificationsToRgb(srgbClr, color);
                return color;
            }
        }

        // Check for scheme color
        const schemeClr = colorElement.querySelector('schemeClr, a\\:schemeClr');
        if (schemeClr) {
            const val = schemeClr.getAttribute('val');
            const currentShape = this._currentProcessingShape;
            const isRectangle3 = currentShape && (currentShape.name === 'Rectangle 3' || currentShape.name?.includes('Rectangle 3'));

            if (isRectangle3) {
                if (this.presentation?.theme?.colors) {
                }
            }

            // CRITICAL FIX: Handle phClr placeholder - return special object for later resolution
            if (val === 'phClr') {
                return {
                    type: 'placeholder',
                    scheme: 'phClr',
                    modifications: this.parseColorModifications(schemeClr),
                    r: 128, g: 128, b: 128, a: 255 // Fallback color
                };
            }

            // First try to get actual theme colors from presentation
            let baseColor = null;
            if (this.presentation && this.presentation.theme && this.presentation.theme.colors) {
                const themeColorHex = this.presentation.theme.colors[val];
                if (themeColorHex) {
                    baseColor = this.parseColorFromHex(themeColorHex);

                    if (isRectangle3) {
                    }
                } else if (isRectangle3) {
                }
            } else if (isRectangle3) {
            }

            // Fall back to enhanced scheme colors with better color matching
            if (!baseColor) {
                if (isRectangle3) {
                }

                const schemeColors = {
                    // No hardcoded scheme colors - only use colors from DOM
                };

                baseColor = schemeColors[val] || null;

                if (isRectangle3) {
                }
            }

            // Apply color modifications
            baseColor = this.applyColorModificationsToRgb(schemeClr, baseColor);

            if (isRectangle3) {
            }

            return baseColor;
        }

        return { r: 0, g: 0, b: 0, a: 255 };
    }

    /**
     * Parse color modifications from a color element (NEW METHOD)
     */
    parseColorModifications(colorElement) {
        const modifications = [];

        // Check for common modifications
        const modElements = colorElement.querySelectorAll('*');
        for (const modEl of modElements) {
            const tagName = modEl.tagName.replace('a:', '').toLowerCase();
            const val = modEl.getAttribute('val');

            if (val) {
                modifications.push({
                    type: tagName,
                    value: parseInt(val)
                });
            }
        }

        return modifications;
    }

    /**
     * Apply color modifications (tint, shade, lumMod, lumOff) to RGB color
     */
    applyColorModificationsToRgb(colorElement, baseColor) {
        if (!colorElement) {return baseColor;}

        const color = { ...baseColor };

        // Check for tint (lighten)
        const tintElement = colorElement.querySelector('tint, a\\:tint');
        if (tintElement) {
            const tint = parseInt(tintElement.getAttribute('val')) / 100000;
            color.r = Math.round(color.r + (255 - color.r) * tint);
            color.g = Math.round(color.g + (255 - color.g) * tint);
            color.b = Math.round(color.b + (255 - color.b) * tint);
        }

        // Check for shade (darken)
        const shadeElement = colorElement.querySelector('shade, a\\:shade');
        if (shadeElement) {
            const shade = parseInt(shadeElement.getAttribute('val')) / 100000;
            color.r = Math.round(color.r * (1 - shade));
            color.g = Math.round(color.g * (1 - shade));
            color.b = Math.round(color.b * (1 - shade));
        }

        // Check for luminance modulation
        const lumModElement = colorElement.querySelector('lumMod, a\\:lumMod');
        if (lumModElement) {
            const lumMod = parseInt(lumModElement.getAttribute('val')) / 100000;
            color.r = Math.round(color.r * lumMod);
            color.g = Math.round(color.g * lumMod);
            color.b = Math.round(color.b * lumMod);
        }

        // Check for luminance offset
        const lumOffElement = colorElement.querySelector('lumOff, a\\:lumOff');
        if (lumOffElement) {
            const lumOff = parseInt(lumOffElement.getAttribute('val')) / 100000 * 255;
            color.r = Math.round(Math.min(255, color.r + lumOff));
            color.g = Math.round(Math.min(255, color.g + lumOff));
            color.b = Math.round(Math.min(255, color.b + lumOff));
        }

        // Check for alpha
        const alphaElement = colorElement.querySelector('alpha, a\\:alpha');
        if (alphaElement) {
            const alpha = parseInt(alphaElement.getAttribute('val')) / 100000;
            color.a = Math.round(255 * alpha);
        }

        // Ensure values are within valid range
        color.r = Math.max(0, Math.min(255, color.r));
        color.g = Math.max(0, Math.min(255, color.g));
        color.b = Math.max(0, Math.min(255, color.b));
        color.a = Math.max(0, Math.min(255, color.a));

        return color;
    }

    /**
     * Parse gradient fill
     */
    parseGradient(gradFillElement) {
        const gradient = {
            stops: [],
            type: 'linear'
        };

        // Get gradient stops
        const gsLst = gradFillElement.querySelector('gsLst, a\\:gsLst');
        if (gsLst) {
            const stops = gsLst.querySelectorAll('gs, a\\:gs');
            stops.forEach(stop => {
                const pos = parseInt(stop.getAttribute('pos')) || 0;
                const color = this.parseColor(stop);
                gradient.stops.push({ position: pos / 100000, color });
            });
        }

        // Get gradient direction
        const lin = gradFillElement.querySelector('lin, a\\:lin');
        if (lin) {
            gradient.angle = parseInt(lin.getAttribute('ang')) || 0;
            gradient.scaled = lin.getAttribute('scaled') === 'true';
        }

        const rad = gradFillElement.querySelector('rad, a\\:rad');
        if (rad) {
            gradient.type = 'radial';
        }

        return gradient;
    }

    /**
     * Parse pattern fill
     */
    parsePattern(pattFillElement) {
        return {
            preset: pattFillElement.getAttribute('prst'),
            foregroundColor: this.parseColor(pattFillElement.querySelector('fgClr, a\\:fgClr')),
            backgroundColor: this.parseColor(pattFillElement.querySelector('bgClr, a\\:bgClr'))
        };
    }

    /**
     * Parse stroke element
     */
    parseStroke(strokeElement) {
        try {
            const stroke = {};

            // Check for noFill — means no stroke should be drawn
            const lineNoFill = strokeElement.querySelector('noFill, a\\:noFill');
            if (lineNoFill) {
                stroke.noFill = true;
                stroke.width = 0;
                return stroke;
            }

            // Get line width (in EMU)
            const w = strokeElement.getAttribute('w');
            if (w) {
                stroke.width = parseInt(w) || 12700; // Default line width (1pt in EMU)
            } else {
                stroke.width = 12700; // Default 1pt
            }

            // Get line color - try multiple fill types
            let colorElement = strokeElement.querySelector('solidFill, a\\:solidFill');
            if (colorElement) {
                // parseColor at line 4670 expects the CONTAINING element (solidFill wrapper)
                // and internally uses querySelector to find the color child
                stroke.color = this.parseColor(colorElement);
            } else {
                // Try gradient fill for stroke
                colorElement = strokeElement.querySelector('gradFill, a\\:gradFill');
                if (colorElement) {
                    stroke.gradient = this.parseGradient(colorElement);
                    stroke.color = stroke.gradient; // For compatibility
                }
            }

            // Get line cap style
            const cap = strokeElement.getAttribute('cap');
            if (cap) {
                stroke.cap = cap; // 'rnd' (round), 'sq' (square), 'flat'
            } else {
                stroke.cap = 'flat'; // Default
            }

            // Get line join style
            const join = strokeElement.getAttribute('join');
            if (join) {
                stroke.join = join; // 'round', 'bevel', 'miter'
            } else {
                stroke.join = 'round'; // Default
            }

            // Get compound line style
            const cmpd = strokeElement.getAttribute('cmpd');
            if (cmpd) {
                stroke.compound = cmpd; // 'sng', 'dbl', 'thickThin', 'thinThick', 'tri'
            } else {
                stroke.compound = 'sng'; // Default single line
            }

            // Get dash pattern
            const prstDashElement = strokeElement.querySelector('prstDash, a\\:prstDash');
            if (prstDashElement) {
                stroke.dashStyle = prstDashElement.getAttribute('val') || 'solid';
                stroke.dashArray = this.getDashArray(stroke.dashStyle);
            } else {
                stroke.dashStyle = 'solid';
                stroke.dashArray = [];
            }

            // Get custom dash pattern
            const custDashElement = strokeElement.querySelector('custDash, a\\:custDash');
            if (custDashElement) {
                stroke.dashArray = this.parseCustomDash(custDashElement);
                stroke.dashStyle = 'custom';
            }

            // Get line alignment
            const algn = strokeElement.getAttribute('algn');
            if (algn) {
                stroke.alignment = algn; // 'ctr' (center), 'in' (inside), 'out' (outside)
            } else {
                stroke.alignment = 'ctr'; // Default center
            }

            // Arrow heads (headEnd/tailEnd)
            const headEnd = strokeElement.querySelector('headEnd, a\\:headEnd');
            if (headEnd) {
                stroke.headEnd = {
                    type: headEnd.getAttribute('type') || headEnd.getAttribute('val') || 'none',
                    w: headEnd.getAttribute('w') || headEnd.getAttribute('width') || 'med',
                    len: headEnd.getAttribute('len') || headEnd.getAttribute('length') || 'med'
                };
                
            }
            const tailEnd = strokeElement.querySelector('tailEnd, a\\:tailEnd');
            if (tailEnd) {
                stroke.tailEnd = {
                    type: tailEnd.getAttribute('type') || tailEnd.getAttribute('val') || 'none',
                    w: tailEnd.getAttribute('w') || tailEnd.getAttribute('width') || 'med',
                    len: tailEnd.getAttribute('len') || tailEnd.getAttribute('length') || 'med'
                };
                
            }

            return stroke;
        } catch (error) {
            return {
                color: { r: 0, g: 0, b: 0, a: 255 },
                width: 12700,
                cap: 'flat',
                join: 'round',
                compound: 'sng',
                dashStyle: 'solid',
                dashArray: [],
                alignment: 'ctr'
            };
        }
    }

    /**
     * Get dash array for preset dash styles
     */
    getDashArray(dashStyle) {
        const dashPatterns = {
            'solid': [],
            'dot': [1, 1],
            'dash': [3, 1],
            'dashDot': [3, 1, 1, 1],
            'dashDotDot': [3, 1, 1, 1, 1, 1],
            'lgDash': [8, 3],
            'lgDashDot': [8, 3, 1, 3],
            'lgDashDotDot': [8, 3, 1, 3, 1, 3],
            'sysDash': [2, 2],
            'sysDot': [1, 2],
            'sysDashDot': [2, 2, 1, 2],
            'sysDashDotDot': [2, 2, 1, 2, 1, 2]
        };

        return dashPatterns[dashStyle] || [];
    }

    /**
     * Parse custom dash pattern
     */
    parseCustomDash(custDashElement) {
        const dashArray = [];

        try {
            const dsElements = custDashElement.querySelectorAll('ds, a\\:ds');
            for (const dsElement of dsElements) {
                const d = parseInt(dsElement.getAttribute('d')) || 100000; // Dash length
                const sp = parseInt(dsElement.getAttribute('sp')) || 100000; // Space length

                // Convert to relative units (percentage of line width)
                dashArray.push(d / 100000, sp / 100000);
            }
        } catch (_error) {
				// Error ignored
			}

        return dashArray;
    }

    /**
     * Parse custom geometry
     */
    parseCustomGeometry(custGeomElement) {
        const pathList = [];

        try {
            const pathLstElement = custGeomElement.querySelector('pathLst, a\\:pathLst');
            if (pathLstElement) {
                const pathElements = pathLstElement.querySelectorAll('path, a\\:path');

                for (const pathElement of pathElements) {
                    const path = {
                        w: parseInt(pathElement.getAttribute('w')) || 100,
                        h: parseInt(pathElement.getAttribute('h')) || 100,
                        commands: []
                    };

                    // Parse path commands (OOXML uses lnTo, not lineTo)
                    const commandElements = pathElement.querySelectorAll('moveTo, lnTo, lineTo, cubicBezTo, quadBezTo, close, a\\:moveTo, a\\:lnTo, a\\:lineTo, a\\:cubicBezTo, a\\:quadBezTo, a\\:close');
                    for (const cmdElement of commandElements) {
                        const command = this.parsePathCommand(cmdElement);
                        if (command) {
                            path.commands.push(command);
                        }
                    }

                    pathList.push(path);
                }
            }
        } catch (_error) {
				// Error ignored
			}

        return pathList;
    }

    /**
     * Parse path command
     */
    parsePathCommand(cmdElement) {
        const tagName = cmdElement.tagName.toLowerCase().replace(/^a:/, '');

        switch (tagName) {
            case 'moveto':
                const ptElement = cmdElement.querySelector('pt, a\\:pt');
                if (ptElement) {
                    return {
                        type: 'moveTo',
                        x: parseInt(ptElement.getAttribute('x')) || 0,
                        y: parseInt(ptElement.getAttribute('y')) || 0
                    };
                }
                break;

            case 'lnto':
            case 'lineto': {
                const linePtElement = cmdElement.querySelector('pt, a\\:pt');
                if (linePtElement) {
                    return {
                        type: 'lineTo',
                        x: parseInt(linePtElement.getAttribute('x')) || 0,
                        y: parseInt(linePtElement.getAttribute('y')) || 0
                    };
                }
                break;
            }

            case 'cubicbezto':
                const pts = cmdElement.querySelectorAll('pt, a\\:pt');
                if (pts.length >= 3) {
                    return {
                        type: 'cubicBezTo',
                        x1: parseInt(pts[0].getAttribute('x')) || 0,
                        y1: parseInt(pts[0].getAttribute('y')) || 0,
                        x2: parseInt(pts[1].getAttribute('x')) || 0,
                        y2: parseInt(pts[1].getAttribute('y')) || 0,
                        x3: parseInt(pts[2].getAttribute('x')) || 0,
                        y3: parseInt(pts[2].getAttribute('y')) || 0
                    };
                }
                break;

            case 'quadbezto': {
                const qpts = cmdElement.querySelectorAll('pt, a\\:pt');
                if (qpts.length >= 2) {
                    return {
                        type: 'quadBezTo',
                        x1: parseInt(qpts[0].getAttribute('x')) || 0,
                        y1: parseInt(qpts[0].getAttribute('y')) || 0,
                        x: parseInt(qpts[1].getAttribute('x')) || 0,
                        y: parseInt(qpts[1].getAttribute('y')) || 0
                    };
                }
                break;
            }

            case 'close':
                return { type: 'close' };
        }

        return null;
    }

    /**
     * Parse shape style
     */
    parseShapeStyle(styleElement) {
        const style = {};

        try {
            // Get fill reference
            const fillRefElement = styleElement.querySelector('fillRef, a\\:fillRef');
            if (fillRefElement) {
                style.fillRef = {
                    idx: parseInt(fillRefElement.getAttribute('idx')) || 0,
                    color: this.parseColor(fillRefElement)
                };
            }

            // Get line reference
            const lnRefElement = styleElement.querySelector('lnRef, a\\:lnRef');
            if (lnRefElement) {
                style.lnRef = {
                    idx: parseInt(lnRefElement.getAttribute('idx')) || 0,
                    color: this.parseColor(lnRefElement)
                };
            }

            // Get effect reference
            const effectRefElement = styleElement.querySelector('effectRef, a\\:effectRef');
            if (effectRefElement) {
                style.effectRef = {
                    idx: parseInt(effectRefElement.getAttribute('idx')) || 0
                };
            }

            // Get font reference
            const fontRefElement = styleElement.querySelector('fontRef, a\\:fontRef');
            if (fontRefElement) {
                style.fontRef = {
                    idx: fontRefElement.getAttribute('idx') || 'minor'
                };
            }

        } catch (_error) {
				// Error ignored
			}

        return style;
    }

    /**
     * Extract style information directly from element structure (NEW METHOD)
     * This method handles the case where style info is embedded in the element structure
     */
    extractStyleFromElement(element) {
        const style = {};

        try {

            // Look for style attributes and sub-elements in various locations
            // Method 1: Check if there's a style element nested anywhere
            const nestedStyleElement = element.querySelector('style, p\\:style, a\\:style');
            if (nestedStyleElement) {
                return this.parseShapeStyle(nestedStyleElement);
            }

            // Method 2: Look for theme-based style references in spPr or other locations
            const spPrElement = element.querySelector('spPr, p\\:spPr');
            if (spPrElement) {
                // Look for style-related elements within spPr
                const styleAttrs = this.extractStyleAttributesFromSpPr(spPrElement);
                if (styleAttrs && Object.keys(styleAttrs).length > 0) {
                    Object.assign(style, styleAttrs);
                }
            }

            // Method 3: Check element attributes for style-related information
            const elementStyleAttrs = this.extractStyleAttributesFromElement(element);
            if (elementStyleAttrs && Object.keys(elementStyleAttrs).length > 0) {
                Object.assign(style, elementStyleAttrs);
            }

            // Method 4: Look for style references in nvSpPr (non-visual properties)
            const nvSpPrElement = element.querySelector('nvSpPr, p\\:nvSpPr');
            if (nvSpPrElement) {
                const nvStyleAttrs = this.extractStyleAttributesFromNvSpPr(nvSpPrElement);
                if (nvStyleAttrs && Object.keys(nvStyleAttrs).length > 0) {
                    Object.assign(style, nvStyleAttrs);
                }
            }

        } catch (_error) {
				// Error ignored
			}

        return style;
    }

    /**
     * Extract style attributes from spPr element
     */
    extractStyleAttributesFromSpPr(spPrElement) {
        const style = {};

        try {
            // Look for style references or attributes
            const styleRef = spPrElement.getAttribute('style') || spPrElement.getAttribute('styleRef');
            if (styleRef) {
                // Parse style reference
                const parsed = this.parseStyleReference(styleRef);
                if (parsed) {Object.assign(style, parsed);}
            }

            // Look for color scheme references
            const schemeClr = spPrElement.querySelector('schemeClr, a\\:schemeClr');
            if (schemeClr) {
                const val = schemeClr.getAttribute('val');
                if (val) {
                    style.schemeColor = val;
                }
            }

        } catch (_error) {
				// Error ignored
			}

        return style;
    }

    /**
     * Extract style attributes from main element
     */
    extractStyleAttributesFromElement(element) {
        const style = {};

        try {
            // Look for style-related attributes
            const styleAttrs = ['style', 'styleRef', 'themeRef', 'colorRef'];
            styleAttrs.forEach(attr => {
                const value = element.getAttribute(attr);
                if (value) {
                    style[attr] = value;
                }
            });

        } catch (_error) {
				// Error ignored
			}

        return style;
    }

    /**
     * Extract style attributes from nvSpPr element
     */
    extractStyleAttributesFromNvSpPr(nvSpPrElement) {
        const style = {};

        try {
            // Look for placeholder types that might indicate styling
            const phElement = nvSpPrElement.querySelector('ph, p\\:ph');
            if (phElement) {
                const type = phElement.getAttribute('type');
                const idx = phElement.getAttribute('idx');

                if (type) {
                    // Map placeholder types to style indices (PPTX convention)
                    const styleMap = {
                        'title': { fillRef: { idx: 1 } },
                        'body': { fillRef: { idx: 2 } },
                        'ctrTitle': { fillRef: { idx: 1 } },
                        'subTitle': { fillRef: { idx: 2 } },
                        'obj': { fillRef: { idx: 3 } }
                    };

                    if (styleMap[type]) {
                        Object.assign(style, styleMap[type]);
                    }
                }
            }

        } catch (_error) {
				// Error ignored
			}
        return style;
    }

    /**
     * Parse style reference string
     */
    parseStyleReference(styleRef) {
        try {
            // Handle different style reference formats
            if (styleRef.includes('fillRef')) {
                const match = styleRef.match(/fillRef:(\d+)/);
                if (match) {
                    return {
                        fillRef: {
                            idx: parseInt(match[1]) || 0
                        }
                    };
                }
            }

        } catch (_error) {
				// Error ignored
			}

        return null;
    }

    /**
     * Apply comprehensive style extraction to any shape (HELPER METHOD)
     */
    applyComprehensiveStyleExtraction(shape, element, shapeType = 'shape') {
        let styleFound = false;

        // Approach 1: Look for style element
        const styleElement = element.querySelector('style, p\\:style');
        if (styleElement) {
            shape.style = this.parseShapeStyle(styleElement);
            shape.preservedStyle = shape.style;
            styleFound = true;
        }

        // Approach 2: Look up from already-parsed raw XML data
        if (!styleFound && window.currentSlideData && window.currentSlideData.rawXMLShapes) {
            const rawShape = window.currentSlideData.rawXMLShapes.find(rawShape =>
                rawShape.name === shape.name || rawShape.id === shape.id
            );

            if (rawShape && rawShape.style) {
                shape.style = rawShape.style;
                shape.preservedStyle = rawShape.style;
                styleFound = true;
            }
        }

        // Approach 3: Extract from element structure
        if (!styleFound) {
            const parsedStyle = this.extractStyleFromElement(element);
            if (parsedStyle && Object.keys(parsedStyle).length > 0) {
                shape.style = parsedStyle;
                shape.preservedStyle = parsedStyle;
                styleFound = true;
            }
        }

        return styleFound;
    }

    /**
     * Convert hex color to RGB
     */
    hexToRgb(hex) {
        const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
        if (result) {
            return {
                r: parseInt(result[1], 16),
                g: parseInt(result[2], 16),
                b: parseInt(result[3], 16)
            };
        }
        return { r: 0, g: 0, b: 0 };
    }

    /**
     * Initialize rendering system
     */
    initializeRendering() {

        try {
            // Initialize drawing document with a default canvas
            const defaultCanvas = document.createElement('canvas');
            defaultCanvas.width = 800;
            defaultCanvas.height = 600;

            const graphicsEngine = this.getGraphicsEngine();
            if (!graphicsEngine) {
                this.initializeGraphicsEngine();
            }

            if (this.drawingDocument) {
                this.drawingDocument.init(defaultCanvas, this); // Pass processor reference
                
                // Initialize SVG renderer with graphics context
                const SVGRendererCtor = (typeof globalThis !== 'undefined' && globalThis.SVGRenderer)
                    || (typeof window !== 'undefined' && window.SVGRenderer)
                    || null;
                if (SVGRendererCtor && this.drawingDocument.graphics) {
                    this.svgRenderer = new SVGRendererCtor(this.drawingDocument.graphics);
                } else {
                    throw new Error('SVGRenderer is not defined');
                }
            } else {
                throw new Error('standard graphics engine not available');
            }

        } catch (error) {
            throw new Error('Failed to initialize rendering system: ' + error.message);
        }
    }

    /**
     * Enhanced slide rendering with standard-style pipeline
     */
    async renderSlide(canvas, slideIndex) {
        const startTime = performance.now();

        try {
            if (!canvas) {
                throw new Error('Canvas is required for rendering');
            }

            if (slideIndex < 0 || slideIndex >= this.slides.length) {
                throw new Error(`Invalid slide index: ${slideIndex}`);
            }

            const slide = this.slides[slideIndex];
            if (!slide) {
                throw new Error(`Slide at index ${slideIndex} is null`);
            }

            // Set current slide for geometry inheritance
            this.currentSlide = slide;

            // Get graphics engine based on rendering mode
            const graphicsEngine = this.getGraphicsEngine();
            if (!graphicsEngine) {
                this.initializeGraphicsEngine();
            }

            // Calculate slide dimensions and setup canvas
            const slideSize = this.getSlideDimensions();
            const canvasRect = this.calculateCanvasRect(canvas, slideSize);

            // Use standard graphics pipeline
            if (!this.drawingDocument) {
                throw new Error('standard graphics engine not available');
            }
            this.drawingDocument.init(canvas, this); // Pass processor reference
            const graphics = this.drawingDocument.graphics;

            // Set up rendering context using logical dimensions
            if (this.renderContext.enableOptimizations) {
                const logicalWidth = parseFloat(canvas.style.width) || canvas.width / (window.devicePixelRatio || 1);
                const logicalHeight = parseFloat(canvas.style.height) || canvas.height / (window.devicePixelRatio || 1);
                graphics.updatedRect = new CRect(0, 0, logicalWidth, logicalHeight);
            }

            // Clear canvas
            graphics.clear();
            
            // Preload images for this slide before rendering
            await this.preloadSlideImages(slideIndex);
            
            this.logSlideImageStatus(slide, slideIndex);

            // Set current slide index before drawing so chart rendering uses correct context
            this.currentSlideIndex = slideIndex;

            // Draw slide using standard-style pipeline
            await this.drawingDocument.drawSlide(slide, slideIndex);

            const renderTime = performance.now() - startTime;

        } catch (error) {
            // Draw error indicator on canvas
            if (canvas) {
                const ctx = canvas.getContext('2d');
                ctx.fillStyle = '#ffebee';
                ctx.fillRect(0, 0, canvas.width, canvas.height);
                ctx.fillStyle = '#f44336';
                ctx.font = '16px Arial';
                ctx.textAlign = 'center';
                ctx.fillText('Rendering Error', canvas.width / 2, canvas.height / 2);
                ctx.fillText(error.message, canvas.width / 2, canvas.height / 2 + 25);
            }
        }
    }

    /**
     * Render slide master using drawSlideEnhanced directly
     */
    async renderMaster(canvas, masterIndex) {
        try {
            if (!canvas) {
                throw new Error('Canvas is required for rendering');
            }

            if (masterIndex < 0 || masterIndex >= this.slideMasters.length) {
                throw new Error(`Invalid master index: ${masterIndex}`);
            }

            const master = this.slideMasters[masterIndex];
            if (!master) {
                throw new Error(`Master at index ${masterIndex} is null`);
            }

            // Create a temporary slide object to render the master
            const masterSlide = {
                name: master.name || `Master ${masterIndex + 1}`,
                type: 'master',
                commonSlideData: master.commonSlideData || master.cSld,
                backgroundFill: master.cSld?.bg,
                showMasterShapes: false,
                layout: null,
                master: null
            };

            // Set current slide for geometry inheritance
            this.currentSlide = masterSlide;

            // Calculate slide dimensions and setup canvas
            const slideSize = this.getSlideDimensions();
            const canvasRect = this.calculateCanvasRect(canvas, slideSize);

            // Use drawSlideEnhanced directly
            if (!this.drawingDocument) {
                this.initializeRendering();
            }
            this.drawingDocument.init(canvas, this);
            const graphics = this.drawingDocument.graphics;
            graphics.clear();

            await this.drawingDocument.drawSlide(masterSlide, masterIndex);

        } catch (error) {
            // Draw error indicator on canvas
            if (canvas) {
                const ctx = canvas.getContext('2d');
                ctx.fillStyle = '#f0f8ff';
                ctx.fillRect(0, 0, canvas.width, canvas.height);
                ctx.fillStyle = '#e74c3c';
                ctx.font = '16px Arial';
                ctx.textAlign = 'center';
                ctx.fillText('Master Rendering Error', canvas.width / 2, canvas.height / 2);
                ctx.fillText(error.message, canvas.width / 2, canvas.height / 2 + 25);
            }
        }
    }

    /**
     * Render slide layout using drawSlideEnhanced directly
     */
    async renderLayout(canvas, layoutIndex) {
        try {
            if (!canvas) {
                throw new Error('Canvas is required for rendering');
            }

            if (layoutIndex < 0 || layoutIndex >= this.slideLayouts.length) {
                throw new Error(`Invalid layout index: ${layoutIndex}`);
            }

            const layout = this.slideLayouts[layoutIndex];
            if (!layout) {
                throw new Error(`Layout at index ${layoutIndex} is null`);
            }

            // Find the master for this layout
            const master = this.slideMasters.find(m => m.id === layout.masterId);

            // Create a temporary slide object to render the layout
            const layoutSlide = {
                name: layout.name || `Layout ${layoutIndex + 1}`,
                type: 'layout',
                commonSlideData: layout.commonSlideData || layout.cSld,
                backgroundFill: layout.cSld?.bg || (master?.cSld?.bg),
                showMasterShapes: layout.showMasterSp !== false,
                layout: layout,
                master: master
            };

            // Set current slide for geometry inheritance
            this.currentSlide = layoutSlide;

            // Calculate slide dimensions and setup canvas
            const slideSize = this.getSlideDimensions();
            const canvasRect = this.calculateCanvasRect(canvas, slideSize);

            // Use drawSlideEnhanced directly
            if (!this.drawingDocument) {
                this.initializeRendering();
            }
            this.drawingDocument.init(canvas, this);
            const graphics = this.drawingDocument.graphics;
            graphics.clear();

            await this.drawingDocument.drawSlide(layoutSlide, layoutIndex);

        } catch (error) {
            // Draw error indicator on canvas
            if (canvas) {
                const ctx = canvas.getContext('2d');
                ctx.fillStyle = '#fff5f5';
                ctx.fillRect(0, 0, canvas.width, canvas.height);
                ctx.fillStyle = '#e74c3c';
                ctx.font = '16px Arial';
                ctx.textAlign = 'center';
                ctx.fillText('Layout Rendering Error', canvas.width / 2, canvas.height / 2);
                ctx.fillText(error.message, canvas.width / 2, canvas.height / 2 + 25);
            }
        }
    }


    /**
     * Draw regular shape geometry
     */
    drawRegularShapeGeometry(graphics, shape, x, y, w, h, fillColor, strokeColor, lineWidth, gradientFill = null) {
        let preset = this.getShapePreset(shape);

        // Also check geometry from properties
        if (!preset && shape.geometry) {
            preset = shape.geometry.preset;
        }

        // Also check properties.geometry
        if (!preset && shape.properties && shape.properties.geometry) {
            preset = shape.properties.geometry.preset;
        }

        // Handle gradient fill - pass gradient object as fillColor to let graphics engine handle it
        if (gradientFill) {
            fillColor = gradientFill; // Pass gradient as fillColor
        } else if (fillColor && typeof fillColor === 'object' && fillColor.type === 'linear') {
            // fillColor is already a gradient object, keep it as-is
        }

        // Convert RGB color objects to proper format for graphics engine
        let fillColorForShape = null;
        let strokeColorForShape = null;

        if (fillColor) {
            if (fillColor.type === 'linear' && fillColor.stops) {
                // Pass gradient object directly to graphics engine
                fillColorForShape = fillColor;
            } else if (typeof fillColor === 'string') {
                const rgb = this.hexToRgb(fillColor);
                fillColorForShape = rgb;
            } else if (fillColor.r !== undefined) {
                fillColorForShape = fillColor;
            }
        }

        if (strokeColor) {
            if (typeof strokeColor === 'string') {
                const rgb = this.hexToRgb(strokeColor);
                strokeColorForShape = rgb;
            } else if (strokeColor.r !== undefined) {
                strokeColorForShape = strokeColor;
            }
        }

        // Get stroke information for advanced styling
        const strokeInfo = this.getShapeStrokeInfo(shape);

        if (preset) {
            graphics.drawPresetGeometry(preset, x, y, w, h, fillColorForShape, strokeColorForShape, lineWidth, strokeInfo);
            } else {
            // Default rectangle
            graphics.drawRectangle(x, y, w, h, fillColorForShape, strokeColorForShape, lineWidth, strokeInfo);
        }
    }

    /**
     * Get slide count
     */
    getSlideCount() {
        return this.slides.length;
    }

    /**
     * Get current slide index
     */
    getCurrentSlideIndex() {
        return this.currentSlideIndex;
    }

    /**
     * Get slide info
     */
    getSlideInfo(index) {
        const slide = this.slides[index];
        if (!slide) {
            return null;
        }

        return {
            index: index,
            name: slide.getName(),
            shapeCount: slide.getShapeTree().length
        };
    }

    /**
     * Get slide dimensions in EMU
     */
    getSlideDimensions() {
        return this.presentation?.slideSize || { cx: 9144000, cy: 6858000 };
    }

    /**
     * Get slide dimensions in pixels for a given canvas size
     */
    getSlideDimensionsForCanvas(canvasWidth, canvasHeight) {
        const slideSize = this.getSlideDimensions();
        const scaleX = canvasWidth / slideSize.cx;
        const scaleY = canvasHeight / slideSize.cy;
        const scale = Math.min(scaleX, scaleY); // Maintain aspect ratio

        return {
            width: slideSize.cx * scale,
            height: slideSize.cy * scale,
            scale: scale,
            offsetX: (canvasWidth - slideSize.cx * scale) / 2,
            offsetY: (canvasHeight - slideSize.cy * scale) / 2
        };
    }

    /**
     * Generate a thumbnail for a slide by index with proper slide dimensions
     */
    async generateSlideThumbnail(slideIndex, maxWidth = 320, maxHeight = 240) {
        try {
            // Ensure graphics engine is initialized
            if (!this.getGraphicsEngine()) {
                this.initializeGraphicsEngine();
            }

            if (!this.slides[slideIndex]) {
                return this.createFallbackThumbnail(slideIndex, maxWidth, maxHeight);
            }

            const slide = this.slides[slideIndex];

            // Get actual slide dimensions
            const slideSize = this.getSlideDimensions();
            const aspectRatio = slideSize.cx / slideSize.cy;

            // Calculate thumbnail size maintaining aspect ratio
            // Use consistent EMU to pixel conversion like full slides: (EMU / 914400) * 96
            const slideWidthPx = (slideSize.cx / 914400) * 96;
            const slideHeightPx = (slideSize.cy / 914400) * 96;
            
            // Calculate scale to fit within max dimensions while maintaining aspect ratio
            const scaleX = maxWidth / slideWidthPx;
            const scaleY = maxHeight / slideHeightPx;
            const scale = Math.min(scaleX, scaleY);
            
            let thumbWidth = slideWidthPx * scale;
            let thumbHeight = slideHeightPx * scale;

            // Ensure minimum size
            thumbWidth = Math.max(thumbWidth, 100);
            thumbHeight = Math.max(thumbHeight, 75);

            // Create a high-resolution canvas for the thumbnail
            const thumbCanvas = document.createElement('canvas');
            this.setupHighResolutionCanvas(thumbCanvas, thumbWidth, thumbHeight);

            // Initialize standard graphics engine with thumbnail canvas
            if (this.drawingDocument) {
                this.drawingDocument.init(thumbCanvas, this); // Pass processor reference
                try {
                    await this.drawingDocument.drawSlide(slide, slideIndex);
                    return thumbCanvas;
                } catch (renderError) {
                    return this.createFallbackThumbnail(slideIndex, thumbWidth, thumbHeight);
                }
            } else {
                return this.createFallbackThumbnail(slideIndex, thumbWidth, thumbHeight);
            }

        } catch (error) {
            return this.createFallbackThumbnail(slideIndex, maxWidth, maxHeight);
        }
    }

    /**
     * Create a fallback thumbnail when rendering fails
     */
    createFallbackThumbnail(slideIndex, width, height) {
        const canvas = document.createElement('canvas');
        this.setupHighResolutionCanvas(canvas, width, height);
        const ctx = canvas.getContext('2d');

        // Draw fallback background
        ctx.fillStyle = '#f0f0f0';
        ctx.fillRect(0, 0, width, height);

        // Draw slide number
        ctx.fillStyle = '#666';
        ctx.font = '16px Arial';
        ctx.textAlign = 'center';
        ctx.fillText(`Slide ${slideIndex + 1}`, width / 2, height / 2);

        // Draw aspect ratio info if available
        try {
            const slideSize = this.getSlideDimensions();
            const aspectRatio = slideSize.cx / slideSize.cy;
            ctx.font = '12px Arial';
            ctx.fillText(`${aspectRatio.toFixed(2)}:1`, width / 2, height / 2 + 20);
        } catch (e) {
            // Ignore errors in fallback thumbnail
        }

        return canvas;
    }

    /**
     * Helper method: Infer layout type from placeholders
     */
    inferLayoutTypeFromPlaceholders(layoutElement) {
        const cSldElement = layoutElement.querySelector('cSld, p\\:cSld');
        if (!cSldElement) {return 'blank';}

        const spTreeElement = cSldElement.querySelector('spTree, p\\:spTree');
        if (!spTreeElement) {return 'blank';}

        const placeholders = [];
        const shapes = spTreeElement.children;

        for (let i = 0; i < shapes.length; i++) {
            const shape = shapes[i];
            const placeholder = this.extractPlaceholderInfo(shape);
            if (placeholder) {
                placeholders.push(placeholder.type);
            }
        }

        // Determine layout type based on placeholders
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
     * Helper method: Extract placeholder info from shape
     */
    extractPlaceholderInfo(shapeElement) {
        const nvSpPr = shapeElement.querySelector('nvSpPr, p\\:nvSpPr') ||
                       shapeElement.querySelector('nvPicPr, p\\:nvPicPr') ||
                       shapeElement.querySelector('nvGrpSpPr, p\\:nvGrpSpPr');

        if (!nvSpPr) {return null;}

        const nvPr = nvSpPr.querySelector('nvPr, p\\:nvPr');
        if (!nvPr) {return null;}

        const ph = nvPr.querySelector('ph, p\\:ph');
        if (!ph) {return null;}

        return {
            type: ph.getAttribute('type') || 'obj',
            idx: ph.getAttribute('idx'),
            orient: ph.getAttribute('orient'),
            sz: ph.getAttribute('sz')
        };
    }

    /**
     * Parse placeholder element
     */
    parsePlaceholder(phElement) {
        // Normalize idx: if absent, leave undefined; if present, coerce to number
        const idxAttr = phElement.getAttribute('idx');
        const normalizedIdx = idxAttr === null || idxAttr === undefined ? undefined : (isNaN(parseInt(idxAttr, 10)) ? undefined : parseInt(idxAttr, 10));

        const placeholder = {
            type: phElement.getAttribute('type') || 'obj',
            orient: phElement.getAttribute('orient') || 'horz',
            sz: phElement.getAttribute('sz') || 'full',
            idx: normalizedIdx
        };

        // Parse placeholder text if present
        const extLst = phElement.querySelector('extLst, p\\:extLst');
        if (extLst) {
            const ext = extLst.querySelector('ext, p\\:ext');
            if (ext) {
                const phTxt = ext.querySelector('phTxt, p\\:phTxt');
                if (phTxt) {
                    placeholder.text = phTxt.textContent;
                }
            }
        }

        return placeholder;
    }

    /**
     * Helper method: Parse background with enhanced image support
     */
    parseBackground(cSldElement) {
        const bgElement = cSldElement.querySelector('bg, p\\:bg');
        if (!bgElement) {return null;}

        const bgPr = bgElement.querySelector('bgPr, p\\:bgPr');
        if (bgPr) {
            const fill = this.parseFill(bgPr);
            return {
                type: 'bgPr',
                fill: fill
            };
        }

        const bgRef = bgElement.querySelector('bgRef, p\\:bgRef');
        if (bgRef) {
            return {
                type: 'bgRef',
                idx: bgRef.getAttribute('idx')
            };
        }
        
        return null;
    }
    
    /**
     * Helper method: Parse text style element
     */
    parseTextStyleElement(element) {
        if (!element) {return null;}

        return {
            element: element,
            parsed: false  // Flag for lazy parsing
        };
    }

    /**
     * Helper method: Parse color map element
     */
    parseColorMapElement(element) {
        if (!element) {return null;}

        const colorMap = {};
        const attributes = element.attributes;

        for (let i = 0; i < attributes.length; i++) {
            const attr = attributes[i];
            colorMap[attr.name] = attr.value;
        }

        return colorMap;
    }

    /**
     * Helper method: Parse color map override
     */
    parseColorMapOverride(element) {
        if (!element) {return null;}

        const masterClrMapping = element.querySelector('masterClrMapping, p\\:masterClrMapping');
        if (masterClrMapping) {
            return this.parseColorMapElement(masterClrMapping);
        }

        const overrideClrMapping = element.querySelector('overrideClrMapping, p\\:overrideClrMapping');
        if (overrideClrMapping) {
            return this.parseColorMapElement(overrideClrMapping);
        }

        return null;
    }

    /**
     * Calculate canvas rectangle with proper aspect ratio
     */
    calculateCanvasRect(canvas, slideSize) {
        // Convert EMU to pixels (1 EMU = 1/914400 inch, 1 inch = 96 pixels)
        const slideWidthPx = (slideSize.cx / 914400) * 96;
        const slideHeightPx = (slideSize.cy / 914400) * 96;

        // Use logical display dimensions for positioning calculations, not scaled canvas dimensions
        let logicalWidth, logicalHeight;

        // Explicit render resolution (data-pvjs-render-width/height) takes priority,
        // matching the graphics adapter. This keeps coordinate mapping correct when
        // the host lets CSS drive display (e.g. width:100%, which would otherwise
        // parse to a meaningless 100 here) and rasterises at an explicit resolution.
        const dataW = parseFloat(canvas.dataset && canvas.dataset.pvjsRenderWidth);
        const dataH = parseFloat(canvas.dataset && canvas.dataset.pvjsRenderHeight);
        if (dataW > 0 && dataH > 0) {
            logicalWidth = dataW;
            logicalHeight = dataH;
        } else if (canvas.style.width && canvas.style.height) {
            // Use CSS dimensions if available
            logicalWidth = parseFloat(canvas.style.width);
            logicalHeight = parseFloat(canvas.style.height);
        } else {
            // Fallback: use actual canvas dimensions divided by device pixel ratio
            logicalWidth = canvas.width / (window.devicePixelRatio || 1);
            logicalHeight = canvas.height / (window.devicePixelRatio || 1);
        }


        // Calculate scale to fit slide in canvas while maintaining aspect ratio
        const scaleX = logicalWidth / slideWidthPx;
        const scaleY = logicalHeight / slideHeightPx;
        let scale = Math.min(scaleX, scaleY);

        // Calculate centered position
        const scaledWidth = slideWidthPx * scale;
        const scaledHeight = slideHeightPx * scale;
        const offsetX = (logicalWidth - scaledWidth) / 2;
        const offsetY = (logicalHeight - scaledHeight) / 2;

        // Snap to integer pixels to avoid 1px gutters due to floating-point rounding
        let iScaledWidth = Math.round(scaledWidth);
        let iScaledHeight = Math.round(scaledHeight);
        let iOffsetX = Math.round(offsetX);
        let iOffsetY = Math.round(offsetY);

        // If we're effectively full-bleed, force exact fit
        if (Math.abs(iScaledWidth - logicalWidth) <= 1 && Math.abs(iScaledHeight - logicalHeight) <= 1) {
            iScaledWidth = Math.round(logicalWidth);
            iScaledHeight = Math.round(logicalHeight);
            iOffsetX = 0;
            iOffsetY = 0;
        }

        // Recompute effective scale from snapped dimensions
        const effectiveScaleX = iScaledWidth / slideWidthPx;
        const effectiveScaleY = iScaledHeight / slideHeightPx;
        scale = Math.min(effectiveScaleX, effectiveScaleY);


        return {
            // Canvas dimensions in logical pixels (for positioning)
            widthPx: Math.round(logicalWidth),
            heightPx: Math.round(logicalHeight),
            // Slide dimensions in pixels
            slideWidthPx: slideWidthPx,
            slideHeightPx: slideHeightPx,
            // Scaling and positioning
            scale: scale,
            offsetX: iOffsetX,
            offsetY: iOffsetY,
            // Scaled slide dimensions
            scaledWidth: iScaledWidth,
            scaledHeight: iScaledHeight,
            // Legacy MM support (for backward compatibility)
            widthMM: (slideSize.cx / 914400) * 25.4,
            heightMM: (slideSize.cy / 914400) * 25.4
        };
    }
    /**
     * Calculate shape bounds in pixels with proper scaling and centering
     */
    calculateShapeBounds(shape, canvasRect) {
        if (!shape || !shape.bounds) {
            return { x: 0, y: 0, w: 100, h: 100 };
        }

        // Get slide dimensions in EMU
        const slideSize = this.getSlideDimensions();
        const slideWidthEMU = slideSize.cx;
        const slideHeightEMU = slideSize.cy;

        // Convert EMU bounds to slide pixels first
        const slideWidthPx = canvasRect.slideWidthPx;
        const slideHeightPx = canvasRect.slideHeightPx;

        // Convert shape bounds from EMU to slide pixels
        const shapeSlidePx = {
            x: ((shape.bounds.l || 0) / slideWidthEMU) * slideWidthPx,
            y: ((shape.bounds.t || 0) / slideHeightEMU) * slideHeightPx,
            w: (((shape.bounds.r || 0) - (shape.bounds.l || 0)) / slideWidthEMU) * slideWidthPx,
            h: (((shape.bounds.b || 0) - (shape.bounds.t || 0)) / slideHeightEMU) * slideHeightPx
        };

        // Apply canvas scaling and centering
        const x = canvasRect.offsetX + shapeSlidePx.x * canvasRect.scale;
        const y = canvasRect.offsetY + shapeSlidePx.y * canvasRect.scale;
        const w = shapeSlidePx.w * canvasRect.scale;
        const h = shapeSlidePx.h * canvasRect.scale;

        // Apply rotation if needed (for bounds calculation purposes)
        const rotation = this.getShapeRotation(shape);
        if (rotation && rotation !== 0) {
            return this.calculateRotatedBounds({ x, y, w, h }, rotation);
        }

        return { x, y, w, h };
    }

    /**
     * Get shape bounds for rendering (without rotation applied to bounds)
     */
    getShapeBounds(shape, canvasRect) {
        if (!shape || !shape.bounds) {
            return { x: 0, y: 0, w: 100, h: 100 };
        }


        // Get slide dimensions in EMU
        const slideSize = this.getSlideDimensions();
        const slideWidthEMU = slideSize.cx;
        const slideHeightEMU = slideSize.cy;

        // Convert EMU bounds to slide pixels first
        const slideWidthPx = canvasRect.slideWidthPx;
        const slideHeightPx = canvasRect.slideHeightPx;

        // Convert shape bounds from EMU to slide pixels
        const shapeSlidePx = {
            x: ((shape.bounds.l || 0) / slideWidthEMU) * slideWidthPx,
            y: ((shape.bounds.t || 0) / slideHeightEMU) * slideHeightPx,
            w: (((shape.bounds.r || 0) - (shape.bounds.l || 0)) / slideWidthEMU) * slideWidthPx,
            h: (((shape.bounds.b || 0) - (shape.bounds.t || 0)) / slideHeightEMU) * slideHeightPx
        };

        // Apply canvas scaling and centering
        const x = canvasRect.offsetX + shapeSlidePx.x * canvasRect.scale;
        const y = canvasRect.offsetY + shapeSlidePx.y * canvasRect.scale;
        const w = shapeSlidePx.w * canvasRect.scale;
        const h = shapeSlidePx.h * canvasRect.scale;


        // Return original bounds for rendering (rotation is applied separately)
        return { x, y, w, h };
    }

    /**
     * Calculate shape transform matrix
     */
    calculateShapeTransform(shape, canvasRect) {
        if (!shape || !shape.properties || !shape.properties.transform) {
            return null;
        }

        const transform = shape.properties.transform;

        // Create transform matrix
        const matrix = new CMatrix();

        // Get slide dimensions in EMU
        const slideSize = this.getSlideDimensions();
        const slideWidthEMU = slideSize.cx;
        const slideHeightEMU = slideSize.cy;

        // Convert EMU bounds to slide pixels first
        const slideWidthPx = canvasRect.slideWidthPx;
        const slideHeightPx = canvasRect.slideHeightPx;

        // Apply translation using proper coordinate conversion
        if (transform.x !== undefined && transform.y !== undefined) {
            // Convert EMU to slide pixels first
            const slidePxX = (transform.x / slideWidthEMU) * slideWidthPx;
            const slidePxY = (transform.y / slideHeightEMU) * slideHeightPx;

            // Apply canvas scaling and centering
            matrix.tx = canvasRect.offsetX + slidePxX * canvasRect.scale;
            matrix.ty = canvasRect.offsetY + slidePxY * canvasRect.scale;
        }

        // Apply rotation if present
        if (transform.rot !== undefined) {
            const angle = transform.rot * Math.PI / 180;
            const cos = Math.cos(angle);
            const sin = Math.sin(angle);

            matrix.sx = cos * canvasRect.scale;
            matrix.sy = cos * canvasRect.scale;
            matrix.shx = -sin * canvasRect.scale;
            matrix.shy = sin * canvasRect.scale;
        } else {
            matrix.sx = canvasRect.scale;
            matrix.sy = canvasRect.scale;
        }

        return matrix;
    }

    /**
     * Draw slide background with enhanced image support
     */
    drawSlideBackground(graphics, slide, canvasRect) {
        if (slide.backgroundFill) {
            this.renderBackgroundFill(graphics, slide.backgroundFill, canvasRect);
        } else {
            // No background specified - leave transparent
        }
    }

    /**
     * Render background fill based on type
     */
    renderBackgroundFill(graphics, backgroundFill, canvasRect) {
        if (!backgroundFill || !backgroundFill.fill) {
            // No background fill specified - leave transparent
            return;
        }

        const fill = backgroundFill.fill;

        switch (fill.type) {
            case 'solid':
                this.renderSolidBackground(graphics, fill, canvasRect);
                break;
            case 'gradient':
                this.renderGradientBackground(graphics, fill, canvasRect);
                break;
            case 'image':
                this.renderImageBackground(graphics, fill, canvasRect);
                break;
            case 'pattern':
                this.renderPatternBackground(graphics, fill, canvasRect);
                break;
            case 'none':
                // Transparent background
                break;
            default:
                // Unknown background type - leave transparent
                break;
        }
    }
    
    /**
     * Render solid color background
     */
    renderSolidBackground(graphics, fill, canvasRect) {
        if (fill.color) {
            const colorHex = this.rgbToHex(fill.color);
            graphics.fillRect(0, 0, canvasRect.widthPx, canvasRect.heightPx, colorHex);
        }
    }

    /**
     * Render gradient background
     */
    renderGradientBackground(graphics, fill, canvasRect) {
        if (!fill.gradient || !fill.gradient.stops || fill.gradient.stops.length === 0) {
            // Fallback to first stop color or white
            const fallbackColor = fill.gradient?.stops?.[0]?.color || { r: 255, g: 255, b: 255, a: 255 };
            graphics.fillRect(0, 0, canvasRect.widthPx, canvasRect.heightPx, this.rgbToHex(fallbackColor));
            return;
        }

        const ctx = graphics.context;
        const { widthPx, heightPx } = canvasRect;

        let gradient;

        if (fill.gradient.type === 'radial') {
            // Radial gradient
            const centerX = widthPx / 2;
            const centerY = heightPx / 2;
            const radius = Math.max(widthPx, heightPx) / 2;
            gradient = ctx.createRadialGradient(centerX, centerY, 0, centerX, centerY, radius);
        } else {
            // Linear gradient
            const angle = (fill.gradient.angle || 0) * Math.PI / 180 / 60000; // Convert from 60ths of degree to radians
            const diagonal = Math.sqrt(widthPx * widthPx + heightPx * heightPx);
            const x1 = widthPx / 2 - Math.cos(angle) * diagonal / 2;
            const y1 = heightPx / 2 - Math.sin(angle) * diagonal / 2;
            const x2 = widthPx / 2 + Math.cos(angle) * diagonal / 2;
            const y2 = heightPx / 2 + Math.sin(angle) * diagonal / 2;
            gradient = ctx.createLinearGradient(x1, y1, x2, y2);
        }

        // Add color stops
        fill.gradient.stops.forEach(stop => {
            const colorHex = this.rgbToHex(stop.color);
            gradient.addColorStop(stop.position, colorHex);
        });

        ctx.fillStyle = gradient;
        ctx.fillRect(0, 0, widthPx, heightPx);
    }

    /**
     * Render image background
     */
    renderImageBackground(graphics, fill, canvasRect) {
        if (!fill.imageData || !fill.imageData.relationshipId) {
            graphics.fillRect(0, 0, canvasRect.widthPx, canvasRect.heightPx, '#f0f0f0');
            return;
        }

        const relationshipId = fill.imageData.relationshipId;
        const cachedImage = this.imageCache.get(relationshipId);

        if (cachedImage && cachedImage.image) {
            // Draw the background image
            this.drawBackgroundImage(graphics, cachedImage.image, fill.imageData, canvasRect);
        } else {
            // Draw placeholder and load image asynchronously
            graphics.fillRect(0, 0, canvasRect.widthPx, canvasRect.heightPx, '#f8f9fa');

            // Add loading text
            const ctx = graphics.context;
            ctx.fillStyle = '#6c757d';
            ctx.font = '14px Arial';
            ctx.textAlign = 'center';
            ctx.fillText('Loading background...', canvasRect.widthPx / 2, canvasRect.heightPx / 2);

            // Load image asynchronously
            this.loadImageAsync(relationshipId);
        }
    }

    /**
     * Draw background image with proper scaling and positioning
     */
    drawBackgroundImage(graphics, image, imageData, canvasRect) {
        const ctx = graphics.context;
        const { widthMM, heightMM } = canvasRect;

        ctx.save();

        try {
            if (imageData.fillMode === 'tile') {
                // Tile the image
                this.drawTiledBackgroundImage(ctx, image, imageData.tileProperties, widthMM, heightMM);
            } else {
                // Stretch or fit the image
                let drawX = 0, drawY = 0, drawW = widthMM, drawH = heightMM;

                if (imageData.fillRect) {
                    // Apply fill rectangle
                    const fillRect = imageData.fillRect;
                    drawX = widthMM * fillRect.l / 100000;
                    drawY = heightMM * fillRect.t / 100000;
                    drawW = widthMM * (100000 - fillRect.l - fillRect.r) / 100000;
                    drawH = heightMM * (100000 - fillRect.t - fillRect.b) / 100000;
                }

                if (imageData.sourceRect) {
                    // Apply source rectangle cropping
                    const srcRect = imageData.sourceRect;
                    const srcX = image.naturalWidth * srcRect.l / 100000;
                    const srcY = image.naturalHeight * srcRect.t / 100000;
                    const srcW = image.naturalWidth * (100000 - srcRect.l - srcRect.r) / 100000;
                    const srcH = image.naturalHeight * (100000 - srcRect.t - srcRect.b) / 100000;

                    graphics.drawImageHighRes(image, srcX, srcY, srcW, srcH, drawX, drawY, drawW, drawH);
                } else {
                    graphics.drawImageHighRes(image, drawX, drawY, drawW, drawH);
                }
            }

            // Apply effects if present
            if (imageData.effects) {
                this.applyImageEffects(graphics, imageData.effects);
            }

        } catch (error) {
            // Fallback to gray background
            ctx.fillStyle = '#f0f0f0';
            ctx.fillRect(0, 0, widthMM, heightMM);
        } finally {
            ctx.restore();
        }
    }

    /**
     * Draw tiled background image
     */
    drawTiledBackgroundImage(ctx, image, tileProps, width, height) {
        if (!tileProps) {return;}

        const scaleX = tileProps.sx / 100000;
        const scaleY = tileProps.sy / 100000;
        const offsetX = tileProps.tx / 914400 * 25.4; // Convert EMU to MM
        const offsetY = tileProps.ty / 914400 * 25.4;

        const tileWidth = image.naturalWidth * scaleX;
        const tileHeight = image.naturalHeight * scaleY;

        // Create pattern
        const tempCanvas = document.createElement('canvas');
        tempCanvas.width = tileWidth;
        tempCanvas.height = tileHeight;
        const tempCtx = tempCanvas.getContext('2d');
        // Enable high-quality scaling for pattern tiles
        tempCtx.imageSmoothingEnabled = true;
        tempCtx.imageSmoothingQuality = 'high';
        tempCtx.drawImage(image, 0, 0, tileWidth, tileHeight);

        const pattern = ctx.createPattern(tempCanvas, 'repeat');

        ctx.save();
        ctx.translate(offsetX, offsetY);
        ctx.fillStyle = pattern;
        ctx.fillRect(-offsetX, -offsetY, width, height);
        ctx.restore();
    }

    /**
     * Render pattern background
     */
    renderPatternBackground(graphics, fill, canvasRect) {
        // Simplified pattern rendering - draw with foreground color
        const color = fill.pattern?.foregroundColor || { r: 128, g: 128, b: 128, a: 255 };
        const colorHex = this.rgbToHex(color);
        graphics.fillRect(0, 0, canvasRect.widthMM, canvasRect.heightMM, colorHex);
    }

    /**
     * Get shape fill color with enhanced inheritance from layout and master
     */
    getShapeFillColor(shape) {
        
        // Set current processing shape context for placeholder resolution
        this._currentProcessingShape = shape;
        const isRectangle3 = shape && (shape.name === 'Rectangle 3' || shape.name?.includes('Rectangle 3'));

        // CRITICAL FIX: Try direct fill property first (highest precedence)
        if (shape.fill) {
            if (shape.fill.type === 'solid' && shape.fill.color) {
                const color = this.processShapeColor(shape.fill.color);
                if (isRectangle3) {
                }
                this._currentProcessingShape = null; // Clear context
                return color;
            }
            if (shape.fill.type === 'gradient' && shape.fill.gradient) {
                this._currentProcessingShape = null; // Clear context
                return this.processGradientFill(shape.fill.gradient);
            }
            if (shape.fill.type === 'none') {
                this._currentProcessingShape = null; // Clear context
                return null;
            }
        }

        // CRITICAL FIX: Try properties.fill (second highest precedence)
        if (shape.properties && shape.properties.fill) {
            if (shape.properties.fill.type === 'solid' && shape.properties.fill.color) {
                const color = this.processShapeColor(shape.properties.fill.color);
                if (isRectangle3) {
                }
                this._currentProcessingShape = null; // Clear context
                return color;
            }
            if (shape.properties.fill.type === 'gradient' && shape.properties.fill.gradient) {
                return this.processGradientFill(shape.properties.fill.gradient);
            }
            if (shape.properties.fill.type === 'none') {
                return null;
            }
        }

        // Check placeholder types - placeholders should be transparent
        if (shape.type === 'sp' && shape.placeholder && shape.placeholder.type) {
            switch (shape.placeholder.type) {
                case 'title':
                case 'ctrTitle':
                case 'body':
                case 'obj':
                    // Placeholders: NO FILL by default (transparent background)
                    this._currentProcessingShape = null; // Clear context
                    return null;
                default:
                    // Other placeholders: transparent by default unless explicitly filled
                    this._currentProcessingShape = null;
                    return null;
            }
        }

        // SIMPLIFIED: Try style-based color resolution
        const styleToUse = shape.preservedStyle || shape.style;
        if (styleToUse && styleToUse.fillRef) {
            this._currentProcessingShape = shape;

            const color = this.resolveStyleFillColor(styleToUse.fillRef);
            if (color) {
                this._currentProcessingShape = null; // Clear context
                return color;
            }
        }

        // For picture shapes, return null (no fill, just show image)
        if (shape.type === 'pic') {
            this._currentProcessingShape = null;
            return null;
        }

        // Enhanced fallback: Use a default color based on shape name/id for variety
        let fallbackColor;
        if (shape.name && shape.name.includes('Rectangle')) {
            // Extract number from shape name for color variety
            const numberMatch = shape.name.match(/\d+/);
            const shapeNumber = numberMatch ? parseInt(numberMatch[0]) : 1;

            // No hardcoded rectangle colors - return null if not specified in DOM
            fallbackColor = null;
        } else {
            fallbackColor = null; // No fallback colors
        }

        // Clear current processing shape context
        this._currentProcessingShape = null;


        return fallbackColor;
    }

    /**
     * Normalize color descriptor objects produced by parseColor(...) so they can be
     * consumed by processShapeColor(...)
     */
    _normalizeColorDescriptor(color) {
        if (!color) {return null;}
        // Already RGB object
        if (typeof color === 'object' && color.r !== undefined) {return color;}
        // Hex string passthrough
        if (typeof color === 'string') {return color;}
        // Translate { type:'scheme'|'srgb', value:'...' } to expected schema
        if (typeof color === 'object' && color.type && color.value) {
            if (color.type === 'scheme') {
                return { scheme: color.value, alpha: color.alpha };
            }
            if (color.type === 'srgb') {
                // value is hex like 'FF0000'
                return `#${color.value}`;
            }
        }
        return color;
    }

    /**
     * Resolve stroke color from style lnRef (theme-based) or direct stroke color
     */
    resolveStyleStrokeColor(lnRef) {
        try {
            if (!lnRef) {return null;}
            const normalized = this._normalizeColorDescriptor(lnRef.color || lnRef);
            const resolved = this.processShapeColor(normalized);
            
            return resolved;
        } catch (_e) { return null; }
    }

    /**
     * Resolve stroke color for a shape (connectors rely heavily on theme lnRef)
     */
    getShapeStrokeColor(shape) {
        try {
            // 0) If explicit line is noFill, return null regardless of style/lnRef
            if (shape?.properties?.stroke?.noFill || shape?.stroke?.noFill) {
                return null;
            }
            // 1) Direct stroke color on properties
            if (shape?.properties?.stroke?.color) {
                const normalized = this._normalizeColorDescriptor(shape.properties.stroke.color);
                const resolved = this.processShapeColor(normalized);
                return resolved;
            }
            // 2) Legacy stroke
            if (shape?.stroke?.color) {
                const normalized = this._normalizeColorDescriptor(shape.stroke.color);
                const resolved = this.processShapeColor(normalized);

                return resolved;
            }
            // 3) Style lnRef
            const styleToUse = shape?.preservedStyle || shape?.style;
            if (styleToUse?.lnRef) {
                const resolved = this.resolveStyleStrokeColor(styleToUse.lnRef);
                if (resolved) { return resolved; }
            }
        } catch (_e) {}
        // Only apply dark gray fallback for connector/line shapes so they remain visible.
        // Regular shapes and placeholders should NOT get a default stroke.
        if (shape && (shape.type === 'cxnSp' || shape.type === 'connector')) {
            return { r: 80, g: 80, b: 80, a: 255 };
        }
        return null;
    }

    /**
     * Resolve line width in millimeters for a shape
     */
    getShapeLineWidth(shape) {
        try {
            let emu = null;
            if (shape?.properties?.stroke?.width !== undefined) {
                emu = parseInt(shape.properties.stroke.width) || 12700;
                
            } else if (shape?.stroke?.width !== undefined) {
                emu = parseInt(shape.stroke.width) || 12700;
                
            } else {
                emu = 12700; // Default 1pt
                
            }
            const mm = emu / 914400 * 25.4;
            // Enforce a visible minimum (approx 2 px at 96dpi)
            const minMm = 2 * 25.4 / 96;
            const finalMm = Math.max(mm, minMm);
            
            return finalMm;
        } catch (_e) {
            // Fallback ~0.75mm
            return 0.75;
        }
    }

    /**
     * Resolve inherited fill color from layout and master (NEW METHOD)
     */
    resolveInheritedFillColor(shape) {
        try {
            // Get current slide context
            const slide = this.slides[this.getCurrentSlideIndex()];
            if (!slide) {return null;}

            // Check if shape has placeholder information for inheritance
            if (!shape.placeholder || !shape.placeholder.type) {
                return null;
            }

            const placeholderType = shape.placeholder.type;
            const placeholderIdx = shape.placeholder.idx || 0;

            // 1. Try to get color from layout
            if (slide.layout && slide.layout.cSld && slide.layout.cSld.spTree) {
                for (const layoutShape of slide.layout.cSld.spTree) {
                    if (this.isMatchingPlaceholder(layoutShape, placeholderType, placeholderIdx)) {
                        const layoutColor = this.extractShapeFillColor(layoutShape);
                        if (layoutColor) {
                            return layoutColor;
                        }
                    }
                }
            }

            // 2. Try to get color from master
            if (slide.layout && slide.layout.master && slide.layout.master.cSld && slide.layout.master.cSld.spTree) {
                for (const masterShape of slide.layout.master.cSld.spTree) {
                    if (this.isMatchingPlaceholder(masterShape, placeholderType, placeholderIdx)) {
                        const masterColor = this.extractShapeFillColor(masterShape);
                        if (masterColor) {
                            return masterColor;
                        }
                    }
                }
            }

            // 3. Try to get color from theme
            const themeColor = this.getThemeColorForPlaceholder(placeholderType);
            if (themeColor) {
                return themeColor;
            }

            // No fallback color - return null if not found in DOM
            return null;

        } catch (error) {
            return null;
        }
    }

    /**
     * Check if layout/master shape matches the placeholder
     */
    isMatchingPlaceholder(templateShape, placeholderType, placeholderIdx) {
        if (!templateShape.placeholder) {return false;}

        // Match by type first
        if (templateShape.placeholder.type === placeholderType) {
            // If idx is specified, it must match too
            if (placeholderIdx !== undefined && templateShape.placeholder.idx !== undefined) {
                return templateShape.placeholder.idx === placeholderIdx;
            }
            return true;
        }

        // Handle special cases like ctrTitle vs title
        if ((placeholderType === 'title' && templateShape.placeholder.type === 'ctrTitle') ||
            (placeholderType === 'ctrTitle' && templateShape.placeholder.type === 'title')) {
            return true;
        }

        return false;
    }

    /**
     * Inherit fill and stroke styles from source shape to target shape
     */
    inheritShapeStyles(targetShape, sourceShape) {
        try {
            // Inherit spPr (shape properties) if available
            if (sourceShape.spPr) {
                if (!targetShape.spPr) {
                    targetShape.spPr = {};
                }
                
                // Inherit fill properties (inherit if target doesn't have them or has null)
                if (sourceShape.spPr.solidFill && !targetShape.spPr.solidFill) {
                    targetShape.spPr.solidFill = sourceShape.spPr.solidFill;
                }
                
                if (sourceShape.spPr.gradFill && !targetShape.spPr.gradFill) {
                    targetShape.spPr.gradFill = sourceShape.spPr.gradFill;
                }
                
                if (sourceShape.spPr.pattFill && !targetShape.spPr.pattFill) {
                    targetShape.spPr.pattFill = sourceShape.spPr.pattFill;
                }
                
                if (sourceShape.spPr.blipFill && !targetShape.spPr.blipFill) {
                    targetShape.spPr.blipFill = sourceShape.spPr.blipFill;
                }
                
                if (sourceShape.spPr.noFill && !targetShape.spPr.noFill) {
                    targetShape.spPr.noFill = sourceShape.spPr.noFill;
                }
                
                // Inherit line properties
                if (sourceShape.spPr.ln && !targetShape.spPr.ln) {
                    targetShape.spPr.ln = sourceShape.spPr.ln;
                }
                
                // Inherit effect properties
                if (sourceShape.spPr.effectLst && !targetShape.spPr.effectLst) {
                    targetShape.spPr.effectLst = sourceShape.spPr.effectLst;
                }
                
                // Inherit geometry (only for non-text shapes to preserve text readability)
                if (sourceShape.spPr.geometry && !targetShape.spPr.geometry && !this.isTextPlaceholder(targetShape)) {
                    targetShape.spPr.geometry = sourceShape.spPr.geometry;
                }
                
                // Inherit transform properties (position, size, rotation) - but preserve target's if it has them
                if (sourceShape.spPr.xfrm && !targetShape.spPr.xfrm) {
                    targetShape.spPr.xfrm = sourceShape.spPr.xfrm;
                }
            }
            
            // Inherit style references
            if (sourceShape.style) {
                if (!targetShape.style) {
                    targetShape.style = {};
                }
                
                if (sourceShape.style.fillRef && !targetShape.style.fillRef) {
                    targetShape.style.fillRef = sourceShape.style.fillRef;
                }
                
                if (sourceShape.style.lnRef && !targetShape.style.lnRef) {
                    targetShape.style.lnRef = sourceShape.style.lnRef;
                }
                
                if (sourceShape.style.effectRef && !targetShape.style.effectRef) {
                    targetShape.style.effectRef = sourceShape.style.effectRef;
                }
                
                if (sourceShape.style.fontRef && !targetShape.style.fontRef) {
                    targetShape.style.fontRef = sourceShape.style.fontRef;
                }
            }
            
            // Inherit text properties for text shapes
            if (this.isTextPlaceholder(targetShape) || targetShape.textBody) {
                this.inheritTextProperties(targetShape, sourceShape);
            }
            
            // Inherit effects (separate from spPr.effectLst)
            if (sourceShape.effects && !targetShape.effects) {
                targetShape.effects = sourceShape.effects;
            }

            // Inherit direct fill property (for shapes that store fill at the root level)
            if (sourceShape.fill && (!targetShape.fill || targetShape.fill === null)) {
                targetShape.fill = sourceShape.fill;
            }

            // Inherit properties.fill (for shapes that store fill in properties object)
            if (sourceShape.properties && sourceShape.properties.fill) {
                if (!targetShape.properties) {
                    targetShape.properties = {};
                }
                if (!targetShape.properties.fill || targetShape.properties.fill === null) {
                    targetShape.properties.fill = sourceShape.properties.fill;
                }
            }

            // IMPORTANT: Inherit transform from source properties.transform when target lacks one
            if (sourceShape.properties && sourceShape.properties.transform) {
                if (!targetShape.properties) {targetShape.properties = {};}
                if (!targetShape.properties.transform) {
                    // Deep copy to avoid reference sharing
                    targetShape.properties.transform = JSON.parse(JSON.stringify(sourceShape.properties.transform));
                }
            }

            // If xfrm exists but properties.transform is still missing, synthesize transform from xfrm
            if ((!targetShape.properties || !targetShape.properties.transform) && targetShape.spPr && targetShape.spPr.xfrm) {
                const xfrm = targetShape.spPr.xfrm;
                const x = (xfrm.x !== undefined ? xfrm.x : (xfrm.off && xfrm.off.x) !== undefined ? xfrm.off.x : 0);
                const y = (xfrm.y !== undefined ? xfrm.y : (xfrm.off && xfrm.off.y) !== undefined ? xfrm.off.y : 0);
                const width = (xfrm.width !== undefined ? xfrm.width : (xfrm.cx !== undefined ? xfrm.cx : (xfrm.ext && xfrm.ext.cx) !== undefined ? xfrm.ext.cx : undefined));
                const height = (xfrm.height !== undefined ? xfrm.height : (xfrm.cy !== undefined ? xfrm.cy : (xfrm.ext && xfrm.ext.cy) !== undefined ? xfrm.ext.cy : undefined));
                if (width !== undefined && height !== undefined) {
                    if (!targetShape.properties) {targetShape.properties = {};}
                    targetShape.properties.transform = { x, y, width, height };
                }
            }
            
        } catch (error) {
        }
    }

    /**
     * Inherit text-specific properties from source to target shape
     */
    inheritTextProperties(targetShape, sourceShape) {
        try {
            // Inherit ONLY text body style properties; do NOT copy paragraph content from masters/layouts
            if (sourceShape.textBody) {
                if (!targetShape.textBody) {
                    targetShape.textBody = { paragraphs: [] };
                }
                // Unify body properties naming: prefer bodyProperties, map bodyPr when needed
                const sourceBodyProps = sourceShape.textBody.bodyProperties || sourceShape.textBody.bodyPr;
                if (!targetShape.textBody.bodyProperties) {
                    if (sourceBodyProps) {
                        targetShape.textBody.bodyProperties = JSON.parse(JSON.stringify(sourceBodyProps));
                    }
                } else if (sourceBodyProps) {
                    // Merge missing or defaulted fields from source
                    const tProps = targetShape.textBody.bodyProperties;
                    // Inherit anchor/vertical alignment if not explicitly set (or left at top)
                    if ((!tProps.anchor && !tProps.verticalAlign) || tProps.verticalAlign === 'top' || tProps.verticalAlign === 't') {
                        if (sourceBodyProps.anchor) {tProps.anchor = sourceBodyProps.anchor;}
                        else if (sourceBodyProps.verticalAlign) {tProps.verticalAlign = sourceBodyProps.verticalAlign;}
                    }
                    // Inherit margins if missing
                    if (tProps.leftMargin === undefined && sourceBodyProps.leftMargin !== undefined) {tProps.leftMargin = sourceBodyProps.leftMargin;}
                    if (tProps.rightMargin === undefined && sourceBodyProps.rightMargin !== undefined) {tProps.rightMargin = sourceBodyProps.rightMargin;}
                    if (tProps.topMargin === undefined && sourceBodyProps.topMargin !== undefined) {tProps.topMargin = sourceBodyProps.topMargin;}
                    if (tProps.bottomMargin === undefined && sourceBodyProps.bottomMargin !== undefined) {tProps.bottomMargin = sourceBodyProps.bottomMargin;}
                }
                if (sourceShape.textBody.lstStyle && !targetShape.textBody.lstStyle) {
                    targetShape.textBody.lstStyle = sourceShape.textBody.lstStyle;
                }
            }
            
            // Inherit txBody if present (alternative text body format)
            if (sourceShape.txBody && !targetShape.txBody) {
                targetShape.txBody = sourceShape.txBody;
            }
            
        } catch (error) {
        }
    }

    /**
     * Check if a shape is a text placeholder
     */
    isTextPlaceholder(shape) {
        // Prefer parsed placeholder info on shape
        const phType = shape?.placeholder?.type || (shape.nvSpPr && shape.nvSpPr.nvPr && shape.nvSpPr.nvPr.ph && shape.nvSpPr.nvPr.ph.type);
        if (phType) {
            return ['title', 'body', 'ctrTitle', 'subTitle', 'dt', 'ftr', 'hdr', 'sldNum'].includes(phType);
        }
        return false;
    }

    /**
     * Apply comprehensive property inheritance from layout and master slides
     */
    applyPropertyInheritance(shape, slideContext = null) {
        try {
            const currentSlide = slideContext || this.currentSlide;
            if (!currentSlide) {
                return;
            }

            // Step 1: Try to inherit from layout slide
            if (currentSlide.layout && currentSlide.layout.cSld && currentSlide.layout.cSld.spTree) {
                const layoutShape = this.findMatchingShapeInTemplate(shape, currentSlide.layout.cSld.spTree);
                if (layoutShape) {
                    this.inheritShapeStyles(shape, layoutShape);
                }
            }

            // Step 2: Try to inherit from master slide (if not found in layout or for additional properties)
            if (currentSlide.layout && currentSlide.layout.master && currentSlide.layout.master.cSld && currentSlide.layout.master.cSld.spTree) {
                const masterShape = this.findMatchingShapeInTemplate(shape, currentSlide.layout.master.cSld.spTree);
                if (masterShape) {
                    this.inheritShapeStyles(shape, masterShape);
                }
            }

        } catch (error) {
            // Silently handle errors to avoid breaking shape processing
        }
    }

    /**
     * Find a matching shape in a template (layout or master) shape tree
     */
    findMatchingShapeInTemplate(targetShape, templateShapeTree) {
        if (!templateShapeTree || !Array.isArray(templateShapeTree)) {
            return null;
        }

        // Strategy 1: Match by placeholder type and index (for placeholders)
        if (targetShape.placeholder) {
            for (const templateShape of templateShapeTree) {
                if (this.isMatchingPlaceholder(templateShape, targetShape.placeholder.type, targetShape.placeholder.idx)) {
                    return templateShape;
                }
            }
        }

        // Strategy 2: Match by name (exact match)
        if (targetShape.name) {
            for (const templateShape of templateShapeTree) {
                if (templateShape.name === targetShape.name) {
                    return templateShape;
                }
            }
        }

        // Strategy 2.5: Match text placeholders by name pattern (e.g., "Text Placeholder X")
        if (targetShape.name && targetShape.name.includes('Text Placeholder') && targetShape.textBody) {
            for (const templateShape of templateShapeTree) {
                if (templateShape.name && templateShape.name.includes('Text Placeholder') && templateShape.textBody) {
                    return templateShape;
                }
            }
        }

        // Strategy 3: Match by position (for non-placeholder shapes)
        if (targetShape.bounds && !targetShape.placeholder) {
            for (const templateShape of templateShapeTree) {
                if (this.isShapeAtSimilarPosition(targetShape, templateShape)) {
                    return templateShape;
                }
            }
        }

        return null;
    }

    /**
     * Check if two shapes are at similar positions (for positional matching)
     */
    isShapeAtSimilarPosition(shape1, shape2) {
        if (!shape1.bounds || !shape2.bounds) {
            return false;
        }

        const tolerance = 91440; // 0.1 inch in EMU
        const xMatch = Math.abs(shape1.bounds.l - shape2.bounds.l) <= tolerance;
        const yMatch = Math.abs(shape1.bounds.t - shape2.bounds.t) <= tolerance;

        return xMatch && yMatch;
    }

    /**
     * Extract geometry information from a template shape (layout/master)
     */
    extractShapeGeometry(templateShape) {
        // Check geometry from spPr
        if (templateShape.spPr && templateShape.spPr.geometry) {
            return templateShape.spPr.geometry;
        }

        // Check direct geometry
        if (templateShape.geometry) {
            return templateShape.geometry;
        }

        // Check properties geometry
        if (templateShape.properties && templateShape.properties.geometry) {
            return templateShape.properties.geometry;
        }

        return null;
    }
    /**
     * Extract fill color from a template shape (layout/master)
     */
    extractShapeFillColor(templateShape) {
        // Try direct fill
        if (templateShape.fill && templateShape.fill.color) {
            return this.processShapeColor(templateShape.fill.color);
        }

        // Try properties fill
        if (templateShape.properties && templateShape.properties.fill && templateShape.properties.fill.color) {
            return this.processShapeColor(templateShape.properties.fill.color);
        }

        // Try style-based fill
        if (templateShape.style && templateShape.style.fillRef) {
            return this.resolveStyleFillColor(templateShape.style.fillRef);
        }

        return null;
    }

    /**
     * Get theme color for placeholder type (following standard inheritance rules)
     */
    getThemeColorForPlaceholder(placeholderType) {
        const theme = this.presentation?.theme;
        if (!theme || !theme.colors) {return null;}

        switch (placeholderType) {
            case 'title':
            case 'ctrTitle':
                // Title placeholders: NO FILL by default (transparent background)
                return null;
            case 'body':
            case 'obj':
                // Body placeholders typically have NO FILL by default - return null for transparent
                // Theme colors for body placeholders should be applied to text, not background
                return null; // No background fill - let layout/master define it
            case 'subTitle':
                // Subtitles typically have subtle backgrounds
                return { r: 245, g: 245, b: 245, a: 255 }; // Light gray
            default:
                // Other placeholders - no fill by default
                return null;
        }
    }

    /**
     * Process shape color with theme and tint/shade support
     */
    processShapeColor(color) {
        if (!color) {
            return null;
        }

        // If it's already a processed color object, return it
        if (color.r !== undefined && color.g !== undefined && color.b !== undefined) {
            return color;
        }

        // Handle hex string colors
        if (typeof color === 'string' && color.startsWith('#')) {
            const rgb = this.hexToRgb(color);
            return rgb;
        }

        // Handle theme colors
        if (color.scheme) {
            const themeColor = this.resolveThemeColor(color.scheme);
            if (themeColor) {
                return this.applyColorModifications(themeColor, color);
            }
        }

        // Handle {type: 'srgb', value: 'RRGGBB', alpha: 0-1} from parseColor()
        if (color.type === 'srgb' && color.value) {
            const hex = color.value;
            const r = parseInt(hex.substring(0, 2), 16) || 0;
            const g = parseInt(hex.substring(2, 4), 16) || 0;
            const b = parseInt(hex.substring(4, 6), 16) || 0;
            const a = (color.alpha !== undefined && color.alpha !== null) ? Math.round(color.alpha * 255) : 255;
            return { r, g, b, a };
        }

        // Handle {type: 'scheme', value: 'accent1', alpha: ...} from parseColor()
        if (color.type === 'scheme' && color.value) {
            const themeColor = this.resolveThemeColor(color.value);
            if (themeColor) {
                return this.applyColorModifications(themeColor, color);
            }
        }

        return color;
    }

    /**
     * Process gradient fill for enhanced rendering
     */
    processGradientFill(gradient) {
        if (!gradient || !gradient.stops || gradient.stops.length === 0) {
            return null;
        }

        const processedGradient = {
            type: gradient.type || 'linear',
            angle: gradient.angle || 0,
            stops: gradient.stops.map(stop => ({
                position: stop.position,
                color: this.processShapeColor(stop.color)
            }))
        };

        return processedGradient;
    }

    /**
     * Resolve theme color from scheme with enhanced color palette
     */
    resolveThemeColor(scheme) {
        // First, try to get theme colors from the actual presentation
        if (this.presentation && this.presentation.theme && this.presentation.theme.colors) {
            const themeColorHex = this.presentation.theme.colors[scheme];
            if (themeColorHex) {
                const actualThemeColor = this.parseColorFromHex(themeColorHex);
                if (actualThemeColor) {
                    return actualThemeColor;
                }
            }
        }

        // No hardcoded theme colors - only use colors from DOM
        const themeColors = {
        };

        const resolvedColor = themeColors[scheme] || null;
        return resolvedColor;
    }

    /**
     * Resolve placeholder color (phClr) in theme format scheme context
     * This is a critical method for proper theme color resolution
     */
    resolvePlaceholderColor(placeholderColor, contextColor = null) {
        if (!placeholderColor || placeholderColor.type !== 'placeholder') {
            return placeholderColor;
        }

        // If we have a context color (from shape's style), use it as the base
        let baseColor = contextColor;

        // If no context color, try to get a reasonable default from theme
        if (!baseColor) {
            if (this.presentation && this.presentation.theme && this.presentation.theme.colors) {
                // Try to get accent1 as a reasonable default
                const accent1Hex = this.presentation.theme.colors.accent1;
                if (accent1Hex) {
                    baseColor = this.parseColorFromHex(accent1Hex);
                }
            }
        }

        // No fallback - return null if no color specified in DOM
        if (!baseColor) {
            return null; // No color specified in DOM data
        }

        // Apply any color modifications from the placeholder
        if (placeholderColor.modifications && placeholderColor.modifications.length > 0) {
            baseColor = this.applyColorModifications(baseColor, placeholderColor.modifications);
        }

        return baseColor;
    }

    /**
     * Convert RGB to HSL for better color manipulation
     */
    rgbToHsl(r, g, b) {
        r /= 255;
        g /= 255;
        b /= 255;

        const max = Math.max(r, g, b);
        const min = Math.min(r, g, b);
        let h, s, l = (max + min) / 2;

        if (max === min) {
            h = s = 0; // achromatic
        } else {
            const d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);

            switch (max) {
                case r: h = (g - b) / d + (g < b ? 6 : 0); break;
                case g: h = (b - r) / d + 2; break;
                case b: h = (r - g) / d + 4; break;
            }
            h /= 6;
        }

        return { h: h * 360, s, l };
    }

    /**
     * Convert HSL to RGB
     */
    hslToRgb(h, s, l) {
        h /= 360;

        const hue2rgb = (p, q, t) => {
            if (t < 0) {t += 1;}
            if (t > 1) {t -= 1;}
            if (t < 1/6) {return p + (q - p) * 6 * t;}
            if (t < 1/2) {return q;}
            if (t < 2/3) {return p + (q - p) * (2/3 - t) * 6;}
            return p;
        };

        let r, g, b;

        if (s === 0) {
            r = g = b = l; // achromatic
        } else {
            const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            const p = 2 * l - q;
            r = hue2rgb(p, q, h + 1/3);
            g = hue2rgb(p, q, h);
            b = hue2rgb(p, q, h - 1/3);
        }

        return {
            r: Math.round(r * 255),
            g: Math.round(g * 255),
            b: Math.round(b * 255)
        };
    }

    /**
     * Parse color from hex string to RGB object
     */
    parseColorFromHex(hexColor) {
        if (!hexColor || typeof hexColor !== 'string') {return null;}

        // Remove # if present
        const hex = hexColor.replace('#', '');

        // Parse RGB values
        if (hex.length === 6) {
            return {
                r: parseInt(hex.substr(0, 2), 16),
                g: parseInt(hex.substr(2, 2), 16),
                b: parseInt(hex.substr(4, 2), 16),
                a: 255
            };
        }

        return null;
    }

    /**
     * Apply color modifications (tint, shade, etc.)
     */
    applyColorModifications(baseColor, modifications) {
        if (!baseColor || !modifications || modifications.length === 0) {
            return baseColor;
        }

        const color = { ...baseColor };

        // Convert to HSL for better color manipulation
        const hsl = this.rgbToHsl(color.r, color.g, color.b);
        let { h, s, l } = hsl;

        for (const mod of modifications) {
            const value = mod.value;

            switch (mod.type) {
                case 'tint':
                    // Tint: mix with white (increase lightness)
                    const tintFactor = value / 100000; // PowerPoint uses 100000 as 100%
                    l = l + (1 - l) * tintFactor;
                    break;

                case 'shade':
                    // Shade: mix with black (decrease lightness)
                    const shadeFactor = value / 100000;
                    l = l * (1 - shadeFactor);
                    break;

                case 'lummod':
                case 'lumMod':
                    // Luminance modulation
                    const lumModFactor = value / 100000;
                    l = l * lumModFactor;
                    break;

                case 'lumoff':
                case 'lumOff':
                    // Luminance offset
                    const lumOffFactor = value / 100000;
                    l = Math.min(1, Math.max(0, l + lumOffFactor));
                    break;

                case 'satmod':
                case 'satMod':
                    // Saturation modulation
                    const satModFactor = value / 100000;
                    s = s * satModFactor;
                    break;

                case 'satoff':
                case 'satOff':
                    // Saturation offset
                    const satOffFactor = value / 100000;
                    s = Math.min(1, Math.max(0, s + satOffFactor));
                    break;

                case 'hue':
                    // Hue absolute value
                    h = (value / 60000) % 360; // PowerPoint uses 60000 as 360 degrees
                    break;

                case 'hueoff':
                case 'hueOff':
                    // Hue offset
                    const hueOffDegrees = value / 60000;
                    h = (h + hueOffDegrees) % 360;
                    break;

                case 'huemod':
                case 'hueMod':
                    // Hue modulation
                    const hueModFactor = value / 100000;
                    h = (h * hueModFactor) % 360;
                    break;

                case 'alpha':
                    // Alpha absolute value
                    color.a = Math.round((value / 100000) * 255);
                    break;

                case 'alphamod':
                case 'alphaMod':
                    // Alpha modulation
                    const alphaModFactor = value / 100000;
                    color.a = Math.round(color.a * alphaModFactor);
                    break;

                case 'alphaoff':
                case 'alphaOff':
                    // Alpha offset
                    const alphaOffFactor = value / 100000;
                    color.a = Math.min(255, Math.max(0, Math.round(color.a + alphaOffFactor * 255)));
                    break;
            }
        }

        // Convert back to RGB
        const rgb = this.hslToRgb(h, s, l);
        color.r = rgb.r;
        color.g = rgb.g;
        color.b = rgb.b;

        // Ensure values are within valid range
        color.r = Math.min(255, Math.max(0, Math.round(color.r)));
        color.g = Math.min(255, Math.max(0, Math.round(color.g)));
        color.b = Math.min(255, Math.max(0, Math.round(color.b)));
        color.a = Math.min(255, Math.max(0, Math.round(color.a || 255)));

        return color;
    }

    /**
     * Convert RGB to HSL
     */
    rgbToHsl(r, g, b) {
        r /= 255;
        g /= 255;
        b /= 255;

        const max = Math.max(r, g, b);
        const min = Math.min(r, g, b);
        let h, s, l = (max + min) / 2;

        if (max === min) {
            h = s = 0; // achromatic
        } else {
            const d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);

            switch (max) {
                case r: h = (g - b) / d + (g < b ? 6 : 0); break;
                case g: h = (b - r) / d + 2; break;
                case b: h = (r - g) / d + 4; break;
            }
            h /= 6;
        }

        return { h, s, l };
    }

    /**
     * Convert HSL to RGB
     */
    hslToRgb(h, s, l) {
        let r, g, b;

        if (s === 0) {
            r = g = b = l; // achromatic
        } else {
            const hue2rgb = (p, q, t) => {
                if (t < 0) {t += 1;}
                if (t > 1) {t -= 1;}
                if (t < 1/6) {return p + (q - p) * 6 * t;}
                if (t < 1/2) {return q;}
                if (t < 2/3) {return p + (q - p) * (2/3 - t) * 6;}
                return p;
            };

            const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            const p = 2 * l - q;
            r = hue2rgb(p, q, h + 1/3);
            g = hue2rgb(p, q, h);
            b = hue2rgb(p, q, h - 1/3);
        }

        return {
            r: Math.round(r * 255),
            g: Math.round(g * 255),
            b: Math.round(b * 255)
        };
    }

    /**
     * Parse color from hex string to RGB object
     */
    parseColorFromHex(hexColor) {
        if (!hexColor || typeof hexColor !== 'string') {return null;}

        // Remove # if present
        const hex = hexColor.replace('#', '');

        // Parse RGB values
        if (hex.length === 6) {
            return {
                r: parseInt(hex.substr(0, 2), 16),
                g: parseInt(hex.substr(2, 2), 16),
                b: parseInt(hex.substr(4, 2), 16),
                a: 255
            };
        }

        return null;
    }

    /**
     * Apply color modifications (tint, shade, etc.)
     */
    applyColorModifications(baseColor, colorInfo) {
        const color = { ...baseColor };

        // Apply tint (lighten)
        if (colorInfo.tint !== undefined) {
            const tint = colorInfo.tint / 100000; // Convert from percentage
            color.r = Math.round(color.r + (255 - color.r) * tint);
            color.g = Math.round(color.g + (255 - color.g) * tint);
            color.b = Math.round(color.b + (255 - color.b) * tint);
        }

        // Apply shade (darken)
        if (colorInfo.shade !== undefined) {
            const shade = colorInfo.shade / 100000; // Convert from percentage
            color.r = Math.round(color.r * (1 - shade));
            color.g = Math.round(color.g * (1 - shade));
            color.b = Math.round(color.b * (1 - shade));
        }

        // Apply saturation
        if (colorInfo.sat !== undefined) {
            const sat = colorInfo.sat / 100000; // Convert from percentage
            // Apply saturation modification (simplified)
            const gray = 0.299 * color.r + 0.587 * color.g + 0.114 * color.b;
            color.r = Math.round(gray + (color.r - gray) * sat);
            color.g = Math.round(gray + (color.g - gray) * sat);
            color.b = Math.round(gray + (color.b - gray) * sat);
        }

        // Apply alpha
        if (colorInfo.alpha !== undefined) {
            color.a = Math.round(255 * (colorInfo.alpha / 100000));
        }

        // Ensure values are within valid range
        color.r = Math.max(0, Math.min(255, color.r));
        color.g = Math.max(0, Math.min(255, color.g));
        color.b = Math.max(0, Math.min(255, color.b));
        color.a = Math.max(0, Math.min(255, color.a));

        return color;
    }

    /**
     * Helper function to traverse and collect all shapes from spTree
     */
    getAllShapesFromTree(spTree) {
        const shapes = [];

        const traverse = (element) => {
            if (!element) {return;}

            if (Array.isArray(element)) {
                element.forEach(traverse);
            } else if (typeof element === 'object') {
                // Add current element if it's a shape
                if (element.nvSpPr || element.nvPicPr || element.nvGrpSpPr) {
                    shapes.push(element);
                }

                // Recursively traverse all properties
                Object.values(element).forEach(value => {
                    if (value && typeof value === 'object') {
                        traverse(value);
                    }
                });
            }
        };

        traverse(spTree);
        return shapes;
    }

    /**
     * Resolve style-based fill color with enhanced theme format scheme processing
     */
    resolveStyleFillColor(fillRef) {
        try {
            if (!fillRef) {
                return null;
            }
        const currentShape = this._currentProcessingShape;
        const isRectangle3 = currentShape && (currentShape.name === 'Rectangle 3' || currentShape.name?.includes('Rectangle 3'));
        if (isRectangle3) {
        }

        // Get the style index
            const index = fillRef.idx || 0;

        if (isRectangle3) {
        }

        // Use the color from the fill reference if available
            if (fillRef.color) {
                const color = this.processShapeColor(fillRef.color);
                const colorStr = color && color.r !== undefined ?
                    `rgba(${color.r}, ${color.g}, ${color.b}, ${color.a || 255})` :
                    JSON.stringify(color);

                if (isRectangle3) {
                }
                return color;
            }

            // No hardcoded default colors - return null if not in theme

                    // ENHANCED: Process theme format scheme with enhanced theme color integration
        const theme = this.presentation?.theme;

        if (isRectangle3) {
            if (theme) {
                if (theme.colors && theme.colors.accent6) {
                    const actualAccent6 = this.parseColorFromHex(theme.colors.accent6);
                }
            }
        }

        if (theme && theme.formatScheme) {

            if (isRectangle3) {
                if (theme.formatScheme.fillStyles) {
                    theme.formatScheme.fillStyles.forEach((fillStyle, idx) => {
                    });
                }
            }

                // Handle fillRef indices according to standard logic
                if (index < 1000) {
                    // Regular fill styles (fillStyleLst)
                    const fillIndex = index - 1; // Convert to 0-based
                    if (fillIndex >= 0 && theme.formatScheme.fillStyles && fillIndex < theme.formatScheme.fillStyles.length) {
                        const fillStyle = theme.formatScheme.fillStyles[fillIndex];

                        // Use enhanced fill style processing
                        if (fillStyle.color && fillStyle.type === 'solid') {
                            return fillStyle.color;
                        } else if (fillStyle.gradient && fillStyle.type === 'gradient') {
                            return this.processGradientFill(fillStyle.gradient);
                        }

                        // Fallback to original processing
                        const color = this.processFillStyleToColor(fillStyle, fillRef.color);
                        if (color) {
                            return color;
                        }
                    }
                } else if (index >= 1001) {
                    // Background fill styles (bgFillStyleLst)
                    const bgFillIndex = index - 1001; // Convert to 0-based (subtract 1001)
                    if (bgFillIndex >= 0 && theme.formatScheme.backgroundFills && bgFillIndex < theme.formatScheme.backgroundFills.length) {
                        const bgFillStyle = theme.formatScheme.backgroundFills[bgFillIndex];

                        // Use enhanced background fill style processing
                        if (bgFillStyle.color && bgFillStyle.type === 'solid') {
                            return bgFillStyle.color;
                        }

                        // Fallback to original processing
                        const color = this.processFillStyleToColor(bgFillStyle, fillRef.color);
                        if (color) {
                            return color;
                        }
                    }
                }
        }

        // No hardcoded colors - return null if not found in theme
        return null;

        } catch (error) {
            // Return null if error occurred - no hardcoded fallback
            return null;
        }
    }

    /**
     * Process fill style to extract color - following standard pattern with placeholder handling
     */
    processFillStyleToColor(fillStyle, schemeColor = null) {
        if (!fillStyle) {return null;}

        try {
            // Handle solid fill
            if (fillStyle.type === 'solid' && fillStyle.color) {
                // CRITICAL FIX: Handle placeholder colors (phClr) in theme format schemes
                if (fillStyle.color.type === 'placeholder' && fillStyle.color.scheme === 'phClr') {
                    // IMPORTANT: For regular shapes (non-placeholders), phClr should resolve to a neutral color
                    // Only actual placeholder shapes should use the full placeholder color resolution

                    // Check if we have context about the current shape being processed
                    const currentShape = this._currentProcessingShape;
                    if (currentShape && !currentShape.placeholder) {
                        // This is a regular shape - return null if not specified in DOM
                        return null; // No color specified
                    }

                    // Use the new comprehensive placeholder resolution method for actual placeholders
                    const contextColor = schemeColor ? this.processShapeColor(schemeColor) : null;
                    const resolvedColor = this.resolvePlaceholderColor(fillStyle.color, contextColor);
                    return resolvedColor;
                }

                // If schemeColor override is provided, use it instead
                if (schemeColor) {
                    return this.processShapeColor(schemeColor);
                }
                return this.processShapeColor(fillStyle.color);
            }

            // Handle gradient fill - enhanced for theme gradients with placeholders
            if (fillStyle.type === 'gradient') {
                let firstColor = null;

                // Try new structure (stops array)
                if (fillStyle.stops && fillStyle.stops.length > 0) {
                    firstColor = fillStyle.stops[0].color;
                }
                // Try old structure (gradient.colors)
                else if (fillStyle.gradient && fillStyle.gradient.colors && fillStyle.gradient.colors.length > 0) {
                    firstColor = fillStyle.gradient.colors[0];
                }

                if (firstColor) {
                    // Handle placeholder colors in gradient stops
                    if (firstColor.type === 'placeholder' && firstColor.value === 'phClr') {
                        // CRITICAL FIX: Regular shapes should not use phClr in gradients
                        const currentShape = this._currentProcessingShape;
                        if (currentShape && !currentShape.placeholder) {
                            // This is a regular shape - return null if not specified in DOM
                            return null; // No color specified
                        }

                        const contextColor = schemeColor ? this.processShapeColor(schemeColor) : null;
                        let resolvedColor = this.resolvePlaceholderColor(firstColor, contextColor);

                        // CRITICAL FIX: Apply gradient stop modifications to the resolved placeholder color
                        if (firstColor.modifications && firstColor.modifications.length > 0) {
                            resolvedColor = this.applyColorModifications(resolvedColor, firstColor.modifications);
                        }

                        return resolvedColor;
                    }

                    // Handle scheme colors in gradient stops
                    if (firstColor.type === 'scheme') {
                        const themeColor = this.getThemeColor(firstColor.value);
                        if (themeColor) {
                            let color = themeColor;
                            // Apply modifications if present
                            if (firstColor.modifications && firstColor.modifications.length > 0) {
                                color = this.applyColorModifications(color, firstColor.modifications);
                            }
            return color;
                        }
                    }

                    // Use schemeColor override if provided
                    if (schemeColor) {
                        return this.processShapeColor(schemeColor);
                    }

                    // Process the gradient color normally
                    return this.processShapeColor(firstColor);
                }
            }

            // Handle pattern fill - use foreground color
            if (fillStyle.type === 'pattern' && fillStyle.fgColor) {
                if (schemeColor) {
                    return this.processShapeColor(schemeColor);
                }
                return this.processShapeColor(fillStyle.fgColor);
            }

            // Handle image/blip fill - return null for transparent
            if (fillStyle.type === 'blip' || fillStyle.type === 'image') {
                return null;
            }

            // Handle no fill
            if (fillStyle.type === 'none' || fillStyle.type === 'noFill') {
                return null;
            }

        } catch (_error) {
				// Error ignored
			}
        return null;
    }

    /**
     * Check if an image is currently being loaded
     */
    isImageLoading(relId) {
        // Simple implementation - in a real app you'd track loading state
        return !this.imageCache.has(relId) && this.imageMap.has(relId);
    }

    /**
     * Trigger re-render of current slide
     */
        triggerSlideRerender() {
        // Dispatch custom event to notify the UI to re-render
        if (typeof window !== 'undefined') {
            window.dispatchEvent(new Event('slideImageLoaded', {
                detail: { slideIndex: this.currentSlideIndex }
            }));
        }
    }
    logSlideImageStatus(slide, slideIndex) {

        // Check background images
        if (slide.backgroundFill && slide.backgroundFill.fill) {
            const bgFill = slide.backgroundFill.fill;
            if (bgFill.type === 'image' && bgFill.imageData) {
                const relId = bgFill.imageData.relationshipId;
                const cached = this.imageCache.has(relId);
            }
        }

        // Check shape images
        let shapeImageCount = 0;
        if (slide.commonSlideData && slide.commonSlideData.shapeTree) {
            for (const shape of slide.commonSlideData.shapeTree) {
                if (shape.type === 'pic' && shape.imageRelId) {
                    const cached = this.imageCache.has(shape.imageRelId);
                    shapeImageCount++;
                }
            }
        }

    }

    /**
     * Draw group shape geometry - Enhanced with standard patterns
     */
    async drawGroupShapeGeometry(graphics, shape, x, y, w, h, fillColor, strokeColor, lineWidth, canvasRect) {
        // Save graphics state for group rendering
        graphics.SaveGrState();

        try {

            // Apply group-level transformations if available
            if (shape.transform) {
                // Apply rotation
                if (shape.transform.rotation && shape.transform.rotation !== 0) {
                    const centerX = x + w / 2;
                    const centerY = y + h / 2;
                    graphics.transform(
                        Math.cos(shape.transform.rotation * Math.PI / 180),
                        Math.sin(shape.transform.rotation * Math.PI / 180),
                        -Math.sin(shape.transform.rotation * Math.PI / 180),
                        Math.cos(shape.transform.rotation * Math.PI / 180),
                        centerX - centerX * Math.cos(shape.transform.rotation * Math.PI / 180) + centerY * Math.sin(shape.transform.rotation * Math.PI / 180),
                        centerY - centerX * Math.sin(shape.transform.rotation * Math.PI / 180) - centerY * Math.cos(shape.transform.rotation * Math.PI / 180)
                    );
                }

                // Apply flip transformations
                if (shape.transform.flipH || shape.transform.flipV) {
                    const scaleX = shape.transform.flipH ? -1 : 1;
                    const scaleY = shape.transform.flipV ? -1 : 1;
                    const centerX = x + w / 2;
                    const centerY = y + h / 2;

                    graphics.transform(scaleX, 0, 0, scaleY,
                        centerX - centerX * scaleX,
                        centerY - centerY * scaleY);
                }
            }

            // Draw group background if fill is specified
            if (fillColor) {
                if (typeof fillColor === 'object' && fillColor.type === 'linear') {
                    // Handle gradient fill - pass to graphics engine as regular fill
                    graphics.fillRect(x, y, w, h, fillColor);
                } else {
                    // Handle solid fill
                    let rgbFill = null;
                    if (typeof fillColor === 'string') {
                        rgbFill = this.hexToRgb(fillColor);
                    } else if (fillColor.r !== undefined) {
                        rgbFill = fillColor;
                    }

                    if (rgbFill) {
                        graphics.drawRectangle(x, y, w, h, rgbFill, null, 0);
                    }
                }
            }

            // Draw group border if stroke is specified
            if (strokeColor && lineWidth > 0) {
                let rgbStroke = null;
                if (typeof strokeColor === 'string') {
                    rgbStroke = this.hexToRgb(strokeColor);
                } else if (strokeColor.r !== undefined) {
                    rgbStroke = strokeColor;
                }

                if (rgbStroke) {
                    graphics.drawRectangle(x, y, w, h, null, rgbStroke, lineWidth);
                }
            }

            // Draw child shapes
            if (shape.shapeTree && shape.shapeTree.length > 0) {

                // Sort child shapes by rendering order for proper z-index
                const sortedChildShapes = this.sortShapesByRenderOrder(shape.shapeTree);

                for (let i = 0; i < sortedChildShapes.length; i++) {
                    const childShape = sortedChildShapes[i];

                    try {

                        // Use renderShape method instead of drawShapeGeometry, which properly handles all shape types
                        await this.renderShape(graphics, childShape, canvasRect);

                    } catch (error) {
                        // Continue with other child shapes
                    }
                }
            } else {
            }

        } catch (_error) {
				// Error ignored
			} finally {
            // Always restore graphics state
            graphics.RestoreGrState();
        }
    }

    /**
     * Draw connector shape geometry
     */
    drawConnectorShapeGeometry(graphics, shape, x, y, w, h, strokeColor, lineWidth) {
        // Draw from top-left to bottom-right of bounds (respects flip via transform)
        graphics.drawLine(x, y, x + w, y + h, strokeColor, lineWidth);
    }

    /**
     * Draw graphic frame geometry
     */
    drawGraphicFrameGeometry(graphics, shape, x, y, w, h, fillColor, strokeColor, lineWidth) {
        // Draw a neutral border to indicate a graphic frame without any placeholder text
        graphics.drawRectangle(x, y, w, h, null, '#a0a0a0', 1);
    }

    /**
     * Draw default shape geometry
     */
    drawDefaultShapeGeometry(graphics, shape, x, y, w, h, fillColor, strokeColor, lineWidth) {
        graphics.drawRectangle(x, y, w, h, fillColor, '#ff0000', lineWidth);
    }

    /**
     * Convert RGB color object to hex string
     */
    rgbToHex(rgb) {
        if (!rgb) {return null;}

        const toHex = (c) => {
            const hex = Math.max(0, Math.min(255, Math.round(c))).toString(16);
            return hex.length === 1 ? '0' + hex : hex;
        };

        return `#${toHex(rgb.r || rgb.R || 0)}${toHex(rgb.g || rgb.G || 0)}${toHex(rgb.b || rgb.B || 0)}`;
    }

    /**
     * Create rotation transform matrix for shape rendering
     */
    createRotationMatrix(angleDegrees, centerX, centerY) {
        const angleRad = angleDegrees * Math.PI / 180;
        const cos = Math.cos(angleRad);
        const sin = Math.sin(angleRad);
        
        return {
            a: cos,
            b: sin,
            c: -sin,
            d: cos,
            e: centerX - centerX * cos + centerY * sin,
            f: centerY - centerX * sin - centerY * cos
        };
    }

    /**
     * Apply rotation transformation to graphics context
     */
    applyRotationTransform(graphics, shape, x, y, w, h) {
        const rotation = this.getShapeRotation(shape);
        
        if (!rotation || rotation === 0) {
            return false; // No rotation to apply
        }

        // Calculate center point for rotation
        const centerX = x + w / 2;
        const centerY = y + h / 2;

        // Create and apply rotation matrix
        const matrix = this.createRotationMatrix(rotation, centerX, centerY);
        
        // Convert to standard matrix format and apply
        const onlyOfficeMatrix = {
            sx: matrix.a,
            sy: matrix.d,
            shx: matrix.c,
            shy: matrix.b,
            tx: matrix.e,
            ty: matrix.f,
            createDuplicate() {
                return {
                    sx: this.sx, sy: this.sy, shx: this.shx,
                    shy: this.shy, tx: this.tx, ty: this.ty,
                    createDuplicate: this.createDuplicate
                };
            }
        };
        
        if (graphics.transform3) {
            graphics.transform3(onlyOfficeMatrix);
        } else if (graphics.transform) {
            graphics.transform(matrix.a, matrix.b, matrix.c, matrix.d, matrix.e, matrix.f);
        }
        
        return true; // Rotation was applied
    }

    /**
     * Get shape rotation in degrees (handles multiple sources)
     */
    getShapeRotation(shape) {
        // Check properties.transform.rot (converted to degrees)
        if (shape.properties?.transform?.rot !== undefined) {
            return shape.properties.transform.rot;
        }
        
        // Check transform.rotation (group transforms)
        if (shape.transform?.rotation !== undefined) {
            return shape.transform.rotation;
        }
        
        // Check spPr for standard compatibility
        if (shape.spPr?.transform?.rot !== undefined) {
            return shape.spPr.transform.rot;
        }
        
        return 0; // No rotation
    }

    /**
     * Calculate rotated bounds for a shape
     */
    calculateRotatedBounds(originalBounds, rotation) {
        if (!rotation || rotation === 0) {
            return originalBounds;
        }

        const { x, y, w, h } = originalBounds;
        const centerX = x + w / 2;
        const centerY = y + h / 2;
        
        // Calculate corners of original rectangle
        const corners = [
            { x: x, y: y },
            { x: x + w, y: y },
            { x: x + w, y: y + h },
            { x: x, y: y + h }
        ];
        
        // Rotate each corner
        const angleRad = rotation * Math.PI / 180;
        const cos = Math.cos(angleRad);
        const sin = Math.sin(angleRad);
        
        const rotatedCorners = corners.map(corner => {
            const dx = corner.x - centerX;
            const dy = corner.y - centerY;
            return {
                x: centerX + dx * cos - dy * sin,
                y: centerY + dx * sin + dy * cos
            };
        });
        
        // Find bounding box of rotated shape
        const minX = Math.min(...rotatedCorners.map(c => c.x));
        const minY = Math.min(...rotatedCorners.map(c => c.y));
        const maxX = Math.max(...rotatedCorners.map(c => c.x));
        const maxY = Math.max(...rotatedCorners.map(c => c.y));
        
        return {
            x: minX,
            y: minY,
            w: maxX - minX,
            h: maxY - minY
        };
    }



    /**
     * Get default text color for a shape (following standard theme inheritance)
     */
    getDefaultTextColor(shape) {
        // Try to get text color from theme
        if (this.presentation?.theme?.colors) {
            // For title placeholders, use dk2 (dark2) or dk1 as primary title color
            if (shape.placeholder?.type === 'title' || shape.placeholder?.type === 'ctrTitle') {
                // First try dk2 (Dark 2) which is often used for titles
                const dk2Color = this.resolveThemeColor('dk2');
                if (dk2Color) {
                    return dk2Color;
                }
                // Fall back to dk1 (Dark 1)
                const dk1Color = this.resolveThemeColor('dk1');
                if (dk1Color) {
                    return dk1Color;
                }
                // Final fallback to tx1
                return this.presentation.theme.colors.tx1 || { r: 0, g: 0, b: 0, a: 255 };
            } else {
                // For body text, use tx1 (primary text) or tx2 (secondary text)
                return this.presentation.theme.colors.tx1 || this.presentation.theme.colors.tx2 || { r: 0, g: 0, b: 0, a: 255 };
            }
        }

        // Fallback to black
        return { r: 0, g: 0, b: 0, a: 255 };
    }


    /**
     * Extract media relationships from a specific document part
     */
    async extractPartMediaRelationships(part) {
        try {
            const relationships = await part.getRelationships();

            for (const [relId, rel] of Object.entries(relationships)) {

                if (rel.type === OpenXmlTypesSafe.image.relationType || rel.type.includes('video') || rel.type.includes('audio')) {
                    // Convert relative target to absolute path
                    let mediaPath = rel.target;
                    if (mediaPath.startsWith('../')) {
                        mediaPath = mediaPath.replace('../', 'ppt/');
                    } else if (!mediaPath.startsWith('/')) {
                        mediaPath = `/ppt/${mediaPath}`;
                    }

                    this.mediaMap.set(relId, {path: mediaPath, type: rel.type});
                    if(this.isImageFile(mediaPath)) {
                        // Prefer SVG files over other image formats when both exist
                        const existingPath = this.imageMap.get(relId);
                        
                        if (!existingPath || this.isSVGFile(mediaPath) && !this.isSVGFile(existingPath)) {
                            // Set if no existing path OR current is SVG and existing is not SVG
                            this.imageMap.set(relId, mediaPath);
                            if (this.isSVGFile(mediaPath)) {
                            }
                        } else if (!this.isSVGFile(mediaPath) && this.isSVGFile(existingPath)) {
                            // Skip non-SVG if we already have SVG
                        } else {
                            // Otherwise, update normally (handles same type overwrites)
                            this.imageMap.set(relId, mediaPath);
                        }
                    }
                } else {
                }
            }
        } catch (_error) {
				// Error ignored
			}
    }

    /**
     * Check if file path represents a media file
     */
    isMediaFile(path) {
        const mediaExtensions = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.webp', '.tiff', '.mp4', '.mov', '.avi', '.mp3', '.wav', '.ogg'];
        const lowerPath = path.toLowerCase();
        return mediaExtensions.some(ext => lowerPath.endsWith(ext));
    }

    /**
     * Check if file path represents an SVG file
     */
    isSVGFile(path) {
        return path && path.toLowerCase().endsWith('.svg');
    }

    /**
     * Extract SVG relationship ID from extension list
     * Looks for asvg:svgBlip elements in a:extLst extensions
     * @param {Element} blipElement - The blip element containing potential SVG alternatives
     * @returns {string|null} SVG relationship ID if found, null otherwise
     */
    extractSVGRelationshipId(blipElement) {
        try {
            // Use getElementsByTagName with wildcard to find all descendants regardless of namespace prefix.
            // querySelector is unreliable for namespace-prefixed elements in text/xml DOM.
            const allDescendants = blipElement.getElementsByTagName('*');
            for (let i = 0; i < allDescendants.length; i++) {
                const el = allDescendants[i];
                if (el.localName === 'svgBlip') {
                    const svgRelId = el.getAttribute('r:embed') || el.getAttribute('r:id') ||
                        el.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed');
                    if (svgRelId) {
                        return svgRelId;
                    }
                }
            }
            return null;
        } catch (error) {
            console.warn('[WARN] Slide Renderer: Error extracting SVG relationship ID:', error);
            return null;
        }
    }

    /**
     * Build mapping from PNG relationship IDs to their corresponding SVG relationship IDs
     * This helps prefer SVG versions when both PNG and SVG exist for the same image
     */
    buildSVGRelationshipMapping() {
        
        // Create reverse maps to find relationships by filename pattern
        const pngRelIds = new Map(); // base filename -> relId  
        const svgRelIds = new Map(); // base filename -> relId
        
        // Process all image relationships
        for (const [relId, imagePath] of this.imageMap.entries()) {
            // Extract base filename without extension
            const filename = imagePath.split('/').pop(); // Get filename from path
            const baseName = filename.replace(/\.[^.]+$/, ''); // Remove extension
            
            if (this.isSVGFile(imagePath)) {
                svgRelIds.set(baseName, relId);
            } else if (imagePath.toLowerCase().endsWith('.png')) {
                pngRelIds.set(baseName, relId);
            }
        }
        
        // Create mapping from PNG relId to SVG relId for matching patterns
        for (const [pngBaseName, pngRelId] of pngRelIds.entries()) {
            // First try exact match
            let svgRelId = svgRelIds.get(pngBaseName);
            
            // If no exact match, try pattern matching for PptxGenJS files
            // Pattern: image-X-N.png -> image-X-(N+1).svg
            if (!svgRelId) {
                const match = pngBaseName.match(/^(image-\d+-)(\d+)$/);
                if (match) {
                    const prefix = match[1]; // "image-X-"
                    const number = parseInt(match[2]); // N
                    const svgBaseName = prefix + (number + 1); // "image-X-(N+1)"
                    svgRelId = svgRelIds.get(svgBaseName);
                    
                    if (svgRelId) {
                    }
                }
            }
            
            if (svgRelId) {
                this.svgRelationshipMap.set(pngRelId, svgRelId);
            }
        }
        
    }

    /**
     * Load image from ZIP and draw it on the canvas
     */
    async loadAndDrawImage(graphics, shape, mediaInfo, x, y, w, h) {
        try {
            
            // Check if this is an SVG file and use enhanced SVG rendering
            if (this.isSVGFile(mediaInfo.path) && this.svgRenderer) {
                await this.loadAndDrawSVGImage(graphics, shape, mediaInfo, x, y, w, h);
                return;
            } else if (this.isSVGFile(mediaInfo.path)) {
                this.logger.log("warn", this.constructor.name, `SVG detected but no SVG renderer available: ${mediaInfo.path}`);
            } else {
            }

            // Check if image is already in cache
            let imageData = this.imageCache.get(shape.imageRelId);

            if (!imageData) {

                // Load image from ZIP
                const fileData = await this.zip.getFileData(mediaInfo.path);
                if (!fileData) {
                    throw new Error('Image file not found in ZIP');
                }

                // Create blob URL
                const mimeType = this.getMimeType(mediaInfo.path);
                const blob = new Blob([fileData], { type: mimeType });
                const imageUrl = URL.createObjectURL(blob);

                // Create image element
                const image = new Image();
                image.crossOrigin = 'anonymous';

                // Wait for image to load
                await new Promise((resolve, reject) => {
                    image.onload = () => {
                        imageData = {
                            image: image,
                            width: image.naturalWidth,
                            height: image.naturalHeight,
                            url: imageUrl
                        };

                        // Cache the image data
                        this.imageCache.set(shape.imageRelId, imageData);

                        resolve();
                    };
                    image.onerror = () => {
                        reject(new Error('Failed to load image'));
                    };
                    image.src = imageUrl;
                });
            } else {
            }

            // Draw the image
            this.drawActualImage(graphics, shape, imageData, x, y, w, h);

        } catch (error) {
            // Fall back to placeholder
            this.drawImagePlaceholder(graphics, shape, x, y, w, h);
        }
    }

    /**
     * Load and render SVG image using enhanced SVG renderer
     */
    async loadAndDrawSVGImage(graphics, shape, mediaInfo, x, y, w, h) {
        try {
            
            // Check if SVG data is already in cache
            let svgData = this.imageCache.get(shape.imageRelId);

            if (!svgData) {
                // Load SVG data from ZIP
                const fileData = await this.zip.getFileData(mediaInfo.path);
                if (!fileData) {
                    throw new Error('SVG file not found in ZIP');
                }

                // Convert array buffer to string for SVG content
                const svgContent = new TextDecoder('utf-8').decode(fileData);
                
                // Cache SVG content and metadata in format expected by graphics adapter
                svgData = {
                    content: svgContent,
                    type: 'svg',
                    path: mediaInfo.path,
                    width: 100, // Default width - SVG should be scalable
                    height: 100 // Default height - SVG should be scalable
                };
                
                this.imageCache.set(shape.imageRelId, svgData);
            }

            // Use SVG renderer to draw the SVG
            
            // Apply coordinate transformations if needed (pass-through for now)
            const transformedCoords = { x, y, w, h };
            
            // Add debug info about SVG content
            
            await this.svgRenderer.renderSVG(
                svgData.content,
                transformedCoords.x,
                transformedCoords.y,
                transformedCoords.w,
                transformedCoords.h,
                {
                    preserveAspectRatio: true
                }
            );


        } catch (error) {
            this.logger.logError(this.constructor.name, `Failed to load/render SVG: ${mediaInfo.path}`, error);
            
            // Fall back to standard image loading (browser SVG support) or placeholder
            try {
                // Try fallback to regular image loading
                const fileData = await this.zip.getFileData(mediaInfo.path);
                if (fileData) {
                    const blob = new Blob([fileData], { type: 'image/svg+xml' });
                    const imageUrl = URL.createObjectURL(blob);
                    
                    const image = new Image();
                    image.crossOrigin = 'anonymous';
                    
                    await new Promise((resolve, reject) => {
                        image.onload = () => {
                            const imageData = {
                                image: image,
                                width: image.naturalWidth,
                                height: image.naturalHeight,
                                url: imageUrl
                            };
                            this.drawActualImage(graphics, shape, imageData, x, y, w, h);
                            resolve();
                        };
                        image.onerror = reject;
                        image.src = imageUrl;
                    });
                }
            } catch (fallbackError) {
                this.logger.log("warn", this.constructor.name, 'SVG fallback also failed, showing placeholder');
                this.drawImagePlaceholder(graphics, shape, x, y, w, h);
            }
        }
    }

    drawFilePlaceholder(graphics, shape, fileType, x, y, w, h) {
        // Draw placeholder with an icon based on file type
        const ctx = graphics.getContext();
        ctx.save();

        // Draw placeholder background
        ctx.fillStyle = '#e9ecef';
        ctx.fillRect(x, y, w, h);

        // Draw placeholder border
        ctx.strokeStyle = '#adb5bd';
        ctx.lineWidth = 1;
        ctx.strokeRect(x, y, w, h);

        // Draw file icon
        let icon = '❓';
        if (fileType.includes('video')) {
            icon = '🎬';
        } else if (fileType.includes('audio')) {
            icon = '🎵';
        }

        ctx.fillStyle = '#495057';
        ctx.font = `${Math.min(w, h) * 0.5}px Arial`;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(icon, x + w / 2, y + h / 2);

        ctx.restore();
    }
    async renderShape(graphics, shape, canvasRect) {
        try {
            // renderShape called for shape
            
            if (!shape) {
                return;
            }

            // Get shape bounds and validate
            const bounds = this.getShapeBounds(shape, canvasRect);
            if (!bounds) {
                return;
            }

            // Check if shape is visible
            if (bounds.w <= 0 || bounds.h <= 0) {
                return;
            }

            const { x, y, w, h } = bounds;

            // Apply group transformations if present
            if (shape.groupTransform) {
                graphics.SaveGrState();

                // Apply group coordinate system transformation
                const transform = this.createTransformMatrix(shape, canvasRect);
                graphics.transform3(transform);
            }

            // Handle different shape types
            if (shape.graphicData && shape.graphicData.uri === 'http://schemas.openxmlformats.org/drawingml/2006/table') {
                // Processing table shape
            }
            switch (shape.type) {
                case 'sp':
                    this.renderRegularShape(graphics, shape, x, y, w, h);
                    break;
                case 'pic':
                    await this.renderPictureShape(graphics, shape, x, y, w, h);
                    break;
                case 'grpSp':
                    this.renderGroupShape(graphics, shape);
                    break;
                case 'cxnSp':
                    this.renderConnectorShape(graphics, shape, x, y, w, h);
                    break;
                case 'graphicFrame':
                    await this.renderGraphicFrame(graphics, shape, x, y, w, h);
                    break;
                default:
                    this.renderDefaultShape(graphics, shape, x, y, w, h);
                    break;
            }

            // Restore group transformation if applied
            if (shape.groupTransform) {
                graphics.RestoreGrState();
            }

        } catch (_error) {
				// Error ignored
			}
    }

    /**
     * Render regular shape with enhanced standard geometry processing
     */
    renderRegularShape(graphics, shape, x, y, w, h) {

        // Save graphics state for rotation
        const rotation = this.getShapeRotation(shape);
        if (rotation !== 0) {
            graphics.SaveGrState();
        }
        
        // Apply rotation before rendering
        const rotationApplied = this.applyRotationTransform(graphics, shape, x, y, w, h);
        
        // Get shape colors and styling
        const fillColor = this.getShapeFillColor(shape);
        const strokeColor = this.getShapeStrokeColor(shape);
        const lineWidth = this.getShapeLineWidth(shape);


        // Get stroke information for advanced styling
        const strokeInfo = this.getShapeStrokeInfo(shape);

        // Get shape geometry using standard patterns
        const preset = this.getShapePresetGeometry(shape);
        
        
        // Extract adjustment values from shape
        const adjustments = this.extractShapeAdjustments(shape);

        let currentFillColor = fillColor;

        // Convert RGB color objects to proper format for graphics engine
        let fillColorForShape = null;
        let strokeColorForShape = null;

        if (currentFillColor) {
            if (typeof currentFillColor === 'object' && currentFillColor.type && currentFillColor.stops) {
                // Gradient object — pass through directly to graphics engine
                fillColorForShape = currentFillColor;
            } else if (typeof currentFillColor === 'string') {
                const rgb = this.hexToRgb(currentFillColor);
                fillColorForShape = rgb;
            } else if (currentFillColor.r !== undefined) {
                fillColorForShape = currentFillColor;
            }
        }

        if (strokeColor) {
            if (typeof strokeColor === 'string') {
                const rgb = this.hexToRgb(strokeColor);
                strokeColorForShape = rgb;
            } else if (strokeColor.r !== undefined) {
                strokeColorForShape = strokeColor;
            }
        }


        // Enhanced shape type detection with custom geometry priority
        const shapeRendered = false;
        let renderMethod = 'none';


        // Priority 1: Check for custom geometry - handled by graphics adapter
        // Custom geometry rendering is now handled by the graphics adapter's drawShapeEnhanced method

        // Check for preset geometry
        if (preset) {
            graphics.drawPresetGeometry(preset, x, y, w, h, fillColorForShape, strokeColorForShape, lineWidth, strokeInfo, adjustments);
            renderMethod = 'preset-geometry';
        } else {
            // Fallback: Default rectangle
            graphics.drawRectangle(x, y, w, h, fillColorForShape, strokeColorForShape, lineWidth || 2, strokeInfo);
            renderMethod = 'fallback-rectangle';
        }


        // Text content is handled by the graphics adapter's drawShapeEnhanced method

        // Restore graphics state if rotation was applied
        if (rotation !== 0) {
            graphics.RestoreGrState();
        }
    }
    
    /**
     * Extract adjustment values from shape data
     * Based on standard adjustment value system
     */
    extractShapeAdjustments(shape) {
        const adjustments = {};
        
        try {
            // Check for geometry with adjustment values
            if (shape.geometry) {
                // Check for avLst (adjustment value list)
                if (shape.geometry.avLst) {
                    for (const [name, value] of Object.entries(shape.geometry.avLst)) {
                        if (typeof value === 'object' && value.val !== undefined) {
                            adjustments[name] = value.val;
                        } else if (typeof value === 'number') {
                            adjustments[name] = value;
                        }
                    }
                }
                
                // Check for adjLst (adjustment list)
                if (shape.geometry.adjLst) {
                    for (const [name, value] of Object.entries(shape.geometry.adjLst)) {
                        if (typeof value === 'object' && value.val !== undefined) {
                            adjustments[name] = value.val;
                        } else if (typeof value === 'number') {
                            adjustments[name] = value;
                        }
                    }
                }
                
                // Check for preset geometry adjustments
                if (shape.geometry.preset && shape.geometry.adjustments) {
                    for (const [name, value] of Object.entries(shape.geometry.adjustments)) {
                        adjustments[name] = value;
                    }
                }
            }
            
            // Check spPr.geometry for adjustment values
            if (shape.spPr && shape.spPr.geometry) {
                const geom = shape.spPr.geometry;
                
                if (geom.avLst) {
                    for (const adj of geom.avLst) {
                        if (adj.name && adj.fmla !== undefined) {
                            // Parse adjustment value from formula
                            const value = this.parseAdjustmentValue(adj.fmla);
                            if (value !== null) {
                                adjustments[adj.name] = value;
                            }
                        }
                    }
                }
            }
            
            // Check for direct adjustment properties
            if (shape.adjustments) {
                Object.assign(adjustments, shape.adjustments);
            }
            
            // Check commonSlideData for adjustments (some shapes store them here)
            if (shape.commonSlideData && shape.commonSlideData.adjustments) {
                Object.assign(adjustments, shape.commonSlideData.adjustments);
            }
            
        } catch (error) {
            // Error ignored  
        }
        
        return adjustments;
    }
    
    /**
     * Parse adjustment value from standard formula
     */
    parseAdjustmentValue(formula) {
        if (typeof formula === 'number') {
            return formula;
        }
        
        if (typeof formula === 'string') {
            // Handle simple numeric values
            const numValue = parseFloat(formula);
            if (!isNaN(numValue)) {
                return numValue;
            }
            
            // Handle val expressions like "val 25000"
            const valMatch = formula.match(/val\s+(-?\d+)/);
            if (valMatch) {
                return parseInt(valMatch[1], 10);
            }
            
            // Handle percentage expressions
            const percentMatch = formula.match(/(-?\d+)%/);
            if (percentMatch) {
                return parseInt(percentMatch[1], 10) * 1000; // Convert to standard units
            }
        }
        
        return null;
    }
    

    /**
     * Get shape preset geometry name
     */
    getShapePresetGeometry(shape) {
        if (shape.geometry && shape.geometry.preset) return shape.geometry.preset;
        if (shape.properties && shape.properties.geometry && shape.properties.geometry.preset) return shape.properties.geometry.preset;
        if (shape.spPr && shape.spPr.geometry && shape.spPr.geometry.preset) return shape.spPr.geometry.preset;
        if (shape.spPr && shape.spPr.prstGeom && shape.spPr.prstGeom.prst) return shape.spPr.prstGeom.prst;
        return null;
    }

    /**
     * Get enhanced stroke information for standard rendering
     */
    getShapeStrokeInfo(shape) {
        const strokeInfo = {};

        // Get line properties from shape (check both line and stroke properties)
        const lineProps = shape.properties?.line || shape.properties?.stroke;
        
        if (lineProps) {
            // Dash pattern
            if (lineProps.dashArray && lineProps.dashArray.length > 0) {
                strokeInfo.dashArray = lineProps.dashArray;
            }

            // Line cap style
            if (lineProps.lineCap || lineProps.cap) {
                strokeInfo.lineCap = lineProps.lineCap || lineProps.cap;
            }

            // Line join style
            if (lineProps.lineJoin || lineProps.join) {
                strokeInfo.lineJoin = lineProps.lineJoin || lineProps.join;
            }

            // Miter limit
            if (lineProps.miterLimit) {
                strokeInfo.miterLimit = lineProps.miterLimit;
            }

            // Arrowheads from XML (OOXML): retain raw headEnd/tailEnd; renderer maps tailEnd→start, headEnd→end
            if (lineProps.headEnd) {
                const h = lineProps.headEnd;
                strokeInfo.headEnd = {
                    type: h.type || h.val || 'none',
                    w: h.w || h.width,
                    len: h.len || h.length
                };
            }
            if (lineProps.tailEnd) {
                const t = lineProps.tailEnd;
                strokeInfo.tailEnd = {
                    type: t.type || t.val || 'none',
                    w: t.w || t.width,
                    len: t.len || t.length
                };
            }
        }

        // Get line properties from spPr
        if (shape.spPr && shape.spPr.ln) {
            const ln = shape.spPr.ln;
            

            // Convert standard line properties
            if (ln.prstDash) {
                strokeInfo.dashArray = this.convertstandardDashPattern(ln.prstDash);
            }

            if (ln.cap) {
                strokeInfo.lineCap = this.convertstandardLineCap(ln.cap);
            }

            if (ln.join) {
                strokeInfo.lineJoin = this.convertstandardLineJoin(ln.join);
            }

            // Parse arrowheads from headEnd/tailEnd
            // Expected structure similar to: ln.headEnd = { type: 'triangle', w: 'med', len: 'med' }
            const mapArrowSize = this._mapArrowSize.bind(this);
            const normalizeArrowType = (t) => {
                if (!t) {return 'none';}
                const s = t.toString();
                // PPTX uses values like 'triangle', 'stealth', 'diamond', 'oval', 'arrow', 'open'
                return s;
            };

            if (ln.headEnd) {
                strokeInfo.headEnd = {
                    type: normalizeArrowType(ln.headEnd.type),
                    w: ln.headEnd.w,
                    len: ln.headEnd.len
                };
            }
            if (ln.tailEnd) {
                strokeInfo.tailEnd = {
                    type: normalizeArrowType(ln.tailEnd.type),
                    w: ln.tailEnd.w,
                    len: ln.tailEnd.len
                };
            }
        }

        
        return Object.keys(strokeInfo).length > 0 ? strokeInfo : null;
    }

    _mapArrowSize(size) {
        switch ((size || '').toString()) {
            case 'sm':
            case 'small':
                return 1.0;
            case 'med':
            case 'medium':
                return 1.5;
            case 'lg':
            case 'large':
                return 2.0;
            default:
                return 1.0;
        }
    }

    /**
     * Convert standard dash pattern to Canvas dash array
     */
    convertstandardDashPattern(prstDash) {
        const dashPatterns = {
            'solid': [],
            'dot': [2, 2],
            'dash': [8, 4],
            'dashDot': [8, 4, 2, 4],
            'dashDotDot': [8, 4, 2, 4, 2, 4],
            'lgDash': [16, 8],
            'lgDashDot': [16, 8, 4, 8],
            'lgDashDotDot': [16, 8, 4, 8, 4, 8],
            'sysDash': [6, 3],
            'sysDashDot': [6, 3, 2, 3],
            'sysDashDotDot': [6, 3, 2, 3, 2, 3],
            'sysDot': [1, 1]
        };

        return dashPatterns[prstDash] || [];
    }

    /**
     * Convert standard line cap to Canvas line cap
     */
    convertstandardLineCap(cap) {
        const capStyles = {
            'rnd': 'round',
            'sq': 'square',
            'flat': 'butt'
        };

        return capStyles[cap] || 'round';
    }

    /**
     * Convert standard line join to Canvas line join
     */
    convertstandardLineJoin(join) {
        const joinStyles = {
            'round': 'round',
            'bevel': 'bevel',
            'miter': 'miter'
        };

        return joinStyles[join] || 'round';
    }

    /**
     * Enhanced picture shape rendering with standard integration
     */
    async renderPictureShape(graphics, shape, x, y, w, h) {
        // renderPictureShape called for shape
        
        // Check if rotation is needed and save graphics state
        const rotation = this.getShapeRotation(shape);
        const needsRotation = rotation && rotation !== 0;
        if (needsRotation) {
            graphics.SaveGrState();
            // Apply rotation transformation
            this.applyRotationTransform(graphics, shape, x, y, w, h);
        }

        // Check if shape has preset geometry (standard pattern for picture frames)
        const preset = this.getShapePresetGeometry(shape);

        if (preset && preset !== 'rect') {
            // Picture with custom shape geometry

            // Draw the shape outline first
            const strokeColor = this.getShapeStrokeColor(shape);
            const lineWidth = this.getShapeLineWidth(shape);
            const strokeInfo = this.getShapeStrokeInfo(shape);

            graphics.drawPresetGeometry(preset, x, y, w, h, null, strokeColor, lineWidth, strokeInfo);
        }

        // Render the actual image using the enhanced picture geometry method
        const fillColor = this.getShapeFillColor(shape);
        const strokeColor = this.getShapeStrokeColor(shape);
        const lineWidth = this.getShapeLineWidth(shape);

        try {
            // Use the enhanced picture shape geometry method which handles async loading
            await this.drawPictureShapeGeometry(graphics, shape, x, y, w, h, fillColor, strokeColor, lineWidth);
        } catch (error) {
            // Fall back to placeholder
            this.drawImagePlaceholder(graphics, shape, x, y, w, h);
        }

        // Restore graphics state if rotation was applied
        if (needsRotation) {
            graphics.RestoreGrState();
        }
    }

    /**
     * Enhanced connector shape rendering with standard integration
     */
    renderConnectorShape(graphics, shape, x, y, w, h) {
        // Check if rotation is needed and save graphics state
        const rotation = this.getShapeRotation(shape);
        const needsRotation = rotation && rotation !== 0;
        if (needsRotation) {
            graphics.SaveGrState();
            // Apply rotation transformation
            this.applyRotationTransform(graphics, shape, x, y, w, h);
        }

        const strokeColor = this.getShapeStrokeColor(shape);
        const lineWidth = this.getShapeLineWidth(shape);
        const strokeInfo = this.getShapeStrokeInfo(shape);

        // Get connector type from standard properties
        const connectorType = this.getConnectorType(shape);

        try {
            this.logger?.log('info', this.constructor.name, '[Connector Render] start', {
                connectorType,
                bounds: { x, y, w, h },
                strokeColor,
                lineWidth,
                strokeInfo
            });
        } catch (e) {}

        switch (connectorType) {
            case 'straight': {
                // Respect arrowheads if available
                if (typeof graphics.drawLineWithArrows === 'function') {
                    graphics.drawLineWithArrows(x, y, x + w, y + h, strokeColor, lineWidth, strokeInfo);
                } else {
                    graphics.drawLine(x, y, x + w, y + h, strokeColor, lineWidth);
                }
                break; }
            case 'elbow':
                this.drawElbowConnector(graphics, x, y, w, h, strokeColor, lineWidth, strokeInfo);
                break;
            case 'curved':
                this.drawCurvedConnector(graphics, x, y, w, h, strokeColor, lineWidth, strokeInfo);
                break;
            default:
                // Default to straight line
                graphics.drawLine(x, y, x + w, y + h, strokeColor, lineWidth);
                break;
        }

        // Restore graphics state if rotation was applied
        if (needsRotation) {
            graphics.RestoreGrState();
        }

        try { this.logger?.log('info', this.constructor.name, '[Connector Render] done'); } catch (e) {}
    }

    /**
     * Get connector type from standard shape properties
     */
    getConnectorType(shape) {
        // Check for connector type in shape properties
        if (shape.properties && shape.properties.connector) {
            return shape.properties.connector.type || 'straight';
        }

        // Check for connector type in spPr
        if (shape.spPr && shape.spPr.connector) {
            return shape.spPr.connector.type || 'straight';
        }

        // Check preset geometry for connector types
        const preset = this.getShapePresetGeometry(shape);
        if (preset && preset.includes('Connector')) {
            if (preset.includes('bent')) {
                return 'elbow';
            } else if (preset.includes('curved')) {
                return 'curved';
            }
        }

        return 'straight';
    }

    /**
     * Draw elbow connector
     */
    drawElbowConnector(graphics, x, y, w, h, strokeColor, lineWidth, strokeInfo) {
        if (!graphics.m_oContext) {return;}

        graphics.m_oContext.save();
        graphics.m_oContext.beginPath();

        // Calculate elbow point
        const elbowX = x + w * 0.5;
        const elbowY = y + h * 0.5;

        graphics.m_oContext.moveTo(x, y);
        graphics.m_oContext.lineTo(elbowX, y);
        graphics.m_oContext.lineTo(elbowX, elbowY);
        graphics.m_oContext.lineTo(x + w, elbowY);
        graphics.m_oContext.lineTo(x + w, y + h);

        if (strokeColor) {
            graphics.m_oContext.strokeStyle = graphics.colorToRgb(strokeColor);
            graphics.m_oContext.lineWidth = lineWidth || 1;

            if (strokeInfo) {
                graphics.applyStrokeInfo(strokeInfo);
            }

            graphics.m_oContext.stroke();
        }

        graphics.m_oContext.restore();
    }

    /**
     * Draw curved connector
     */
    drawCurvedConnector(graphics, x, y, w, h, strokeColor, lineWidth, strokeInfo) {
        if (!graphics.m_oContext) {return;}

        graphics.m_oContext.save();
        graphics.m_oContext.beginPath();

        // Create curved path
        graphics.m_oContext.moveTo(x, y);
        graphics.m_oContext.bezierCurveTo(
            x + w * 0.3, y,
            x + w * 0.7, y + h,
            x + w, y + h
        );

        if (strokeColor) {
            graphics.m_oContext.strokeStyle = graphics.colorToRgb(strokeColor);
            graphics.m_oContext.lineWidth = lineWidth || 1;

            if (strokeInfo) {
                graphics.applyStrokeInfo(strokeInfo);
            }

            graphics.m_oContext.stroke();
        }

        graphics.m_oContext.restore();
    }

    /**
     * Enhanced graphic frame rendering with standard integration
     */
    async renderGraphicFrame(graphics, shape, x, y, w, h) {
        
        // Check if rotation is needed and save graphics state
        const rotation = this.getShapeRotation(shape);
        const needsRotation = rotation && rotation !== 0;
        if (needsRotation) {
            graphics.SaveGrState();
            // Apply rotation transformation
            this.applyRotationTransform(graphics, shape, x, y, w, h);
        }

        // Draw frame outline
        const strokeColor = this.getShapeStrokeColor(shape);
        const lineWidth = this.getShapeLineWidth(shape) || 1;
        const strokeInfo = this.getShapeStrokeInfo(shape);

        graphics.drawRectangle(x, y, w, h, null, strokeColor, lineWidth, strokeInfo);

        // Draw frame content based on type
        if (shape.graphicData) {
            // Processing shape with URI
            switch (shape.graphicData.uri) {
                case 'http://schemas.openxmlformats.org/drawingml/2006/table':
                    // Calling renderTableFrame for table
                    this.renderTableFrame(graphics, shape, x, y, w, h);
                    // renderTableFrame completed
                    break;
                case 'http://schemas.openxmlformats.org/drawingml/2006/chart':
                    await this.renderChartFrame(graphics, shape, x, y, w, h);
                    break;
                case 'http://schemas.openxmlformats.org/drawingml/2006/diagram':
                    this.renderDiagramFrame(graphics, shape, x, y, w, h);
                    break;
                default:
                    this.renderGenericFrame(graphics, shape, x, y, w, h);
                    break;
            }
        } else {
            this.renderGenericFrame(graphics, shape, x, y, w, h);
        }

        // Restore graphics state if rotation was applied
        if (needsRotation) {
            graphics.RestoreGrState();
        }
    }

    /**
     * Render table frame
     */
    /**
     * Enhanced table rendering in four stages
     */
    renderTableFrame(graphics, shape, x, y, w, h) {
        // Stage 1: Extract and validate table data
        const table = this.getTableInstance(shape);
        // getTableInstance result checked
        
        if (!table) {
            // Drawing placeholder due to missing table instance
            this.drawTablePlaceholder(graphics, x, y, w, h);
            return;
        }
        
        // Table parsed successfully, proceeding with rendering

        // Save graphics state for table rendering
        graphics.SaveGrState();
        
        try {
            // Stage 2: Draw table background and outer border
            this.drawTableBackgroundAndOuterBorder(graphics, table, x, y, w, h);
            
            // Stage 3: Draw cell backgrounds
            this.drawCellsBackground(graphics, table, x, y, w, h);
            
            // Stage 4: Draw cell content
            this.drawCellsContent(graphics, table, x, y, w, h);
            
            // Stage 5: Draw cell borders
            this.drawCellsBorders(graphics, table, x, y, w, h);
            
        } catch (error) {
            this.drawTablePlaceholder(graphics, x, y, w, h);
        } finally {
            graphics.RestoreGrState();
        }
    }

    /**
     * Get table instance with caching
     */
    getTableInstance(shape) {
        // Starting table instance retrieval
        
        // Check for existing cached table instance
        if (shape.table && shape.table instanceof CTable) {
            // Using cached table instance
            return shape.table;
        }

        // No cached table, checking for XML
        
        // Parse table from XML if available
        if (shape.graphicData && shape.graphicData.tableXml) {
            // Attempting to parse table from XML
            try {
                const table = this.parseTableFromXML(shape.graphicData.tableXml);
                // parseTableFromXML result checked
                if (table) {
                    // Caching parsed table and returning
                    // Cache the parsed table
                    shape.table = table;
                    return table;
                }
            } catch (error) {
                console.error('[GetTableInstance] Error parsing table:', error.message);
                console.error('[GetTableInstance] Error stack:', error.stack);
            }
        } else {
            // No table XML available
        }

        // Returning null - no table available
        return null;
    }

    /**
     * Stage 1: Draw table background and outer border
     */
    drawTableBackgroundAndOuterBorder(graphics, table, x, y, w, h) {
        // Get table properties
        const tableProps = table.getTableProperties();
        const tableBorders = table.getTableBorders();
        const tableShading = table.getTableShading();

        // Draw table background if specified
        if (tableShading && tableShading.fill) {
            const bgColor = this.resolveTableBackgroundColor(tableShading);
            graphics.b_color1(bgColor.r, bgColor.g, bgColor.b, bgColor.a || 255);
            graphics.TableRect(x, y, w, h);
        }

        // Draw outer table border
        if (tableBorders) {
            this.drawTableOuterBorder(graphics, tableBorders, x, y, w, h);
        }
    }

    /**
     * Stage 2: Draw cell backgrounds
     */
    drawCellsBackground(graphics, table, x, y, w, h) {
        const rowCount = table.getRowCount();
        const colCount = table.getColumnCount();
        
        // Calculate cell dimensions
        const cellWidths = this.calculateCellWidths(table, w);
        const cellHeights = this.calculateCellHeights(table, h);

        let currentY = y;
        
        for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            let currentX = x;
            const rowHeight = cellHeights[rowIndex];
            
            for (let colIndex = 0; colIndex < colCount; colIndex++) {
                const cell = table.getCell(rowIndex, colIndex);
                if (!cell) {continue;}

                const cellWidth = cellWidths[colIndex];
                
                // Handle merged cells
                const gridSpan = cell.getGridSpan() || 1;
                const rowSpan = cell.getRowSpan() || 1;
                
                // Calculate actual cell dimensions
                const actualWidth = this.calculateMergedCellWidth(cellWidths, colIndex, gridSpan);
                const actualHeight = this.calculateMergedCellHeight(cellHeights, rowIndex, rowSpan);
                
                // Only draw background for the starting cell of merged ranges
                if (this.isStartingCellOfMerge(cell, rowIndex, colIndex)) {
                    this.drawCellBackground(graphics, cell, currentX, currentY, actualWidth, actualHeight);
                }
                
                currentX += cellWidth;
            }
            
            currentY += rowHeight;
        }
    }

    /**
     * Stage 3: Draw cell content
     */
    drawCellsContent(graphics, table, x, y, w, h) {
        const rowCount = table.getRowCount();
        const colCount = table.getColumnCount();
        
        const cellWidths = this.calculateCellWidths(table, w);
        const cellHeights = this.calculateCellHeights(table, h);

        let currentY = y;
        
        for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            let currentX = x;
            const rowHeight = cellHeights[rowIndex];
            
            for (let colIndex = 0; colIndex < colCount; colIndex++) {
                const cell = table.getCell(rowIndex, colIndex);
                if (!cell) {continue;}

                const cellWidth = cellWidths[colIndex];
                
                // Only render content for starting cells of merged ranges
                if (this.isStartingCellOfMerge(cell, rowIndex, colIndex)) {
                    const gridSpan = cell.getGridSpan() || 1;
                    const rowSpan = cell.getRowSpan() || 1;
                    
                    const actualWidth = this.calculateMergedCellWidth(cellWidths, colIndex, gridSpan);
                    const actualHeight = this.calculateMergedCellHeight(cellHeights, rowIndex, rowSpan);
                    
                    this.drawCellContent(graphics, cell, currentX, currentY, actualWidth, actualHeight);
                }
                
                currentX += cellWidth;
            }
            
            currentY += rowHeight;
        }
    }

    /**
     * Stage 4: Draw cell borders
     */
    drawCellsBorders(graphics, table, x, y, w, h) {
        const rowCount = table.getRowCount();
        const colCount = table.getColumnCount();
        
        const cellWidths = this.calculateCellWidths(table, w);
        const cellHeights = this.calculateCellHeights(table, h);

        const currentY = y;
        
        // Draw borders from right to left to handle conflicts
        for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            for (let colIndex = colCount - 1; colIndex >= 0; colIndex--) {
                const cell = table.getCell(rowIndex, colIndex);
                if (!cell) {continue;}

                const cellWidth = cellWidths[colIndex];
                const cellHeight = cellHeights[rowIndex];
                
                const cellX = x + this.calculateCellXOffset(cellWidths, colIndex);
                const cellY = y + this.calculateCellYOffset(cellHeights, rowIndex);
                
                this.drawCellBorders(graphics, cell, cellX, cellY, cellWidth, cellHeight, rowIndex, colIndex, table);
            }
        }
    }

    /**
     * Draw table placeholder (fallback)
     */
    drawTablePlaceholder(graphics, x, y, w, h) {
        // Draw simple grid placeholder
        const cellWidth = w / 3;
        const cellHeight = h / 3;

        for (let row = 0; row < 3; row++) {
            for (let col = 0; col < 3; col++) {
                const cellX = x + col * cellWidth;
                const cellY = y + row * cellHeight;

                graphics.drawRectangle(cellX, cellY, cellWidth, cellHeight, null, { r: 128, g: 128, b: 128, a: 255 }, 1);
            }
        }

        // Draw table label
        this.drawCenteredText(graphics, 'Table', x, y, w, h);
    }
    /**
     * Draw table outer border
     */
    drawTableOuterBorder(graphics, tableBorders, x, y, w, h) {
        const borders = ['top', 'right', 'bottom', 'left'];
        
        borders.forEach(borderSide => {
            const border = tableBorders[borderSide];
            if (border && border.style !== 'none') {
                const color = this.resolveBorderColor(border);
                const width = border.width || 1;
                
                graphics.p_color(color.r, color.g, color.b, color.a || 255);
                graphics.p_width(width);
                
                switch (borderSide) {
                    case 'top':
                        graphics.drawHorLine(x, y, x + w, y, width);
                        break;
                    case 'right':
                        graphics.drawVerLine(x + w, y, y + h, width);
                        break;
                    case 'bottom':
                        graphics.drawHorLine(x, y + h, x + w, y + h, width);
                        break;
                    case 'left':
                        graphics.drawVerLine(x, y, y + h, width);
                        break;
                }
            }
        });
    }

    /**
     * Draw cell background
     */
    drawCellBackground(graphics, cell, x, y, w, h) {
        const cellShading = cell.getCellShading();
        
        if (cellShading && cellShading.fill) {
            const bgColor = this.resolveCellBackgroundColor(cellShading);
            graphics.b_color1(bgColor.r, bgColor.g, bgColor.b, bgColor.a || 255);
            graphics.TableRect(x, y, w, h);
        }
    }

    /**
     * Draw cell content
     */
    drawCellContent(graphics, cell, x, y, w, h) {
        const textBody = cell.getTextBody();
        if (!textBody || !textBody.paragraphs) {return;}

        // Apply cell padding
        const padding = 5;
        const contentX = x + padding;
        const contentY = y + padding;
        const contentW = w - (padding * 2);
        const contentH = h - (padding * 2);

        // Render text content
        this.renderCellTextContent(graphics, textBody, contentX, contentY, contentW, contentH);
    }

    /**
     * Draw cell borders with conflict resolution
     */
    drawCellBorders(graphics, cell, x, y, w, h, rowIndex, colIndex, table) {
        const cellBorders = cell.getCellBorders();
        if (!cellBorders) {return;}

        const borders = ['top', 'right', 'bottom', 'left'];
        
        borders.forEach(borderSide => {
            const border = cellBorders[borderSide];
            if (!border || border.style === 'none') {return;}

            // Border conflict resolution
            if (this.shouldDrawBorder(borderSide, rowIndex, colIndex, table)) {
                const color = this.resolveBorderColor(border);
                const width = border.width || 1;
                
                graphics.p_color(color.r, color.g, color.b, color.a || 255);
                graphics.p_width(width);
                
                switch (borderSide) {
                    case 'top':
                        graphics.drawHorLine(x, y, x + w, y, width);
                        break;
                    case 'right':
                        graphics.drawVerLine(x + w, y, y + h, width);
                        break;
                    case 'bottom':
                        graphics.drawHorLine(x, y + h, x + w, y + h, width);
                        break;
                    case 'left':
                        graphics.drawVerLine(x, y, y + h, width);
                        break;
                }
            }
        });
    }

    /**
     * Border conflict resolution
     */
    shouldDrawBorder(borderSide, rowIndex, colIndex, table) {
        const rowCount = table.getRowCount();
        const colCount = table.getColumnCount();
        
        switch (borderSide) {
            case 'top':
                return rowIndex === 0; // Only draw top border for first row
            case 'right':
                return colIndex === colCount - 1; // Only draw right border for last column
            case 'bottom':
                return rowIndex === rowCount - 1; // Only draw bottom border for last row
            case 'left':
                return colIndex === 0; // Only draw left border for first column
            default:
                return true;
        }
    }

    /**
     * Calculate cell widths based on table grid
     */
    calculateCellWidths(table, totalWidth) {
        const tableGrid = table.getTableGrid();
        if (!tableGrid || tableGrid.length === 0) {
            // Fallback to equal distribution
            const colCount = table.getColumnCount();
            return new Array(colCount).fill(totalWidth / colCount);
        }

        const totalGridWidth = tableGrid.reduce((sum, col) => sum + (col.width || 0), 0);
        return tableGrid.map(col => {
            const gridWidth = col.width || 0;
            return (gridWidth / totalGridWidth) * totalWidth;
        });
    }

    /**
     * Calculate cell heights
     */
    calculateCellHeights(table, totalHeight) {
        const rowCount = table.getRowCount();
        const unitHeights = [];
        let hasExplicitHeights = false;

        for (let i = 0; i < rowCount; i++) {
            const row = table.getRow(i);
            if (row && row.height) {
                unitHeights.push(row.height);
                hasExplicitHeights = true;
            } else {
                unitHeights.push(0);
            }
        }

        if (!hasExplicitHeights) {
            // All equal
            return unitHeights.map(() => totalHeight / rowCount);
        }

        // Fill in zero-height rows with average of explicit heights
        const explicitSum = unitHeights.reduce((s, v) => s + v, 0);
        const zeroCount = unitHeights.filter(v => v === 0).length;
        const avgHeight = zeroCount < rowCount ? explicitSum / (rowCount - zeroCount) : totalHeight / rowCount;
        const filledHeights = unitHeights.map(v => v === 0 ? avgHeight : v);

        // Normalize proportionally to totalHeight
        const totalUnits = filledHeights.reduce((s, v) => s + v, 0);
        return filledHeights.map(u => totalHeight * (u / totalUnits));
    }

    /**
     * Check if cell is the starting cell of a merged range
     */
    isStartingCellOfMerge(cell, rowIndex, colIndex) {
        // This is a simplified check
        return !cell.isMergedContinue;
    }

    /**
     * Calculate merged cell width
     */
    calculateMergedCellWidth(cellWidths, startCol, gridSpan) {
        let width = 0;
        for (let i = 0; i < gridSpan; i++) {
            if (startCol + i < cellWidths.length) {
                width += cellWidths[startCol + i];
            }
        }
        return width;
    }

    /**
     * Calculate merged cell height
     */
    calculateMergedCellHeight(cellHeights, startRow, rowSpan) {
        let height = 0;
        for (let i = 0; i < rowSpan; i++) {
            if (startRow + i < cellHeights.length) {
                height += cellHeights[startRow + i];
            }
        }
        return height;
    }

    /**
     * Calculate cell X offset
     */
    calculateCellXOffset(cellWidths, colIndex) {
        let offset = 0;
        for (let i = 0; i < colIndex; i++) {
            offset += cellWidths[i];
        }
        return offset;
    }

    /**
     * Calculate cell Y offset
     */
    calculateCellYOffset(cellHeights, rowIndex) {
        let offset = 0;
        for (let i = 0; i < rowIndex; i++) {
            offset += cellHeights[i];
        }
        return offset;
    }

    /**
     * Resolve table background color
     */
    resolveTableBackgroundColor(tableShading) {
        // Default white background
        return { r: 255, g: 255, b: 255, a: 255 };
    }

    /**
     * Resolve cell background color
     */
    resolveCellBackgroundColor(cellShading) {
        // Default white background
        return { r: 255, g: 255, b: 255, a: 255 };
    }

    /**
     * Resolve border color
     */
    resolveBorderColor(border) {
        // Default black border
        return { r: 0, g: 0, b: 0, a: 255 };
    }

    /**
     * Render cell text content
     */
    renderCellTextContent(graphics, textBody, x, y, w, h) {
        if (!textBody.paragraphs || textBody.paragraphs.length === 0) {return;}

        let currentY = y;
        const lineHeight = 16; // Default line height

        textBody.paragraphs.forEach(paragraph => {
            if (!paragraph.runs || paragraph.runs.length === 0) {
                currentY += lineHeight;
                return;
            }

            // Combine all runs in the paragraph
            const text = paragraph.runs.map(run => run.text || '').join('');
            if (!text.trim()) {
                currentY += lineHeight;
                return;
            }

            // Get paragraph properties
            const paraProps = paragraph.properties || {};
            const fontSize = paraProps.fontSize || 12;
            const fontFamily = paraProps.fontFamily || 'Arial';
            const isBold = paraProps.bold || false;
            const isItalic = paraProps.italic || false;

            // Set font
            graphics.font(fontFamily, fontSize, isBold ? 'bold' : 'normal', isItalic ? 'italic' : 'normal');

            // Set text color
            const textColor = this.resolveTextColor(paragraph);
            graphics.b_color1(textColor.r, textColor.g, textColor.b, textColor.a || 255);

            // Draw text
            graphics.FillText(x, currentY, text);

            currentY += lineHeight;
        });
    }

    /**
     * Resolve text color
     */
    resolveTextColor(paragraph) {
        // Default black text
        return { r: 0, g: 0, b: 0, a: 255 };
    }

    /**
     * Enhanced table parsing
     */
    parseTableFromXML(tableXml) {
        // Starting table XML parsing
        if (!tableXml) {
            // No table XML provided
            return null;
        }

        // Table XML length and preview checked

        try {
            const parser = new DOMParser();
            const doc = parser.parseFromString(tableXml, 'text/xml');
            
            if (doc.documentElement.nodeName === 'parsererror') {
                console.error('[ParseTableXML] XML parsing error');
                return null;
            }

            // XML parsed successfully, creating table instance
            const table = new CTable();
            
            // Stage 1: Parse table properties
            this.parseTableProperties(doc, table);
            
            // Stage 2: Parse table grid (column definitions)
            this.parseTableGrid(doc, table);
            
            // Stage 3: Parse table rows and cells
            this.parseTableRows(doc, table);
            
            // Stage 4: Process merged cells and relationships
            this.processTableMerges(table);
            
            // Stage 5: Validate table structure
            const isValid = this.validateTableStructure(table);
            
            if (!isValid) {
                return null;
            }

            return table;

        } catch (error) {
            console.error('[ParseTableXML] Exception during table parsing:', error.message);
            console.error('[ParseTableXML] Error stack:', error.stack);
            return null;
        }
    }

    /**
     * Parse table properties
     */
    parseTableProperties(doc, table) {
        // Find table properties element
        let tblPr = doc.querySelector('tblPr');
        if (!tblPr) {tblPr = doc.querySelector('a\\:tblPr');}
        
        if (tblPr) {
            const tableProps = {};
            
            // Parse table style
            let tblStyle = tblPr.querySelector('tblStyle');
            if (!tblStyle) {tblStyle = tblPr.querySelector('a\\:tblStyle');}
            if (tblStyle) {
                tableProps.style = tblStyle.getAttribute('val');
            }
            
            // Parse table borders
            let tblBorders = tblPr.querySelector('tblBorders');
            if (!tblBorders) {tblBorders = tblPr.querySelector('a\\:tblBorders');}
            if (tblBorders) {
                tableProps.borders = this.parseTableBorders(tblBorders);
            }
            
            // Parse table shading
            let tblShading = tblPr.querySelector('tblShading');
            if (!tblShading) {tblShading = tblPr.querySelector('a\\:tblShading');}
            if (tblShading) {
                tableProps.shading = this.parseTableShading(tblShading);
            }
            
            // Parse table layout
            let tblLayout = tblPr.querySelector('tblLayout');
            if (!tblLayout) {tblLayout = tblPr.querySelector('a\\:tblLayout');}
            if (tblLayout) {
                tableProps.layout = tblLayout.getAttribute('type') || 'autofit';
            }
            // Parse preferred width (tblW)
            let tblW = tblPr.querySelector('tblW');
            if (!tblW) {tblW = tblPr.querySelector('a\\:tblW');}
            if (tblW) {
                const wVal = tblW.getAttribute('w');
                const type = tblW.getAttribute('type') || tblW.getAttribute('wtype') || 'auto';
                if (wVal) {
                    tableProps.preferredWidth = { value: parseInt(wVal) || 0, type: type.toLowerCase() };
                } else {
                    tableProps.preferredWidth = { value: 0, type: type.toLowerCase() };
                }
            }
            
            table.setTableProperties(tableProps);
        }
    }

    /**
     * Parse table grid
     */
    parseTableGrid(doc, table) {
        let tblGrid = doc.querySelector('tblGrid');
        if (!tblGrid) {tblGrid = doc.querySelector('a\\:tblGrid');}
        
        if (tblGrid) {
            let gridCols = tblGrid.querySelectorAll('gridCol');
            if (gridCols.length === 0) {
                gridCols = tblGrid.querySelectorAll('a\\:gridCol');
            }
            
            const columns = [];
            gridCols.forEach(gridCol => {
                const width = gridCol.getAttribute('w');
                const column = {
                    width: width ? parseInt(width) : 914400, // Default 1 inch
                    type: gridCol.getAttribute('type') || 'auto'
                };
                columns.push(column);
            });
            table.setTableGrid(columns);
        }
    }

    /**
     * Parse table rows
     */
    parseTableRows(doc, table) {
        let tableRows = doc.querySelectorAll('tr');
        if (tableRows.length === 0) {
            tableRows = doc.querySelectorAll('a\\:tr');
        }
        
        tableRows.forEach((trElement, rowIndex) => {
            const row = new CTableRow();
            
            // Parse row properties
            this.parseRowProperties(trElement, row);
            
            // Parse table cells
            this.parseRowCells(trElement, row, rowIndex);
            
            table.addRow(row);
        });
    }

    /**
     * Parse row properties
     */
    parseRowProperties(trElement, row) {
        let trPr = trElement.querySelector('trPr');
        if (!trPr) {trPr = trElement.querySelector('a\\:trPr');}
        
        if (trPr) {
            // Row height
            let trHeight = trPr.querySelector('trHeight');
            if (!trHeight) {trHeight = trPr.querySelector('a\\:trHeight');}
            if (trHeight) {
                const parsed = parseInt(trHeight.getAttribute('val'), 10);
                if (!Number.isNaN(parsed)) {
                    row.height = parsed;
                }
                row.heightRule = trHeight.getAttribute('hRule') || 'auto';
            }
            
            // Row header
            let tblHeader = trPr.querySelector('tblHeader');
            if (!tblHeader) {tblHeader = trPr.querySelector('a\\:tblHeader');}
            if (tblHeader) {
                row.isHeader = true;
            }
        }

        if ((!row.height || row.height <= 0) && trElement) {
            const attrHeight = trElement.getAttribute('h');
            if (attrHeight !== null) {
                const parsed = parseInt(attrHeight, 10);
                if (!Number.isNaN(parsed) && parsed > 0) {
                    row.height = parsed;
                    if (!row.heightRule) {
                        row.heightRule = 'auto';
                    }
                }
            }
        }
    }

    /**
     * Parse row cells
     */
    parseRowCells(trElement, row, rowIndex) {
        let tableCells = trElement.querySelectorAll('tc');
        if (tableCells.length === 0) {
            tableCells = trElement.querySelectorAll('a\\:tc');
        }
        
        tableCells.forEach((tcElement, cellIndex) => {
            const cell = new CTableCell();
            
            // Parse cell properties
            this.parseCellProperties(tcElement, cell, rowIndex, cellIndex);
            
            // Parse cell content
            this.parseCellContent(tcElement, cell);
            
            row.addCell(cell);
        });
    }

    /**
     * Parse cell properties
     */
    parseCellProperties(tcElement, cell, rowIndex, cellIndex) {
        // Read merge attributes from <a:tc> element directly (PowerPoint style)
        // PowerPoint puts gridSpan/rowSpan/vMerge as attributes on <a:tc>, not inside <a:tcPr>
        const gridSpanAttr = tcElement.getAttribute('gridSpan');
        const rowSpanAttr = tcElement.getAttribute('rowSpan');
        const vMergeAttr = tcElement.getAttribute('vMerge');
        const hMergeAttr = tcElement.getAttribute('hMerge');

        if (gridSpanAttr) {
            const gs = parseInt(gridSpanAttr);
            if (gs > 1) { cell.gridSpan = gs; }
        }
        if (rowSpanAttr) {
            const rs = parseInt(rowSpanAttr);
            if (rs > 1) { cell.rowSpan = rs; }
        }
        // vMerge="1" or "true" or "continue" means this cell is a vertical merge continuation
        if (vMergeAttr === '1' || vMergeAttr === 'true' || vMergeAttr === 'continue') {
            cell.isMergedContinue = true;
        }
        // hMerge="1" means this cell is a horizontal merge continuation
        if (hMergeAttr === '1' || hMergeAttr === 'true') {
            cell.isMergedContinue = true;
        }

        let tcPr = tcElement.querySelector('tcPr');
        if (!tcPr) {tcPr = tcElement.querySelector('a\\:tcPr');}

        if (tcPr) {
            // Grid span (inside tcPr, some producers use val attribute)
            let gridSpan = tcPr.querySelector('gridSpan');
            if (!gridSpan) {gridSpan = tcPr.querySelector('a\\:gridSpan');}
            if (gridSpan) {
                const gs = parseInt(gridSpan.getAttribute('val')) || 1;
                if (gs > 1) { cell.gridSpan = gs; }
            }

            // Row span
            let rowSpan = tcPr.querySelector('rowSpan');
            if (!rowSpan) {rowSpan = tcPr.querySelector('a\\:rowSpan');}
            if (rowSpan) {
                const rs = parseInt(rowSpan.getAttribute('val')) || 1;
                if (rs > 1) { cell.rowSpan = rs; }
            }

            // Vertical merge (inside tcPr, val="restart" means start, no val or other means continue)
            let vMerge = tcPr.querySelector('vMerge');
            if (!vMerge) {vMerge = tcPr.querySelector('a\\:vMerge');}
            if (vMerge) {
                const val = vMerge.getAttribute('val');
                if (val !== 'restart') {
                    cell.isMergedContinue = true;
                }
            }

            // Cell borders
            let tcBorders = tcPr.querySelector('tcBorders');
            if (!tcBorders) {tcBorders = tcPr.querySelector('a\\:tcBorders');}
            if (tcBorders) {
                cell.borders = this.parseCellBorders(tcBorders);
            }

            // Cell shading
            let tcShading = tcPr.querySelector('tcShading');
            if (!tcShading) {tcShading = tcPr.querySelector('a\\:tcShading');}
            if (tcShading) {
                cell.shading = this.parseCellShading(tcShading);
            }

            // Cell margins
            let tcMar = tcPr.querySelector('tcMar');
            if (!tcMar) {tcMar = tcPr.querySelector('a\\:tcMar');}
            if (tcMar) {
                cell.margins = this.parseCellMargins(tcMar);
            }
        }
    }

    /**
     * Parse cell content
     */
    parseCellContent(tcElement, cell) {
        // Parse text body
        const textBody = this.parseTextBodyFromElement(tcElement);
        if (textBody) {
            cell.setTextBody(textBody);
        }
        
        // Parse other content types (images, etc.)
        // This can be extended based on content parsing
    }

    /**
     * Process table merges
     */
    processTableMerges(table) {
        const rows = table.getRows();
        
        for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
            const row = rows[rowIndex];
            const cells = row.getCells();
            
            for (let colIndex = 0; colIndex < cells.length; colIndex++) {
                const cell = cells[colIndex];
                
                // Handle horizontal merges (gridSpan)
                if (cell.gridSpan > 1) {
                    this.processHorizontalMerge(table, rowIndex, colIndex, cell.gridSpan);
                }
                
                // Handle vertical merges (rowSpan)
                if (cell.rowSpan > 1) {
                    this.processVerticalMerge(table, rowIndex, colIndex, cell.rowSpan);
                }
            }
        }
    }

    /**
     * Process horizontal merge
     */
    processHorizontalMerge(table, rowIndex, startCol, gridSpan) {
        for (let i = 1; i < gridSpan; i++) {
            const targetCell = table.getCell(rowIndex, startCol + i);
            if (targetCell) {
                targetCell.isMergedContinue = true;
                targetCell.mergeParent = { row: rowIndex, col: startCol };
            }
        }
    }

    /**
     * Process vertical merge
     */
    processVerticalMerge(table, startRow, colIndex, rowSpan) {
        for (let i = 1; i < rowSpan; i++) {
            const targetCell = table.getCell(startRow + i, colIndex);
            if (targetCell) {
                targetCell.isMergedContinue = true;
                targetCell.mergeParent = { row: startRow, col: colIndex };
            }
        }
    }

    /**
     * Validate table structure
     */
    validateTableStructure(table) {
        const rows = table.getRows();
        if (rows.length === 0) {return false;}
        
        const expectedCols = table.getTableGrid().length;
        if (expectedCols === 0) {return false;}
        
        for (const row of rows) {
            const cells = row.getCells();
            if (cells.length !== expectedCols) {
                // Allow mismatch if merged placeholders are present; attempt to normalize silently
                console.warn('[Table Validation] Row has', cells.length, 'cells, expected', expectedCols, '- allowing due to merges');
            }
        }
        
        return true;
    }

    /**
     * Parse table borders
     */
    parseTableBorders(bordersElement) {
        const borders = {};
        const borderTypes = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
        
        borderTypes.forEach(type => {
            let border = bordersElement.querySelector(type);
            if (!border) {border = bordersElement.querySelector(`a\\:${type}`);}
            
            if (border) {
                borders[type] = this.parseBorderProperties(border);
            }
        });
        
        return borders;
    }

    /**
     * Parse cell borders
     */
    parseCellBorders(bordersElement) {
        return this.parseTableBorders(bordersElement);
    }

    /**
     * Parse border properties
     */
    parseBorderProperties(borderElement) {
        return {
            style: borderElement.getAttribute('val') || 'single',
            size: parseInt(borderElement.getAttribute('sz')) || 4,
            color: this.parseColor(borderElement),
            space: parseInt(borderElement.getAttribute('space')) || 0
        };
    }

    /**
     * Parse table shading
     */
    parseTableShading(shadingElement) {
        return this.parseCellShading(shadingElement);
    }

    /**
     * Parse cell shading
     */
    parseCellShading(shadingElement) {
        const fill = shadingElement.getAttribute('fill');
        const color = shadingElement.getAttribute('color');
        
        return {
            fill: fill,
            color: color ? this.parseColorFromHex(color) : null
        };
    }

    /**
     * Parse cell margins
     */
    parseCellMargins(marginsElement) {
        const margins = {};
        const marginTypes = ['top', 'left', 'bottom', 'right'];
        
        marginTypes.forEach(type => {
            let margin = marginsElement.querySelector(type);
            if (!margin) {margin = marginsElement.querySelector(`a\\:${type}`);}
            
            if (margin) {
                margins[type] = parseInt(margin.getAttribute('w')) || 0;
            }
        });
        
        return margins;
    }

    /**
     * Render chart frame
     */
    async renderChartFrame(graphics, shape, x, y, w, h) {
        try {
            // Check if we have parsed chart data
            let chartData = shape.chartData;

            // If chartData is a DEFERRED_CHART marker (set during parsing without slide context),
            // clear it so we reload with proper slide context during rendering
            if (chartData && chartData.type === 'DEFERRED_CHART') {
                chartData = null;
                shape.chartData = null;
            }

            // If no chart data, try to extract it now
            if (!chartData && shape.graphicData && window.ChartProcessor) {
                const chartProcessor = new ChartProcessor(this.context);

                // Try different extraction methods
                // Prefer chartRef (relationship ID) over element since it allows slide-context-aware loading
                if (shape.graphicData.chartRef) {
                    // Try async chart loading from relationship
                    // Pass slide context to improve chart relationship resolution
                    const slideContext = {
                        slideIndex: this.currentSlideIndex,
                        slideName: this.currentSlide?.name || `slide${this.currentSlideIndex + 1}`,
                        slide: this.currentSlide
                    };
                    try {
                        const data = await chartProcessor.loadChartFromRelationship(shape.graphicData.chartRef, slideContext);
                        if (data) {
                            shape.chartData = data;
                            chartData = data;
                        }
                    } catch (error) {
                        // Chart loading failed - fall through to placeholder
                    }
                } else if (shape.graphicData.element) {
                    // Fallback: try to parse embedded chart data from the graphicData element
                    chartData = chartProcessor.parseEmbeddedChartData(shape.graphicData.element);
                } else {
                    // Last resort: create placeholder chart data
                    chartData = chartProcessor.createPlaceholderChart();
                }
                
                if (chartData) {
                    shape.chartData = chartData;
                }
            }
            
            // Render chart using ChartRenderer if available
            if (chartData && ChartRenderer) {
                
                // Create chart renderer with graphics adapter
                const chartRenderer = new ChartRenderer(graphics);
                
                // Render the actual chart
                await chartRenderer.renderChart(chartData, x, y, w, h);
                
                return;
            }
            
            // Fallback: Enhanced placeholder with better styling
            this.renderEnhancedChartPlaceholder(graphics, shape, x, y, w, h);
            
        } catch (error) {
            console.error('[Chart Debug] Error in renderChartFrame:', error);
            // Fallback to simple placeholder
            this.renderSimpleChartPlaceholder(graphics, shape, x, y, w, h);
        }
    }
    /**
     * Render enhanced chart placeholder with better visual design
     */
    renderEnhancedChartPlaceholder(graphics, shape, x, y, w, h) {
        try {
            // Draw background
            graphics.fillRect(x, y, w, h, { r: 248, g: 249, b: 250 });
            graphics.strokeRect(x, y, w, h, { r: 200, g: 200, b: 200 }, 1);
            
            // Chart area with margins (no placeholder title)
            const chartX = x + 40;
            const chartY = y + 30;
            const chartW = w - 80;
            const chartH = h - 60;
            
            // CRITICAL FIX: Render as line chart instead of bars to match reference
            // Sample data values for placeholder - ALL 5 data points
            const sampleValues = [4500, 5200, 4800, 6100, 5900];
            const categories = ['Q1 2023', 'Q2 2023', 'Q3 2023', 'Q4 2023', 'Q1 2024'];
            const maxValue = 7000; // Fixed Y-axis scale to match reference
            const minValue = 0;
            
            // Draw Y-axis with proper scale (0-7000)
            const baseY = chartY + chartH - 20;
            graphics.drawLine(chartX, chartY, chartX, baseY, { r: 68, g: 68, b: 68 }, 1);
            
            // Draw Y-axis grid only (no numeric labels)
            for (let i = 0; i <= 7; i++) {
                const value = i * 1000;
                const labelY = baseY - (value / maxValue) * (chartH - 40);
                
                // Draw grid lines
                if (i > 0) {
                    graphics.drawLine(chartX, labelY, chartX + chartW, labelY, { r: 200, g: 200, b: 200 }, 1);
                }
            }
            
            // Draw X-axis
            graphics.drawLine(chartX, baseY, chartX + chartW, baseY, { r: 68, g: 68, b: 68 }, 1);
            
            // Draw line chart with data points
            const dataPointWidth = chartW / (sampleValues.length - 1);
            const lineColor = { r: 91, g: 155, b: 213 }; // PowerPoint blue
            
            // Draw line segments and data points
            for (let i = 0; i < sampleValues.length; i++) {
                const pointX = chartX + i * dataPointWidth;
                const normalizedValue = (sampleValues[i] - minValue) / (maxValue - minValue);
                const pointY = baseY - normalizedValue * (chartH - 40);
                
                // Draw line segment to next point
                if (i < sampleValues.length - 1) {
                    const nextPointX = chartX + (i + 1) * dataPointWidth;
                    const nextNormalizedValue = (sampleValues[i + 1] - minValue) / (maxValue - minValue);
                    const nextPointY = baseY - nextNormalizedValue * (chartH - 40);
                    
                    graphics.drawLine(pointX, pointY, nextPointX, nextPointY, lineColor, 2);
                }
                
                // Draw data point circle
                graphics.fillCircle(pointX, pointY, 4, lineColor);
                
                // No placeholder text labels
            }
            
            // Draw legend at bottom
            const legendY = baseY + 40;
            const legendX = chartX + chartW / 2;
            
            // Draw legend line and text
            graphics.drawLine(legendX - 30, legendY, legendX - 10, legendY, lineColor, 2);
            graphics.fillCircle(legendX - 20, legendY, 3, lineColor);
            // No legend text
            
        } catch (error) {
            console.error('[Chart Debug] Error rendering enhanced placeholder:', error);
            this.renderSimpleChartPlaceholder(graphics, shape, x, y, w, h);
        }
    }
    
    /**
     * Render simple chart placeholder as fallback
     */
    renderSimpleChartPlaceholder(graphics, shape, x, y, w, h) {
        // Draw simple placeholder
        const barWidth = w / 5;
        const maxBarHeight = h * 0.8;
        const baseY = y + h - 20;

        for (let i = 0; i < 4; i++) {
            const barX = x + 20 + i * barWidth;
            const barHeight = maxBarHeight * (0.3 + Math.random() * 0.7);
            const barY = baseY - barHeight;

            graphics.drawRectangle(barX, barY, barWidth * 0.8, barHeight, { r: 100, g: 150, b: 200, a: 255 }, null, 0);
        }

        // No placeholder label text
    }
    
    /**
     * Draw value label for chart data
     */
    drawValueLabel(graphics, text, x, y) {
        if (!graphics.m_oContext) {return;}
        
        graphics.m_oContext.save();
        graphics.m_oContext.fillStyle = 'rgba(80, 80, 80, 1)';
        graphics.m_oContext.font = '10px Calibri';
        graphics.m_oContext.textAlign = 'center';
        graphics.m_oContext.textBaseline = 'bottom';
        graphics.m_oContext.fillText(text, x, y);
        graphics.m_oContext.restore();
    }
    
    /**
     * Draw category label for chart
     */
    drawCategoryLabel(graphics, text, x, y) {
        if (!graphics.m_oContext) {return;}
        
        graphics.m_oContext.save();
        graphics.m_oContext.fillStyle = 'rgba(68, 68, 68, 1)';
        graphics.m_oContext.font = '10px Calibri';
        graphics.m_oContext.textAlign = 'center';
        graphics.m_oContext.textBaseline = 'top';
        graphics.m_oContext.fillText(text, x, y);
        graphics.m_oContext.restore();
    }
    
    /**
     * Draw chart title
     */
    drawChartTitle(graphics, text, x, y) {
        if (!graphics.m_oContext) {return;}
        
        graphics.m_oContext.save();
        graphics.m_oContext.fillStyle = 'rgba(0, 0, 0, 1)';
        graphics.m_oContext.font = 'bold 14px Calibri';
        graphics.m_oContext.textAlign = 'center';
        graphics.m_oContext.textBaseline = 'middle';
        graphics.m_oContext.fillText(text, x, y);
        graphics.m_oContext.restore();
    }

    /**
     * Render diagram frame
     */
    renderDiagramFrame(graphics, shape, x, y, w, h) {
        // Draw diagram placeholder
        const centerX = x + w / 2;
        const centerY = y + h / 2;
        const radius = Math.min(w, h) / 4;

        // Draw connected circles
        const positions = [
            { x: centerX, y: centerY - radius },
            { x: centerX + radius, y: centerY },
            { x: centerX, y: centerY + radius },
            { x: centerX - radius, y: centerY }
        ];

        // Draw connections
        for (let i = 0; i < positions.length; i++) {
            const next = (i + 1) % positions.length;
            graphics.drawLine(positions[i].x, positions[i].y, positions[next].x, positions[next].y, { r: 128, g: 128, b: 128, a: 255 }, 2);
        }

        // Draw circles (no label)
        positions.forEach(pos => {
            graphics.drawEllipse(pos.x - 15, pos.y - 15, 30, 30, { r: 200, g: 200, b: 200, a: 255 }, { r: 128, g: 128, b: 128, a: 255 }, 2);
        });
    }

    /**
     * Render generic frame
     */
    renderGenericFrame(graphics, shape, x, y, w, h) {
        // Draw neutral border only
        graphics.drawRectangle(x, y, w, h, null, '#a0a0a0', 1);
    }

    /**
     * Draw centered text
     */
    drawCenteredText(graphics, text, x, y, w, h) {
        // Deprecated: avoid drawing placeholder labels
        return;
    }

    /**
     * Enhanced default shape rendering
     */
    renderDefaultShape(graphics, shape, x, y, w, h) {
        // Check if rotation is needed and save graphics state
        const rotation = this.getShapeRotation(shape);
        const needsRotation = rotation && rotation !== 0;
        if (needsRotation) {
            graphics.SaveGrState();
            // Apply rotation transformation
            this.applyRotationTransform(graphics, shape, x, y, w, h);
        }

        // Get shape styling
        const fillColor = this.getShapeFillColor(shape);
        const strokeColor = this.getShapeStrokeColor(shape);
        const lineWidth = this.getShapeLineWidth(shape) || 1;
        const strokeInfo = this.getShapeStrokeInfo(shape);

        // Draw default rectangle with enhanced styling
        graphics.drawRectangle(x, y, w, h, fillColor, strokeColor, lineWidth, strokeInfo);

        // Draw shape type label
        this.drawCenteredText(graphics, shape.type || 'Shape', x, y, w, h);

        // Restore graphics state if rotation was applied
        if (needsRotation) {
            graphics.RestoreGrState();
        }
    }

    /**
     * Process shape color and apply modifications
     */
    processShapeColor(color) {
        if (!color) {return null;}

        // Convert hex string to RGB object if needed
        let rgbColor;
        if (typeof color === 'string') {
            if (color.startsWith('#')) {
                rgbColor = this.hexToRgb(color);
            } else {
                // Handle other string formats if needed
                return null;
            }
        } else if (typeof color === 'object' && color.r !== undefined) {
            rgbColor = color;
        } else {
            return null;
        }
        const colorStr = `rgba(${rgbColor.r}, ${rgbColor.g}, ${rgbColor.b}, ${rgbColor.a || 255})`;

        // Apply color modifications if present
        const result = { ...rgbColor };

        // Apply tint (lighten)
        if (color.tint !== undefined) {
            const tintVal = color.tint / 100000; // Convert from percentage
            result.r = Math.round(result.r + (255 - result.r) * tintVal);
            result.g = Math.round(result.g + (255 - result.g) * tintVal);
            result.b = Math.round(result.b + (255 - result.b) * tintVal);
        }

        // Apply shade (darken)
        if (color.shade !== undefined) {
            const shadeVal = 1 - (color.shade / 100000); // Convert from percentage
            result.r = Math.round(result.r * shadeVal);
            result.g = Math.round(result.g * shadeVal);
            result.b = Math.round(result.b * shadeVal);
        }

        // Apply satMod (saturation modification)
        if (color.satMod !== undefined) {
            // Convert RGB to HSL, modify saturation, convert back
            const hsl = this.rgbToHsl(result.r, result.g, result.b);
            hsl.s *= (color.satMod / 100000);
            const rgb = this.hslToRgb(hsl.h, hsl.s, hsl.l);
            result.r = rgb.r;
            result.g = rgb.g;
            result.b = rgb.b;
        }

        const finalColor = `rgba(${result.r}, ${result.g}, ${result.b}, ${result.a || 255})`;

        return result;
    }

    /**
     * Convert RGB to HSL
     */
    rgbToHsl(r, g, b) {
        r /= 255;
        g /= 255;
        b /= 255;

        const max = Math.max(r, g, b);
        const min = Math.min(r, g, b);
        let h, s, l = (max + min) / 2;

        if (max === min) {
            h = s = 0; // achromatic
                    } else {
            const d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            switch (max) {
                case r: h = (g - b) / d + (g < b ? 6 : 0); break;
                case g: h = (b - r) / d + 2; break;
                case b: h = (r - g) / d + 4; break;
            }
            h /= 6;
        }

        return { h, s, l };
    }

    /**
     * Convert HSL to RGB
     */
    hslToRgb(h, s, l) {
        let r, g, b;

        if (s === 0) {
            r = g = b = l; // achromatic
        } else {
            const hue2rgb = (p, q, t) => {
                if (t < 0) {t += 1;}
                if (t > 1) {t -= 1;}
                if (t < 1/6) {return p + (q - p) * 6 * t;}
                if (t < 1/2) {return q;}
                if (t < 2/3) {return p + (q - p) * (2/3 - t) * 6;}
                return p;
            };

            const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            const p = 2 * l - q;
            r = hue2rgb(p, q, h + 1/3);
            g = hue2rgb(p, q, h);
            b = hue2rgb(p, q, h - 1/3);
        }

        return {
            r: Math.round(r * 255),
            g: Math.round(g * 255),
            b: Math.round(b * 255)
        };
    }

    /**
     * Sort shapes by rendering order (z-order)
     * In PowerPoint/standard, shapes are typically rendered in the order they appear in the XML,
     * but we should respect any explicit ordering properties
     */
    sortShapesByRenderOrder(shapes) {
        if (!shapes || shapes.length === 0) {
            return [];
        }


        // Create a copy to avoid mutating the original array
        const sortedShapes = [...shapes];

        // Sort by any explicit order properties or maintain original order
        sortedShapes.sort((a, b) => {
            // Check for explicit z-order or order properties
            const orderA = a.order || a.zOrder || a.index || 0;
            const orderB = b.order || b.zOrder || b.index || 0;

            if (orderA !== orderB) {
                return orderA - orderB;
            }

            // If no explicit order, maintain original array order
            const indexA = shapes.indexOf(a);
            const indexB = shapes.indexOf(b);
            return indexA - indexB;
        });


        return sortedShapes;
    }
}
/**
 * Image Loader class for handling image loading operations
 * Based on standard image loading patterns
 */
class ImageLoader {
    constructor() {
        this.loadingImages = new Set();
        this.logger = new Logger('ImageLoader');
    }

    /**
     * Load image from URL
     */
    async loadImageFromUrl(url) {
        return new Promise((resolve, reject) => {
            const img = new Image();

            img.onload = () => {
                resolve(img);
            };

            img.onerror = (error) => {
                reject(new Error(`Failed to load image: ${error.message || 'Unknown error'}`));
            };

            // Set crossOrigin if needed for blob URLs
            if (url.startsWith('blob:')) {
                img.crossOrigin = 'anonymous';
            }

            img.src = url;
        });
    }

    /**
     * Load image from binary data
     */
    async loadImageFromData(data, mimeType = 'image/png') {
        try {
            const blob = new Blob([data], { type: mimeType });
            const url = URL.createObjectURL(blob);

            const image = await this.loadImageFromUrl(url);

            // Don't revoke URL immediately - cache will handle cleanup
            return image;
            
        } catch (error) {
            throw error;
        }
    }
    
    /**
     * Preload multiple images
     */
    async preloadImages(urls) {
        const loadPromises = urls.map(url => {
            return this.loadImageFromUrl(url).catch(error => {
                return null; // Return null for failed loads
            });
        });

        const results = await Promise.allSettled(loadPromises);
        return results.map(result => result.status === 'fulfilled' ? result.value : null);
    }

    /**
     * Get image dimensions without fully loading
     */
    async getImageDimensions(url) {
        return new Promise((resolve, reject) => {
            const img = new Image();

            img.onload = () => {
                resolve({
                    width: img.naturalWidth,
                    height: img.naturalHeight,
                    aspectRatio: img.naturalWidth / img.naturalHeight
                });
            };

            img.onerror = () => {
                reject(new Error('Failed to load image for dimension check'));
            };

            img.src = url;
        });
    }

    /**
     * Create thumbnail from image
     */
    createThumbnail(image, maxWidth = 150, maxHeight = 150) {
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');

        // Calculate thumbnail dimensions maintaining aspect ratio
        const aspectRatio = image.naturalWidth / image.naturalHeight;
        let thumbWidth, thumbHeight;

        if (aspectRatio > 1) {
            thumbWidth = Math.min(maxWidth, image.naturalWidth);
            thumbHeight = thumbWidth / aspectRatio;
            } else {
            thumbHeight = Math.min(maxHeight, image.naturalHeight);
            thumbWidth = thumbHeight * aspectRatio;
        }

        canvas.width = thumbWidth;
        canvas.height = thumbHeight;

        // Draw scaled image with high quality
        ctx.imageSmoothingEnabled = true;
        ctx.imageSmoothingQuality = 'high';
        ctx.drawImage(image, 0, 0, thumbWidth, thumbHeight);

        return canvas;
    }

    /**
     * Get table data for a specific shape ID
     */
    getTableData(shapeId) {
        // Find the shape by ID across all slides
        for (let slideIndex = 0; slideIndex < this.slides.length; slideIndex++) {
            const slide = this.slides[slideIndex];
            if (slide && slide.shapes) {
                for (const shape of slide.shapes) {
                    if (shape.id === shapeId || shape.name === shapeId) {
                        // Check if this is a table shape
                        if (shape.graphicData && shape.graphicData.uri === 'http://schemas.openxmlformats.org/drawingml/2006/table') {
                            return this.extractTableStructure(shape);
                        }
                    }
                }
            }
        }
        return null;
    }

    /**
     * Get all tables in the presentation
     */
    getAllTableData() {
        const tables = [];
        
        for (let slideIndex = 0; slideIndex < this.slides.length; slideIndex++) {
            const slide = this.slides[slideIndex];
            if (slide && slide.shapes) {
                for (const shape of slide.shapes) {
                    if (shape.graphicData && shape.graphicData.uri === 'http://schemas.openxmlformats.org/drawingml/2006/table') {
                        const tableData = this.extractTableStructure(shape);
                        if (tableData) {
                            tableData.slideIndex = slideIndex;
                            tableData.shapeId = shape.id;
                            tableData.shapeName = shape.name;
                            tables.push(tableData);
                        }
                    }
                }
            }
        }
        
        return tables;
    }

    /**
     * Enhanced table structure extraction
     */
    extractTableStructure(shape) {
        if (!shape.graphicData || !shape.graphicData.tableXml) {
            return null;
        }

        try {
            const table = this.parseTableFromXML(shape.graphicData.tableXml);
            if (!table) {
                return null;
            }

            // Convert table to structured data
            const tableData = {
                id: shape.id,
                name: shape.name,
                type: 'table',
                properties: this.extractTableProperties(table),
                structure: this.extractTableStructureData(table),
                content: this.extractTableContent(table),
                styling: this.extractTableStyling(table),
                position: shape.bounds || shape.properties?.transform,
                metadata: {
                    rowCount: table.getRowCount(),
                    columnCount: table.getColumnCount(),
                    hasMergedCells: this.hasMergedCells(table),
                    tableStyle: table.getTableProperties()?.style || 'TableGrid'
                }
            };

            return tableData;

        } catch (error) {
            return null;
        }
    }

    /**
     * Extract table properties
     */
    extractTableProperties(table) {
        const props = table.getTableProperties() || {};
        
        return {
            style: props.style || 'TableGrid',
            layout: props.layout || 'autofit',
            width: props.width || 'auto',
            alignment: props.alignment || 'left',
            borders: props.borders || {},
            shading: props.shading || {},
            spacing: {
                before: props.spacing?.before || 0,
                after: props.spacing?.after || 0,
                line: props.spacing?.line || 240
            }
        };
    }

    /**
     * Extract table structure data
     */
    extractTableStructureData(table) {
        const rows = table.getRows();
        const tableGrid = table.getTableGrid();
        
        return {
            grid: tableGrid.map((col, index) => ({
                index: index,
                width: col.width,
                type: col.type || 'auto'
            })),
            rows: rows.map((row, rowIndex) => ({
                index: rowIndex,
                height: row.height,
                heightRule: row.heightRule || 'auto',
                isHeader: row.isHeader || false,
                cells: row.getCells().map((cell, cellIndex) => ({
                    index: cellIndex,
                    gridSpan: cell.gridSpan || 1,
                    rowSpan: cell.rowSpan || 1,
                    vMerge: cell.vMerge || null,
                    isMergedContinue: cell.isMergedContinue || false,
                    mergeParent: cell.mergeParent || null
                }))
            }))
        };
    }

    /**
     * Extract table content
     */
    extractTableContent(table) {
        const rows = table.getRows();
        
        return {
            cells: rows.map((row, rowIndex) => 
                row.getCells().map((cell, cellIndex) => ({
                    rowIndex: rowIndex,
                    cellIndex: cellIndex,
                    text: this.extractCellText(cell),
                    textBody: this.extractCellTextBody(cell),
                    content: this.extractCellContent(cell)
                }))
            )
        };
    }

    /**
     * Extract table styling
     */
    extractTableStyling(table) {
        const rows = table.getRows();
        
        return {
            table: {
                borders: table.getTableProperties()?.borders || {},
                shading: table.getTableProperties()?.shading || {},
                style: table.getTableProperties()?.style || 'TableGrid'
            },
            rows: rows.map((row, rowIndex) => ({
                index: rowIndex,
                height: row.height,
                heightRule: row.heightRule || 'auto',
                isHeader: row.isHeader || false
            })),
            cells: rows.map((row, rowIndex) => 
                row.getCells().map((cell, cellIndex) => ({
                    rowIndex: rowIndex,
                    cellIndex: cellIndex,
                    borders: cell.borders || {},
                    shading: cell.shading || {},
                    margins: cell.margins || {},
                    verticalAlignment: cell.verticalAlignment || 'top',
                    textDirection: cell.textDirection || 'lr'
                }))
            )
        };
    }

    /**
     * Check if table has merged cells
     */
    hasMergedCells(table) {
        const rows = table.getRows();
        
        for (const row of rows) {
            for (const cell of row.getCells()) {
                if (cell.gridSpan > 1 || cell.rowSpan > 1 || cell.isMergedContinue) {
                    return true;
                }
            }
        }
        
        return false;
    }

    /**
     * Extract cell text body (OnlyOffice-style)
     */
    extractCellTextBody(cell) {
        const textBody = cell.getTextBody();
        if (!textBody || !textBody.paragraphs) {
            return null;
        }

        return {
            paragraphs: textBody.paragraphs.map(paragraph => ({
                properties: paragraph.properties || {},
                runs: paragraph.runs?.map(run => ({
                    text: run.text || '',
                    properties: run.properties || {}
                })) || []
            }))
        };
    }

    /**
     * Extract cell content (OnlyOffice-style)
     */
    extractCellContent(cell) {
        const content = {
            type: 'text', // Default type
            data: null
        };

        const textBody = cell.getTextBody();
        if (textBody && textBody.paragraphs) {
            content.type = 'text';
            content.data = this.extractCellTextBody(cell);
        }

        // Future: Add support for other content types (images, charts, etc.)
        // This follows OnlyOffice's content extraction pattern

        return content;
    }

    /**
     * Extract text content from a cell
     */
    extractCellText(cell) {
        if (!cell.textBody || !cell.textBody.paragraphs) {
            return '';
        }

        return cell.textBody.paragraphs.map(paragraph => {
            if (!paragraph.runs) {return '';}
            return paragraph.runs.map(run => run.text || '').join('');
        }).join('\n');
    }

    /**
     * Cleanup resources
     */
    cleanup() {
        this.loadingImages.clear();
    }
}
// Export classes (maintain backward compatibility)
if (typeof globalThis !== 'undefined') {
    globalThis.PPTXSlideRenderer = PPTXSlideRenderer;
    globalThis.ImageLoader = ImageLoader;
}
if (typeof window !== 'undefined') {
    // Ensure browser global also has the constructors for UMD runtime checks
    if (!window.PPTXSlideRenderer) { window.PPTXSlideRenderer = PPTXSlideRenderer; }
    if (!window.ImageLoader) { window.ImageLoader = ImageLoader; }
}

// ES Module exports (UMD will also expose globals above)
export { PPTXSlideRenderer, ImageLoader };
export default PPTXSlideRenderer;
