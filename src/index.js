/**
 * PptxViewJS Library
 * Main entry point for the PowerPoint viewer library
 * 
 * Copyright 2025 gptsci.com
 * Licensed under the MIT License
 */

// Side-effect imports to include all modules that register globals
import './utils/utils.js';
import './utils/font-config.js';
// Import as side-effect AND named to retain symbols in bundles
import './parsers/zip-processor.js';
// Ensure XML parser symbols are retained and available on globals
import './parsers/xml-parser.js';
// Ensure DOM classes are retained and available on globals
import './dom/presentation-dom.js';
import './processors/table-processor.js';
import './processors/chart-processor.js';
import './renderers/chart-renderer.js';
import './renderers/chartjs-renderer.js';
import './graphics/graphics-adapter.js';
import './graphics/graphics-engine.js';
import './graphics/font-engine.js';
import './processors/media-processor.js';
import './processors/svg-renderer.js';
import './processors/theme-processor.js';
import PPTXSlideRendererClass from './processors/slide-renderer.js';
import PPTXProcessorClass from './processors/PPTXProcessor.js';
// Bundle externals for UMD consumers
import ChartModule from 'chart.js/auto';
import JSZipModule from 'jszip';
// Avoid importing package.json to maintain compatibility with parsers/linters
let LIB_VERSION = '1.1.8';
try {
    if (typeof process !== 'undefined' && process.env && process.env.npm_package_version) {
        LIB_VERSION = process.env.npm_package_version;
    }
} catch (_err) {}

const resolveDependency = (moduleValue, globalName) => {
    if (moduleValue) {return moduleValue;}
    if (typeof globalThis !== 'undefined' && globalThis[globalName]) {return globalThis[globalName];}
    if (typeof window !== 'undefined' && window[globalName]) {return window[globalName];}
    throw new Error(`${globalName} is required but was not found. Install and import "${globalName === 'Chart' ? 'chart.js' : 'jszip'}" in your application.`);
};

const Chart = resolveDependency(ChartModule, 'Chart');
const JSZip = resolveDependency(JSZipModule, 'JSZip');
// Note: side-effect imports only; classes are exposed on window
// and used via globals for UMD and script-tag compatibility.
// import { PPTXProcessor as PPTXProcessorClass } from './processors/PPTXProcessor.js';
// import { DefaultSlideRenderer } from './processors/slide-renderer.js';


// Provide minimal fallbacks for required globals if missing
if (typeof globalThis !== 'undefined') {
    if (!globalThis.Logger) {
        class LoggerFallback {
            constructor() {}
            log() {}
            info() {}
            warn() {}
            error() {}
            debug() {}
            trace() {}
            logError() {}
            markTiming() {}
            measureTiming() {}
        }
        globalThis.Logger = LoggerFallback;
    }
    if (!globalThis.PPTXSlideRenderer && typeof window !== 'undefined' && window.PPTXSlideRenderer) {
        globalThis.PPTXSlideRenderer = window.PPTXSlideRenderer;
    }
    if (typeof window !== 'undefined') {
        // Mirror to window in case UMD runtime checks expect it
        if (!window.PPTXSlideRenderer && globalThis.PPTXSlideRenderer) {
            window.PPTXSlideRenderer = globalThis.PPTXSlideRenderer;
        }
    }
    // Mirror from window in case globals attached by side-effect modules
    if (!globalThis.FontEngine && typeof window !== 'undefined' && window.FontEngine) { globalThis.FontEngine = window.FontEngine; }
    if (!globalThis.SimplifiedFontManager && typeof window !== 'undefined' && window.SimplifiedFontManager) { globalThis.SimplifiedFontManager = window.SimplifiedFontManager; }
    if (!globalThis.ZLib && typeof window !== 'undefined' && window.ZLib) { globalThis.ZLib = window.ZLib; }
    if (!globalThis.OpenXmlPackage && typeof window !== 'undefined' && window.OpenXmlPackage) { globalThis.OpenXmlPackage = window.OpenXmlPackage; }
    if (!globalThis.OpenXmlPart && typeof window !== 'undefined' && window.OpenXmlPart) { globalThis.OpenXmlPart = window.OpenXmlPart; }
    if (!globalThis.OpenXmlTypes && typeof window !== 'undefined' && window.OpenXmlTypes) { globalThis.OpenXmlTypes = window.OpenXmlTypes; }
    if (!globalThis.SVGRenderer && typeof window !== 'undefined' && window.SVGRenderer) { globalThis.SVGRenderer = window.SVGRenderer; }
    if (!globalThis.XmlParserContext && typeof window !== 'undefined' && window.XmlParserContext) { globalThis.XmlParserContext = window.XmlParserContext; }
    if (!globalThis.StaxParser && typeof window !== 'undefined' && window.StaxParser) { globalThis.StaxParser = window.StaxParser; }
    if (!globalThis.CPresentation && typeof window !== 'undefined' && window.CPresentation) { globalThis.CPresentation = window.CPresentation; }
    if (!globalThis.CSlide && typeof window !== 'undefined' && window.CSlide) { globalThis.CSlide = window.CSlide; }
    if (!globalThis.CSld && typeof window !== 'undefined' && window.CSld) { globalThis.CSld = window.CSld; }
    if (!globalThis.CShape && typeof window !== 'undefined' && window.CShape) { globalThis.CShape = window.CShape; }
    if (!globalThis.CGroupShape && typeof window !== 'undefined' && window.CGroupShape) { globalThis.CGroupShape = window.CGroupShape; }
    if (!globalThis.CSlideMaster && typeof window !== 'undefined' && window.CSlideMaster) { globalThis.CSlideMaster = window.CSlideMaster; }
    if (!globalThis.CSlideLayout && typeof window !== 'undefined' && window.CSlideLayout) { globalThis.CSlideLayout = window.CSlideLayout; }
    if (!globalThis.CTheme && typeof window !== 'undefined' && window.CTheme) { globalThis.CTheme = window.CTheme; }
    if (!globalThis.CTable && typeof window !== 'undefined' && window.CTable) { globalThis.CTable = window.CTable; }
    // Expose bundled externals if not already present
    if (!globalThis.JSZip && typeof JSZip !== 'undefined') {
        globalThis.JSZip = JSZip;
    }
    // Ensure PPTXProcessor and factory are exposed on globals
    if (!globalThis.PPTXProcessor && typeof PPTXProcessorClass === 'function') {
        globalThis.PPTXProcessor = PPTXProcessorClass;
    }
    if (!globalThis.createPPTXProcessor && typeof PPTXProcessorClass === 'function') {
        globalThis.createPPTXProcessor = (options = {}) => new PPTXProcessorClass(options);
    }
}

if (typeof globalThis !== 'undefined') {
    if (typeof window !== 'undefined') {
        window.CGraphics = window.CGraphics || (globalThis.CGraphics || undefined);
        if (!window.Chart && typeof Chart !== 'undefined') {
            window.Chart = Chart;
        }
        if (!window.PPTXSlideRenderer && typeof globalThis.PPTXSlideRenderer !== 'undefined') {
            window.PPTXSlideRenderer = globalThis.PPTXSlideRenderer;
        }
        if (!window.PPTXProcessor && typeof globalThis.PPTXProcessor !== 'undefined') {
            window.PPTXProcessor = globalThis.PPTXProcessor;
        }
        if (!window.createPPTXProcessor && typeof globalThis.createPPTXProcessor !== 'undefined') {
            window.createPPTXProcessor = globalThis.createPPTXProcessor;
        }
    }
    if (!globalThis.CGraphics && typeof window !== 'undefined' && window.CGraphics) {
        globalThis.CGraphics = window.CGraphics;
    }
}

class PPTXViewer {
    constructor(options = {}) {
        const userOptions = options || {};
        this.options = {
            canvas: userOptions.canvas ?? null,
            debug: userOptions.debug ?? false,
            enableThumbnails: userOptions.enableThumbnails ?? true,
            slideSizeMode: userOptions.slideSizeMode ?? 'fit',
            backgroundColor: userOptions.backgroundColor ?? '#ffffff',
            autoRenderFirstSlide: userOptions.autoRenderFirstSlide ?? true,
            logger: userOptions.logger ?? (typeof window !== 'undefined' ? window.console : console),
            // New: simplify integration by auto-exposing globals for chart relationship resolution
            autoExposeGlobals: userOptions.autoExposeGlobals ?? true,
            // New: schedule one delayed re-render to catch async chart parsing
            autoChartRerenderDelayMs: userOptions.autoChartRerenderDelayMs ?? 200
        };
        Object.assign(this.options, userOptions);

        this.processor = null;
        this.presentation = null;
        this.currentSlideIndex = 0;
        this.slideCount = 0;
        this.isLoaded = false;
        this.eventListeners = {};
        this._initPromise = null;
        this._scheduledPostLoadRerender = false;
    }

    async _initializeProcessor() {
        if (this._initPromise) {return this._initPromise;}
        this._initPromise = (async () => {
            const g = (typeof globalThis !== 'undefined') ? globalThis : (typeof window !== 'undefined' ? window : null);
            const tryResolve = () => {
                const ctor = g ? (g.PPTXProcessor || null) : null;
                const factory = g ? (g.createPPTXProcessor || null) : null;
                return { ctor, factory };
            };
            try {
                // Small retry loop to avoid race conditions in UMD init order
                let attempts = 0;
                let resolved = tryResolve();
                while (!resolved.ctor && !resolved.factory && attempts < 100) {
                    await new Promise(r => setTimeout(r, 20));
                    attempts += 1;
                    resolved = tryResolve();
                }

                if (resolved.ctor && typeof resolved.ctor === 'function') {
                    this.processor = new resolved.ctor();
                } else if (resolved.factory && typeof resolved.factory === 'function') {
                    this.processor = resolved.factory();
                } else {
                    this.processor = null;
                }

                if (!this.processor) {
                    throw new Error('PPTXProcessor is not available. Ensure processors are loaded.');
                }
                if (typeof this.processor.initialize === 'function') {
                    await this.processor.initialize();
                }
                this.emit('processorReady');
                return this.processor;
            } catch (error) {
                this.emit('initError', error);
                throw error;
            }
        })();
        return this._initPromise;
    }

    async loadFile(input, options = {}) {
        // Reset any existing processor state before loading a new file to avoid stale presentations
        if (this.isLoaded && this.processor) {
            try {
                if (typeof this.processor.destroy === 'function') {
                    this.processor.destroy();
                } else if (typeof this.processor.reset === 'function') {
                    this.processor.reset();
                }
            } catch (_resetErr) {}
            this.processor = null;
            this._initPromise = null;
        }

        await this._initializeProcessor();
        try {
            this.emit('loadStart');
            let arrayBuffer;
            if (typeof File !== 'undefined' && input instanceof File) {
                arrayBuffer = await input.arrayBuffer();
            } else if (input instanceof ArrayBuffer) {
                arrayBuffer = input;
            } else if (input instanceof Uint8Array) {
                arrayBuffer = input.buffer.slice(input.byteOffset, input.byteOffset + input.byteLength);
            } else {
                throw new Error('Input must be File, ArrayBuffer, or Uint8Array');
            }
            await this.processor.processFile(arrayBuffer, options);
            this.presentation = this.processor.presentation;
            this.slideCount = this.processor.getSlidesCount();
            this.currentSlideIndex = 0;
            this.isLoaded = true;
            this._scheduledPostLoadRerender = false;
            // Auto expose globals for charts so host pages don't need boilerplate
            if (this.options.autoExposeGlobals) {
                try { this._exposeGlobalsForCharts(); } catch(_e) {}
            }
            this.emit('loadComplete', { slideCount: this.slideCount, presentation: this.presentation });
            return this;
        } catch (error) {
            this.emit('loadError', error);
            throw error;
        }
    }

    async loadFromUrl(url) {
        this.emit('loadStart');
        try {
            const response = await fetch(url);
            if (!response.ok) {throw new Error(`Failed to fetch PPTX: ${response.status} ${response.statusText}`);}            
            const arrayBuffer = await response.arrayBuffer();
            return this.loadFile(arrayBuffer);
        } catch (error) {
            this.emit('loadError', error);
            throw error;
        }
    }

    async render(canvas, options = {}) {
        if (!this.isLoaded) {throw new Error('No PPTX loaded. Call loadFile() first.');}
        const targetCanvas = canvas || this.options.canvas;
        if (!targetCanvas) {throw new Error('Canvas element required. Provide canvas parameter or set in constructor.');}
        const slideIndex = options.slideIndex !== undefined ? options.slideIndex : this.currentSlideIndex;
        if (slideIndex < 0 || slideIndex >= this.slideCount) {throw new Error(`Invalid slide index: ${slideIndex}.`);}        
        try {
            this.emit('renderStart', slideIndex);
            await this.processor.renderSlide(targetCanvas, slideIndex, options);
            this.currentSlideIndex = slideIndex;
            this.emit('renderComplete', slideIndex);
            this.emit('slideChanged', slideIndex);
            // After first render following load, optionally re-render once to catch async chart parsing
            if (this.options.autoChartRerenderDelayMs > 0 && !this._scheduledPostLoadRerender) {
                this._scheduledPostLoadRerender = true;
                const delay = this.options.autoChartRerenderDelayMs;
                setTimeout(() => {
                    try { this.render(targetCanvas, { ...options, slideIndex }); } catch(_e) {}
                }, delay);
            }
            return this;
        } catch (error) {
            this.emit('renderError', error);
            throw error;
        }
    }

    async renderSlide(slideIndex, canvas = null, options = {}) {
        return this.render(canvas, { ...options, slideIndex });
    }

    async nextSlide(canvas = null) {
        if (this.currentSlideIndex < this.slideCount - 1) {
            await this.render(canvas, { slideIndex: this.currentSlideIndex + 1 });
        }
        return this;
    }

    async previousSlide(canvas = null) {
        if (this.currentSlideIndex > 0) {
            await this.render(canvas, { slideIndex: this.currentSlideIndex - 1 });
        }
        return this;
    }

    async goToSlide(slideIndex, canvas = null) {
        return this.render(canvas, { slideIndex });
    }

    getSlideCount() {
        return this.slideCount;
    }

    getCurrentSlideIndex() {
        return this.currentSlideIndex;
    }

    setCanvas(canvas) {
        this.options.canvas = canvas;
        return this;
    }

    on(event, callback) {
        if (!this.eventListeners[event]) {this.eventListeners[event] = [];}        
        this.eventListeners[event].push(callback);
    }

    off(event, callback) {
        if (!this.eventListeners[event]) {return;}
        const index = this.eventListeners[event].indexOf(callback);
        if (index > -1) {this.eventListeners[event].splice(index, 1);}        
    }

    emit(event, ...args) {
        if (this.eventListeners[event]) {
            this.eventListeners[event].forEach(cb => {
                try { cb(...args); } catch (_err) {}
            });
        }
    }

    destroy() {
        this.processor = null;
        this.eventListeners = {};
        this.isLoaded = false;
        this.currentSlideIndex = 0;
        this.slideCount = 0;
    }

    /**
     * Expose globals used by chart processing to simplify host integration
     * - window.currentProcessor: { processor, zip, package, reRenderShape }
     * - window.currentZipData alias
     */
    _exposeGlobalsForCharts() {
        try {
            const proc = this.processor;
            if (!proc) { return; }
            // Resolve effective zip/package across possible processor shapes
            const effectiveZip = proc.zip || proc.processor?.zip || proc.zipProcessor?.zip || null;
            const effectivePackage = proc.package || proc.processor?.package || proc.zipProcessor?.package || null;
            const reRenderShape = () => {
                const slideIndex = (typeof this.getCurrentSlideIndex === 'function') ? this.getCurrentSlideIndex() : (this.currentSlideIndex || 0);
                const canvas = this.options.canvas;
                if (canvas) {
                    this.render(canvas, { slideIndex }).catch(() => {});
                }
            };
            const current = {
                processor: proc,
                zip: effectiveZip,
                package: effectivePackage,
                reRenderShape
            };
            if (typeof window !== 'undefined') {
                window.currentProcessor = current;
                if (effectiveZip) { window.currentZipData = effectiveZip; }
                if (window.PPTXSlideRenderer && effectiveZip) {
                    try { window.PPTXSlideRenderer.currentZip = current.zip; } catch(_e) {}
                }
                // Also mirror onto nested processor object for compatibility with public demo
                try {
                    if (window.currentProcessor && window.currentProcessor.processor) {
                        if (!window.currentProcessor.processor.zip && effectiveZip) {
                            window.currentProcessor.processor.zip = effectiveZip;
                        }
                        if (!window.currentProcessor.processor.package && effectivePackage) {
                            window.currentProcessor.processor.package = effectivePackage;
                        }
                        // Also mirror zipProcessor shape for maximum compatibility
                        if (!window.currentProcessor.processor.zipProcessor) {
                            window.currentProcessor.processor.zipProcessor = {};
                        }
                        if (effectiveZip && !window.currentProcessor.processor.zipProcessor.zip) {
                            window.currentProcessor.processor.zipProcessor.zip = effectiveZip;
                        }
                        if (effectivePackage && !window.currentProcessor.processor.zipProcessor.package) {
                            window.currentProcessor.processor.zipProcessor.package = effectivePackage;
                        }
                    }
                } catch(_e) {}
            }
        } catch(_e) {}
    }
}

// Simple helper to mount a basic viewer with file input and navigation
function mountSimpleViewer({ canvas, fileInput, prevBtn, nextBtn, statusEl, options = {} } = {}) {
    if (!canvas) { throw new Error('canvas is required'); }
    const viewer = new PPTXViewer({ canvas, ...options });
    const updateNav = () => {
        try {
            const total = viewer.getSlideCount();
            const idx = (typeof viewer.getCurrentSlideIndex === 'function') ? viewer.getCurrentSlideIndex() : 0;
            if (prevBtn) { prevBtn.disabled = !(idx > 0); }
            if (nextBtn) { nextBtn.disabled = !(idx < total - 1); }
            if (statusEl) { statusEl.textContent = total ? `Slide ${idx + 1} / ${total}` : ''; }
        } catch(_e) {}
    };
    viewer.on('renderComplete', () => updateNav());
    viewer.on('loadComplete', () => {
        // Ensure globals are exposed for charts in environments where consumers rely on them
        try {
            if (viewer && viewer.processor && options.autoExposeGlobals !== false) {
                const proc = viewer.processor;
                const current = {
                    processor: proc,
                    zip: proc.zip,
                    package: proc.package,
                    reRenderShape: () => {
                        const idx = (typeof viewer.getCurrentSlideIndex === 'function') ? viewer.getCurrentSlideIndex() : 0;
                        viewer.render(canvas, { slideIndex: idx }).catch(() => {});
                    }
                };
                if (typeof window !== 'undefined') {
                    window.currentProcessor = current;
                    if (proc.zip) { window.currentZipData = proc.zip; }
                    if (window.PPTXSlideRenderer && current.zip) {
                        try { window.PPTXSlideRenderer.currentZip = current.zip; } catch(_e) {}
                    }
                }
            }
        } catch(_e) {}
        // Update nav after load
        updateNav();
    });
    if (fileInput) {
        fileInput.addEventListener('change', async () => {
            const file = fileInput.files && fileInput.files[0];
            if (!file) { return; }
            if (statusEl) { statusEl.textContent = 'Loading...'; }
            try {
                await viewer.loadFile(file);
                // Extra safety: expose globals before first render
                try {
                    if (viewer && viewer.processor && options.autoExposeGlobals !== false) {
                        const proc = viewer.processor;
                        const effectiveZip = proc.zip || proc.processor?.zip || proc.zipProcessor?.zip || null;
                        const effectivePackage = proc.package || proc.processor?.package || proc.zipProcessor?.package || null;
                        if (typeof window !== 'undefined') {
                            window.currentProcessor = {
                                processor: proc,
                                zip: effectiveZip,
                                package: effectivePackage,
                                reRenderShape: () => {
                                    const idx = (typeof viewer.getCurrentSlideIndex === 'function') ? viewer.getCurrentSlideIndex() : 0;
                                    viewer.render(canvas, { slideIndex: idx }).catch(() => {});
                                }
                            };
                            if (effectiveZip) { window.currentZipData = effectiveZip; }
                            // Ensure nested alias for compatibility
                            try {
                                if (window.currentProcessor && window.currentProcessor.processor) {
                                    if (!window.currentProcessor.processor.zip && effectiveZip) {
                                        window.currentProcessor.processor.zip = effectiveZip;
                                    }
                                    if (!window.currentProcessor.processor.package && effectivePackage) {
                                        window.currentProcessor.processor.package = effectivePackage;
                                    }
                                    if (!window.currentProcessor.processor.zipProcessor) {
                                        window.currentProcessor.processor.zipProcessor = {};
                                    }
                                    if (effectiveZip && !window.currentProcessor.processor.zipProcessor.zip) {
                                        window.currentProcessor.processor.zipProcessor.zip = effectiveZip;
                                    }
                                    if (effectivePackage && !window.currentProcessor.processor.zipProcessor.package) {
                                        window.currentProcessor.processor.zipProcessor.package = effectivePackage;
                                    }
                                }
                            } catch(_e) {}
                        }
                    }
                } catch(_e) {}
                await viewer.render(canvas, { slideIndex: 0 });
                // Schedule a short delayed re-render to catch charts resolved asynchronously
                const delay = (typeof options.autoChartRerenderDelayMs === 'number') ? options.autoChartRerenderDelayMs : 200;
                if (delay > 0) {
                    setTimeout(() => {
                        const idx = (typeof viewer.getCurrentSlideIndex === 'function') ? viewer.getCurrentSlideIndex() : 0;
                        viewer.render(canvas, { slideIndex: idx }).catch(() => {});
                    }, delay);
                }
            } catch (err) {
                if (statusEl) { statusEl.textContent = 'Error: ' + (err?.message || 'Failed to load'); }
            }
        });
    }
    if (prevBtn) { prevBtn.addEventListener('click', async () => { await viewer.previousSlide(canvas); }); }
    if (nextBtn) { nextBtn.addEventListener('click', async () => { await viewer.nextSlide(canvas); }); }
    // Re-render current slide whenever async chart parsing completes
    if (typeof window !== 'undefined') {
        window.addEventListener('chartRenderingComplete', () => {
            if (!viewer.isLoaded) { return; }
            const idx = (typeof viewer.getCurrentSlideIndex === 'function') ? viewer.getCurrentSlideIndex() : 0;
            viewer.render(canvas, { slideIndex: idx }).catch(() => {});
        });
    }
    return viewer;
}

const api = { PPTXViewer, version: LIB_VERSION, mountSimpleViewer };

// Optionally expose internal classes/utilities on the public namespace for script consumers
// This mirrors what's available on globals (window/globalThis) so that
// PptxViewJS.min.js contains and exposes the same classes as the non-minified build.
const maybeExposeOnApi = (name) => {
    try {
        if (typeof globalThis !== 'undefined' && globalThis[name] && !api[name]) {
            api[name] = globalThis[name];
        }
    } catch (_err) {}
};

[
    // Graphics / Engines
    'CGraphics',
    'FontEngine',
    'SimplifiedFontManager',
    // OpenXML related
    'ZLib',
    'OpenXmlPackage',
    'OpenXmlPart',
    'OpenXmlTypes',
    // XML / Parsing
    'SVGRenderer',
    'XmlParserContext',
    'StaxParser',
    // Presentation DOM classes
    'CPresentation',
    'CSlide',
    'CSld',
    'CShape',
    'CGroupShape',
    'CSlideMaster',
    'CSlideLayout',
    'CTheme',
    'CTable',
    // Processors / Renderers
    'PPTXProcessor',
    'PPTXSlideRenderer',
    // Bundled externals
    'Chart',
    'JSZip'
].forEach(maybeExposeOnApi);

if (typeof globalThis !== 'undefined') {
    // Expose for non-module usage; UMD will also expose under library name
    globalThis.PptxViewJS = api;
    if (!globalThis.PPTXProcessor && typeof PPTXProcessorClass !== 'undefined') {
        globalThis.PPTXProcessor = PPTXProcessorClass;
    }
    if (!globalThis.createPPTXProcessor && typeof PPTXProcessorClass !== 'undefined') {
        globalThis.createPPTXProcessor = (options = {}) => new PPTXProcessorClass(options);
    }
}
    // Ensure SlideRenderer is exposed on globals
    if (!globalThis.PPTXSlideRenderer && typeof PPTXSlideRendererClass === 'function') {
        globalThis.PPTXSlideRenderer = PPTXSlideRendererClass;
    }

export { PPTXViewer };
export default api;
