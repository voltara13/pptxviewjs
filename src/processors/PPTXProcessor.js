/**
 * PPTXProcessor - Enhanced wrapper for the slide renderer functionality
 * Provides a unified interface for PPTX processing with comprehensive error handling
 * 
 * Copyright 2025 gptsci.com
 * Licensed under the MIT License
 */

// Import validation and error handling frameworks
// These should be loaded before this module

// Note: Dependencies are loaded via script tags in index.html
// - utils.js
// - zip-processor.js  
// - xml-parser.js
// - presentation-dom.js
// - graphics-adapter.js
// - graphics-engine.js
// - font-engine.js
// - slide-renderer.js

class PPTXProcessor {
    constructor(options = {}) {
        // Input validation for constructor options
        const validator = this._getValidator();
        const validatedOptions = this._validateConstructorOptions(options);
        
        // Initialize error boundary
        this.errorBoundary = new (this._getErrorBoundary())({
            context: 'PPTXProcessor',
            enableLogging: validatedOptions.enableLogging !== false,
            errorHandler: validatedOptions.errorHandler
        });
        
        // Core properties
        this.processor = null;
        this.presentation = null;
        this.isInitialized = false;
        
        // Enhanced configuration
        this.config = {
            maxFileSize: validatedOptions.maxFileSize || 100 * 1024 * 1024, // 100MB
            timeout: validatedOptions.timeout || 80000, // 80 seconds
            enableValidation: validatedOptions.enableValidation !== false,
            strictMode: validatedOptions.strictMode === true,
            retryOptions: {
                maxAttempts: validatedOptions.maxRetryAttempts || 3,
                baseDelay: validatedOptions.retryDelay || 1000
            },
            ...validatedOptions
        };
        
        // Performance monitoring
        this.performanceMarks = new Map();
        this.memoryUsage = { peak: 0, current: 0 };
        
        // State tracking
        this.processingState = 'idle'; // idle, processing, error, completed
        this.lastError = null;
        this.currentSlideIndex = 0; // Track current slide index
        
        // Initialize monitoring
        this._initializeMonitoring();
    }
    
    /**
     * Initialize the processor with enhanced error handling
     */
    async initialize() {
        if (this.isInitialized) {
            return;
        }
        
        return this.errorBoundary.wrap(async () => {
            this._markPerformance('initialize_start');
            
            try {
                // Validate environment
                this._validateEnvironment();
                
                // Create processor instance with retry logic
                await this._initializeWithRetry();
                
                this.isInitialized = true;
                this.processingState = 'idle';
                this._markPerformance('initialize_end');
                
                if (window.logger) {
                    window.logger.info('PPTXProcessor', 'Processor initialized successfully');
                }
                
            } catch (error) {
                this.processingState = 'error';
                this.lastError = error;
                throw this._enhanceError(error, 'initialization');
            }
        }, { context: 'initialize' })();
    }
    
    /**
     * Ensure processor is initialized before any operation
     */
    async _ensureInitialized() {
        if (!this.isInitialized) {
            await this.initialize();
        }
        if (!this.processor) {
            throw this._createError('Failed to initialize PPTX processor', {
                name: 'ProcessingError',
                code: 'INITIALIZATION_FAILED',
                recoverable: true
            });
        }
    }
    
    /**
     * Process PPTX file from ArrayBuffer with comprehensive validation and error handling
     */
    async processFile(arrayBuffer, options = {}) {
        // Input validation
        const validationResult = this._validateFileInput(arrayBuffer, options);
        if (!validationResult.valid) {
            throw this._createValidationError('Invalid file input', validationResult.errors);
        }
        
        return this.errorBoundary.wrap(async () => {
            this._markPerformance('processFile_start');
            this.processingState = 'processing';
            
            try {
                await this._ensureInitialized();
                
                // Monitor memory usage
                this._updateMemoryUsage();
                
                // Process with timeout and retry logic
                const result = await this._processWithTimeout(
                    () => this.processor.processFile(arrayBuffer),
                    this.config.timeout
                );
                
                this.presentation = this.processor.presentation;
                this.processingState = 'completed';
                this._markPerformance('processFile_end');
                
                // Validate output
                if (this.config.enableValidation) {
                    this._validateProcessingResult(this.presentation);
                }
                
                if (window.logger) {
                    const timing = this._getPerformanceTiming('processFile_start', 'processFile_end');
                    window.logger.info('PPTXProcessor', `File processed successfully in ${timing}ms`);
                }
                
                return this.presentation;
                
            } catch (error) {
                this.processingState = 'error';
                this.lastError = error;
                throw this._enhanceError(error, 'file_processing', { arrayBufferSize: arrayBuffer.byteLength });
            }
        }, { context: 'processFile' })();
    }
    
    /**
     * Load PPTX from ArrayBuffer (alias for compatibility)
     */
    async loadFromArrayBuffer(arrayBuffer) {
        return this.processFile(arrayBuffer);
    }
    
    /**
     * Render slide to canvas with enhanced validation and error handling
     */
    async renderSlide(canvas, slideIndex, options = {}) {
        // Input validation
        const validationResult = this._validateRenderInput(canvas, slideIndex, options);
        if (!validationResult.valid) {
            throw this._createValidationError('Invalid render input', validationResult.errors);
        }
        
        return this.errorBoundary.wrap(async () => {
            this._markPerformance('renderSlide_start');
            
            try {
                await this._ensureInitialized();
                
                // Validate slide index bounds
                const slideCount = this.getSlidesCount();
                if (slideIndex < 0 || slideIndex >= slideCount) {
                    throw this._createError(`Slide index ${slideIndex} out of bounds (0-${slideCount - 1})`, {
                        name: 'ValidationError',
                        code: 'INVALID_SLIDE_INDEX'
                    });
                }
                
                // Monitor memory during rendering
                this._updateMemoryUsage();
                
                // Render with timeout
                const result = await this._processWithTimeout(
                    () => this.processor.renderSlide(canvas, slideIndex),
                    options.timeout || this.config.timeout
                );
                
                // Update current slide index after successful render
                this.currentSlideIndex = slideIndex;
                
                this._markPerformance('renderSlide_end');
                
                if (window.logger) {
                    const timing = this._getPerformanceTiming('renderSlide_start', 'renderSlide_end');
                    window.logger.trace('PPTXProcessor', `Slide ${slideIndex} rendered in ${timing}ms`);
                }
                
                return result;
                
            } catch (error) {
                throw this._enhanceError(error, 'slide_rendering', { slideIndex, canvasSize: `${canvas.width}x${canvas.height}` });
            }
        }, { context: 'renderSlide' })();
    }
    
    /**
     * Get slides count
     */
    getSlidesCount() {
        if (!this.processor || !this.processor.slides) {return 0;}
        return this.processor.slides.length;
    }
    
    /**
     * Get current slide index
     */
    getCurrentSlideIndex() {
        // Return the current slide index (default to 0 if not set)
        return this.currentSlideIndex !== undefined ? this.currentSlideIndex : 0;
    }
    
    /**
     * Get current slide
     */
    getCurrentSlide(slideIndex = null) {
        if (!this.processor || !this.processor.slides) {return null;}
        const index = slideIndex !== null ? slideIndex : this.getCurrentSlideIndex();
        return this.processor.slides[index] || null;
    }
    
    /**
     * Get slide dimensions
     */
    getSlideDimensions() {
        // Return slide dimensions from the underlying processor
        if (!this.processor) {
            // Default PowerPoint slide size (10" x 7.5" in EMUs)
            return { cx: 9144000, cy: 6858000 };
        }
        
        // Try to get from the processor's method
        if (typeof this.processor.getSlideDimensions === 'function') {
            return this.processor.getSlideDimensions();
        }
        
        // Try to get from presentation object
        if (this.processor.presentation && this.processor.presentation.slideSize) {
            return this.processor.presentation.slideSize;
        }
        
        // Default dimensions
        return { cx: 9144000, cy: 6858000 };
    }
    
    /**
     * Clear all caches in the processor
     * This should be called when loading a new PPTX file to prevent data contamination
     */
    clearCaches() {
        this._ensureInitialized();
        
        if (this.processor) {
            // Clear image cache
            if (this.processor.imageCache) {
                // Revoke blob URLs before clearing
                for (const [relId, imageData] of this.processor.imageCache) {
                    if (imageData.url && imageData.url.startsWith('blob:')) {
                        URL.revokeObjectURL(imageData.url);
                    }
                }
                this.processor.imageCache.clear();
            }
            
            // Clear image and media maps
            if (this.processor.imageMap) {
                this.processor.imageMap.clear();
            }
            
            if (this.processor.mediaMap) {
                this.processor.mediaMap.clear();
            }
            
            // Clear any other processor-specific caches
            if (this.processor.clearCaches && typeof this.processor.clearCaches === 'function') {
                this.processor.clearCaches();
            }
            
            // Reset graphics adapter state if available
            if (this.processor.drawingDocument && this.processor.drawingDocument.graphics) {
                this.processor.drawingDocument.graphics.resetState();
            }
            
            // Clear slide renderer presentation and theme data
            if (this.processor.presentation) {
                this.processor.presentation = null;
            }
        }
        
        // Reset presentation reference
        this.presentation = null;
    }

    /**
     * Get presentation object with validation
     */
    getPresentation() {
        if (!this.presentation && this.config.strictMode) {
            throw this._createError('No presentation loaded', {
                name: 'ProcessingError',
                code: 'NO_PRESENTATION_LOADED'
            });
        }
        return this.presentation;
    }
    
    /**
     * Get processing statistics and health metrics
     */
    getProcessingStats() {
        return {
            state: this.processingState,
            isInitialized: this.isInitialized,
            slidesCount: this.getSlidesCount(),
            memoryUsage: { ...this.memoryUsage },
            lastError: this.lastError ? {
                name: this.lastError.name,
                message: this.lastError.message,
                timestamp: this.lastError.timestamp
            } : null,
            performanceMarks: Array.from(this.performanceMarks.entries())
        };
    }
    
    // ===== PRIVATE HELPER METHODS =====
    
    /**
     * Get validator instance (with fallback)
     */
    _getValidator() {
        return (typeof window !== 'undefined' && window.DataValidator) ? 
            window.DataValidator : class { validate() { return { valid: true, errors: [] }; } };
    }
    
    /**
     * Get error boundary class (with fallback)
     */
    _getErrorBoundary() {
        return (typeof window !== 'undefined' && window.ErrorBoundary) ? 
            window.ErrorBoundary : class { 
                constructor() {} 
                wrap(fn) { return fn; }
            };
    }
    
    /**
     * Validate constructor options
     */
    _validateConstructorOptions(options) {
        const validator = this._getValidator();
        
        // Define schema for constructor options
        const schema = {
            type: 'object',
            properties: {
                maxFileSize: { type: 'number', minimum: 1024, maximum: 500 * 1024 * 1024 },
                timeout: { type: 'number', minimum: 1000, maximum: 300000 },
                enableLogging: { type: 'boolean' },
                enableValidation: { type: 'boolean' },
                strictMode: { type: 'boolean' },
                maxRetryAttempts: { type: 'number', minimum: 0, maximum: 10 },
                retryDelay: { type: 'number', minimum: 100, maximum: 10000 }
            }
        };
        
        try {
            const result = new validator({ throwOnError: false }).validate(options, schema);
            if (!result.valid && window.logger) {
                window.logger.warn('PPTXProcessor', 'Invalid constructor options:', result.errors);
            }
        } catch (error) {
            // Validation unavailable, continue with defaults
        }
        
        return options;
    }
    
    /**
     * Validate environment and dependencies
     */
    _validateEnvironment() {
        const requiredGlobals = [
            'Logger', 'PPTXSlideRenderer'
        ];
        
        const missingDependencies = requiredGlobals.filter(dep => 
            typeof globalThis[dep] === 'undefined'
        );
        
        if (missingDependencies.length > 0) {
            throw this._createError(`Missing required dependencies: ${missingDependencies.join(', ')}`, {
                name: 'ProcessingError',
                code: 'MISSING_DEPENDENCIES',
                details: { missingDependencies }
            });
        }
    }
    
    /**
     * Initialize with retry logic
     */
    async _initializeWithRetry() {
        const { maxAttempts, baseDelay } = this.config.retryOptions;
        
        const ErrorRecovery = (typeof window !== 'undefined' && window.ErrorRecovery) ? 
            window.ErrorRecovery : {
                retry: async (operation, options) => {
                    let lastError;
                    for (let i = 0; i < options.maxAttempts; i++) {
                        try {
                            return await operation(i + 1);
                        } catch (error) {
                            lastError = error;
                            if (i < options.maxAttempts - 1) {
                                await new Promise(resolve => setTimeout(resolve, options.baseDelay));
                            }
                        }
                    }
                    throw lastError;
                }
            };
        
        return ErrorRecovery.retry(
            async (attempt) => {
                if (window.logger) {
                    window.logger.trace('PPTXProcessor', `Initialization attempt ${attempt}`);
                }
                
                this.processor = new globalThis.PPTXSlideRenderer();
                
                // Verify processor was created successfully
                if (!this.processor || typeof this.processor.processFile !== 'function') {
                    throw new Error('PPTXSlideRenderer instance is invalid');
                }
                
                return this.processor;
            },
            { maxAttempts, baseDelay }
        );
    }
    
    /**
     * Validate file input
     */
    _validateFileInput(arrayBuffer, options) {
        const errors = [];
        
        // Basic ArrayBuffer validation
        if (!(arrayBuffer instanceof ArrayBuffer)) {
            errors.push('Input must be an ArrayBuffer');
        } else {
            if (arrayBuffer.byteLength === 0) {
                errors.push('ArrayBuffer cannot be empty');
            }
            
            if (arrayBuffer.byteLength > this.config.maxFileSize) {
                errors.push(`File size ${arrayBuffer.byteLength} exceeds maximum ${this.config.maxFileSize}`);
            }
            
            // Check for PPTX file signature (PK)
            const view = new Uint8Array(arrayBuffer, 0, Math.min(4, arrayBuffer.byteLength));
            if (view.length >= 2 && (view[0] !== 0x50 || view[1] !== 0x4B)) {
                errors.push('File does not appear to be a valid ZIP/PPTX file');
            }
        }
        
        // Options validation
        if (options && typeof options !== 'object') {
            errors.push('Options must be an object');
        }
        
        return {
            valid: errors.length === 0,
            errors
        };
    }
    
    /**
     * Validate render input
     */
    _validateRenderInput(canvas, slideIndex, options) {
        const errors = [];
        
        // Canvas validation
        if (!canvas) {
            errors.push('Canvas is required');
        } else if (!(canvas instanceof HTMLCanvasElement)) {
            errors.push('Canvas must be an HTMLCanvasElement');
        } else {
            if (canvas.width <= 0 || canvas.height <= 0) {
                errors.push('Canvas must have positive dimensions');
            }
        }
        
        // Slide index validation
        if (typeof slideIndex !== 'number' || !Number.isInteger(slideIndex)) {
            errors.push('Slide index must be an integer');
        }
        
        // Options validation
        if (options && typeof options !== 'object') {
            errors.push('Options must be an object');
        }
        
        return {
            valid: errors.length === 0,
            errors
        };
    }
    
    /**
     * Validate processing result
     */
    _validateProcessingResult(presentation) {
        if (!presentation) {
            throw this._createError('Processing completed but no presentation was created', {
                name: 'ProcessingError',
                code: 'NO_PRESENTATION_RESULT'
            });
        }
        
        // Additional validation can be added here
        if (this.config.strictMode) {
            if (!this.processor.slides || this.processor.slides.length === 0) {
                window.logger?.warn('PPTXProcessor', 'No slides found in presentation');
            }
        }
    }
    
    /**
     * Process with timeout wrapper
     */
    async _processWithTimeout(operation, timeoutMs) {
        return new Promise((resolve, reject) => {
            const timeout = setTimeout(() => {
                reject(this._createError(`Operation timed out after ${timeoutMs}ms`, {
                    name: 'TimeoutError',
                    code: 'OPERATION_TIMEOUT'
                }));
            }, timeoutMs);
            
            Promise.resolve(operation())
                .then(result => {
                    clearTimeout(timeout);
                    resolve(result);
                })
                .catch(error => {
                    clearTimeout(timeout);
                    reject(error);
                });
        });
    }
    
    /**
     * Initialize performance monitoring
     */
    _initializeMonitoring() {
        // Performance monitoring setup
        if (typeof performance !== 'undefined' && performance.mark) {
            this.performanceEnabled = true;
        }
        
        // Memory monitoring setup
        if (typeof performance !== 'undefined' && performance.memory) {
            this.memoryMonitoringEnabled = true;
        }
    }
    
    /**
     * Mark performance timing
     */
    _markPerformance(name) {
        if (this.performanceEnabled) {
            try {
                performance.mark(name);
                this.performanceMarks.set(name, performance.now());
            } catch (error) {
                // Ignore performance marking errors
            }
        } else {
            this.performanceMarks.set(name, Date.now());
        }
    }
    
    /**
     * Get performance timing between marks
     */
    _getPerformanceTiming(startMark, endMark) {
        const start = this.performanceMarks.get(startMark);
        const end = this.performanceMarks.get(endMark);
        
        if (start && end) {
            return Math.round(end - start);
        }
        
        return 0;
    }
    
    /**
     * Update memory usage tracking
     */
    _updateMemoryUsage() {
        if (this.memoryMonitoringEnabled && performance.memory) {
            try {
                const current = performance.memory.usedJSHeapSize;
                this.memoryUsage.current = current;
                this.memoryUsage.peak = Math.max(this.memoryUsage.peak, current);
            } catch (error) {
                // Ignore memory monitoring errors
            }
        }
    }
    
    /**
     * Create enhanced error with context
     */
    _createError(message, options = {}) {
        const ErrorHandlerFactory = (typeof window !== 'undefined' && window.ErrorHandlerFactory) ? 
            window.ErrorHandlerFactory : {
                createFromError: (error) => new Error(error.message || error)
            };
        
        const error = new Error(message);
        error.name = options.name || 'ProcessingError';
        error.code = options.code || null;
        error.context = 'PPTXProcessor';
        error.timestamp = new Date().toISOString();
        error.recoverable = options.recoverable !== false;
        
        if (options.details) {
            error.details = options.details;
        }
        
        return error;
    }
    
    /**
     * Create validation error
     */
    _createValidationError(message, errors) {
        const error = this._createError(message, {
            name: 'ValidationError',
            code: 'INPUT_VALIDATION_FAILED'
        });
        
        error.validationErrors = errors;
        return error;
    }
    
    /**
     * Enhance error with additional context
     */
    _enhanceError(error, operation, context = {}) {
        if (error.enhanced) {
            return error; // Already enhanced
        }
        
        const enhanced = error instanceof Error ? error : new Error(String(error));
        enhanced.operation = operation;
        enhanced.processorContext = context;
        enhanced.timestamp = new Date().toISOString();
        enhanced.memoryUsage = { ...this.memoryUsage };
        enhanced.enhanced = true;
        
        return enhanced;
    }
}

// Enhanced global export with error boundary wrapper
// Expose factory on both window and globalThis for robustness
(function exposeProcessorGlobal(){
    const g = (typeof globalThis !== 'undefined') ? globalThis : (typeof window !== 'undefined' ? window : null);
    if (!g) {return;}
    const create = (options = {}) => {
        try { return new PPTXProcessor(options); } catch (error) {
            try { (g.logger || window?.logger)?.error?.('PPTXProcessor', 'Failed to create processor:', error); } catch(_e) {}
            throw error;
        }
    };
    if (!g.createPPTXProcessor) {g.createPPTXProcessor = create;}
    if (!g.PPTXProcessor) {g.PPTXProcessor = PPTXProcessor;}
})();

// ES module exports (also exposing on globals via IIFE above)
export { PPTXProcessor };
export default PPTXProcessor;
