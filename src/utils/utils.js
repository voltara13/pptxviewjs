/**
 * PPTX PPTX Pipeline Demo - Utility Classes
 * Provides logging, progress tracking, and common utility functions
 */

// ===== LOGGING SYSTEM =====
class Logger {
    constructor(category = 'Default') {
        this.category = category;
        this.logElement = null;
        this.logHistory = [];
    }

    init() {
        this.logElement = document.getElementById('log');
    }

    log(level, category, message, data = null) {
        const timestamp = new Date().toISOString().split('T')[1].split('.')[0];
        const logEntry = {
            timestamp,
            level: level.toUpperCase(),
            category,
            message,
            data
        };

        this.logHistory.push(logEntry);

        // Also log to console for compatibility with tests
        const logMessage = `[${category}] ${message}`;
        const SUPPRESSED_CONSOLE_CATEGORIES = new Set(['PPTXSlideRenderer', 'PPTXProcessor', 'Thumbnail Debug']);
        if (!SUPPRESSED_CONSOLE_CATEGORIES.has(category)) {
            switch (level.toLowerCase()) {
                case 'error':
                    console.error(logMessage, data);
                    break;
                case 'warn':
                    console.warn(logMessage, data);
                    break;
                case 'info':
                    console.info(logMessage, data);
                    break;
                case 'debug':
                    console.debug(logMessage, data);
                    break;
                default:
                    console.log(logMessage, data);
            }
        }

        if (this.logElement) {
            const dataStr = data ? ` | ${JSON.stringify(data)}` : '';
            const logLine = `[${timestamp}] [${level.toUpperCase()}] [${category}] ${message}${dataStr}\n`;
            this.logElement.textContent += logLine;
            this.logElement.scrollTop = this.logElement.scrollHeight;
        }


    }

    // Convenience methods for different log levels
    debug(category, message, data = null) {
        this.log('debug', category, message, data);
    }

    info(category, message, data = null) {
        this.log('info', category, message, data);
    }

    warn(category, message, data = null) {
        this.log('warn', category, message, data);
    }

    error(category, message, data = null) {
        this.log('error', category, message, data);
    }



    trace(category, message, data = null) {
        this.log('trace', category, message, data);
    }

    logError(category, error) {
        let message = 'Unknown error';
        let errorData = {};

        if (error && typeof error === 'object') {
            message = error.message || error.toString() || 'Unknown error';
            errorData = {
                stack: error.stack,
                name: error.name || 'Error'
            };
        } else if (typeof error === 'string') {
            message = error;
        } else {
            message = String(error);
        }

        this.log('error', category, message, errorData);
    }

    markTiming(name) {
        // Enhanced timing functionality with performance monitoring
        if (typeof window !== 'undefined' && window.getGlobalMonitor) {
            const monitor = window.getGlobalMonitor();
            monitor.start(name, { category: this.category });
        } else if (typeof performance !== 'undefined' && performance.mark) {
            try {
                performance.mark(name);
            } catch (error) {
                // Fallback to console timing
                console.time(name);
            }
        }
    }

    measureTiming(name) {
        // Enhanced timing measurement with performance monitoring
        if (typeof window !== 'undefined' && window.getGlobalMonitor) {
            const monitor = window.getGlobalMonitor();
            const measure = monitor.end(name, { category: this.category });
            if (measure) {
                this.log('trace', this.category, `${name} completed in ${Math.round(measure.duration)}ms`);
            }
        } else if (typeof performance !== 'undefined' && performance.mark) {
            try {
                performance.measure(name);
                const entries = performance.getEntriesByName(name, 'measure');
                if (entries.length > 0) {
                    const duration = Math.round(entries[entries.length - 1].duration);
                    this.log('trace', this.category, `${name} completed in ${duration}ms`);
                }
            } catch (error) {
                // Fallback to console timing
                console.timeEnd(name);
            }
        }
    }
}

// Global logger instance
const logger = new Logger();

// ===== PROGRESS TRACKING =====
class ProgressTracker {
    constructor() {
        this.steps = {
            1: { name: 'ZIP Processing', progress: 0 },
            2: { name: 'XML Parsing', progress: 0 },
            3: { name: 'DOM Creation', progress: 0 },
            4: { name: 'Canvas Rendering', progress: 0 }
        };
    }

    updateProgress(stepNum, progress) {
        if (this.steps[stepNum]) {
            this.steps[stepNum].progress = Math.min(100, Math.max(0, progress));

            const progressBar = document.getElementById(`progress${stepNum}`);
            const stepElement = document.getElementById(`step${stepNum}`);

            if (progressBar) {
                progressBar.style.width = `${this.steps[stepNum].progress}%`;
            }

            if (stepElement) {
                stepElement.classList.remove('active', 'processing');
                if (progress >= 100) {
                    stepElement.classList.add('active');
                } else if (progress > 0) {
                    stepElement.classList.add('processing');
                }
            }

        }
    }

    reset() {
        for (const stepNum in this.steps) {
            this.updateProgress(parseInt(stepNum), 0);
        }
    }
}

// Global progress tracker
const progressTracker = new ProgressTracker();

// ===== FILE READER UTILITIES =====
class FileReaderAsync {
    /**
     * Read file as ArrayBuffer with enhanced error handling and validation
     * @param {File|Blob} file - File to read
     * @param {Object} options - Reading options
     * @returns {Promise<ArrayBuffer>} File content as ArrayBuffer
     */
    static async readAsArrayBuffer(file, options = {}) {
        // Input validation
        if (!file || !(file instanceof File || file instanceof Blob)) {
            throw new Error('Invalid file parameter: must be File or Blob instance');
        }

        if (file.size === 0) {
            throw new Error('Cannot read empty file');
        }

        const { maxSize = 100 * 1024 * 1024, timeout = 80000 } = options; // 100MB default, 80s timeout

        if (file.size > maxSize) {
            throw new Error(`File too large: ${file.size} bytes (max: ${maxSize} bytes)`);
        }

        return new Promise((resolve, reject) => {
            const reader = new window.FileReader();
            
            // Set up timeout
            const timeoutId = setTimeout(() => {
                reader.abort();
                reject(new Error(`File reading timed out after ${timeout}ms`));
            }, timeout);

            reader.onload = (event) => {
                clearTimeout(timeoutId);
                const result = event.target?.result;
                if (result instanceof ArrayBuffer) {
                    resolve(result);
                } else {
                    reject(new Error('Failed to read file as ArrayBuffer'));
                }
            };

            reader.onerror = () => {
                clearTimeout(timeoutId);
                const error = reader.error || new Error('Unknown file reading error');
                reject(new Error(`Failed to read file: ${error.message}`));
            };

            reader.onabort = () => {
                clearTimeout(timeoutId);
                reject(new Error('File reading was aborted'));
            };

            try {
                reader.readAsArrayBuffer(file);
            } catch (error) {
                clearTimeout(timeoutId);
                reject(new Error(`Failed to start file reading: ${error.message}`));
            }
        });
    }

    /**
     * Read file as text with encoding support and validation
     * @param {File|Blob} file - File to read
     * @param {Object} options - Reading options
     * @returns {Promise<string>} File content as text
     */
    static async readAsText(file, options = {}) {
        // Input validation
        if (!file || !(file instanceof File || file instanceof Blob)) {
            throw new Error('Invalid file parameter: must be File or Blob instance');
        }

        const { 
            encoding = 'UTF-8', 
            maxSize = 50 * 1024 * 1024, 
            timeout = 80000,
            validateText = true 
        } = options;

        if (file.size > maxSize) {
            throw new Error(`File too large: ${file.size} bytes (max: ${maxSize} bytes)`);
        }

        return new Promise((resolve, reject) => {
            const reader = new window.FileReader();
            
            // Set up timeout
            const timeoutId = setTimeout(() => {
                reader.abort();
                reject(new Error(`File reading timed out after ${timeout}ms`));
            }, timeout);

            reader.onload = (event) => {
                clearTimeout(timeoutId);
                const result = event.target?.result;
                
                if (typeof result === 'string') {
                    // Optional text validation
                    if (validateText && result.includes('\uFFFD')) {
                        reject(new Error('File contains invalid characters - may not be a text file'));
                        return;
                    }
                    resolve(result);
                } else {
                    reject(new Error('Failed to read file as text'));
                }
            };

            reader.onerror = () => {
                clearTimeout(timeoutId);
                const error = reader.error || new Error('Unknown file reading error');
                reject(new Error(`Failed to read file as text: ${error.message}`));
            };

            reader.onabort = () => {
                clearTimeout(timeoutId);
                reject(new Error('File reading was aborted'));
            };

            try {
                reader.readAsText(file, encoding);
            } catch (error) {
                clearTimeout(timeoutId);
                reject(new Error(`Failed to start file reading: ${error.message}`));
            }
        });
    }

    /**
     * Read file as Data URL with validation
     * @param {File|Blob} file - File to read
     * @param {Object} options - Reading options
     * @returns {Promise<string>} File content as Data URL
     */
    static async readAsDataURL(file, options = {}) {
        // Input validation
        if (!file || !(file instanceof File || file instanceof Blob)) {
            throw new Error('Invalid file parameter: must be File or Blob instance');
        }

        const { maxSize = 10 * 1024 * 1024, timeout = 80000 } = options; // 10MB default for data URLs

        if (file.size > maxSize) {
            throw new Error(`File too large for Data URL: ${file.size} bytes (max: ${maxSize} bytes)`);
        }

        return new Promise((resolve, reject) => {
            const reader = new window.FileReader();
            
            // Set up timeout
            const timeoutId = setTimeout(() => {
                reader.abort();
                reject(new Error(`File reading timed out after ${timeout}ms`));
            }, timeout);

            reader.onload = (event) => {
                clearTimeout(timeoutId);
                const result = event.target?.result;
                
                if (typeof result === 'string' && result.startsWith('data:')) {
                    resolve(result);
                } else {
                    reject(new Error('Failed to read file as Data URL'));
                }
            };

            reader.onerror = () => {
                clearTimeout(timeoutId);
                const error = reader.error || new Error('Unknown file reading error');
                reject(new Error(`Failed to read file as Data URL: ${error.message}`));
            };

            reader.onabort = () => {
                clearTimeout(timeoutId);
                reject(new Error('File reading was aborted'));
            };

            try {
                reader.readAsDataURL(file);
            } catch (error) {
                clearTimeout(timeoutId);
                reject(new Error(`Failed to start file reading: ${error.message}`));
            }
        });
    }

    /**
     * Read multiple files in parallel with progress tracking
     * @param {File[]} files - Array of files to read
     * @param {string} readType - Type of reading ('arrayBuffer', 'text', 'dataURL')
     * @param {Object} options - Reading options
     * @returns {Promise<Array>} Array of file contents
     */
    static async readMultipleFiles(files, readType = 'arrayBuffer', options = {}) {
        if (!Array.isArray(files) || files.length === 0) {
            throw new Error('Invalid files parameter: must be non-empty array');
        }

        const { 
            onProgress = null,
            maxConcurrent = 3,
            ...readOptions 
        } = options;

        const readMethod = this[`readAs${readType.charAt(0).toUpperCase() + readType.slice(1)}`];
        if (!readMethod) {
            throw new Error(`Invalid read type: ${readType}`);
        }

        const results = [];
        let completed = 0;

        // Process files in batches to avoid overwhelming the system
        for (let i = 0; i < files.length; i += maxConcurrent) {
            const batch = files.slice(i, i + maxConcurrent);
            const batchPromises = batch.map(async (file, index) => {
                try {
                    const result = await readMethod.call(this, file, readOptions);
                    completed++;
                    
                    if (onProgress) {
                        onProgress({
                            completed,
                            total: files.length,
                            progress: (completed / files.length) * 100,
                            currentFile: file.name || `File ${i + index + 1}`
                        });
                    }
                    
                    return result;
                } catch (error) {
                    completed++;
                    if (onProgress) {
                        onProgress({
                            completed,
                            total: files.length,
                            progress: (completed / files.length) * 100,
                            currentFile: file.name || `File ${i + index + 1}`,
                            error: error.message
                        });
                    }
                    throw error;
                }
            });

            const batchResults = await Promise.all(batchPromises);
            results.push(...batchResults);
        }

        return results;
    }
}

// Backward compatibility alias
const FileReader = FileReaderAsync;

// ===== ERROR HANDLING =====
class ErrorHandler {
    // Enhanced error types for better categorization
    static ERROR_TYPES = {
        VALIDATION: 'ValidationError',
        FILE_IO: 'FileIOError',
        PARSING: 'ParsingError',
        RENDERING: 'RenderingError',
        NETWORK: 'NetworkError',
        TIMEOUT: 'TimeoutError',
        MEMORY: 'MemoryError',
        UNKNOWN: 'UnknownError'
    };

    static SEVERITY_LEVELS = {
        LOW: 'low',
        MEDIUM: 'medium',
        HIGH: 'high',
        CRITICAL: 'critical'
    };

    /**
     * Enhanced error handling with categorization and severity
     * @param {Error|string} error - Error to handle
     * @param {Object} options - Error handling options
     */
    static handle(error, options = {}) {
        const {
            context = 'Unknown',
            severity = this.SEVERITY_LEVELS.MEDIUM,
            displayToUser = true,
            logToConsole = true,
            throwAfterHandle = false,
            customHandler = null
        } = options;

        // Normalize error object
        const errorInfo = this.normalizeError(error, context);
        
        // Log to console if enabled
        if (logToConsole) {
            this.logError(errorInfo, severity);
        }

        // Display to user if enabled
        if (displayToUser) {
            this.displayError(errorInfo, severity);
        }

        // Call custom handler if provided
        if (customHandler && typeof customHandler === 'function') {
            try {
                customHandler(errorInfo, severity);
            } catch (handlerError) {
                console.error('Error in custom error handler:', handlerError);
            }
        }

        // Throw if requested
        if (throwAfterHandle) {
            throw error instanceof Error ? error : new Error(errorInfo.message);
        }

        return errorInfo;
    }

    /**
     * Normalize error to consistent format
     * @param {Error|string|*} error - Error to normalize
     * @param {string} context - Error context
     * @returns {Object} Normalized error info
     */
    static normalizeError(error, context) {
        let message = 'Unknown error occurred';
        let stack = null;
        let name = this.ERROR_TYPES.UNKNOWN;
        const originalError = error;

        if (error instanceof Error) {
            message = error.message || 'Error occurred';
            stack = error.stack;
            name = error.name || this.categorizeError(error);
        } else if (typeof error === 'string') {
            message = error;
        } else if (error && typeof error === 'object') {
            message = error.message || error.toString() || 'Object error';
            stack = error.stack;
            name = error.name || this.ERROR_TYPES.UNKNOWN;
        } else {
            message = String(error);
        }

        return {
            message,
            context,
            stack,
            name,
            timestamp: new Date().toISOString(),
            originalError,
            userAgent: navigator?.userAgent || 'Unknown'
        };
    }

    /**
     * Categorize error by message content
     * @param {Error} error - Error to categorize
     * @returns {string} Error category
     */
    static categorizeError(error) {
        const message = error.message?.toLowerCase() || '';
        
        if (message.includes('validation') || message.includes('invalid')) {
            return this.ERROR_TYPES.VALIDATION;
        }
        if (message.includes('file') || message.includes('read') || message.includes('write')) {
            return this.ERROR_TYPES.FILE_IO;
        }
        if (message.includes('parse') || message.includes('xml') || message.includes('json')) {
            return this.ERROR_TYPES.PARSING;
        }
        if (message.includes('render') || message.includes('canvas') || message.includes('draw')) {
            return this.ERROR_TYPES.RENDERING;
        }
        if (message.includes('network') || message.includes('fetch') || message.includes('load')) {
            return this.ERROR_TYPES.NETWORK;
        }
        if (message.includes('timeout') || message.includes('timed out')) {
            return this.ERROR_TYPES.TIMEOUT;
        }
        if (message.includes('memory') || message.includes('allocation')) {
            return this.ERROR_TYPES.MEMORY;
        }
        
        return this.ERROR_TYPES.UNKNOWN;
    }

    /**
     * Log error to console with appropriate level
     * @param {Object} errorInfo - Normalized error info
     * @param {string} severity - Error severity
     */
    static logError(errorInfo, severity) {
        const logMessage = `[${errorInfo.context}] ${errorInfo.message}`;
        
        switch (severity) {
            case this.SEVERITY_LEVELS.LOW:
                console.info('🔵', logMessage, errorInfo);
                break;
            case this.SEVERITY_LEVELS.MEDIUM:
                console.warn('🟡', logMessage, errorInfo);
                break;
            case this.SEVERITY_LEVELS.HIGH:
                console.error('🔴', logMessage, errorInfo);
                break;
            case this.SEVERITY_LEVELS.CRITICAL:
                console.error('💥', logMessage, errorInfo);
                if (errorInfo.stack) {
                    console.error('Stack trace:', errorInfo.stack);
                }
                break;
            default:
                console.log(logMessage, errorInfo);
        }
    }

    /**
     * Display error to user with enhanced UI
     * @param {Object} errorInfo - Normalized error info
     * @param {string} severity - Error severity
     */
    static displayError(errorInfo, severity) {
        const output = document.getElementById('output');
        if (!output) {return;}

        const errorDiv = document.createElement('div');
        errorDiv.className = `error error--${severity}`;
        
        // Create error content with better structure
        const errorHeader = document.createElement('div');
        errorHeader.className = 'error__header';
        errorHeader.innerHTML = `
            <span class="error__icon">${this.getSeverityIcon(severity)}</span>
            <span class="error__title">Error in ${errorInfo.context}</span>
            <span class="error__type">${errorInfo.name}</span>
        `;

        const errorMessage = document.createElement('div');
        errorMessage.className = 'error__message';
        errorMessage.textContent = errorInfo.message;

        const errorTime = document.createElement('div');
        errorTime.className = 'error__time';
        errorTime.textContent = new Date(errorInfo.timestamp).toLocaleTimeString();

        errorDiv.appendChild(errorHeader);
        errorDiv.appendChild(errorMessage);
        errorDiv.appendChild(errorTime);

        // Add dismiss functionality
        const dismissBtn = document.createElement('button');
        dismissBtn.className = 'error__dismiss';
        dismissBtn.innerHTML = '×';
        dismissBtn.onclick = () => errorDiv.remove();
        errorDiv.appendChild(dismissBtn);

        // Insert at top with animation
        errorDiv.style.transform = 'translateY(-100%)';
        errorDiv.style.transition = 'transform 0.3s ease-in-out';
        output.insertBefore(errorDiv, output.firstChild);
        
        // Trigger animation
        requestAnimationFrame(() => {
            errorDiv.style.transform = 'translateY(0)';
        });

        // Auto-dismiss based on severity
        const dismissTime = this.getAutoDismissTime(severity);
        if (dismissTime > 0) {
            setTimeout(() => {
                if (errorDiv.parentNode) {
                    errorDiv.style.transform = 'translateY(-100%)';
                    setTimeout(() => errorDiv.remove(), 300);
                }
            }, dismissTime);
        }
    }

    /**
     * Get icon for severity level
     * @param {string} severity - Severity level
     * @returns {string} Icon
     */
    static getSeverityIcon(severity) {
        switch (severity) {
            case this.SEVERITY_LEVELS.LOW: return 'ℹ️';
            case this.SEVERITY_LEVELS.MEDIUM: return '⚠️';
            case this.SEVERITY_LEVELS.HIGH: return '❌';
            case this.SEVERITY_LEVELS.CRITICAL: return '💥';
            default: return '❓';
        }
    }

    /**
     * Get auto-dismiss time based on severity
     * @param {string} severity - Severity level
     * @returns {number} Dismiss time in milliseconds (0 = no auto-dismiss)
     */
    static getAutoDismissTime(severity) {
        switch (severity) {
            case this.SEVERITY_LEVELS.LOW: return 3000;
            case this.SEVERITY_LEVELS.MEDIUM: return 5000;
            case this.SEVERITY_LEVELS.HIGH: return 8000;
            case this.SEVERITY_LEVELS.CRITICAL: return 0; // Manual dismiss only
            default: return 5000;
        }
    }

    /**
     * Create and throw a custom error with enhanced properties
     * @param {string} message - Error message
     * @param {Object} options - Error options
     * @returns {Error} Custom error
     */
    static createError(message, options = {}) {
        const {
            name = this.ERROR_TYPES.UNKNOWN,
            context = 'Unknown',
            cause = null,
            code = null,
            details = null
        } = options;

        const error = new Error(message);
        error.name = name;
        error.context = context;
        error.timestamp = new Date().toISOString();
        
        if (cause) {error.cause = cause;}
        if (code) {error.code = code;}
        if (details) {error.details = details;}

        return error;
    }
}

// ===== COORDINATE TRANSFORMATION UTILITIES =====
class CoordinateTransform {
    // Enhanced EMU (English Metric Units) conversion with validation
    static UNITS = {
        EMU_PER_INCH: 914400,
        EMU_PER_CM: 360000,
        EMU_PER_MM: 36000,
        EMU_PER_POINT: 12700,
        INCHES_PER_MM: 1 / 25.4,
        MM_PER_INCH: 25.4,
        POINTS_PER_INCH: 72
    };

    static VALIDATION = {
        MAX_EMU: 2147483647, // Max 32-bit signed integer
        MIN_EMU: -2147483648,
        MAX_PIXELS: 32767,
        MAX_MM: 59652, // Roughly 2km
        MIN_DPI: 1,
        MAX_DPI: 600
    };

    /**
     * Validate numeric input for coordinate transformations
     * @param {number} value - Value to validate
     * @param {string} unit - Unit type for context
     * @param {Object} bounds - Min/max bounds
     * @returns {number} Validated value
     */
    static validateInput(value, unit, bounds = {}) {
        if (typeof value !== 'number' || !isFinite(value)) {
            throw ErrorHandler.createError(
                `Invalid ${unit} value: must be a finite number`,
                { name: ErrorHandler.ERROR_TYPES.VALIDATION, context: 'CoordinateTransform' }
            );
        }

        const { min = -Infinity, max = Infinity } = bounds;
        
        if (value < min || value > max) {
            throw ErrorHandler.createError(
                `${unit} value out of range: ${value} (valid range: ${min} to ${max})`,
                { name: ErrorHandler.ERROR_TYPES.VALIDATION, context: 'CoordinateTransform' }
            );
        }

        return value;
    }

    /**
     * Convert EMU to millimeters with validation
     * @param {number} emu - EMU value
     * @returns {number} Millimeters
     */
    static emuToMM(emu) {
        const validEmu = this.validateInput(emu, 'EMU', {
            min: this.VALIDATION.MIN_EMU,
            max: this.VALIDATION.MAX_EMU
        });
        
        return validEmu / this.UNITS.EMU_PER_MM;
    }

    /**
     * Convert millimeters to EMU with validation
     * @param {number} mm - Millimeter value
     * @returns {number} EMU value
     */
    static mmToEMU(mm) {
        const validMM = this.validateInput(mm, 'MM', {
            min: -this.VALIDATION.MAX_MM,
            max: this.VALIDATION.MAX_MM
        });
        
        const result = validMM * this.UNITS.EMU_PER_MM;
        
        // Ensure result doesn't overflow EMU bounds
        return Math.max(this.VALIDATION.MIN_EMU, 
               Math.min(this.VALIDATION.MAX_EMU, Math.round(result)));
    }

    /**
     * Convert EMU to points with validation
     * @param {number} emu - EMU value
     * @returns {number} Points
     */
    static emuToPoints(emu) {
        const validEmu = this.validateInput(emu, 'EMU', {
            min: this.VALIDATION.MIN_EMU,
            max: this.VALIDATION.MAX_EMU
        });
        
        return validEmu / this.UNITS.EMU_PER_POINT;
    }

    /**
     * Convert points to EMU with validation
     * @param {number} points - Points value
     * @returns {number} EMU value
     */
    static pointsToEMU(points) {
        const validPoints = this.validateInput(points, 'Points', {
            min: -30000, // Reasonable bounds for document coordinates
            max: 30000
        });
        
        const result = validPoints * this.UNITS.EMU_PER_POINT;
        
        return Math.max(this.VALIDATION.MIN_EMU, 
               Math.min(this.VALIDATION.MAX_EMU, Math.round(result)));
    }

    /**
     * Convert EMU to pixels with enhanced DPI validation
     * @param {number} emu - EMU value
     * @param {number} dpi - DPI value (default: 96)
     * @returns {number} Pixels
     */
    static emuToPixels(emu, dpi = 96) {
        const validEmu = this.validateInput(emu, 'EMU', {
            min: this.VALIDATION.MIN_EMU,
            max: this.VALIDATION.MAX_EMU
        });
        
        const validDpi = this.validateInput(dpi, 'DPI', {
            min: this.VALIDATION.MIN_DPI,
            max: this.VALIDATION.MAX_DPI
        });
        
        const inches = validEmu / this.UNITS.EMU_PER_INCH;
        return Math.round(inches * validDpi);
    }

    /**
     * Convert pixels to EMU with enhanced validation
     * @param {number} pixels - Pixel value
     * @param {number} dpi - DPI value (default: 96)
     * @returns {number} EMU value
     */
    static pixelsToEMU(pixels, dpi = 96) {
        const validPixels = this.validateInput(pixels, 'Pixels', {
            min: -this.VALIDATION.MAX_PIXELS,
            max: this.VALIDATION.MAX_PIXELS
        });
        
        const validDpi = this.validateInput(dpi, 'DPI', {
            min: this.VALIDATION.MIN_DPI,
            max: this.VALIDATION.MAX_DPI
        });
        
        const inches = validPixels / validDpi;
        const result = inches * this.UNITS.EMU_PER_INCH;
        
        return Math.max(this.VALIDATION.MIN_EMU, 
               Math.min(this.VALIDATION.MAX_EMU, Math.round(result)));
    }

    /**
     * Convert millimeters to pixels with validation
     * @param {number} mm - Millimeter value
     * @param {number} dpi - DPI value (default: 96)
     * @returns {number} Pixels
     */
    static mmToPixels(mm, dpi = 96) {
        const validMM = this.validateInput(mm, 'MM', {
            min: -this.VALIDATION.MAX_MM,
            max: this.VALIDATION.MAX_MM
        });
        
        const validDpi = this.validateInput(dpi, 'DPI', {
            min: this.VALIDATION.MIN_DPI,
            max: this.VALIDATION.MAX_DPI
        });
        
        const inches = validMM * this.UNITS.INCHES_PER_MM;
        return Math.round(inches * validDpi);
    }

    /**
     * Convert pixels to millimeters with validation
     * @param {number} pixels - Pixel value
     * @param {number} dpi - DPI value (default: 96)
     * @returns {number} Millimeters
     */
    static pixelsToMM(pixels, dpi = 96) {
        const validPixels = this.validateInput(pixels, 'Pixels', {
            min: -this.VALIDATION.MAX_PIXELS,
            max: this.VALIDATION.MAX_PIXELS
        });
        
        const validDpi = this.validateInput(dpi, 'DPI', {
            min: this.VALIDATION.MIN_DPI,
            max: this.VALIDATION.MAX_DPI
        });
        
        const inches = validPixels / validDpi;
        return inches * this.UNITS.MM_PER_INCH;
    }

    /**
     * Convert between arbitrary units with validation
     * @param {number} value - Value to convert
     * @param {string} fromUnit - Source unit (emu, mm, pixels, points, inches)
     * @param {string} toUnit - Target unit
     * @param {number} dpi - DPI for pixel conversions (default: 96)
     * @returns {number} Converted value
     */
    static convert(value, fromUnit, toUnit, dpi = 96) {
        if (fromUnit === toUnit) {return value;}

        const normalizedFrom = fromUnit.toLowerCase();
        const normalizedTo = toUnit.toLowerCase();

        // Convert to EMU as intermediate unit
        let emuValue;
        
        switch (normalizedFrom) {
            case 'emu':
                emuValue = value;
                break;
            case 'mm':
                emuValue = this.mmToEMU(value);
                break;
            case 'pixels':
            case 'px':
                emuValue = this.pixelsToEMU(value, dpi);
                break;
            case 'points':
            case 'pt':
                emuValue = this.pointsToEMU(value);
                break;
            case 'inches':
            case 'in':
                const validValue = this.validateInput(value, 'Inches', { min: -100, max: 100 });
                emuValue = Math.round(validValue * this.UNITS.EMU_PER_INCH);
                break;
            default:
                throw ErrorHandler.createError(
                    `Unsupported source unit: ${fromUnit}`,
                    { name: ErrorHandler.ERROR_TYPES.VALIDATION, context: 'CoordinateTransform' }
                );
        }

        // Convert from EMU to target unit
        switch (normalizedTo) {
            case 'emu':
                return emuValue;
            case 'mm':
                return this.emuToMM(emuValue);
            case 'pixels':
            case 'px':
                return this.emuToPixels(emuValue, dpi);
            case 'points':
            case 'pt':
                return this.emuToPoints(emuValue);
            case 'inches':
            case 'in':
                return emuValue / this.UNITS.EMU_PER_INCH;
            default:
                throw ErrorHandler.createError(
                    `Unsupported target unit: ${toUnit}`,
                    { name: ErrorHandler.ERROR_TYPES.VALIDATION, context: 'CoordinateTransform' }
                );
        }
    }

    /**
     * Batch convert multiple values
     * @param {number[]} values - Array of values to convert
     * @param {string} fromUnit - Source unit
     * @param {string} toUnit - Target unit
     * @param {number} dpi - DPI for pixel conversions
     * @returns {number[]} Array of converted values
     */
    static convertBatch(values, fromUnit, toUnit, dpi = 96) {
        if (!Array.isArray(values)) {
            throw ErrorHandler.createError(
                'Values must be an array',
                { name: ErrorHandler.ERROR_TYPES.VALIDATION, context: 'CoordinateTransform' }
            );
        }

        return values.map(value => this.convert(value, fromUnit, toUnit, dpi));
    }

    /**
     * Get conversion factor between two units
     * @param {string} fromUnit - Source unit
     * @param {string} toUnit - Target unit
     * @param {number} dpi - DPI for pixel conversions
     * @returns {number} Conversion factor
     */
    static getConversionFactor(fromUnit, toUnit, dpi = 96) {
        // Use a reference value of 1 to calculate the factor
        return this.convert(1, fromUnit, toUnit, dpi);
    }
}

// ===== STRING UTILITIES =====
class StringUtils {
    static decodeXMLEntities(str) {
        if (!str) {return str;}

        const entityMap = {
            '&amp;': '&',
            '&lt;': '<',
            '&gt;': '>',
            '&quot;': '"',
            '&apos;': "'"
        };

        return str.replace(/&[a-zA-Z0-9#]+;/g, (match) => {
            return entityMap[match] || match;
        });
    }

    static encodeXMLEntities(str) {
        if (!str) {return str;}

        return str
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    static extractNamespace(nodeName) {
        const colonIndex = nodeName.indexOf(':');
        return colonIndex !== -1 ? nodeName.substring(0, colonIndex) : '';
    }

    static extractLocalName(nodeName) {
        const colonIndex = nodeName.indexOf(':');
        return colonIndex !== -1 ? nodeName.substring(colonIndex + 1) : nodeName;
    }
}

// ===== DOM UTILITIES =====
class DOMUtils {
    static getAttributeValue(element, attributeName) {
        return element.getAttribute ? element.getAttribute(attributeName) : null;
    }

    static getChildElements(element) {
        if (!element.children) {return [];}
        return Array.from(element.children);
    }

    static getElementByTagName(element, tagName) {
        if (!element.getElementsByTagName) {return null;}
        const elements = element.getElementsByTagName(tagName);
        return elements.length > 0 ? elements[0] : null;
    }

    static getElementsByTagName(element, tagName) {
        if (!element.getElementsByTagName) {return [];}
        return Array.from(element.getElementsByTagName(tagName));
    }

    static getTextContent(element) {
        return element.textContent || element.innerText || '';
    }
}

// ===== ANIMATION UTILITIES =====
class AnimationUtils {
    static easeInOut(t) {
        return t < 0.5 ? 2 * t * t : -1 + (4 - 2 * t) * t;
    }

    static animate(duration, callback, onComplete = null) {
        const startTime = performance.now();

        function frame(currentTime) {
            const elapsed = currentTime - startTime;
            const progress = Math.min(elapsed / duration, 1);

            callback(progress);

            if (progress < 1) {
                requestAnimationFrame(frame);
            } else if (onComplete) {
                onComplete();
            }
        }

        requestAnimationFrame(frame);
    }
}

// ===== INITIALIZATION =====
document.addEventListener('DOMContentLoaded', function() {
    logger.init();
});

// Initialize performance monitoring
let performanceMonitor = null;
if (typeof window !== 'undefined') {
    // Try to load performance monitoring
    try {
        // Performance monitoring will be available if performance.js is loaded
        if (window.PerformanceMonitor) {
            performanceMonitor = new window.PerformanceMonitor({
                enabled: true,
                collectMemoryData: true,
                enableConsoleReports: false, // Use logger instead
                maxHistoryEntries: 500
            });
            window.performanceMonitor = performanceMonitor;
        }
    } catch (error) {
    }
}

// Enhanced logger with performance integration
class EnhancedLogger extends Logger {
    constructor(category) {
        super(category);
        this.performanceMonitor = performanceMonitor;
    }
    
    /**
     * Start performance measurement
     * @param {string} operation - Operation name
     * @param {Object} metadata - Additional metadata
     */
    startPerformance(operation, metadata = {}) {
        if (this.performanceMonitor) {
            this.performanceMonitor.start(operation, {
                ...metadata,
                category: this.category
            });
        }
        this.markTiming(operation);
    }
    
    /**
     * End performance measurement
     * @param {string} operation - Operation name
     * @param {Object} metadata - Additional metadata
     */
    endPerformance(operation, metadata = {}) {
        if (this.performanceMonitor) {
            const measure = this.performanceMonitor.end(operation, {
                ...metadata,
                category: this.category
            });
            
            if (measure) {
                const duration = Math.round(measure.duration * 100) / 100;
                this.trace(this.category, `${operation} completed in ${duration}ms`, {
                    duration,
                    memoryUsage: measure.memorySnapshot?.jsHeapSize
                });
                
                // Log slow operations as warnings
                if (duration > 1000) {
                    this.warn(this.category, `Slow operation detected: ${operation} took ${duration}ms`);
                }
            }
        }
        this.measureTiming(operation);
    }
    
    /**
     * Measure a synchronous operation
     * @param {string} operation - Operation name
     * @param {Function} fn - Function to measure
     * @param {Object} metadata - Additional metadata
     */
    measureSync(operation, fn, metadata = {}) {
        if (this.performanceMonitor) {
            const { result, measure } = this.performanceMonitor.measure(operation, fn, {
                ...metadata,
                category: this.category
            });
            
            if (measure) {
                const duration = Math.round(measure.duration * 100) / 100;
                this.trace(this.category, `${operation} completed in ${duration}ms`);
            }
            
            return result;
        } else {
            this.startPerformance(operation, metadata);
            try {
                const result = fn();
                this.endPerformance(operation, metadata);
                return result;
            } catch (error) {
                this.endPerformance(operation, { ...metadata, error: true });
                throw error;
            }
        }
    }
    
    /**
     * Measure an asynchronous operation
     * @param {string} operation - Operation name
     * @param {Function} fn - Async function to measure
     * @param {Object} metadata - Additional metadata
     */
    async measureAsync(operation, fn, metadata = {}) {
        if (this.performanceMonitor) {
            const { result, measure } = await this.performanceMonitor.measureAsync(operation, fn, {
                ...metadata,
                category: this.category
            });
            
            if (measure) {
                const duration = Math.round(measure.duration * 100) / 100;
                this.trace(this.category, `${operation} completed in ${duration}ms`);
            }
            
            return result;
        } else {
            this.startPerformance(operation, metadata);
            try {
                const result = await fn();
                this.endPerformance(operation, metadata);
                return result;
            } catch (error) {
                this.endPerformance(operation, { ...metadata, error: true });
                throw error;
            }
        }
    }
    
    /**
     * Get performance statistics for this logger's category
     */
    getPerformanceStats() {
        if (!this.performanceMonitor) {
            return null;
        }
        
        const allHistory = this.performanceMonitor.getHistory();
        const categoryHistory = allHistory.filter(op => 
            op.startMetadata?.category === this.category || 
            op.endMetadata?.category === this.category
        );
        
        if (categoryHistory.length === 0) {
            return { category: this.category, operations: 0 };
        }
        
        const totalDuration = categoryHistory.reduce((sum, op) => sum + op.duration, 0);
        const averageDuration = totalDuration / categoryHistory.length;
        const slowest = categoryHistory.reduce((max, op) => op.duration > max.duration ? op : max);
        const fastest = categoryHistory.reduce((min, op) => op.duration < min.duration ? op : min);
        
        return {
            category: this.category,
            operations: categoryHistory.length,
            totalDuration: Math.round(totalDuration),
            averageDuration: Math.round(averageDuration * 100) / 100,
            slowestOperation: {
                name: slowest.name,
                duration: Math.round(slowest.duration * 100) / 100
            },
            fastestOperation: {
                name: fastest.name,
                duration: Math.round(fastest.duration * 100) / 100
            }
        };
    }
}

// Create enhanced global logger instance
const enhancedLogger = new EnhancedLogger();

// Export for use in other modules (maintain backward compatibility)
if (typeof window !== 'undefined') {
    window.Logger = Logger;
    window.EnhancedLogger = EnhancedLogger;
    window.ProgressTracker = ProgressTracker;
    window.FileReader = FileReader;
    window.ErrorHandler = ErrorHandler;
    window.CoordinateTransform = CoordinateTransform;
    window.StringUtils = StringUtils;
    window.DOMUtils = DOMUtils;
    window.AnimationUtils = AnimationUtils;
    window.logger = logger;
    window.enhancedLogger = enhancedLogger;
    window.progressTracker = progressTracker;
    if (performanceMonitor) {
        window.performanceMonitor = performanceMonitor;
    }
}

// ES Module exports (disabled for script-tag compatibility)
// export {
//     Logger,
//     EnhancedLogger,
//     ProgressTracker,
//     FileReaderAsync as FileReader,
//     ErrorHandler,
//     CoordinateTransform,
//     StringUtils,
//     DOMUtils,
//     AnimationUtils
// };