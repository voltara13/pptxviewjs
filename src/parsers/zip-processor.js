/**
 * ZIP Processing Module
 * Simplified version of standard ZLib functionality using JSZip
 * Enhanced with comprehensive error handling and input validation
 */

// Resolve JSZip in both script-tag and bundler environments
async function resolveJSZip() {
    // Browser global (root index.html uses CDN)
    if (typeof window !== 'undefined' && window.JSZip && typeof window.JSZip.loadAsync === 'function') {
        return window.JSZip;
    }
    // GlobalThis fallback
    if (typeof globalThis !== 'undefined' && globalThis.JSZip && typeof globalThis.JSZip.loadAsync === 'function') {
        return globalThis.JSZip;
    }
    // Try CommonJS require when bundled/running under Node or bundler
    try {
         
        const mod = require('jszip');
        if (mod && typeof mod.loadAsync === 'function') {return mod;}
        if (mod && mod.default && typeof mod.default.loadAsync === 'function') {return mod.default;}
    } catch (_e) {}
    // As a last resort in browser, dynamically load from CDN
    if (typeof document !== 'undefined') {
        await new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
            script.onload = () => resolve();
            script.onerror = () => reject(new Error('Failed to load JSZip'));
            document.head.appendChild(script);
        });
        if (window.JSZip) {return window.JSZip;}
    }
    throw new Error('JSZip is not available');
}

/**
 * Simplified ZLib implementation using JSZip
 * Mimics the standard ZLib interface with enhanced error handling
 */
class ZLib {
    constructor() {
        this.engine = null;
        this.files = {};
        this.isModuleInit = true; // Always true for browser implementation
        // Create logger with fallback
        
        if (typeof Logger !== 'undefined') {
            this.logger = new Logger('ZLib');
        } else if (typeof window !== 'undefined' && window.Logger) {
            this.logger = new window.Logger('ZLib');
        } else if (typeof globalThis !== 'undefined' && globalThis.Logger) {
            this.logger = new globalThis.Logger('ZLib');
        } else {
            // Fallback logger (silent)
            this.logger = {
                info: () => {},
                debug: () => {},
                warn: () => {},
                error: () => {},
                log: () => {},
                logError: () => {}
            };
        }
        this.isOpen = false;
        this.lastError = null;
    }

    /**
     * Validate input data for ZIP operations
     * @param {*} data - Data to validate
     * @param {string} operation - Operation name for error messages
     * @returns {boolean} True if valid
     */
    validateInput(data, operation) {
        if (data === null || data === undefined) {
            this.lastError = new Error(`${operation}: Input data is null or undefined`);
            return false;
        }

        if (!(data instanceof ArrayBuffer) && !(data instanceof Uint8Array) && !(data instanceof Blob)) {
            this.lastError = new Error(`${operation}: Invalid data type. Expected ArrayBuffer, Uint8Array, or Blob`);
            return false;
        }

        if (data.byteLength === 0 || (data.size !== undefined && data.size === 0)) {
            this.lastError = new Error(`${operation}: Input data is empty`);
            return false;
        }

        return true;
    }

    /**
     * Open archive from bytes
     * @param {ArrayBuffer|Uint8Array|Blob} data - ZIP file data
     * @returns {Promise<boolean>} success or not
     */
    async open(data) {
        try {
            // Reset state
            this.isOpen = false;
            this.lastError = null;
            
            // Validate input
            if (!this.validateInput(data, 'open')) {
                this.logger.logError("ZLib", 'Failed to open ZIP:', this.lastError.message);
                return false;
            }

            // Resolve JSZip implementation
            const JSZip = await resolveJSZip();
            // Convert data to appropriate format if needed
            let zipData = data;
            if (data instanceof Uint8Array) {
                zipData = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
            }

            // Load the ZIP file
            this.engine = await JSZip.loadAsync(zipData);

            if (!this.engine) {
                this.lastError = new Error('Failed to create JSZip instance');
                this.logger.logError("ZLib", 'Failed to open ZIP:', this.lastError.message);
                return false;
            }
            
            

            // Populate files object similar to standard implementation
            this.files = {};
            let fileCount = 0;
            
            this.engine.forEach((relativePath, file) => {
                if (!file.dir) {
                    this.files[relativePath] = null; // Lazy loading like standard
                    fileCount++;
                }
            });

            // Validate that we have a proper ZIP structure
            if (fileCount === 0) {
                this.lastError = new Error('ZIP file contains no readable files');
                this.logger.log("warn", "ZLib", 'Opened ZIP file has no readable files');
            }

            this.isOpen = true;
            return true;

        } catch (error) {
            this.lastError = error;
            this.isOpen = false;
            this.logger.logError("ZLib", 'Failed to open ZIP:', error.message);
            
            // Provide more specific error messages
            if (error.message.includes('End of central directory not found')) {
                this.logger.logError("ZLib", 'Invalid ZIP file: corrupted or not a ZIP file');
            } else if (error.message.includes('encrypted')) {
                this.logger.logError("ZLib", 'Cannot open encrypted ZIP files');
            } else if (error.message.includes('Unsupported compression')) {
                this.logger.logError("ZLib", 'ZIP file uses unsupported compression method');
            }
            
            return false;
        }
    }

    /**
     * Validate file path
     * @param {string} path - File path to validate
     * @param {string} operation - Operation name for error messages
     * @returns {boolean} True if valid
     */
    validatePath(path, operation) {
        if (typeof path !== 'string') {
            this.lastError = new Error(`${operation}: Path must be a string`);
            return false;
        }

        if (path.length === 0) {
            this.lastError = new Error(`${operation}: Path cannot be empty`);
            return false;
        }

        // Check for dangerous path patterns
        if (path.includes('..') || path.includes('//')) {
            this.lastError = new Error(`${operation}: Invalid path contains dangerous patterns`);
            return false;
        }

        return true;
    }

    /**
     * Check if ZIP is properly opened
     * @param {string} operation - Operation name for error messages
     * @returns {boolean} True if ready
     */
    checkReady(operation) {
        if (!this.isOpen || !this.engine) {
            this.lastError = new Error(`${operation}: ZIP archive is not open`);
            return false;
        }
        return true;
    }

    /**
     * Get file from archive
     * @param {string} path - File path in archive
     * @returns {Promise<Uint8Array|null>} File content as bytes
     */
    async getFile(path) {
        try {
            // Validate state and input
            if (!this.checkReady('getFile')) {
                this.logger.logError("ZLib", 'Cannot get file:', this.lastError.message);
                return null;
            }

            if (!this.validatePath(path, 'getFile')) {
                this.logger.logError("ZLib", 'Invalid file path:', this.lastError.message);
                return null;
            }

            // Try both raw and normalized paths regardless of index, since some bundles
            // may not populate the index consistently after minification
            let file = this.engine.file(path);
            if (!file && path.startsWith('/')) {
                file = this.engine.file(path.substring(1));
            }
            if (!file) {
                this.logger.log("warn", "ZLib", `File not found in ZIP: ${path}`);
                return null;
            }

            const arrayBuffer = await file.async('arraybuffer');
            if (!arrayBuffer || arrayBuffer.byteLength === 0) {
                this.logger.log("warn", "ZLib", `File is empty: ${path}`);
                return new Uint8Array(0);
            }

            return new Uint8Array(arrayBuffer);

        } catch (error) {
            this.lastError = error;
            this.logger.logError("ZLib", `Failed to get file ${path}:`, error.message);
            return null;
        }
    }

    /**
     * Get file as text
     * @param {string} path - File path in archive
     * @returns {Promise<string|null>} File content as text
     */
    async getFileText(path) {
        try {
            // Validate state and input
            if (!this.checkReady('getFileText')) {
                this.logger.logError("ZLib", 'Cannot get file text:', this.lastError.message);
                return null;
            }

            if (!this.validatePath(path, 'getFileText')) {
                this.logger.logError("ZLib", 'Invalid file path:', this.lastError.message);
                return null;
            }

            // Try both raw and normalized paths regardless of index
            let file = this.engine.file(path);
            if (!file && path.startsWith('/')) {
                file = this.engine.file(path.substring(1));
            }
            if (!file) {
                this.logger.log("warn", "ZLib", `File not found in ZIP: ${path}`);
                return null;
            }

            const textContent = await file.async('text');
            
            // Validate text content
            if (typeof textContent !== 'string') {
                this.logger.log("warn", "ZLib", `File ${path} did not return valid text content`);
                return null;
            }

            return textContent;

        } catch (error) {
            this.lastError = error;
            this.logger.logError("ZLib", `Failed to get file text ${path}:`, error.message);
            
            // Provide more specific error messages
            if (error.message.includes('binary')) {
                this.logger.logError("ZLib", `File ${path} appears to be binary, cannot read as text`);
            } else if (error.message.includes('encoding')) {
                this.logger.logError("ZLib", `File ${path} has encoding issues`);
            }
            
            return null;
        }
    }

    /**
     * Get all file paths in archive
     * @returns {Array<string>} Array of file paths
     */
    getPaths() {
        try {
            if (!this.checkReady('getPaths')) {
                this.logger.logError("ZLib", 'Cannot get paths:', this.lastError.message);
                return [];
            }

            // Prefer the underlying engine's file list to avoid index inconsistencies
            const paths = this.engine && this.engine.files ? Object.keys(this.engine.files) : Object.keys(this.files);
            return paths;

        } catch (error) {
            this.lastError = error;
            this.logger.logError("ZLib", 'Failed to get file paths:', error.message);
            return [];
        }
    }

    /**
     * Check if file exists in archive
     * @param {string} path - File path to check
     * @returns {boolean} True if file exists
     */
    fileExists(path) {
        try {
            if (!this.validatePath(path, 'fileExists')) {
                return false;
            }

            if (!this.checkReady('fileExists')) {
                return false;
            }

            return this.files.hasOwnProperty(path);

        } catch (error) {
            this.lastError = error;
            this.logger.logError("ZLib", `Error checking file existence for ${path}:`, error.message);
            return false;
        }
    }

    /**
     * Get file size
     * @param {string} path - File path
     * @returns {number} File size in bytes, -1 if error
     */
    getFileSize(path) {
        try {
            if (!this.validatePath(path, 'getFileSize')) {
                return -1;
            }

            if (!this.checkReady('getFileSize')) {
                return -1;
            }

            const file = this.engine.file(path);
            if (!file) {
                return -1;
            }

            // Return the uncompressed size
            return file._data ? file._data.uncompressedSize : 0;

        } catch (error) {
            this.lastError = error;
            this.logger.logError("ZLib", `Error getting file size for ${path}:`, error.message);
            return -1;
        }
    }

    /**
     * Get last error
     * @returns {Error|null} Last error or null
     */
    getLastError() {
        return this.lastError;
    }

    /**
     * Close and cleanup
     */
    close() {
        try {
            if (this.engine) {
                // Clear references to help garbage collection
                this.engine = null;
            }
            
            this.files = {};
            this.isOpen = false;
            this.lastError = null;
            

        } catch (error) {
            this.logger.logError("ZLib", 'Error during ZIP cleanup:', error.message);
        }
    }

    /**
     * Get file data as ArrayBuffer (enhanced version with comprehensive validation)
     * @param {string} path - File path in archive
     * @returns {Promise<ArrayBuffer>} file data
     */
    async getFileData(path) {
        try {
            // Validate state and input
            if (!this.checkReady('getFileData')) {
                throw this.lastError;
            }

            if (!this.validatePath(path, 'getFileData')) {
                throw this.lastError;
            }

            // Normalize path
            const normalizedPath = path.startsWith('/') ? path.substring(1) : path;

            const file = this.engine.file(normalizedPath);
            if (!file) {
                throw new Error(`File not found: ${path}`);
            }

            const arrayBuffer = await file.async('arraybuffer');
            
            if (!arrayBuffer) {
                throw new Error(`Failed to read file data: ${path}`);
            }

            return arrayBuffer;

        } catch (error) {
            this.lastError = error;
            this.logger.logError("ZLib", `Failed to get file data for ${path}:`, error.message);
            throw new Error(`Failed to extract file data: ${path} - ${error.message}`);
        }
    }

    /**
     * Get archive statistics
     * @returns {Object} Archive statistics
     */
    getStats() {
        try {
            if (!this.checkReady('getStats')) {
                return {
                    isOpen: false,
                    fileCount: 0,
                    totalSize: 0,
                    error: this.lastError?.message
                };
            }

            const paths = Object.keys(this.files);
            let totalSize = 0;

            // Calculate total uncompressed size
            paths.forEach(path => {
                const size = this.getFileSize(path);
                if (size > 0) {
                    totalSize += size;
                }
            });

            return {
                isOpen: this.isOpen,
                fileCount: paths.length,
                totalSize: totalSize,
                files: paths,
                lastError: this.lastError?.message || null
            };

        } catch (error) {
            this.lastError = error;
            this.logger.logError("ZLib", 'Error getting archive stats:', error.message);
            return {
                isOpen: false,
                fileCount: 0,
                totalSize: 0,
                error: error.message
            };
        }
    }
}

/**
 * OpenXML Package processor
 * Simplified version of standard OpenXmlPackage
 */
class OpenXmlPackage {
    constructor(zip) {
        this.zip = zip;
        this.parts = {};
        this.contentTypes = {};
        this.relationships = {};
    }

    /**
     * Initialize package by reading content types and relationships
     */
    async initialize() {
        // Read [Content_Types].xml
        const contentTypesXml = await this.zip.getFileText('[Content_Types].xml');
        if (contentTypesXml) {
            this.parseContentTypes(contentTypesXml);
        }

        // Read main relationships
        const mainRelsXml = await this.zip.getFileText('_rels/.rels');
        if (mainRelsXml) {
            this.relationships[''] = this.parseRelationships(mainRelsXml);
        }

        // Load all relationship files for individual parts
        const paths = this.zip.getPaths();
        for (const path of paths) {
            if (path.startsWith('_rels/') && path.endsWith('.rels') && path !== '_rels/.rels') {
                // Extract the part URI from the relationship file path
                // e.g., "ppt/slides/_rels/slide1.xml.rels" -> "/ppt/slides/slide1.xml"
                const relsMatch = path.match(/^(.+)\/_rels\/(.+)\.rels$/);
                if (relsMatch) {
                    const partDir = relsMatch[1];
                    const partFile = relsMatch[2];
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
                this.parts[partUri] = new OpenXmlPart(this, partUri, contentType);
            }
        }
    }

    /**
     * Parse Content_Types.xml
     */
    parseContentTypes(xml) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(xml, 'text/xml');

        const byLocal = (name) => {
            const all = doc.getElementsByTagName('*');
            const out = [];
            for (let i = 0; i < all.length; i++) {
                if (all[i].localName === name) {out.push(all[i]);}
            }
            return out;
        };

        // Parse Default elements (namespaced safe)
        byLocal('Default').forEach(def => {
            const extension = def.getAttribute('Extension');
            const contentType = def.getAttribute('ContentType');
            if (extension && contentType) {
                this.contentTypes[`ext:${extension}`] = contentType;
            }
        });

        // Parse Override elements (namespaced safe)
        byLocal('Override').forEach(override => {
            const partName = override.getAttribute('PartName');
            const contentType = override.getAttribute('ContentType');
            if (partName && contentType) {
                this.contentTypes[partName] = contentType;
            }
        });
    }

    /**
     * Parse relationships XML
     */
    parseRelationships(xml) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(xml, 'text/xml');
        const relationships = {};

        const byLocal = (name) => {
            const all = doc.getElementsByTagName('*');
            const out = [];
            for (let i = 0; i < all.length; i++) {
                if (all[i].localName === name) {out.push(all[i]);}
            }
            return out;
        };

        const rels = byLocal('Relationship');
        rels.forEach(rel => {
            const id = rel.getAttribute('Id');
            const type = rel.getAttribute('Type');
            const target = rel.getAttribute('Target');
            if (!id) {return;}
            relationships[id] = {
                type: type,
                target: target,
                targetMode: rel.getAttribute('TargetMode') || 'Internal'
            };
        });

        return relationships;
    }

    /**
     * Get content type for a part
     */
    getContentType(partName) {
        // Check for direct override
        if (this.contentTypes[partName]) {
            return this.contentTypes[partName];
        }

        // Check by extension
        const extension = partName.split('.').pop();
        if (extension && this.contentTypes[`ext:${extension}`]) {
            return this.contentTypes[`ext:${extension}`];
        }

        return 'application/octet-stream';
    }

    /**
     * Get part by relationship type
     */
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

    /**
     * Get part by URI
     */
    getPartByUri(uri) {
        return this.parts[uri];
    }
}

/**
 * OpenXML Part
 */
class OpenXmlPart {
    constructor(pkg, uri, contentType) {
        this.package = pkg;
        this.uri = uri;
        this.contentType = contentType;
        this._content = null;
    }

    /**
     * Get document content as text
     */
    async getDocumentContent() {
        if (this._content === null) {
            // Remove leading slash for ZIP path
            const zipPath = this.uri.startsWith('/') ? this.uri.substring(1) : this.uri;
            this._content = await this.package.zip.getFileText(zipPath);
        }
        return this._content;
    }

    /**
     * Get relationships for this part
     */
    async getRelationships() {
        // Return pre-loaded relationships if available
        if (this.package.relationships[this.uri]) {
            return this.package.relationships[this.uri];
        }

        // Fallback: try to load relationships file
        const relsPath = this.uri.replace(/\/([^\/]+)$/, '/_rels/$1.rels');
        const zipPath = relsPath.startsWith('/') ? relsPath.substring(1) : relsPath;

        const relsXml = await this.package.zip.getFileText(zipPath);
        if (relsXml) {
            return this.package.parseRelationships(relsXml);
        }
        return {};
    }
}

/**
 * OpenXML Types - simplified version of standard types
 */
const OpenXmlTypes = {
    presentation: {
        relationType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
    },
    extendedFileProperties: {
        relationType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'
    },
    coreFileProperties: {
        relationType: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
    },
    customFileProperties: {
        relationType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
    },
    presentationProperties: {
        relationType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps'
    },
    slide: {
        relationType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
    },
    slideLayout: {
        relationType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
    },
    slideMaster: {
        relationType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'
    },
    theme: {
        relationType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
    },
    image: {
        relationType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
    }
};

// Export classes (maintain backward compatibility)
// Expose to both window and globalThis for bundler/UMD environments
if (typeof window !== 'undefined') {
    window.ZLib = ZLib;
    window.OpenXmlPackage = OpenXmlPackage;
    window.OpenXmlPart = OpenXmlPart;
    window.OpenXmlTypes = OpenXmlTypes;
}
if (typeof globalThis !== 'undefined') {
    globalThis.ZLib = ZLib;
    globalThis.OpenXmlPackage = OpenXmlPackage;
    globalThis.OpenXmlPart = OpenXmlPart;
    globalThis.OpenXmlTypes = OpenXmlTypes;
}

// Intentionally no ES module exports here to support classic <script> usage in root demo
