/**
 * Media Processor Module
 * Handles video, audio, and embedded media elements in PPTX presentations
 * Supports PptxGenJS media types and YouTube embeds
 */

/**
 * Media Information Structure
 */
class MediaInfo {
    constructor() {
        this.type = '';           // 'video', 'audio', 'youtube', 'image'
        this.source = '';         // File path or URL
        this.mimeType = '';       // MIME type (e.g., 'video/mp4', 'audio/wav')
        this.duration = 0;        // Duration in seconds
        this.width = 0;           // Original width
        this.height = 0;          // Original height
        this.poster = '';         // Poster image for video
        this.autoplay = false;    // Auto-play flag
        this.loop = false;        // Loop flag
        this.controls = true;     // Show controls
        this.muted = false;       // Muted flag
        this.volume = 1.0;        // Volume (0.0 to 1.0)
        this.startTime = 0;       // Start time in seconds
        this.endTime = 0;         // End time in seconds
        this.relationship = null; // Relationship ID for embedded files
        this.embeddedData = null; // Base64 encoded data for embedded media
        this.thumbnail = null;    // Thumbnail image data
    }
}

/**
 * YouTube Embed Information
 */
class YouTubeEmbed {
    constructor() {
        this.videoId = '';        // YouTube video ID
        this.startTime = 0;       // Start time in seconds
        this.endTime = 0;         // End time in seconds
        this.autoplay = false;    // Auto-play flag
        this.showControls = true; // Show player controls
        this.showInfo = false;    // Show video info
        this.loop = false;        // Loop flag
        this.fullUrl = '';        // Full YouTube URL
        this.embedUrl = '';       // Embed URL
        this.thumbnailUrl = '';   // Thumbnail URL
    }
}

/**
 * Media Processor - Main class for handling media elements
 */
class MediaProcessor {
    constructor(context) {
        this.context = context;
        this.logger = new Logger('MediaProcessor');
        this.supportedVideoTypes = ['mp4', 'webm', 'ogg', 'avi', 'mov', 'wmv'];
        this.supportedAudioTypes = ['mp3', 'wav', 'ogg', 'aac', 'm4a', 'wma'];
        this.supportedImageTypes = ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'svg', 'webp'];
    }

    /**
     * Parse media from graphic frame or shape
     * @param {Element} element - The element containing media information
     * @return {MediaInfo|null} Parsed media information
     */
    parseMediaFromElement(element) {
        try {
            // Check for different media element types
            if (this.isVideoElement(element)) {
                return this.parseVideoElement(element);
            } else if (this.isAudioElement(element)) {
                return this.parseAudioElement(element);
            } else if (this.isImageElement(element)) {
                return this.parseImageElement(element);
            } else if (this.isYouTubeEmbed(element)) {
                return this.parseYouTubeEmbed(element);
            }

            return null;
        } catch (error) {
            this.logger.logError(this.constructor.name, 'Error parsing media element:', error);
            return null;
        }
    }

    /**
     * Check if element is a video element
     * @param {Element} element - Element to check
     * @return {boolean} True if video element
     */
    isVideoElement(element) {
        // Check for video-specific attributes or child elements
        const videoTags = ['video', 'p:video', 'a:video'];
        return videoTags.some(tag => element.querySelector(tag)) ||
               this.hasVideoMimeType(element);
    }

    /**
     * Check if element is an audio element
     * @param {Element} element - Element to check
     * @return {boolean} True if audio element
     */
    isAudioElement(element) {
        const audioTags = ['audio', 'p:audio', 'a:audio'];
        return audioTags.some(tag => element.querySelector(tag)) ||
               this.hasAudioMimeType(element);
    }

    /**
     * Check if element is an image element
     * @param {Element} element - Element to check
     * @return {boolean} True if image element
     */
    isImageElement(element) {
        const imageTags = ['pic', 'p:pic', 'image', 'img'];
        return imageTags.some(tag => element.querySelector(tag)) ||
               this.hasImageMimeType(element);
    }

    /**
     * Check if element is a YouTube embed
     * @param {Element} element - Element to check
     * @return {boolean} True if YouTube embed
     */
    isYouTubeEmbed(element) {
        const textContent = element.textContent || '';
        const htmlContent = element.innerHTML || '';
        
        return textContent.includes('youtube.com') ||
               textContent.includes('youtu.be') ||
               htmlContent.includes('youtube.com') ||
               htmlContent.includes('youtu.be');
    }

    /**
     * Parse video element
     * @param {Element} element - Video element
     * @return {MediaInfo} Video information
     */
    parseVideoElement(element) {
        const mediaInfo = new MediaInfo();
        mediaInfo.type = 'video';

        // Parse common media attributes
        this.parseCommonMediaAttributes(element, mediaInfo);

        // Parse video-specific attributes
        const videoElement = element.querySelector('video, p\\:video, a\\:video');
        if (videoElement) {
            mediaInfo.poster = videoElement.getAttribute('poster') || '';
            mediaInfo.autoplay = videoElement.hasAttribute('autoplay');
            mediaInfo.loop = videoElement.hasAttribute('loop');
            mediaInfo.controls = !videoElement.hasAttribute('controls') || 
                               videoElement.getAttribute('controls') !== 'false';
            mediaInfo.muted = videoElement.hasAttribute('muted');
        }

        // Try to extract video source
        this.extractMediaSource(element, mediaInfo);

        return mediaInfo;
    }

    /**
     * Parse audio element
     * @param {Element} element - Audio element
     * @return {MediaInfo} Audio information
     */
    parseAudioElement(element) {
        const mediaInfo = new MediaInfo();
        mediaInfo.type = 'audio';

        // Parse common media attributes
        this.parseCommonMediaAttributes(element, mediaInfo);

        // Parse audio-specific attributes
        const audioElement = element.querySelector('audio, p\\:audio, a\\:audio');
        if (audioElement) {
            mediaInfo.autoplay = audioElement.hasAttribute('autoplay');
            mediaInfo.loop = audioElement.hasAttribute('loop');
            mediaInfo.controls = !audioElement.hasAttribute('controls') || 
                               audioElement.getAttribute('controls') !== 'false';
            mediaInfo.muted = audioElement.hasAttribute('muted');
            
            const volumeAttr = audioElement.getAttribute('volume');
            if (volumeAttr) {
                mediaInfo.volume = parseFloat(volumeAttr) || 1.0;
            }
        }

        // Try to extract audio source
        this.extractMediaSource(element, mediaInfo);

        return mediaInfo;
    }

    /**
     * Parse image element
     * @param {Element} element - Image element
     * @return {MediaInfo} Image information
     */
    parseImageElement(element) {
        const mediaInfo = new MediaInfo();
        mediaInfo.type = 'image';

        // Parse common media attributes
        this.parseCommonMediaAttributes(element, mediaInfo);

        // Try to extract image source
        this.extractMediaSource(element, mediaInfo);

        return mediaInfo;
    }

    /**
     * Parse YouTube embed
     * @param {Element} element - Element containing YouTube reference
     * @return {MediaInfo} YouTube media information
     */
    parseYouTubeEmbed(element) {
        const mediaInfo = new MediaInfo();
        mediaInfo.type = 'youtube';

        const youtubeEmbed = new YouTubeEmbed();
        
        // Extract YouTube URL or video ID
        const textContent = element.textContent || '';
        const htmlContent = element.innerHTML || '';
        const content = textContent + ' ' + htmlContent;

        // Extract video ID from various YouTube URL formats
        const videoId = this.extractYouTubeVideoId(content);
        if (videoId) {
            youtubeEmbed.videoId = videoId;
            youtubeEmbed.fullUrl = `https://www.youtube.com/watch?v=${videoId}`;
            youtubeEmbed.embedUrl = `https://www.youtube.com/embed/${videoId}`;
            youtubeEmbed.thumbnailUrl = `https://img.youtube.com/vi/${videoId}/maxresdefault.jpg`;
        }

        // Extract start and end times if present
        const timeMatch = content.match(/[?&]t=(\d+)/);
        if (timeMatch) {
            youtubeEmbed.startTime = parseInt(timeMatch[1]);
        }

        mediaInfo.source = youtubeEmbed.embedUrl;
        mediaInfo.embeddedData = youtubeEmbed;

        return mediaInfo;
    }

    /**
     * Parse common media attributes
     * @param {Element} element - Media element
     * @param {MediaInfo} mediaInfo - Media info to populate
     */
    parseCommonMediaAttributes(element, mediaInfo) {
        // Try to get dimensions
        const width = element.getAttribute('width') || element.getAttribute('w');
        const height = element.getAttribute('height') || element.getAttribute('h');
        
        if (width) {mediaInfo.width = parseInt(width) || 0;}
        if (height) {mediaInfo.height = parseInt(height) || 0;}

        // Try to get duration
        const duration = element.getAttribute('duration');
        if (duration) {
            mediaInfo.duration = parseFloat(duration) || 0;
        }

        // Try to get relationship ID for embedded files
        const rId = element.getAttribute('r:id') || element.getAttribute('rid');
        if (rId) {
            mediaInfo.relationship = rId;
        }
    }

    /**
     * Extract media source from element
     * @param {Element} element - Media element
     * @param {MediaInfo} mediaInfo - Media info to populate
     */
    extractMediaSource(element, mediaInfo) {
        // Try different methods to find source
        
        // Method 1: Direct src attribute
        const src = element.getAttribute('src');
        if (src) {
            mediaInfo.source = src;
            mediaInfo.mimeType = this.getMimeTypeFromExtension(src);
            return;
        }

        // Method 2: Child source element
        const sourceElement = element.querySelector('source');
        if (sourceElement) {
            mediaInfo.source = sourceElement.getAttribute('src') || '';
            mediaInfo.mimeType = sourceElement.getAttribute('type') || 
                               this.getMimeTypeFromExtension(mediaInfo.source);
            return;
        }

        // Method 3: Relationship reference
        if (mediaInfo.relationship && this.context) {
            // Try to resolve relationship to actual file
            const resolvedSource = this.resolveRelationship(mediaInfo.relationship);
            if (resolvedSource) {
                mediaInfo.source = resolvedSource;
                mediaInfo.mimeType = this.getMimeTypeFromExtension(resolvedSource);
                
                // For SVG files, try to create data URI if we have access to ZIP
                if (this.getMimeTypeFromExtension(resolvedSource) === 'image/svg+xml') {
                    try {
                        const zipPackage = this.getZipPackage();
                        if (zipPackage && zipPackage.getFileData) {
                            // Try to load the SVG file data and create data URI
                            zipPackage.getFileData(resolvedSource).then(fileData => {
                                if (fileData) {
                                    const svgContent = new TextDecoder('utf-8').decode(fileData);
                                    const base64Data = btoa(svgContent);
                                    mediaInfo.embeddedData = `data:image/svg+xml;base64,${base64Data}`;
                                }
                            }).catch(error => {
                                this.logger.log("warn", this.constructor.name, `Failed to create SVG data URI for ${resolvedSource}:`, error);
                            });
                        }
                    } catch (error) {
                        this.logger.log("warn", this.constructor.name, `Error creating SVG data URI for ${resolvedSource}:`, error);
                    }
                }
                return;
            }
        }

        // Method 4: Embedded data
        const embeddedData = this.extractEmbeddedData(element);
        if (embeddedData) {
            mediaInfo.embeddedData = embeddedData;
            mediaInfo.mimeType = this.getMimeTypeFromData(embeddedData);
        }
    }

    /**
     * Get ZIP package from context
     * @return {Object|null} ZIP package object
     */
    getZipPackage() {
        if (this.context && this.context.zip) {
            return this.context.zip;
        } else if (this.context && this.context.package) {
            return this.context.package;
        } else if (window.currentProcessor && window.currentProcessor.package) {
            return window.currentProcessor.package;
        } else if (window.currentProcessor && window.currentProcessor.zip) {
            return window.currentProcessor.zip;
        }
        return null;
    }

    /**
     * Extract YouTube video ID from various URL formats
     * @param {string} content - Content containing YouTube URL
     * @return {string|null} Video ID or null if not found
     */
    extractYouTubeVideoId(content) {
        const patterns = [
            /(?:youtube\.com\/watch\?v=|youtu\.be\/)([^&\n?#]+)/,
            /youtube\.com\/embed\/([^&\n?#]+)/,
            /youtube\.com\/v\/([^&\n?#]+)/
        ];

        for (const pattern of patterns) {
            const match = content.match(pattern);
            if (match && match[1]) {
                return match[1];
            }
        }

        return null;
    }

    /**
     * Get MIME type from file extension
     * @param {string} filename - File name or path
     * @return {string} MIME type
     */
    getMimeTypeFromExtension(filename) {
        if (!filename) {return '';}

        const extension = filename.split('.').pop().toLowerCase();
        
        const mimeTypes = {
            // Video
            'mp4': 'video/mp4',
            'webm': 'video/webm',
            'ogg': 'video/ogg',
            'avi': 'video/x-msvideo',
            'mov': 'video/quicktime',
            'wmv': 'video/x-ms-wmv',
            
            // Audio
            'mp3': 'audio/mpeg',
            'wav': 'audio/wav',
            'aac': 'audio/aac',
            'm4a': 'audio/mp4',
            'wma': 'audio/x-ms-wma',
            
            // Image
            'jpg': 'image/jpeg',
            'jpeg': 'image/jpeg',
            'png': 'image/png',
            'gif': 'image/gif',
            'bmp': 'image/bmp',
            'svg': 'image/svg+xml',
            'webp': 'image/webp'
        };

        return mimeTypes[extension] || '';
    }

    /**
     * Resolve relationship ID to actual file path
     * @param {string} relationshipId - Relationship ID
     * @return {string|null} Resolved file path
     */
    resolveRelationship(relationshipId) {
        try {
            if (!relationshipId || !this.context) {
                return null;
            }

            // Try to get the ZIP package from context
            let zipPackage = null;
            if (this.context.zip) {
                zipPackage = this.context.zip;
            } else if (this.context.package) {
                zipPackage = this.context.package;
            } else if (window.currentProcessor && window.currentProcessor.package) {
                zipPackage = window.currentProcessor.package;
            }

            if (!zipPackage || !zipPackage.relationships) {
                return null;
            }

            // Look for the relationship in all available relationship maps
            const relationshipMaps = Object.values(zipPackage.relationships);
            
            for (const relMap of relationshipMaps) {
                if (relMap && relMap[relationshipId]) {
                    const rel = relMap[relationshipId];
                    let mediaPath = rel.target;
                    
                    // Convert relative path to absolute path
                    if (mediaPath.startsWith('../')) {
                        mediaPath = mediaPath.replace('../', 'ppt/');
                    } else if (!mediaPath.startsWith('/')) {
                        mediaPath = `/ppt/${mediaPath}`;
                    }
                    
                    return mediaPath;
                }
            }

            this.logger.log("warn", this.constructor.name, `Could not resolve relationship ID: ${relationshipId}`);
            return null;
        } catch (error) {
            this.logger.logError(this.constructor.name, `Error resolving relationship ${relationshipId}:`, error);
            return null;
        }
    }

    /**
     * Extract embedded data from element
     * @param {Element} element - Element to check
     * @return {string|null} Base64 encoded data
     */
    extractEmbeddedData(element) {
        // Look for base64 data in various attributes
        const dataAttrs = ['data', 'data-src', 'data-url', 'src', 'href'];
        
        for (const attr of dataAttrs) {
            const data = element.getAttribute(attr);
            if (data && data.startsWith('data:')) {
                return data;
            }
        }

        // Also check text content for embedded data URIs (common in PPTX)
        const textContent = element.textContent || '';
        const dataUriMatch = textContent.match(/data:[^;]+;base64,[A-Za-z0-9+/]+=*/g);
        if (dataUriMatch && dataUriMatch.length > 0) {
            return dataUriMatch[0];
        }

        return null;
    }

    /**
     * Get MIME type from base64 data
     * @param {string} data - Base64 data string
     * @return {string} MIME type
     */
    getMimeTypeFromData(data) {
        if (!data || !data.startsWith('data:')) {return '';}
        
        const match = data.match(/^data:([^;]+)/);
        return match ? match[1] : '';
    }

    /**
     * Check if data URI contains SVG content
     * @param {string} dataUri - Data URI string
     * @return {boolean} True if contains SVG
     */
    isSVGDataUri(dataUri) {
        if (!dataUri || !dataUri.startsWith('data:')) {return false;}
        
        const mimeType = this.getMimeTypeFromData(dataUri);
        return mimeType === 'image/svg+xml';
    }

    /**
     * Extract SVG content from base64 data URI
     * @param {string} dataUri - SVG data URI
     * @return {string|null} SVG content string
     */
    extractSVGFromDataUri(dataUri) {
        if (!this.isSVGDataUri(dataUri)) {return null;}
        
        try {
            const base64Data = dataUri.split(',')[1];
            if (!base64Data) {return null;}
            
            const svgContent = atob(base64Data);
            return svgContent;
        } catch (error) {
            return null;
        }
    }

    /**
     * Check if element has video MIME type
     * @param {Element} element - Element to check
     * @return {boolean} True if has video MIME type
     */
    hasVideoMimeType(element) {
        const type = element.getAttribute('type') || '';
        return type.startsWith('video/');
    }

    /**
     * Check if element has audio MIME type
     * @param {Element} element - Element to check
     * @return {boolean} True if has audio MIME type
     */
    hasAudioMimeType(element) {
        const type = element.getAttribute('type') || '';
        return type.startsWith('audio/');
    }

    /**
     * Check if element has image MIME type
     * @param {Element} element - Element to check
     * @return {boolean} True if has image MIME type
     */
    hasImageMimeType(element) {
        const type = element.getAttribute('type') || '';
        return type.startsWith('image/');
    }

    /**
     * Create media placeholder information
     * @param {string} type - Media type
     * @return {MediaInfo} Placeholder media info
     */
    createPlaceholder(type = 'media') {
        const mediaInfo = new MediaInfo();
        mediaInfo.type = type;
        mediaInfo.source = '';
        mediaInfo.width = 320;
        mediaInfo.height = 240;
        mediaInfo.controls = true;
        
        return mediaInfo;
    }
}

// Export classes

// Export classes (maintain backward compatibility)
if (typeof window !== 'undefined') {
    window.MediaInfo = MediaInfo;
    window.YouTubeEmbed = YouTubeEmbed;
    window.MediaProcessor = MediaProcessor;
}

// ES Module exports (disabled for script-tag compatibility)
// export { MediaInfo,YouTubeEmbed,MediaProcessor };
