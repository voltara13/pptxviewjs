/**
 * Graphics Adapter for Slide Editor
 * Integrates CGraphics and text rendering with the PPTX slide editor
 * Based on modern graphics rendering components
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

// Import required modules
// import { ChartRenderer } from '../renderers/chart-renderer.js';

// CoordinateTransform will be available from utils.js loaded globally

// Import graphics components if they exist
if (typeof AscCommon === 'undefined') {
    window.AscCommon = {};
}

// Defer adapter prototype augmentation until after class is defined
function __augmentAdapterWithArrows__() {
    if (typeof CanvasGraphicsAdapter === 'undefined') { return; }
    if (CanvasGraphicsAdapter.prototype.drawLineWithArrows) { return; }
    CanvasGraphicsAdapter.prototype.drawLineWithArrows = function(x1, y1, x2, y2, strokeColor, lineWidth = 1, strokeInfo = null) {
        const ctx = this._context;
        if (!ctx) {return;}

        ctx.save();
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);

        const color = (typeof this.colorToRgb === 'function' ? this.colorToRgb(strokeColor) : null) || 'rgba(0,0,0,1)';
        const pxWidth = Math.max(CoordinateTransform.mmToPixels(lineWidth || 1), 1);
        ctx.strokeStyle = color;
        ctx.lineWidth = pxWidth;
        if (strokeInfo && strokeInfo.dashArray && strokeInfo.dashArray.length > 0) {
            const scaledDashArray = strokeInfo.dashArray.map(dash => dash * pxWidth);
            ctx.setLineDash(scaledDashArray);
        }
        if (strokeInfo && typeof this.applyStrokeStyle === 'function') {
            try { this.applyStrokeStyle(ctx, strokeInfo, pxWidth); } catch(_e) {}
        }

        ctx.stroke();

        // Render arrowheads using XML semantics: tailEnd → start, headEnd → end
        // Map PPTX arrow size attributes to scale factors
        // PPTX uses: sm (small), med (medium), lg (large) for both w (width) and len (length)
        const mapSize = (s) => {
            if (typeof s === 'number') {return s;}
            const v = (s || '').toString().toLowerCase();
            // PPTX specification size mappings
            // These values are based on PowerPoint's rendering behavior
            if (v === 'lg' || v === 'large') {return 3.0;}  // Large arrows are ~3x base size
            if (v === 'med' || v === 'medium') {return 2.0;} // Medium arrows are ~2x base size
            if (v === 'sm' || v === 'small') {return 1.0;}   // Small arrows are 1x base size
            return 2.0; // Default to medium if not specified
        };
        
        // Per request: headEnd applies at start (x1,y1), tailEnd at end (x2,y2)
        const startDef = strokeInfo && strokeInfo.headEnd;
        const endDef = strokeInfo && strokeInfo.tailEnd;

        // Suppress 'none'
        const sType = startDef && startDef.type;
        const eType = endDef && endDef.type;
        const hasStart = !!(startDef && sType && sType !== 'none');
        const hasEnd = !!(endDef && eType && eType !== 'none');
        if (hasStart || hasEnd) {
            const angle = Math.atan2(y2 - y1, x2 - x1);
            // Get current transformation scale to make arrow heads proportional to display size
            const transform = ctx.getTransform();
            const currentScale = Math.sqrt(transform.a * transform.a + transform.b * transform.b);
            
            // Calculate arrow head size from line width
            // PPTX arrow heads scale with line width: length ~= 3*lineWidth, width ~= 2.5*lineWidth
            // These multipliers match PowerPoint's actual rendering behavior more closely
            const headLen = pxWidth * 3.0;  // Length proportional to line width
            const headWid = pxWidth * 2.5;  // Width proportional to line width
            

            const drawHead = (x, y, ang, def) => {
                if (!def) {return;}
                // Apply PPTX size multipliers (w and len attributes)
                const lengthScale = def.lengthScale || mapSize(def.len);
                const widthScale = def.widthScale || mapSize(def.w);
                const length = headLen * lengthScale;
                const width = headWid * widthScale;
                const type = (def.type || def.val || 'arrow').toString().toLowerCase();
                const halfW = width / 2;
                
                ctx.save();
                ctx.translate(x, y);
                ctx.rotate(ang);
                ctx.beginPath();
                
                // Draw different arrow types based on PPTX type attribute
                // Types: triangle, stealth, diamond, oval, arrow, open
                if (type === 'open') {
                    // Open arrow - two lines forming a V shape
                    ctx.moveTo(0, 0);
                    ctx.lineTo(-length, -halfW);
                    ctx.moveTo(0, 0);
                    ctx.lineTo(-length, halfW);
                    ctx.strokeStyle = color;
                    ctx.lineWidth = Math.max(1, pxWidth);
                    ctx.stroke();
                } else if (type === 'stealth') {
                    // Stealth arrow - pointed/streamlined triangle
                    ctx.moveTo(0, 0);
                    ctx.lineTo(-length, -(halfW * 0.5));
                    ctx.lineTo(-length * 0.85, 0); // Concave back for stealth look
                    ctx.lineTo(-length, (halfW * 0.5));
                    ctx.closePath();
                    ctx.fillStyle = color;
                    ctx.fill();
                    ctx.strokeStyle = color;
                    ctx.lineWidth = Math.max(0.5, pxWidth * 0.3);
                    ctx.stroke();
                } else if (type === 'diamond') {
                    // Diamond arrow - rhombus shape
                    ctx.moveTo(0, 0);
                    ctx.lineTo(-length / 2, -halfW);
                    ctx.lineTo(-length, 0);
                    ctx.lineTo(-length / 2, halfW);
                    ctx.closePath();
                    ctx.fillStyle = color;
                    ctx.fill();
                    ctx.strokeStyle = color;
                    ctx.lineWidth = Math.max(0.5, pxWidth * 0.3);
                    ctx.stroke();
                } else if (type === 'oval') {
                    // Oval arrow - ellipse/circle
                    ctx.ellipse(-length / 2, 0, length / 2, halfW, 0, 0, Math.PI * 2);
                    ctx.fillStyle = color;
                    ctx.fill();
                    ctx.strokeStyle = color;
                    ctx.lineWidth = Math.max(0.5, pxWidth * 0.3);
                    ctx.stroke();
                } else {
                    // Default: triangle/arrow - standard filled triangle
                    ctx.moveTo(0, 0);
                    ctx.lineTo(-length, -halfW);
                    ctx.lineTo(-length, halfW);
                    ctx.closePath();
                    ctx.fillStyle = color;
                    ctx.fill();
                    ctx.strokeStyle = color;
                    ctx.lineWidth = Math.max(0.5, pxWidth * 0.3);
                    ctx.stroke();
                }
                ctx.restore();
            };

            // Draw arrow heads at appropriate ends
            if (hasStart) { drawHead(x1, y1, angle + Math.PI, startDef); }
            if (hasEnd) { drawHead(x2, y2, angle, endDef); }
        }

        ctx.restore();
    };
}
if (typeof AscFormat === 'undefined') {
    window.AscFormat = {};
}

/**
 * Graphics Matrix Implementation - Based on GraphicsMatrix
 */
class GraphicsMatrix {
    constructor() {
        this.sx = 1.0;   // Scale X
        this.sy = 1.0;   // Scale Y
        this.shx = 0.0;  // Shear X
        this.shy = 0.0;  // Shear Y
        this.tx = 0.0;   // Translate X
        this.ty = 0.0;   // Translate Y
    }

    TransformPointX(x, y) {
        return x * this.sx + y * this.shx + this.tx;
    }

    TransformPointY(x, y) {
        return x * this.shy + y * this.sy + this.ty;
    }

    Invert() {
        const det = this.sx * this.sy - this.shx * this.shy;
        if (Math.abs(det) < 1e-12) {
            return;
        }

        const invDet = 1.0 / det;
        const newSx = this.sy * invDet;
        const newSy = this.sx * invDet;
        const newShx = -this.shx * invDet;
        const newShy = -this.shy * invDet;
        const newTx = (this.shx * this.ty - this.sy * this.tx) * invDet;
        const newTy = (this.shy * this.tx - this.sx * this.ty) * invDet;

        this.sx = newSx;
        this.sy = newSy;
        this.shx = newShx;
        this.shy = newShy;
        this.tx = newTx;
        this.ty = newTy;
    }

    createDuplicate() {
        const matrix = new GraphicsMatrix();
        matrix.sx = this.sx;
        matrix.sy = this.sy;
        matrix.shx = this.shx;
        matrix.shy = this.shy;
        matrix.tx = this.tx;
        matrix.ty = this.ty;
        return matrix;
    }

    reset() {
        this.sx = 1.0;
        this.sy = 1.0;
        this.shx = 0.0;
        this.shy = 0.0;
        this.tx = 0.0;
        this.ty = 0.0;
    }
}

/**
 * Brush Implementation
 */
if (typeof CBrush === 'undefined') {
    window.CBrush = class CBrush {
        constructor() {
            this.Color1 = { R: 0, G: 0, B: 0, A: 255 };
            this.Color2 = { R: 0, G: 0, B: 0, A: 255 };
            this.Type = 0; // 0 = solid, 1 = gradient, etc.
        }

        GetType() {
            return this.Type;
        }
    };
}

/**
 * Pen Implementation
 */
if (typeof CPen === 'undefined') {
    window.CPen = class CPen {
        constructor() {
            this.Color = { R: 0, G: 0, B: 0, A: 255 };
            this.Size = 1.0;
            this.DashStyle = 0; // 0 = solid
        }
    };
}

/**
 * Font Manager Mock
 */
if (typeof CFontManager === 'undefined') {
    window.CFontManager = class CFontManager {
    constructor() {
        this.m_oGlyphString = {
            m_fX: 0,
            m_fY: 0,
            m_pGlyphsBuffer: []
        };
    }

    LoadString4C(code, x, y) {
        // Mock implementation - would normally load font glyphs
        this.m_oGlyphString.m_fX = x;
        this.m_oGlyphString.m_fY = y;
        return x + 10; // Mock advance width
    }

    LoadString2(text, x, y) {
        // Mock implementation for text strings
        this.m_oGlyphString.m_fX = x;
        this.m_oGlyphString.m_fY = y;
    }

    GetNextChar2() {
        // Mock implementation
        return null;
    }
    };
}

/**
 * Canvas Graphics Implementation - Based on sdkjs CGraphics
 */
class CanvasGraphicsAdapter {
    constructor() {
        this.canvas = null;
        this._context = null;  // Use private property for the actual context
        this.widthPx = 0;
        this.heightPx = 0;
        this.widthMM = 0;
        this.heightMM = 0;
        this.dpiX = 96;
        this.dpiY = 96;
        this.transform = new GraphicsMatrix();
        this.grStateStack = [];
        this.clipStack = [];
        this.font = null;
        this.brush = new CBrush();
        this.pen = new CPen();
        this.currentPath = null;
        this.isPathStarted = false;
        
        // Enhanced custom geometry support
        this.customGeometryProcessor = new CustomGeometryProcessor();
    }

    /**
     * Get the canvas context (for compatibility with graphics adapter expectations)
     */
    get context() {
        return this._context;
    }

    /**
     * Initialize graphics context
     */
    init(context, widthPx, heightPx, widthMM, heightMM) {
        this._context = context;
        this.widthPx = widthPx;
        this.heightPx = heightPx;
        this.widthMM = widthMM || (widthPx * 25.4 / 96); // Default to 96 DPI if not provided
        this.heightMM = heightMM || (heightPx * 25.4 / 96);

        this.dpiX = 96.0; // Use standard DPI
        this.dpiY = 96.0;

        // Set up coordinate transformation - use 1:1 pixel mapping
        this.transform.sx = 1.0;
        this.transform.sy = 1.0;
        this.transform.shx = 0.0;
        this.transform.shy = 0.0;
        this.transform.tx = 0.0;
        this.transform.ty = 0.0;

        this.recalculateTransforms();
    }

    /**
     * Recalculate transforms
     */
    recalculateTransforms() {
        // Combine coordinate transform and local transform
        this.transform = this.transform.createDuplicate();
        // Apply local transform
        const sx = this.transform.sx * this.transform.sx + this.transform.shx * this.transform.shy;
        const sy = this.transform.sy * this.transform.sy + this.transform.shy * this.transform.shx;
        const shx = this.transform.sx * this.transform.shx + this.transform.shx * this.transform.sy;
        const shy = this.transform.sy * this.transform.shy + this.transform.shy * this.transform.sx;
        const tx = this.transform.sx * this.transform.tx + this.transform.shx * this.transform.ty + this.transform.tx;
        const ty = this.transform.sy * this.transform.ty + this.transform.shy * this.transform.tx + this.transform.ty;

        this.transform.sx = sx;
        this.transform.sy = sy;
        this.transform.shx = shx;
        this.transform.shy = shy;
        this.transform.tx = tx;
        this.transform.ty = ty;

        // Create inverted transform
        this.transform = this.transform.createDuplicate();
        this.transform.Invert();
    }

    /**
     * Reset graphics adapter state for new file load
     * This clears persistent state that could interfere with new presentations
     */
    resetState() {
        // Clear coordinate system that persists between file loads
        this.coordinateSystem = null;
        
        // Reset transform to identity
        this.transform.reset();
        
        // Clear graphics state stacks
        this.grStateStack = [];
        this.clipStack = [];
        
        // Reset current rendering shape reference
        this.currentRenderingShape = null;
        
        // Clear any cached geometry or path data
        if (this.customGeometryProcessor) {
            this.customGeometryProcessor.cache.clear();
        }
        
        // Reset path state
        this.currentPath = null;
        this.isPathStarted = false;
        
        // Reset brush and pen to defaults
        this.brush = new CBrush();
        this.pen = new CPen();
    }

    /**
     * Save state
     */
    SaveGrState() {
        if (this._context) {
            this._context.save();
            this.grStateStack.push({
                transform: this.transform.createDuplicate(),
                fullTransform: this.transform.createDuplicate(),
                invertFullTransform: this.transform.createDuplicate()
            });
        }
    }

    /**
     * Restore state
     */
    RestoreGrState() {
        if (this._context) {
            this._context.restore();
            if (this.grStateStack.length > 0) {
                const state = this.grStateStack.pop();
                this.transform = state.transform;
                this.transform = state.fullTransform;
                this.transform = state.invertFullTransform;
            }
        }
    }

    /**
     * Set transformation matrix
     */
    transform3(matrix) {
        this.transform = matrix.createDuplicate();
        this.recalculateTransforms();

        if (this._context) {
            this._context.setTransform(
                this.transform.sx,
                this.transform.shy,
                this.transform.shx,
                this.transform.sy,
                this.transform.tx,
                this.transform.ty
            );
        }
    }

    /**
     * Add clipping rectangle
     */
    AddClipRect(x, y, w, h) {
        if (this._context) {
            this._context.save();
            this._context.beginPath();
            this._context.rect(x, y, w, h);
            this._context.clip();
        }
    }

    /**
     * Clear canvas
     */
    clear() {
        if (this._context) {
            this._context.save();
            this._context.setTransform(1, 0, 0, 1, 0, 0);
            this._context.clearRect(0, 0, this.widthPx, this.heightPx);
            this._context.restore();
        }
    }



    /**
     * Text rendering - t implementation (text string)
     * Following the pattern from sdkjs/word/Drawing/Graphics.js t method
     */
    t(text, x, y, isBounds) {
        if (this.bIsBreak || !this._context) {
            return;
        }

        // Use coordinates directly without inversion for text rendering
        const _x = x;
        const _y = y;

        // Set text style
        this._context.fillStyle = this.colorToRgb(this.brush.Color1);
        this._context.textBaseline = 'alphabetic';

        // Render text character by character for proper glyph handling
        let currentX = _x;
        let bounds = null;

        if (isBounds) {
            bounds = { x: _x, y: _y, r: _x, b: _y };
        }

        // Render each character individually
        for (let i = 0; i < text.length; i++) {
            const char = text.charAt(i);

            // Render character
            this._context.fillText(char, currentX, _y);

            // Measure character for advancement and bounds
            const metrics = this._context.measureText(char);

            if (isBounds) {
                // Update bounds
                bounds.r = Math.max(bounds.r, currentX + metrics.width);
                if (metrics.actualBoundingBoxAscent !== undefined) {
                    bounds.y = Math.min(bounds.y, _y - metrics.actualBoundingBoxAscent);
                }
                if (metrics.actualBoundingBoxDescent !== undefined) {
                    bounds.b = Math.max(bounds.b, _y + metrics.actualBoundingBoxDescent);
                }
            }

            // Advance position
            currentX += metrics.width;
        }

        return bounds;
    }

    /**
     * Set brush color
     */
    b_color1(r, g, b, a = 255) {
        this.brush.Color1.R = r;
        this.brush.Color1.G = g;
        this.brush.Color1.B = b;
        this.brush.Color1.A = a;
        this.bBrushColorInit = true;
    }

    /**
     * Set pen color
     */
    p_color(r, g, b, a = 255) {
        this.pen.Color.R = r;
        this.pen.Color.G = g;
        this.pen.Color.B = b;
        this.pen.Color.A = a;
        this.bPenColorInit = true;
    }

    /**
     * Set pen width
     */
    p_width(w) {
        this.pen.Size = w;
        if (this._context) {
            this._context.lineWidth = w;
        }
    }

    /**
     * Fill rectangle
     */
    fillRect(x, y, w, h, color) {
        if (!this._context || !color) {return;}

        this._context.save();
        this._context.fillStyle = this.colorToRgb(color);
        this._context.fillRect(x, y, w, h);
        this._context.restore();
    }

    /**
     * Stroke rectangle
     */
    strokeRect(x, y, w, h, color, lineWidth = 1) {
        if (!this._context) {return;}

        this._context.save();
        this._context.strokeStyle = this.colorToRgb(color) || this.colorToRgb(this.pen.Color);
        this._context.lineWidth = lineWidth;
        this._context.strokeRect(x, y, w, h);
        this._context.restore();
    }

    /**
     * Draw a line between two points
     */
    drawLine(x1, y1, x2, y2, color, lineWidth = 1) {
        if (!this._context) {return;}

        this._context.save();
        this._context.strokeStyle = this.colorToRgb(color);
        this._context.lineWidth = lineWidth;
        this._context.beginPath();
        this._context.moveTo(x1, y1);
        this._context.lineTo(x2, y2);
        this._context.stroke();
        this._context.restore();
    }

    /**
     * Fill a circle
     */
    fillCircle(x, y, radius, color) {
        if (!this._context) {return;}

        this._context.save();
        this._context.fillStyle = this.colorToRgb(color);
        this._context.beginPath();
        this._context.arc(x, y, radius, 0, 2 * Math.PI);
        this._context.fill();
        this._context.restore();
    }

    /**
     * Stroke a circle
     */
    strokeCircle(x, y, radius, color, lineWidth = 1) {
        if (!this._context) {return;}

        this._context.save();
        this._context.strokeStyle = this.colorToRgb(color);
        this._context.lineWidth = lineWidth;
        this._context.beginPath();
        this._context.arc(x, y, radius, 0, 2 * Math.PI);
        this._context.stroke();
        this._context.restore();
    }

    /**
     * Fill text with specified style
     */
    fillText(text, x, y, style = {}) {
        if (!this._context) {return;}

        this._context.save();
        
        // Set font
        const fontSize = style.fontSize || style.size || 12;
        const fontFamily = style.fontFamily || style.family || 'Arial';
        const fontWeight = style.bold ? 'bold' : 'normal';
        const fontStyle = style.italic ? 'italic' : 'normal';
        this._context.font = `${fontStyle} ${fontWeight} ${fontSize}px ${fontFamily}`;
        
        // Set color
        this._context.fillStyle = this.colorToRgb(style.color || { r: 0, g: 0, b: 0 });
        
        // Set text alignment - use consistent baseline for accurate positioning
        this._context.textAlign = style.textAlign || 'left';
        this._context.textBaseline = style.textBaseline || 'alphabetic'; // Consistent baseline
        
        // Fill text
        this._context.fillText(text, x, y);
        
        this._context.restore();
    }

    /**
     * Convert color object to CSS color string
     */
    colorToRgb(color) {
        if (typeof color === 'string') {
            return color;
        }
        if (color && typeof color === 'object') {
            // Handle both uppercase and lowercase color properties
            const r = color.r !== undefined ? color.r : (color.R !== undefined ? color.R : 0);
            const g = color.g !== undefined ? color.g : (color.G !== undefined ? color.G : 0);
            const b = color.b !== undefined ? color.b : (color.B !== undefined ? color.B : 0);
            const a = color.a !== undefined ? color.a : (color.A !== undefined ? color.A : 255);

            if (a !== undefined && a !== 255) {
                return `rgba(${r}, ${g}, ${b}, ${a / 255})`;
            }
            return `rgb(${r}, ${g}, ${b})`;
        }
        return 'rgb(0, 0, 0)';
    }

    /**
     * Command support
     */
    Start_Command(commandId) {
        this.commandStack.push(commandId);
    }

    End_Command(commandId) {
        if (this.commandStack.length > 0) {
            this.commandStack.pop();
        }
    }

    /**
     * Get context for external use
     */
    get context() {
        return this._context;
    }

    /**
     * Bounds checker support
     */
    isBoundsChecker() {
        return false; // For now, not implementing bounds checking
    }
    /**
     * Draw preset geometry shapes
     */
    drawPresetGeometry(shapeType, x, y, width, height, fillColor, strokeColor, lineWidth = 1, strokeInfo = null, adjustments = {}) {

        if (!this._context) {return;}

        const ctx = this._context;
        ctx.save();

        // Apply opacity from fill color alpha - handles all rendering paths
        if (fillColor && typeof fillColor === 'object' && fillColor.a !== undefined && fillColor.a < 255) {
            ctx.globalAlpha = fillColor.a / 255;
            fillColor = { ...fillColor, a: 255 };
        }

        // Set fill color
        let isGradientFill = false;
        if (fillColor && typeof fillColor === 'object' && fillColor.type === 'linear' && fillColor.stops) {
            // Handle gradient fill - create canvas gradient
            isGradientFill = true;
            const bounds = { x, y, w: width, h: height };
            const gradient = this.createCanvasGradient ? this.createCanvasGradient(ctx, bounds, fillColor) : null;
            if (gradient) {
                ctx.fillStyle = gradient;
            } else if (fillColor.stops[0]?.color) {
                ctx.fillStyle = this.colorToRgb(fillColor.stops[0].color);
            }
        } else if (fillColor) {
            ctx.fillStyle = this.colorToRgb(fillColor);
        }

        // Set stroke color and width
        if (strokeColor) {
            ctx.strokeStyle = this.colorToRgb(strokeColor);
            ctx.lineWidth = CoordinateTransform.mmToPixels(lineWidth || 1);
        }

        // Draw the shape based on type
        ctx.beginPath();

        switch (shapeType) {
            case 'teardrop':
                this.drawTeardropPath(ctx, x, y, width, height);
                break;
            case 'rect':
            case 'rectangle':
                ctx.rect(x, y, width, height);
                break;
            case 'ellipse':
            case 'oval':
                ctx.ellipse(x + width/2, y + height/2, width/2, height/2, 0, 0, 2 * Math.PI);
                break;
            case 'roundRect': {
                const adjVal = adjustments.adj !== undefined ? adjustments.adj : 16667;
                const clampedAdj = Math.min(50000, Math.max(0, adjVal));
                const radius = (clampedAdj / 100000) * Math.min(width, height);
                this.drawRoundRectPath(ctx, x, y, width, height, radius);
                break;
            }
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
            case 'star4':
                this.drawStar4Path(ctx, x, y, width, height);
                break;
            case 'star6':
                this.drawStar6Path(ctx, x, y, width, height);
                break;
            case 'star8':
                this.drawStar8Path(ctx, x, y, width, height);
                break;
            case 'star10':
                this.drawStar10Path(ctx, x, y, width, height);
                break;
            case 'star12':
                this.drawStar12Path(ctx, x, y, width, height);
                break;
            case 'star16':
                this.drawStar16Path(ctx, x, y, width, height);
                break;
            case 'star24':
                this.drawStar24Path(ctx, x, y, width, height);
                break;
            case 'star32':
                this.drawStar32Path(ctx, x, y, width, height);
                break;
            case 'heart':
                this.drawHeartPath(ctx, x, y, width, height);
                break;
            // Flowchart shapes implementation
            case 'flowChartProcess':
                this.drawFlowChartProcessPath(ctx, x, y, width, height);
                break;
            case 'flowChartDecision':
                this.drawFlowChartDecisionPath(ctx, x, y, width, height);
                break;
            case 'flowChartInputOutput':
                this.drawFlowChartInputOutputPath(ctx, x, y, width, height);
                break;
            case 'flowChartPredefinedProcess':
                this.drawFlowChartPredefinedProcessPath(ctx, x, y, width, height);
                break;
            case 'flowChartInternalStorage':
                this.drawFlowChartInternalStoragePath(ctx, x, y, width, height);
                break;
            case 'flowChartDocument':
                this.drawFlowChartDocumentPath(ctx, x, y, width, height);
                break;
            case 'flowChartMultidocument':
                this.drawFlowChartMultidocumentPath(ctx, x, y, width, height);
                break;
            case 'flowChartTerminator':
                this.drawRoundRectPath(ctx, x, y, width, height, Math.min(width, height) * 0.2);
                break;
            case 'flowChartPreparation':
                this.drawFlowChartPreparationPath(ctx, x, y, width, height);
                break;
            case 'flowChartManualInput':
                this.drawFlowChartManualInputPath(ctx, x, y, width, height);
                break;
            case 'flowChartManualOperation':
                this.drawFlowChartManualOperationPath(ctx, x, y, width, height);
                break;
            case 'flowChartConnector':
                ctx.ellipse(x + width/2, y + height/2, width/2, height/2, 0, 0, 2 * Math.PI);
                break;
            case 'flowChartOffpageConnector':
                this.drawFlowChartOffpageConnectorPath(ctx, x, y, width, height);
                break;
            case 'flowChartPunchedCard':
                this.drawFlowChartPunchedCardPath(ctx, x, y, width, height);
                break;
            case 'flowChartPunchedTape':
                this.drawFlowChartPunchedTapePath(ctx, x, y, width, height);
                break;
            case 'flowChartSummingJunction':
                this.drawFlowChartSummingJunctionPath(ctx, x, y, width, height);
                break;
            case 'flowChartOr':
                this.drawFlowChartOrPath(ctx, x, y, width, height);
                break;
            case 'flowChartCollate':
                this.drawFlowChartCollatePath(ctx, x, y, width, height);
                break;
            case 'flowChartSort':
                this.drawFlowChartSortPath(ctx, x, y, width, height);
                break;
            case 'flowChartExtract':
                this.drawFlowChartExtractPath(ctx, x, y, width, height);
                break;
            case 'flowChartMerge':
                this.drawFlowChartMergePath(ctx, x, y, width, height);
                break;
            case 'flowChartStoredData':
                this.drawFlowChartStoredDataPath(ctx, x, y, width, height);
                break;
            case 'flowChartDelay':
                this.drawFlowChartDelayPath(ctx, x, y, width, height);
                break;
            case 'flowChartMagneticTape':
                this.drawFlowChartMagneticTapePath(ctx, x, y, width, height);
                break;
            case 'flowChartMagneticDisk':
                this.drawFlowChartMagneticDiskPath(ctx, x, y, width, height);
                break;
            case 'flowChartMagneticDrum':
                this.drawFlowChartMagneticDrumPath(ctx, x, y, width, height);
                break;
            case 'flowChartDisplay':
                this.drawFlowChartDisplayPath(ctx, x, y, width, height);
                break;
            // Arrow shapes
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
            case 'leftRightArrow':
                this.drawLeftRightArrowPath(ctx, x, y, width, height);
                break;
            case 'upDownArrow':
                this.drawUpDownArrowPath(ctx, x, y, width, height);
                break;
            case 'quadArrow':
                this.drawQuadArrowPath(ctx, x, y, width, height);
                break;
            case 'bentArrow':
                this.drawBentArrowPath(ctx, x, y, width, height);
                break;
            case 'uturnArrow':
                this.drawUturnArrowPath(ctx, x, y, width, height);
                break;
            case 'curvedRightArrow':
                this.drawCurvedRightArrowPath(ctx, x, y, width, height);
                break;
            case 'curvedLeftArrow':
                this.drawCurvedLeftArrowPath(ctx, x, y, width, height);
                break;
            // Callout shapes
            case 'callout1':
                this.drawCallout1Path(ctx, x, y, width, height);
                break;
            case 'callout2':
                this.drawCallout2Path(ctx, x, y, width, height);
                break;
            case 'callout3':
                this.drawCallout3Path(ctx, x, y, width, height);
                break;
            case 'accentCallout1':
                this.drawAccentCallout1Path(ctx, x, y, width, height);
                break;
            case 'accentCallout2':
                this.drawAccentCallout2Path(ctx, x, y, width, height);
                break;
            case 'accentCallout3':
                this.drawAccentCallout3Path(ctx, x, y, width, height);
                break;
            case 'borderCallout1':
                this.drawBorderCallout1Path(ctx, x, y, width, height);
                break;
            case 'borderCallout2':
                this.drawBorderCallout2Path(ctx, x, y, width, height);
                break;
            case 'borderCallout3':
                this.drawBorderCallout3Path(ctx, x, y, width, height);
                break;
            case 'accentBorderCallout1':
                this.drawAccentCallout1Path(ctx, x, y, width, height);
                break;
            case 'accentBorderCallout2':
                this.drawAccentCallout2Path(ctx, x, y, width, height);
                break;
            case 'accentBorderCallout3':
                this.drawAccentCallout3Path(ctx, x, y, width, height);
                break;
            // Special shapes
            case 'lightningBolt':
                this.drawLightningBoltPath(ctx, x, y, width, height);
                break;
            case 'sun':
                this.drawSunPath(ctx, x, y, width, height);
                break;
            case 'moon':
                this.drawMoonPath(ctx, x, y, width, height);
                break;
            case 'cloud':
                this.drawCloudPath(ctx, x, y, width, height);
                break;
            case 'wave':
                this.drawWavePath(ctx, x, y, width, height);
                break;
            case 'doubleWave':
                this.drawDoubleWavePath(ctx, x, y, width, height);
                break;
            case 'smileyFace':
                this.drawSmileyFacePath(ctx, x, y, width, height);
                break;
            case 'noSmoking':
                this.drawNoSmokingPath(ctx, x, y, width, height);
                break;
            case 'blockArc':
                this.drawBlockArcPath(ctx, x, y, width, height);
                break;
            case 'foldedCorner':
                this.drawFoldedCornerPath(ctx, x, y, width, height);
                break;
            case 'bevel':
                this.drawBevelPath(ctx, x, y, width, height);
                break;
            case 'donut':
                this.drawDonutPath(ctx, x, y, width, height);
                break;
            case 'noSymbol':
                this.drawNoSymbolPath(ctx, x, y, width, height);
                break;
            // 3D shapes
            case 'cube':
                this.drawCubePath(ctx, x, y, width, height);
                break;
            case 'can':
                this.drawCanPath(ctx, x, y, width, height);
                break;
            case 'cone':
                this.drawConePath(ctx, x, y, width, height);
                break;
            case 'pyramid':
                this.drawPyramidPath(ctx, x, y, width, height);
                break;
            // Mathematical symbols
            case 'plus':
                this.drawPlusPath(ctx, x, y, width, height);
                break;
            case 'minus':
                this.drawMinusPath(ctx, x, y, width, height);
                break;
            case 'multiply':
                this.drawMultiplyPath(ctx, x, y, width, height);
                break;
            case 'divide':
                this.drawDividePath(ctx, x, y, width, height);
                break;
            case 'equal':
                this.drawEqualPath(ctx, x, y, width, height);
                break;
            case 'notEqual':
                this.drawNotEqualPath(ctx, x, y, width, height);
                break;
            // Line shapes
            case 'line':
                this.drawLinePath(ctx, x, y, width, height);
                break;
            case 'straightConnector1':
                this.drawLinePath(ctx, x, y, width, height);
                break;
            case 'bentConnector2':
                this.drawBentConnector2Path(ctx, x, y, width, height);
                break;
            case 'bentConnector3':
                this.drawBentConnector3Path(ctx, x, y, width, height);
                break;
            case 'bentConnector4':
                this.drawBentConnector4Path(ctx, x, y, width, height);
                break;
            case 'bentConnector5':
                this.drawBentConnector5Path(ctx, x, y, width, height);
                break;
            case 'curvedConnector2':
                this.drawCurvedConnector2Path(ctx, x, y, width, height);
                break;
            case 'curvedConnector3':
                this.drawCurvedConnector3Path(ctx, x, y, width, height);
                break;
            case 'curvedConnector4':
                this.drawCurvedConnector4Path(ctx, x, y, width, height);
                break;
            case 'curvedConnector5':
                this.drawCurvedConnector5Path(ctx, x, y, width, height);
                break;
            case 'rtTriangle': {
                ctx.moveTo(x, y + height);
                ctx.lineTo(x + width, y + height);
                ctx.lineTo(x, y);
                ctx.closePath();
                break;
            }
            case 'arc': {
                // OOXML arc: closed pie/sector — adj1=stAng, adj2=enAng in 1/60000ths of degree (clockwise from east)
                const stAngDeg = (adjustments.adj1 !== undefined ? adjustments.adj1 : 0) / 60000;
                const enAngDeg = (adjustments.adj2 !== undefined ? adjustments.adj2 : 21600000) / 60000;
                const stAngRad = stAngDeg * Math.PI / 180;
                const enAngRad = enAngDeg * Math.PI / 180;
                const arcCx = x + width / 2, arcCy = y + height / 2;
                const arcRx = width / 2, arcRy = height / 2;
                // moveTo center, then ellipse auto-connects center→arcStart, then arc; closePath goes back to center
                ctx.moveTo(arcCx, arcCy);
                ctx.ellipse(arcCx, arcCy, arcRx, arcRy, 0, stAngRad, enAngRad);
                ctx.closePath();
                break;
            }
            case 'pie':
            case 'pieWedge': {
                const cx = x + width/2, cy = y + height/2;
                ctx.moveTo(cx, cy);
                ctx.arc(cx, cy, Math.min(width, height)/2, -Math.PI/2, 0);
                ctx.closePath();
                break;
            }
            case 'chord': {
                ctx.arc(x + width/2, y + height/2, Math.min(width, height)/2, 0, Math.PI * 1.5, true);
                ctx.closePath();
                break;
            }
            case 'frame': {
                const t = Math.min(width, height) * 0.15;
                ctx.rect(x, y, width, height);
                ctx.moveTo(x + t, y + t);
                ctx.rect(x + t, y + t, width - 2*t, height - 2*t);
                break;
            }
            case 'parallelogram': {
                const offset = width * 0.25;
                ctx.moveTo(x + offset, y);
                ctx.lineTo(x + width, y);
                ctx.lineTo(x + width - offset, y + height);
                ctx.lineTo(x, y + height);
                ctx.closePath();
                break;
            }
            case 'trapezoid': {
                const inset = width * 0.2;
                ctx.moveTo(x + inset, y);
                ctx.lineTo(x + width - inset, y);
                ctx.lineTo(x + width, y + height);
                ctx.lineTo(x, y + height);
                ctx.closePath();
                break;
            }
            case 'cross':
                this.drawPlusPath(ctx, x, y, width, height);
                break;
            case 'chevron': {
                const notch = width * 0.25;
                ctx.moveTo(x, y);
                ctx.lineTo(x + width - notch, y);
                ctx.lineTo(x + width, y + height/2);
                ctx.lineTo(x + width - notch, y + height);
                ctx.lineTo(x, y + height);
                ctx.lineTo(x + notch, y + height/2);
                ctx.closePath();
                break;
            }
            case 'homePlate': {
                const arrow = width * 0.2;
                ctx.moveTo(x, y);
                ctx.lineTo(x + width - arrow, y);
                ctx.lineTo(x + width, y + height/2);
                ctx.lineTo(x + width - arrow, y + height);
                ctx.lineTo(x, y + height);
                ctx.closePath();
                break;
            }
            case 'notchedRightArrow':
                this.drawRightArrowPath(ctx, x, y, width, height);
                break;
            case 'stripedRightArrow':
                this.drawRightArrowPath(ctx, x, y, width, height);
                break;
            case 'leftBracket': {
                const r = height * 0.15;
                ctx.moveTo(x + width * 0.4, y);
                ctx.quadraticCurveTo(x, y, x, y + r);
                ctx.lineTo(x, y + height - r);
                ctx.quadraticCurveTo(x, y + height, x + width * 0.4, y + height);
                break;
            }
            case 'rightBracket': {
                const r2 = height * 0.15;
                ctx.moveTo(x + width * 0.6, y);
                ctx.quadraticCurveTo(x + width, y, x + width, y + r2);
                ctx.lineTo(x + width, y + height - r2);
                ctx.quadraticCurveTo(x + width, y + height, x + width * 0.6, y + height);
                break;
            }
            case 'leftBrace': {
                const mid = y + height / 2;
                const bw = width * 0.4;
                ctx.moveTo(x + width, y);
                ctx.quadraticCurveTo(x + bw, y, x + bw, mid - height*0.1);
                ctx.quadraticCurveTo(x + bw, mid, x, mid);
                ctx.quadraticCurveTo(x + bw, mid, x + bw, mid + height*0.1);
                ctx.quadraticCurveTo(x + bw, y + height, x + width, y + height);
                break;
            }
            case 'rightBrace': {
                const mid2 = y + height / 2;
                const bw2 = width * 0.6;
                ctx.moveTo(x, y);
                ctx.quadraticCurveTo(x + bw2, y, x + bw2, mid2 - height*0.1);
                ctx.quadraticCurveTo(x + bw2, mid2, x + width, mid2);
                ctx.quadraticCurveTo(x + bw2, mid2, x + bw2, mid2 + height*0.1);
                ctx.quadraticCurveTo(x + bw2, y + height, x, y + height);
                break;
            }
            case 'bracketPair': {
                const br = Math.min(width, height) * 0.15;
                ctx.moveTo(x + br, y);
                ctx.arc(x + br, y + br, br, -Math.PI/2, Math.PI, true);
                ctx.lineTo(x, y + height - br);
                ctx.arc(x + br, y + height - br, br, Math.PI, Math.PI/2, true);
                ctx.moveTo(x + width - br, y);
                ctx.arc(x + width - br, y + br, br, -Math.PI/2, 0);
                ctx.lineTo(x + width, y + height - br);
                ctx.arc(x + width - br, y + height - br, br, 0, Math.PI/2);
                break;
            }
            case 'bracePair': {
                // Approximate as bracket pair
                const br2 = Math.min(width, height) * 0.15;
                ctx.moveTo(x + br2, y);
                ctx.arc(x + br2, y + br2, br2, -Math.PI/2, Math.PI, true);
                ctx.lineTo(x, y + height - br2);
                ctx.arc(x + br2, y + height - br2, br2, Math.PI, Math.PI/2, true);
                ctx.moveTo(x + width - br2, y);
                ctx.arc(x + width - br2, y + br2, br2, -Math.PI/2, 0);
                ctx.lineTo(x + width, y + height - br2);
                ctx.arc(x + width - br2, y + height - br2, br2, 0, Math.PI/2);
                break;
            }
            case 'diagStripe': {
                ctx.moveTo(x, y);
                ctx.lineTo(x + width, y);
                ctx.lineTo(x, y + height);
                ctx.closePath();
                break;
            }
            case 'corner': {
                const cw = width * 0.4, ch = height * 0.4;
                ctx.moveTo(x, y);
                ctx.lineTo(x + cw, y);
                ctx.lineTo(x + cw, y + height - ch);
                ctx.lineTo(x + width, y + height - ch);
                ctx.lineTo(x + width, y + height);
                ctx.lineTo(x, y + height);
                ctx.closePath();
                break;
            }
            case 'halfFrame': {
                const fw = width * 0.2, fh = height * 0.2;
                ctx.moveTo(x, y);
                ctx.lineTo(x + width, y);
                ctx.lineTo(x + width - fw, y + fh);
                ctx.lineTo(x + fw, y + fh);
                ctx.lineTo(x + fw, y + height);
                ctx.lineTo(x, y + height);
                ctx.closePath();
                break;
            }
            case 'snip1Rect':
            case 'snip2SameRect':
            case 'snip2DiagRect':
            case 'snipRoundRect': {
                const snip = Math.min(width, height) * 0.15;
                ctx.moveTo(x + snip, y);
                ctx.lineTo(x + width - snip, y);
                ctx.lineTo(x + width, y + snip);
                ctx.lineTo(x + width, y + height);
                ctx.lineTo(x, y + height);
                ctx.lineTo(x, y + snip);
                ctx.closePath();
                break;
            }
            case 'round1Rect':
            case 'round2SameRect':
            case 'round2DiagRect': {
                this.drawRoundRectPath(ctx, x, y, width, height, Math.min(width, height) * 0.15);
                break;
            }
            case 'ribbon':
            case 'ribbon2': {
                ctx.moveTo(x, y + height * 0.3);
                ctx.lineTo(x + width * 0.15, y + height * 0.15);
                ctx.lineTo(x + width * 0.15, y);
                ctx.lineTo(x + width * 0.85, y);
                ctx.lineTo(x + width * 0.85, y + height * 0.15);
                ctx.lineTo(x + width, y + height * 0.3);
                ctx.lineTo(x + width * 0.85, y + height * 0.45);
                ctx.lineTo(x + width * 0.85, y + height);
                ctx.lineTo(x + width * 0.15, y + height);
                ctx.lineTo(x + width * 0.15, y + height * 0.45);
                ctx.closePath();
                break;
            }
            case 'mathPlus':
                this.drawPlusPath(ctx, x, y, width, height);
                break;
            case 'mathMinus':
                this.drawMinusPath(ctx, x, y, width, height);
                break;
            case 'mathMultiply':
                this.drawMultiplyPath(ctx, x, y, width, height);
                break;
            case 'mathDivide':
                this.drawDividePath(ctx, x, y, width, height);
                break;
            case 'mathEqual':
                this.drawEqualPath(ctx, x, y, width, height);
                break;
            case 'mathNotEqual':
                this.drawNotEqualPath(ctx, x, y, width, height);
                break;
            case 'circularArrow':
                // Approximate as a curved arrow
                ctx.arc(x + width/2, y + height/2, Math.min(width, height)/2, 0, Math.PI * 1.5, true);
                break;
            default:
                // Default to rectangle for unknown shapes
                ctx.rect(x, y, width, height);
                break;
        }

        // Fill the shape
        if (fillColor) {
            ctx.fill();

        }

        // Stroke the shape - always at full opacity (not affected by fill transparency)
        if (strokeColor) {
            // Reset globalAlpha so stroke is fully opaque regardless of fill transparency
            ctx.globalAlpha = 1;
            // Apply dash pattern if available
            if (strokeInfo && strokeInfo.dashArray && strokeInfo.dashArray.length > 0) {
                const lineWidthPixels = CoordinateTransform.mmToPixels(lineWidth || 1);
                const scaledDashArray = strokeInfo.dashArray.map(dash => dash * lineWidthPixels);
                ctx.setLineDash(scaledDashArray);
            }
            ctx.stroke();

        }

        ctx.restore();

    }

    /**
     * Draw teardrop path - PowerPoint compatible implementation
     * Based on reference SVG path: M 711 7 C 475 8 442 8 423 12 C 312 31 221 77 144 154 C 53 247 4 363 4 492 C 4 565 16 625 46 689 C 82 769 143 842 219 894 C 258 921 307 945 359 961 C 418 980 509 987 577 977 C 720 956 845 878 922 758 C 953 710 977 654 988 603 C 1001 544 1001 546 1001 265 L 1001 7 L 989 6 C 983 6 857 6 711 7 Z
     */
    drawTeardropPath(ctx, x, y, width, height) {
        // Reference coordinates are based on 1001x992 viewBox
        // Scale factors to transform to actual dimensions
        const scaleX = width / 1001;
        const scaleY = height / 992;
        
        // Helper function to scale coordinates
        const sx = (coord) => x + (coord * scaleX);
        const sy = (coord) => y + (coord * scaleY);
        
        // Start from the reference path: M 711 7
        ctx.moveTo(sx(711), sy(7));
        
        // First curve: C 475 8 442 8 423 12
        ctx.bezierCurveTo(sx(475), sy(8), sx(442), sy(8), sx(423), sy(12));
        
        // Second curve: C 312 31 221 77 144 154
        ctx.bezierCurveTo(sx(312), sy(31), sx(221), sy(77), sx(144), sy(154));
        
        // Third curve: C 53 247 4 363 4 492
        ctx.bezierCurveTo(sx(53), sy(247), sx(4), sy(363), sx(4), sy(492));
        
        // Fourth curve: C 4 565 16 625 46 689
        ctx.bezierCurveTo(sx(4), sy(565), sx(16), sy(625), sx(46), sy(689));
        
        // Fifth curve: C 82 769 143 842 219 894
        ctx.bezierCurveTo(sx(82), sy(769), sx(143), sy(842), sx(219), sy(894));
        
        // Sixth curve: C 258 921 307 945 359 961
        ctx.bezierCurveTo(sx(258), sy(921), sx(307), sy(945), sx(359), sy(961));
        
        // Seventh curve: C 418 980 509 987 577 977
        ctx.bezierCurveTo(sx(418), sy(980), sx(509), sy(987), sx(577), sy(977));
        
        // Eighth curve: C 720 956 845 878 922 758
        ctx.bezierCurveTo(sx(720), sy(956), sx(845), sy(878), sx(922), sy(758));
        
        // Ninth curve: C 953 710 977 654 988 603
        ctx.bezierCurveTo(sx(953), sy(710), sx(977), sy(654), sx(988), sy(603));
        
        // Tenth curve: C 1001 544 1001 546 1001 265
        ctx.bezierCurveTo(sx(1001), sy(544), sx(1001), sy(546), sx(1001), sy(265));
        
        // Line to: L 1001 7
        ctx.lineTo(sx(1001), sy(7));
        
        // Line to: L 989 6
        ctx.lineTo(sx(989), sy(6));
        
        // Final curve: C 983 6 857 6 711 7 (back to start)
        ctx.bezierCurveTo(sx(983), sy(6), sx(857), sy(6), sx(711), sy(7));
        
        // Close path: Z
        ctx.closePath();
    }

    /**
     * Draw rounded rectangle path
     */
    drawRoundRectPath(ctx, x, y, width, height, radius) {
        ctx.moveTo(x + radius, y);
        ctx.lineTo(x + width - radius, y);
        ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
        ctx.lineTo(x + width, y + height - radius);
        ctx.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
        ctx.lineTo(x + radius, y + height);
        ctx.quadraticCurveTo(x, y + height, x, y + height - radius);
        ctx.lineTo(x, y + radius);
        ctx.quadraticCurveTo(x, y, x + radius, y);
        ctx.closePath();
    }

    /**
     * Draw triangle path
     */
    drawTrianglePath(ctx, x, y, width, height) {
        ctx.moveTo(x + width / 2, y);
        ctx.lineTo(x + width, y + height);
        ctx.lineTo(x, y + height);
        ctx.closePath();
    }

    /**
     * Draw diamond path
     */
    drawDiamondPath(ctx, x, y, width, height) {
        ctx.moveTo(x + width / 2, y);
        ctx.lineTo(x + width, y + height / 2);
        ctx.lineTo(x + width / 2, y + height);
        ctx.lineTo(x, y + height / 2);
        ctx.closePath();
    }

    /**
     * Draw pentagon path
     */
    drawPentagonPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const radius = Math.min(width, height) / 2;

        for (let i = 0; i < 5; i++) {
            const angle = (i * 2 * Math.PI) / 5 - Math.PI / 2;
            const pointX = centerX + radius * Math.cos(angle);
            const pointY = centerY + radius * Math.sin(angle);

            if (i === 0) {
                ctx.moveTo(pointX, pointY);
            } else {
                ctx.lineTo(pointX, pointY);
            }
        }
        ctx.closePath();
    }

    /**
     * Draw hexagon path
     */
    drawHexagonPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const radius = Math.min(width, height) / 2;

        for (let i = 0; i < 6; i++) {
            const angle = (i * 2 * Math.PI) / 6;
            const pointX = centerX + radius * Math.cos(angle);
            const pointY = centerY + radius * Math.sin(angle);

            if (i === 0) {
                ctx.moveTo(pointX, pointY);
            } else {
                ctx.lineTo(pointX, pointY);
            }
        }
        ctx.closePath();
    }

    /**
     * Draw star path
     */
    drawStarPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const outerRadius = Math.min(width, height) / 2;
        const innerRadius = outerRadius * 0.4;

        for (let i = 0; i < 10; i++) {
            const angle = (i * Math.PI) / 5 - Math.PI / 2;
            const radius = i % 2 === 0 ? outerRadius : innerRadius;
            const pointX = centerX + radius * Math.cos(angle);
            const pointY = centerY + radius * Math.sin(angle);

            if (i === 0) {
                ctx.moveTo(pointX, pointY);
            } else {
                ctx.lineTo(pointX, pointY);
            }
        }
        ctx.closePath();
    }

    /**
     * Draw heart path - PowerPoint-compatible implementation
     */
    drawHeartPath(ctx, x, y, width, height) {
        // PowerPoint-compatible heart - fills entire bounding box with no margins
        // Based on standard heart SVG but scaled to use full width/height
        
        // Heart path actual bounds: x: 2 to 22 (20 units), y: 3 to 21.35 (18.35 units)
        const heartWidth = 20;   // 22 - 2
        const heartHeight = 18.35; // 21.35 - 3
        const heartMinX = 2;
        const heartMinY = 3;
        
        const scaleX = width / heartWidth;
        const scaleY = height / heartHeight;
        
        // Transform function to map heart coordinates to canvas coordinates
        const transformX = (heartX) => x + (heartX - heartMinX) * scaleX;
        const transformY = (heartY) => y + (heartY - heartMinY) * scaleY;
        
        // Start at bottom center point (12, 21.35)
        ctx.moveTo(transformX(12), transformY(21.35));
        
        // Left curve to top-left (reducing by 1.45, 1.32 for the curve)
        ctx.lineTo(transformX(10.55), transformY(20.03));
        
        // Left side curve C5.4,15.36,2,12.28,2,8.5
        ctx.bezierCurveTo(
            transformX(5.4), transformY(15.36),
            transformX(2), transformY(12.28),
            transformX(2), transformY(8.5)
        );
        
        // Left lobe top C2,5.42,4.42,3,7.5,3
        ctx.bezierCurveTo(
            transformX(2), transformY(5.42),
            transformX(4.42), transformY(3),
            transformX(7.5), transformY(3)
        );
        
        // Left lobe to center valley
        ctx.bezierCurveTo(
            transformX(9.24), transformY(3),     // 7.5 + 1.74
            transformX(10.91), transformY(3.81), // 7.5 + 3.41
            transformX(12), transformY(5.09)     // center at 4.5 + 2.09
        );
        
        // Center valley to right lobe
        ctx.bezierCurveTo(
            transformX(13.09), transformY(3.81),
            transformX(14.76), transformY(3),
            transformX(16.5), transformY(3)
        );
        
        // Right lobe top C19.58,3,22,5.42,22,8.5
        ctx.bezierCurveTo(
            transformX(19.58), transformY(3),
            transformX(22), transformY(5.42),
            transformX(22), transformY(8.5)
        );
        
        // Right side curve
        ctx.bezierCurveTo(
            transformX(22), transformY(12.28),
            transformX(18.6), transformY(15.36),
            transformX(13.45), transformY(20.04)
        );
        
        // Right curve back to bottom
        ctx.lineTo(transformX(12), transformY(21.35));
        
        ctx.closePath();
    }

    /**
     * Draw octagon path
     */
    drawOctagonPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const radius = Math.min(width, height) / 2;

        for (let i = 0; i < 8; i++) {
            const angle = (i * 2 * Math.PI) / 8;
            const pointX = centerX + radius * Math.cos(angle);
            const pointY = centerY + radius * Math.sin(angle);

            if (i === 0) {
                ctx.moveTo(pointX, pointY);
            } else {
                ctx.lineTo(pointX, pointY);
            }
        }
        ctx.closePath();
    }

    /**
     * Draw star paths with different point counts
     */
    drawStar4Path(ctx, x, y, width, height) {
        this.drawStarGeneric(ctx, x, y, width, height, 4);
    }

    drawStar6Path(ctx, x, y, width, height) {
        this.drawStarGeneric(ctx, x, y, width, height, 6);
    }

    drawStar8Path(ctx, x, y, width, height) {
        this.drawStarGeneric(ctx, x, y, width, height, 8);
    }

    drawStar10Path(ctx, x, y, width, height) {
        this.drawStarGeneric(ctx, x, y, width, height, 10);
    }

    drawStar12Path(ctx, x, y, width, height) {
        this.drawStarGeneric(ctx, x, y, width, height, 12);
    }

    drawStar16Path(ctx, x, y, width, height) {
        this.drawStarGeneric(ctx, x, y, width, height, 16);
    }

    drawStar24Path(ctx, x, y, width, height) {
        this.drawStarGeneric(ctx, x, y, width, height, 24);
    }

    drawStar32Path(ctx, x, y, width, height) {
        this.drawStarGeneric(ctx, x, y, width, height, 32);
    }

    /**
     * Generic star drawing function
     */
    drawStarGeneric(ctx, x, y, width, height, points) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const outerRadius = Math.min(width, height) / 2;
        const innerRadius = outerRadius * 0.4;

        for (let i = 0; i < points * 2; i++) {
            const angle = (i * Math.PI) / points - Math.PI / 2;
            const radius = i % 2 === 0 ? outerRadius : innerRadius;
            const pointX = centerX + radius * Math.cos(angle);
            const pointY = centerY + radius * Math.sin(angle);

            if (i === 0) {
                ctx.moveTo(pointX, pointY);
            } else {
                ctx.lineTo(pointX, pointY);
            }
        }
        ctx.closePath();
    }

    /**
     * Flowchart shapes
     */
    drawFlowChartProcessPath(ctx, x, y, width, height) {
        ctx.rect(x, y, width, height);
    }

    drawFlowChartDecisionPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;

        ctx.moveTo(centerX, y);
        ctx.lineTo(x + width, centerY);
        ctx.lineTo(centerX, y + height);
        ctx.lineTo(x, centerY);
        ctx.closePath();
    }

    drawFlowChartInputOutputPath(ctx, x, y, width, height) {
        const skew = width * 0.2;

        ctx.moveTo(x + skew, y);
        ctx.lineTo(x + width, y);
        ctx.lineTo(x + width - skew, y + height);
        ctx.lineTo(x, y + height);
        ctx.closePath();
    }

    drawFlowChartPredefinedProcessPath(ctx, x, y, width, height) {
        const margin = width * 0.12;

        // Main rectangle
        ctx.rect(x, y, width, height);

        // Left line
        ctx.moveTo(x + margin, y);
        ctx.lineTo(x + margin, y + height);

        // Right line
        ctx.moveTo(x + width - margin, y);
        ctx.lineTo(x + width - margin, y + height);
    }

    drawFlowChartInternalStoragePath(ctx, x, y, width, height) {
        const margin = Math.min(width, height) * 0.2;

        // Main rectangle
        ctx.rect(x, y, width, height);

        // Top line
        ctx.moveTo(x, y + margin);
        ctx.lineTo(x + width, y + margin);

        // Left line
        ctx.moveTo(x + margin, y);
        ctx.lineTo(x + margin, y + height);
    }

    drawFlowChartDocumentPath(ctx, x, y, width, height) {
        const waveHeight = height * 0.2;
        const y1 = y + height - waveHeight;

        ctx.moveTo(x, y);
        ctx.lineTo(x + width, y);
        ctx.lineTo(x + width, y1);

        // Wavy bottom
        ctx.bezierCurveTo(x + width * 0.5, y + height + waveHeight, x + width * 0.5, y + height - waveHeight, x, y1);
        ctx.closePath();
    }

    drawFlowChartMultidocumentPath(ctx, x, y, width, height) {
        const offset = Math.min(width, height) * 0.1;

        // Back documents
        ctx.rect(x + offset * 2, y, width - offset * 2, height - offset * 2);
        ctx.rect(x + offset, y + offset, width - offset, height - offset);

        // Front document
        this.drawFlowChartDocumentPath(ctx, x, y + offset * 2, width - offset * 2, height - offset * 2);
    }

    drawFlowChartPreparationPath(ctx, x, y, width, height) {
        const skew = width * 0.2;

        ctx.moveTo(x, y + height / 2);
        ctx.lineTo(x + skew, y);
        ctx.lineTo(x + width - skew, y);
        ctx.lineTo(x + width, y + height / 2);
        ctx.lineTo(x + width - skew, y + height);
        ctx.lineTo(x + skew, y + height);
        ctx.closePath();
    }

    drawFlowChartManualInputPath(ctx, x, y, width, height) {
        const slant = height * 0.2;

        ctx.moveTo(x, y + slant);
        ctx.lineTo(x + width, y);
        ctx.lineTo(x + width, y + height);
        ctx.lineTo(x, y + height);
        ctx.closePath();
    }
    drawFlowChartManualOperationPath(ctx, x, y, width, height) {
        const indent = width * 0.2;

        ctx.moveTo(x, y);
        ctx.lineTo(x + width, y);
        ctx.lineTo(x + width - indent, y + height);
        ctx.lineTo(x + indent, y + height);
        ctx.closePath();
    }
    drawFlowChartOffpageConnectorPath(ctx, x, y, width, height) {
        const pointHeight = height * 0.2;

        ctx.moveTo(x, y);
        ctx.lineTo(x + width, y);
        ctx.lineTo(x + width, y + height - pointHeight);
        ctx.lineTo(x + width / 2, y + height);
        ctx.lineTo(x, y + height - pointHeight);
        ctx.closePath();
    }

    drawFlowChartPunchedCardPath(ctx, x, y, width, height) {
        const corner = Math.min(width, height) * 0.1;

        ctx.moveTo(x + corner, y);
        ctx.lineTo(x + width, y);
        ctx.lineTo(x + width, y + height);
        ctx.lineTo(x, y + height);
        ctx.lineTo(x, y + corner);
        ctx.closePath();
    }

    drawFlowChartPunchedTapePath(ctx, x, y, width, height) {
        const waveHeight = height * 0.1;

        // Top wavy line
        ctx.moveTo(x, y);
        ctx.bezierCurveTo(x + width * 0.25, y + waveHeight, x + width * 0.75, y - waveHeight, x + width, y);
        ctx.lineTo(x + width, y + height);

        // Bottom wavy line
        ctx.bezierCurveTo(x + width * 0.75, y + height - waveHeight, x + width * 0.25, y + height + waveHeight, x, y + height);
        ctx.closePath();
    }

    drawFlowChartSummingJunctionPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const radius = Math.min(width, height) / 2;

        // Circle
        ctx.arc(centerX, centerY, radius, 0, 2 * Math.PI);

        // Plus sign
        const lineLength = radius * 0.6;
        ctx.moveTo(centerX - lineLength, centerY);
        ctx.lineTo(centerX + lineLength, centerY);
        ctx.moveTo(centerX, centerY - lineLength);
        ctx.lineTo(centerX, centerY + lineLength);
    }

    drawFlowChartOrPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const radius = Math.min(width, height) / 2;

        // Circle
        ctx.arc(centerX, centerY, radius, 0, 2 * Math.PI);

        // OR lines (simplified)
        const lineLength = radius * 0.6;
        ctx.moveTo(centerX - lineLength, centerY - lineLength);
        ctx.lineTo(centerX + lineLength, centerY + lineLength);
        ctx.moveTo(centerX + lineLength, centerY - lineLength);
        ctx.lineTo(centerX - lineLength, centerY + lineLength);
    }

    drawFlowChartCollatePath(ctx, x, y, width, height) {
        ctx.moveTo(x, y);
        ctx.lineTo(x + width, y + height);
        ctx.lineTo(x, y + height);
        ctx.lineTo(x + width, y);
        ctx.closePath();
    }

    drawFlowChartSortPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;

        ctx.moveTo(x, y);
        ctx.lineTo(x + width, y);
        ctx.lineTo(centerX, y + height);
        ctx.closePath();

        // Horizontal line
        ctx.moveTo(x + width * 0.25, y + height * 0.5);
        ctx.lineTo(x + width * 0.75, y + height * 0.5);
    }

    drawFlowChartExtractPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;

        ctx.moveTo(x, y + height);
        ctx.lineTo(centerX, y);
        ctx.lineTo(x + width, y + height);
        ctx.closePath();
    }

    drawFlowChartMergePath(ctx, x, y, width, height) {
        const centerX = x + width / 2;

        ctx.moveTo(x, y);
        ctx.lineTo(x + width, y);
        ctx.lineTo(centerX, y + height);
        ctx.closePath();
    }

    drawFlowChartStoredDataPath(ctx, x, y, width, height) {
        const arcWidth = width * 0.1;

        // Left arc
        ctx.arc(x + arcWidth, y + height / 2, height / 2, Math.PI / 2, -Math.PI / 2);

        // Top line
        ctx.lineTo(x + width - arcWidth, y);

        // Right arc
        ctx.arc(x + width - arcWidth, y + height / 2, height / 2, -Math.PI / 2, Math.PI / 2);

        // Bottom line
        ctx.lineTo(x + arcWidth, y + height);
    }

    drawFlowChartDelayPath(ctx, x, y, width, height) {
        const arcWidth = width * 0.2;

        ctx.moveTo(x, y);
        ctx.lineTo(x + width - arcWidth, y);
        ctx.arc(x + width - arcWidth, y + height / 2, height / 2, -Math.PI / 2, Math.PI / 2);
        ctx.lineTo(x, y + height);
        ctx.closePath();
    }

    drawFlowChartMagneticTapePath(ctx, x, y, width, height) {
        const waveHeight = height * 0.15;

        // Top wavy line
        ctx.moveTo(x, y + waveHeight);
        ctx.bezierCurveTo(x + width * 0.33, y, x + width * 0.67, y + waveHeight * 2, x + width, y + waveHeight);
        ctx.lineTo(x + width, y + height - waveHeight);

        // Bottom wavy line
        ctx.bezierCurveTo(x + width * 0.67, y + height - waveHeight * 2, x + width * 0.33, y + height, x, y + height - waveHeight);
        ctx.closePath();
    }

    drawFlowChartMagneticDiskPath(ctx, x, y, width, height) {
        const diskHeight = height * 0.2;

        // Top ellipse
        ctx.ellipse(x + width / 2, y + diskHeight / 2, width / 2, diskHeight / 2, 0, 0, 2 * Math.PI);

        // Side lines
        ctx.moveTo(x, y + diskHeight / 2);
        ctx.lineTo(x, y + height - diskHeight / 2);
        ctx.moveTo(x + width, y + diskHeight / 2);
        ctx.lineTo(x + width, y + height - diskHeight / 2);

        // Bottom ellipse
        ctx.ellipse(x + width / 2, y + height - diskHeight / 2, width / 2, diskHeight / 2, 0, 0, 2 * Math.PI);
    }

    drawFlowChartMagneticDrumPath(ctx, x, y, width, height) {
        const drumWidth = width * 0.2;

        // Left ellipse
        ctx.ellipse(x + drumWidth / 2, y + height / 2, drumWidth / 2, height / 2, 0, 0, 2 * Math.PI);

        // Top line
        ctx.moveTo(x + drumWidth / 2, y);
        ctx.lineTo(x + width, y);

        // Bottom line
        ctx.moveTo(x + drumWidth / 2, y + height);
        ctx.lineTo(x + width, y + height);

        // Right line
        ctx.moveTo(x + width, y);
        ctx.lineTo(x + width, y + height);
    }

    drawFlowChartDisplayPath(ctx, x, y, width, height) {
        const arcWidth = width * 0.2;

        ctx.moveTo(x, y + height / 2);
        ctx.lineTo(x + arcWidth, y);
        ctx.lineTo(x + width - arcWidth, y);
        ctx.arc(x + width - arcWidth, y + height / 2, height / 2, -Math.PI / 2, Math.PI / 2);
        ctx.lineTo(x + arcWidth, y + height);
        ctx.closePath();
    }

    /**
     * Arrow shapes
     */
    drawRightArrowPath(ctx, x, y, width, height) {
        const arrowWidth = height * 0.6;
        const arrowY = y + (height - arrowWidth) / 2;
        const pointX = x + width * 0.7;

        ctx.moveTo(x, arrowY);
        ctx.lineTo(pointX, arrowY);
        ctx.lineTo(pointX, y);
        ctx.lineTo(x + width, y + height / 2);
        ctx.lineTo(pointX, y + height);
        ctx.lineTo(pointX, arrowY + arrowWidth);
        ctx.lineTo(x, arrowY + arrowWidth);
        ctx.closePath();
    }

    drawLeftArrowPath(ctx, x, y, width, height) {
        const arrowWidth = height * 0.6;
        const arrowY = y + (height - arrowWidth) / 2;
        const pointX = x + width * 0.3;

        ctx.moveTo(x + width, arrowY);
        ctx.lineTo(pointX, arrowY);
        ctx.lineTo(pointX, y);
        ctx.lineTo(x, y + height / 2);
        ctx.lineTo(pointX, y + height);
        ctx.lineTo(pointX, arrowY + arrowWidth);
        ctx.lineTo(x + width, arrowY + arrowWidth);
        ctx.closePath();
    }

    drawUpArrowPath(ctx, x, y, width, height) {
        const arrowWidth = width * 0.6;
        const arrowX = x + (width - arrowWidth) / 2;
        const pointY = y + height * 0.3;

        ctx.moveTo(arrowX, y + height);
        ctx.lineTo(arrowX, pointY);
        ctx.lineTo(x, pointY);
        ctx.lineTo(x + width / 2, y);
        ctx.lineTo(x + width, pointY);
        ctx.lineTo(arrowX + arrowWidth, pointY);
        ctx.lineTo(arrowX + arrowWidth, y + height);
        ctx.closePath();
    }

    drawDownArrowPath(ctx, x, y, width, height) {
        const arrowWidth = width * 0.6;
        const arrowX = x + (width - arrowWidth) / 2;
        const pointY = y + height * 0.7;

        ctx.moveTo(arrowX, y);
        ctx.lineTo(arrowX, pointY);
        ctx.lineTo(x, pointY);
        ctx.lineTo(x + width / 2, y + height);
        ctx.lineTo(x + width, pointY);
        ctx.lineTo(arrowX + arrowWidth, pointY);
        ctx.lineTo(arrowX + arrowWidth, y);
        ctx.closePath();
    }

    drawLeftRightArrowPath(ctx, x, y, width, height) {
        const arrowWidth = height * 0.6;
        const arrowY = y + (height - arrowWidth) / 2;
        const leftPoint = x + width * 0.2;
        const rightPoint = x + width * 0.8;

        ctx.moveTo(leftPoint, arrowY);
        ctx.lineTo(rightPoint, arrowY);
        ctx.lineTo(rightPoint, y);
        ctx.lineTo(x + width, y + height / 2);
        ctx.lineTo(rightPoint, y + height);
        ctx.lineTo(rightPoint, arrowY + arrowWidth);
        ctx.lineTo(leftPoint, arrowY + arrowWidth);
        ctx.lineTo(leftPoint, y + height);
        ctx.lineTo(x, y + height / 2);
        ctx.lineTo(leftPoint, y);
        ctx.closePath();
    }

    drawUpDownArrowPath(ctx, x, y, width, height) {
        const arrowWidth = width * 0.6;
        const arrowX = x + (width - arrowWidth) / 2;
        const topPoint = y + height * 0.2;
        const bottomPoint = y + height * 0.8;

        ctx.moveTo(arrowX, topPoint);
        ctx.lineTo(arrowX, bottomPoint);
        ctx.lineTo(x, bottomPoint);
        ctx.lineTo(x + width / 2, y + height);
        ctx.lineTo(x + width, bottomPoint);
        ctx.lineTo(arrowX + arrowWidth, bottomPoint);
        ctx.lineTo(arrowX + arrowWidth, topPoint);
        ctx.lineTo(x + width, topPoint);
        ctx.lineTo(x + width / 2, y);
        ctx.lineTo(x, topPoint);
        ctx.closePath();
    }

    drawQuadArrowPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const arrowSize = Math.min(width, height) * 0.3;

        // Top arrow
        ctx.moveTo(centerX, y);
        ctx.lineTo(centerX + arrowSize / 2, y + arrowSize);
        ctx.lineTo(centerX + arrowSize / 4, y + arrowSize);
        ctx.lineTo(centerX + arrowSize / 4, centerY - arrowSize / 4);

        // Right arrow
        ctx.lineTo(centerX + arrowSize / 4, centerY - arrowSize / 4);
        ctx.lineTo(x + width - arrowSize, centerY - arrowSize / 4);
        ctx.lineTo(x + width - arrowSize, centerY - arrowSize / 2);
        ctx.lineTo(x + width, centerY);
        ctx.lineTo(x + width - arrowSize, centerY + arrowSize / 2);
        ctx.lineTo(x + width - arrowSize, centerY + arrowSize / 4);

        // Bottom arrow
        ctx.lineTo(centerX + arrowSize / 4, centerY + arrowSize / 4);
        ctx.lineTo(centerX + arrowSize / 4, y + height - arrowSize);
        ctx.lineTo(centerX + arrowSize / 2, y + height - arrowSize);
        ctx.lineTo(centerX, y + height);
        ctx.lineTo(centerX - arrowSize / 2, y + height - arrowSize);
        ctx.lineTo(centerX - arrowSize / 4, y + height - arrowSize);

        // Left arrow
        ctx.lineTo(centerX - arrowSize / 4, centerY + arrowSize / 4);
        ctx.lineTo(x + arrowSize, centerY + arrowSize / 4);
        ctx.lineTo(x + arrowSize, centerY + arrowSize / 2);
        ctx.lineTo(x, centerY);
        ctx.lineTo(x + arrowSize, centerY - arrowSize / 2);
        ctx.lineTo(x + arrowSize, centerY - arrowSize / 4);
        ctx.lineTo(centerX - arrowSize / 4, centerY - arrowSize / 4);
        ctx.lineTo(centerX - arrowSize / 4, y + arrowSize);
        ctx.lineTo(centerX - arrowSize / 2, y + arrowSize);
        ctx.closePath();
    }

    drawBentArrowPath(ctx, x, y, width, height) {
        const cornerX = x + width * 0.7;
        const cornerY = y + height * 0.3;
        const arrowWidth = height * 0.4;

        ctx.moveTo(x, y + height / 2 - arrowWidth / 2);
        ctx.lineTo(cornerX, y + height / 2 - arrowWidth / 2);
        ctx.lineTo(cornerX, cornerY);
        ctx.lineTo(cornerX - arrowWidth / 2, cornerY);
        ctx.lineTo(x + width, y);
        ctx.lineTo(cornerX + arrowWidth / 2, cornerY);
        ctx.lineTo(cornerX, cornerY);
        ctx.lineTo(cornerX, y + height / 2 + arrowWidth / 2);
        ctx.lineTo(x, y + height / 2 + arrowWidth / 2);
        ctx.closePath();
    }

    drawUturnArrowPath(ctx, x, y, width, height) {
        const arrowWidth = width * 0.3;
        const centerX = x + width / 2;
        const radius = width * 0.3;

        // Shaft
        ctx.moveTo(x, y + height);
        ctx.lineTo(x, y + height / 2);
        ctx.arc(centerX, y + height / 2, radius, Math.PI, 0);
        ctx.lineTo(x + width - arrowWidth, y + height / 2 - radius);
        ctx.lineTo(x + width - arrowWidth, y);
        ctx.lineTo(x + width, y + height / 4);
        ctx.lineTo(x + width - arrowWidth * 2, y + height / 4);
        ctx.lineTo(x + width - arrowWidth * 2, y + height / 2 - radius);
        ctx.arc(centerX, y + height / 2, radius - arrowWidth, 0, Math.PI, true);
        ctx.lineTo(x + arrowWidth, y + height);
        ctx.closePath();
    }

    drawCurvedRightArrowPath(ctx, x, y, width, height) {
        const controlX = x + width * 0.8;
        const controlY = y + height * 0.2;
        const arrowSize = Math.min(width, height) * 0.2;

        ctx.moveTo(x, y + height / 2);
        ctx.quadraticCurveTo(controlX, controlY, x + width - arrowSize, y + height / 2);
        ctx.lineTo(x + width - arrowSize, y + height / 2 - arrowSize);
        ctx.lineTo(x + width, y + height / 2);
        ctx.lineTo(x + width - arrowSize, y + height / 2 + arrowSize);
        ctx.lineTo(x + width - arrowSize, y + height / 2 + arrowSize / 2);
        ctx.quadraticCurveTo(controlX, controlY + arrowSize, x, y + height / 2 + arrowSize / 2);
        ctx.closePath();
    }

    drawCurvedLeftArrowPath(ctx, x, y, width, height) {
        const controlX = x + width * 0.2;
        const controlY = y + height * 0.2;
        const arrowSize = Math.min(width, height) * 0.2;

        ctx.moveTo(x + width, y + height / 2);
        ctx.quadraticCurveTo(controlX, controlY, x + arrowSize, y + height / 2);
        ctx.lineTo(x + arrowSize, y + height / 2 - arrowSize);
        ctx.lineTo(x, y + height / 2);
        ctx.lineTo(x + arrowSize, y + height / 2 + arrowSize);
        ctx.lineTo(x + arrowSize, y + height / 2 + arrowSize / 2);
        ctx.quadraticCurveTo(controlX, controlY + arrowSize, x + width, y + height / 2 + arrowSize / 2);
        ctx.closePath();
    }

    /**
     * Callout shapes - simplified implementations
     */
    drawCallout1Path(ctx, x, y, width, height) {
        const calloutX = x + width * 0.8;
        const calloutY = y + height * 0.8;
        const calloutSize = Math.min(width, height) * 0.2;

        ctx.rect(x, y, width * 0.7, height * 0.7);

        // Callout pointer
        ctx.moveTo(x + width * 0.6, y + height * 0.6);
        ctx.lineTo(calloutX, calloutY);
        ctx.lineTo(x + width * 0.5, y + height * 0.7);
    }

    drawCallout2Path(ctx, x, y, width, height) {
        this.drawCallout1Path(ctx, x, y, width, height);

        // Second pointer
        const calloutX2 = x + width * 0.9;
        const calloutY2 = y + height * 0.9;

        ctx.moveTo(x + width * 0.65, y + height * 0.65);
        ctx.lineTo(calloutX2, calloutY2);
        ctx.lineTo(x + width * 0.55, y + height * 0.75);
    }

    drawCallout3Path(ctx, x, y, width, height) {
        this.drawCallout2Path(ctx, x, y, width, height);

        // Third pointer
        const calloutX3 = x + width * 0.7;
        const calloutY3 = y + height * 0.95;

        ctx.moveTo(x + width * 0.7, y + height * 0.7);
        ctx.lineTo(calloutX3, calloutY3);
        ctx.lineTo(x + width * 0.6, y + height * 0.8);
    }

    drawAccentCallout1Path(ctx, x, y, width, height) {
        // Rectangle with accent border
        ctx.rect(x, y, width * 0.7, height * 0.7);
        ctx.rect(x - 2, y - 2, width * 0.7 + 4, height * 0.7 + 4);
        this.drawCallout1Path(ctx, x, y, width, height);
    }

    drawAccentCallout2Path(ctx, x, y, width, height) {
        ctx.rect(x, y, width * 0.7, height * 0.7);
        ctx.rect(x - 2, y - 2, width * 0.7 + 4, height * 0.7 + 4);
        this.drawCallout2Path(ctx, x, y, width, height);
    }

    drawAccentCallout3Path(ctx, x, y, width, height) {
        ctx.rect(x, y, width * 0.7, height * 0.7);
        ctx.rect(x - 2, y - 2, width * 0.7 + 4, height * 0.7 + 4);
        this.drawCallout3Path(ctx, x, y, width, height);
    }

    drawBorderCallout1Path(ctx, x, y, width, height) {
        // Just rectangle with border for simplicity
        ctx.rect(x, y, width * 0.7, height * 0.7);
        this.drawCallout1Path(ctx, x, y, width, height);
    }

    drawBorderCallout2Path(ctx, x, y, width, height) {
        ctx.rect(x, y, width * 0.7, height * 0.7);
        this.drawCallout2Path(ctx, x, y, width, height);
    }

    drawBorderCallout3Path(ctx, x, y, width, height) {
        ctx.rect(x, y, width * 0.7, height * 0.7);
        this.drawCallout3Path(ctx, x, y, width, height);
    }


    /**
     * Special shapes
     */
    drawLightningBoltPath(ctx, x, y, width, height) {
        ctx.moveTo(x + width * 0.3, y);
        ctx.lineTo(x + width * 0.7, y);
        ctx.lineTo(x + width * 0.4, y + height * 0.4);
        ctx.lineTo(x + width * 0.8, y + height * 0.4);
        ctx.lineTo(x + width * 0.2, y + height);
        ctx.lineTo(x + width * 0.5, y + height * 0.6);
        ctx.lineTo(x + width * 0.1, y + height * 0.6);
        ctx.closePath();
    }

    drawSunPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const innerRadius = Math.min(width, height) * 0.3;
        const outerRadius = Math.min(width, height) * 0.45;

        // Sun rays
        for (let i = 0; i < 16; i++) {
            const angle = (i * 2 * Math.PI) / 16;
            const x1 = centerX + innerRadius * Math.cos(angle);
            const y1 = centerY + innerRadius * Math.sin(angle);
            const x2 = centerX + outerRadius * Math.cos(angle);
            const y2 = centerY + outerRadius * Math.sin(angle);

            ctx.moveTo(x1, y1);
            ctx.lineTo(x2, y2);
        }

        // Center circle
        ctx.arc(centerX, centerY, innerRadius * 0.7, 0, 2 * Math.PI);
    }

    drawMoonPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const radius = Math.min(width, height) / 2;

        ctx.arc(centerX, centerY, radius, 0, 2 * Math.PI);
        ctx.arc(centerX + radius * 0.3, centerY - radius * 0.3, radius * 0.8, 0, 2 * Math.PI, true);
    }

    drawCloudPath(ctx, x, y, width, height) {
        // PowerPoint-compatible cloud - fills entire bounding box with no margins
        // Multiple connected bulges forming a continuous cloud outline
        
        const w = width;
        const h = height;
        
        // Start from leftmost point
        ctx.moveTo(x, y + h * 0.5);
        
        // Left side bulge (bottom part)
        ctx.bezierCurveTo(
            x, y + h * 0.7,
            x + w * 0.05, y + h * 0.85,
            x + w * 0.15, y + h * 0.8
        );
        
        // Bottom left to center transition
        ctx.bezierCurveTo(
            x + w * 0.25, y + h * 0.95,
            x + w * 0.35, y + h * 0.95,
            x + w * 0.45, y + h * 0.85
        );
        
        // Bottom center bulge
        ctx.bezierCurveTo(
            x + w * 0.55, y + h,
            x + w * 0.65, y + h,
            x + w * 0.75, y + h * 0.85
        );
        
        // Bottom right bulge
        ctx.bezierCurveTo(
            x + w * 0.85, y + h * 0.95,
            x + w * 0.95, y + h * 0.8,
            x + w, y + h * 0.65
        );
        
        // Right side bulge
        ctx.bezierCurveTo(
            x + w, y + h * 0.5,
            x + w * 0.95, y + h * 0.35,
            x + w * 0.9, y + h * 0.3
        );
        
        // Top right bulge
        ctx.bezierCurveTo(
            x + w * 0.95, y + h * 0.15,
            x + w * 0.85, y + h * 0.05,
            x + w * 0.75, y + h * 0.1
        );
        
        // Top center-right bulge
        ctx.bezierCurveTo(
            x + w * 0.65, y,
            x + w * 0.55, y,
            x + w * 0.5, y + h * 0.08
        );
        
        // Top center-left bulge (largest)
        ctx.bezierCurveTo(
            x + w * 0.45, y,
            x + w * 0.3, y,
            x + w * 0.25, y + h * 0.15
        );
        
        // Top left bulge
        ctx.bezierCurveTo(
            x + w * 0.15, y + h * 0.05,
            x + w * 0.05, y + h * 0.2,
            x + w * 0.08, y + h * 0.3
        );
        
        // Left side bulge (top part) - back to start
        ctx.bezierCurveTo(
            x, y + h * 0.35,
            x, y + h * 0.4,
            x, y + h * 0.5
        );
        
        ctx.closePath();
    }

    drawWavePath(ctx, x, y, width, height) {
        const waveHeight = height * 0.3;
        const centerY = y + height / 2;

        ctx.moveTo(x, centerY);
        ctx.bezierCurveTo(x + width * 0.25, centerY - waveHeight, x + width * 0.75, centerY + waveHeight, x + width, centerY);
    }

    drawDoubleWavePath(ctx, x, y, width, height) {
        const waveHeight = height * 0.2;
        const y1 = y + height * 0.3;
        const y2 = y + height * 0.7;

        ctx.moveTo(x, y1);
        ctx.bezierCurveTo(x + width * 0.25, y1 - waveHeight, x + width * 0.75, y1 + waveHeight, x + width, y1);

        ctx.moveTo(x, y2);
        ctx.bezierCurveTo(x + width * 0.25, y2 - waveHeight, x + width * 0.75, y2 + waveHeight, x + width, y2);
    }

    drawSmileyFacePath(ctx, x, y, width, height) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const radius = Math.min(width, height) / 2;

        // Face
        ctx.arc(centerX, centerY, radius, 0, 2 * Math.PI);

        // Eyes
        ctx.arc(centerX - radius * 0.3, centerY - radius * 0.3, radius * 0.1, 0, 2 * Math.PI);
        ctx.arc(centerX + radius * 0.3, centerY - radius * 0.3, radius * 0.1, 0, 2 * Math.PI);

        // Smile
        ctx.arc(centerX, centerY, radius * 0.5, 0, Math.PI);
    }

    drawNoSmokingPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const radius = Math.min(width, height) / 2;

        // Circle
        ctx.arc(centerX, centerY, radius, 0, 2 * Math.PI);

        // Diagonal line
        ctx.moveTo(centerX - radius * 0.7, centerY - radius * 0.7);
        ctx.lineTo(centerX + radius * 0.7, centerY + radius * 0.7);
    }

    drawBlockArcPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const outerRadius = Math.min(width, height) / 2;
        const innerRadius = outerRadius * 0.6;

        ctx.arc(centerX, centerY, outerRadius, 0, Math.PI);
        ctx.arc(centerX, centerY, innerRadius, Math.PI, 0, true);
        ctx.closePath();
    }

    drawFoldedCornerPath(ctx, x, y, width, height) {
        const foldSize = Math.min(width, height) * 0.2;

        ctx.moveTo(x, y);
        ctx.lineTo(x + width - foldSize, y);
        ctx.lineTo(x + width, y + foldSize);
        ctx.lineTo(x + width, y + height);
        ctx.lineTo(x, y + height);
        ctx.closePath();

        // Fold line
        ctx.moveTo(x + width - foldSize, y);
        ctx.lineTo(x + width - foldSize, y + foldSize);
        ctx.lineTo(x + width, y + foldSize);
    }

    drawBevelPath(ctx, x, y, width, height) {
        const bevelSize = Math.min(width, height) * 0.1;

        ctx.moveTo(x + bevelSize, y);
        ctx.lineTo(x + width - bevelSize, y);
        ctx.lineTo(x + width, y + bevelSize);
        ctx.lineTo(x + width, y + height - bevelSize);
        ctx.lineTo(x + width - bevelSize, y + height);
        ctx.lineTo(x + bevelSize, y + height);
        ctx.lineTo(x, y + height - bevelSize);
        ctx.lineTo(x, y + bevelSize);
        ctx.closePath();
    }

    drawDonutPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const outerRadius = Math.min(width, height) / 2;
        const innerRadius = outerRadius * 0.4;

        ctx.arc(centerX, centerY, outerRadius, 0, 2 * Math.PI);
        ctx.arc(centerX, centerY, innerRadius, 0, 2 * Math.PI, true);
    }

    drawNoSymbolPath(ctx, x, y, width, height) {
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const radius = Math.min(width, height) / 2;

        // Circle
        ctx.arc(centerX, centerY, radius, 0, 2 * Math.PI);

        // Diagonal line
        ctx.moveTo(centerX - radius * 0.7, centerY + radius * 0.7);
        ctx.lineTo(centerX + radius * 0.7, centerY - radius * 0.7);
    }

    /**
     * 3D shapes
     */
    drawCubePath(ctx, x, y, width, height) {
        const depth = Math.min(width, height) * 0.3;

        // Front face
        ctx.rect(x, y + depth, width - depth, height - depth);

        // Top face
        ctx.moveTo(x, y + depth);
        ctx.lineTo(x + depth, y);
        ctx.lineTo(x + width, y);
        ctx.lineTo(x + width - depth, y + depth);
        ctx.closePath();

        // Right face
        ctx.moveTo(x + width - depth, y + depth);
        ctx.lineTo(x + width, y);
        ctx.lineTo(x + width, y + height - depth);
        ctx.lineTo(x + width - depth, y + height);
        ctx.closePath();
    }

    drawCanPath(ctx, x, y, width, height) {
        const ellipseHeight = height * 0.15;

        // Top ellipse
        ctx.ellipse(x + width / 2, y + ellipseHeight / 2, width / 2, ellipseHeight / 2, 0, 0, 2 * Math.PI);

        // Side lines
        ctx.moveTo(x, y + ellipseHeight / 2);
        ctx.lineTo(x, y + height - ellipseHeight / 2);
        ctx.moveTo(x + width, y + ellipseHeight / 2);
        ctx.lineTo(x + width, y + height - ellipseHeight / 2);

        // Bottom ellipse
        ctx.ellipse(x + width / 2, y + height - ellipseHeight / 2, width / 2, ellipseHeight / 2, 0, 0, 2 * Math.PI);
    }

    drawConePath(ctx, x, y, width, height) {
        const baseHeight = height * 0.2;

        // Cone sides
        ctx.moveTo(x + width / 2, y);
        ctx.lineTo(x, y + height - baseHeight);
        ctx.ellipse(x + width / 2, y + height - baseHeight / 2, width / 2, baseHeight / 2, 0, Math.PI, 0);
        ctx.lineTo(x + width / 2, y);
        ctx.closePath();
    }

    drawPyramidPath(ctx, x, y, width, height) {
        const baseDepth = width * 0.3;

        // Front face
        ctx.moveTo(x + width / 2, y);
        ctx.lineTo(x, y + height);
        ctx.lineTo(x + width - baseDepth, y + height);
        ctx.closePath();

        // Right face
        ctx.moveTo(x + width / 2, y);
        ctx.lineTo(x + width - baseDepth, y + height);
        ctx.lineTo(x + width, y + height - baseDepth);
        ctx.closePath();

        // Base edge
        ctx.moveTo(x, y + height);
        ctx.lineTo(x + baseDepth, y + height - baseDepth);
        ctx.lineTo(x + width, y + height - baseDepth);
        ctx.lineTo(x + width - baseDepth, y + height);
    }
    /**
     * Mathematical symbols
     */
    drawPlusPath(ctx, x, y, width, height) {
        const lineWidth = Math.min(width, height) * 0.2;
        const centerX = x + width / 2;
        const centerY = y + height / 2;

        // Horizontal line
        ctx.rect(x + lineWidth, centerY - lineWidth / 2, width - lineWidth * 2, lineWidth);

        // Vertical line
        ctx.rect(centerX - lineWidth / 2, y + lineWidth, lineWidth, height - lineWidth * 2);
    }
    drawMinusPath(ctx, x, y, width, height) {
        const lineWidth = Math.min(width, height) * 0.2;
        const centerY = y + height / 2;

        ctx.rect(x + lineWidth, centerY - lineWidth / 2, width - lineWidth * 2, lineWidth);
    }

    drawMultiplyPath(ctx, x, y, width, height) {
        const lineWidth = Math.min(width, height) * 0.1;

        // Diagonal line 1
        ctx.moveTo(x + lineWidth, y + lineWidth);
        ctx.lineTo(x + width - lineWidth, y + height - lineWidth);

        // Diagonal line 2
        ctx.moveTo(x + width - lineWidth, y + lineWidth);
        ctx.lineTo(x + lineWidth, y + height - lineWidth);
    }

    drawDividePath(ctx, x, y, width, height) {
        const lineWidth = Math.min(width, height) * 0.1;
        const centerX = x + width / 2;
        const centerY = y + height / 2;
        const dotRadius = lineWidth;

        // Top dot
        ctx.arc(centerX, centerY - height * 0.25, dotRadius, 0, 2 * Math.PI);

        // Line
        ctx.rect(x + lineWidth, centerY - lineWidth / 2, width - lineWidth * 2, lineWidth);

        // Bottom dot
        ctx.arc(centerX, centerY + height * 0.25, dotRadius, 0, 2 * Math.PI);
    }

    drawEqualPath(ctx, x, y, width, height) {
        const lineWidth = Math.min(width, height) * 0.1;
        const centerY = y + height / 2;
        const spacing = height * 0.15;

        // Top line
        ctx.rect(x + lineWidth, centerY - spacing - lineWidth / 2, width - lineWidth * 2, lineWidth);

        // Bottom line
        ctx.rect(x + lineWidth, centerY + spacing - lineWidth / 2, width - lineWidth * 2, lineWidth);
    }

    drawNotEqualPath(ctx, x, y, width, height) {
        this.drawEqualPath(ctx, x, y, width, height);

        // Diagonal line
        const lineWidth = Math.min(width, height) * 0.1;
        ctx.moveTo(x + width * 0.3, y + height * 0.8);
        ctx.lineTo(x + width * 0.7, y + height * 0.2);
    }

    /**
     * Line and connector shapes
     */
    drawLinePath(ctx, x, y, width, height) {
        // Draw diagonal line from top-left to bottom-right of bounds
        // For horizontal lines (height=0), this draws left to right
        // For vertical lines (width=0), this draws top to bottom
        ctx.moveTo(x, y);
        ctx.lineTo(x + width, y + height);
    }

    drawBentConnector2Path(ctx, x, y, width, height) {
        const midX = x + width / 2;

        ctx.moveTo(x, y + height);
        ctx.lineTo(x, y + height / 2);
        ctx.lineTo(midX, y + height / 2);
        ctx.lineTo(midX, y);
        ctx.lineTo(x + width, y);
    }

    drawBentConnector3Path(ctx, x, y, width, height) {
        const midX = x + width / 2;
        const midY = y + height / 2;

        ctx.moveTo(x, y + height);
        ctx.lineTo(x, midY);
        ctx.lineTo(midX, midY);
        ctx.lineTo(midX, y);
        ctx.lineTo(x + width, y);
    }

    drawBentConnector4Path(ctx, x, y, width, height) {
        const midX1 = x + width * 0.33;
        const midX2 = x + width * 0.67;
        const midY = y + height / 2;

        ctx.moveTo(x, y + height);
        ctx.lineTo(x, midY);
        ctx.lineTo(midX1, midY);
        ctx.lineTo(midX1, y);
        ctx.lineTo(midX2, y);
        ctx.lineTo(midX2, midY);
        ctx.lineTo(x + width, midY);
        ctx.lineTo(x + width, y + height);
    }

    drawBentConnector5Path(ctx, x, y, width, height) {
        const midX1 = x + width * 0.25;
        const midX2 = x + width * 0.5;
        const midX3 = x + width * 0.75;
        const midY1 = y + height * 0.33;
        const midY2 = y + height * 0.67;

        ctx.moveTo(x, y + height);
        ctx.lineTo(x, midY2);
        ctx.lineTo(midX1, midY2);
        ctx.lineTo(midX1, midY1);
        ctx.lineTo(midX2, midY1);
        ctx.lineTo(midX2, y);
        ctx.lineTo(midX3, y);
        ctx.lineTo(midX3, midY1);
        ctx.lineTo(x + width, midY1);
        ctx.lineTo(x + width, y + height);
    }

    drawCurvedConnector2Path(ctx, x, y, width, height) {
        ctx.moveTo(x, y + height);
        ctx.quadraticCurveTo(x + width / 2, y, x + width, y);
    }

    drawCurvedConnector3Path(ctx, x, y, width, height) {
        ctx.moveTo(x, y + height);
        ctx.bezierCurveTo(x, y + height / 2, x + width / 2, y + height / 2, x + width, y);
    }

    drawCurvedConnector4Path(ctx, x, y, width, height) {
        ctx.moveTo(x, y + height);
        ctx.bezierCurveTo(x + width * 0.25, y + height * 0.75, x + width * 0.75, y + height * 0.25, x + width, y);
    }

    drawCurvedConnector5Path(ctx, x, y, width, height) {
        ctx.moveTo(x, y + height);
        ctx.bezierCurveTo(x + width * 0.2, y + height * 0.8, x + width * 0.4, y + height * 0.6, x + width * 0.5, y + height * 0.5);
        ctx.bezierCurveTo(x + width * 0.6, y + height * 0.4, x + width * 0.8, y + height * 0.2, x + width, y);
    }


    /**
     * Draw custom geometry using path system
     */
    drawCustomGeometry(geometry, x, y, width, height, fillColor, strokeColor, lineWidth) {
        
        try {
            // Handle both pathLst and pathList property names
            const pathList = geometry.pathLst || geometry.pathList;
            if (!geometry || !pathList || pathList.length === 0) {
                return;
            }

            const ctx = this.context;
            ctx.save();

            // Calculate coordinate bounds from actual path data
            let coordWidth, coordHeight;
            
            if (geometry.coordSize?.width && geometry.coordSize?.height) {
                coordWidth = geometry.coordSize.width;
                coordHeight = geometry.coordSize.height;
            } else {
                // Calculate coordinate bounds from actual path data
                let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
                
                pathList.forEach(path => {
                    if (path.commands) {
                        path.commands.forEach(cmd => {
                            if (cmd.x !== undefined) {
                                minX = Math.min(minX, cmd.x);
                                maxX = Math.max(maxX, cmd.x);
                            }
                            if (cmd.y !== undefined) {
                                minY = Math.min(minY, cmd.y);
                                maxY = Math.max(maxY, cmd.y);
                            }
                            // Check control points too
                            if (cmd.x1 !== undefined) {
                                minX = Math.min(minX, cmd.x1);
                                maxX = Math.max(maxX, cmd.x1);
                                minY = Math.min(minY, cmd.y1);
                                maxY = Math.max(maxY, cmd.y1);
                            }
                            if (cmd.x2 !== undefined) {
                                minX = Math.min(minX, cmd.x2);
                                maxX = Math.max(maxX, cmd.x2);
                                minY = Math.min(minY, cmd.y2);
                                maxY = Math.max(maxY, cmd.y2);
                            }
                        });
                    }
                });
                
                coordWidth = maxX - minX;
                coordHeight = maxY - minY;
                
                // Handle degenerate cases (single points, lines)
                if (coordWidth === 0) {coordWidth = 100;}
                if (coordHeight === 0) {coordHeight = 100;}
            }

            ctx.restore();

            // Process each path in the path list
            
            pathList.forEach((path, index) => {
                
                // Calculate min coordinates for this specific path
                let pathMinX = Infinity, pathMinY = Infinity;
                path.commands.forEach(cmd => {
                    if (cmd.x !== undefined) { pathMinX = Math.min(pathMinX, cmd.x); pathMinY = Math.min(pathMinY, cmd.y); }
                    if (cmd.x1 !== undefined) { pathMinX = Math.min(pathMinX, cmd.x1); pathMinY = Math.min(pathMinY, cmd.y1); }
                    if (cmd.x2 !== undefined) { pathMinX = Math.min(pathMinX, cmd.x2); pathMinY = Math.min(pathMinY, cmd.y2); }
                    if (cmd.x3 !== undefined) { pathMinX = Math.min(pathMinX, cmd.x3); pathMinY = Math.min(pathMinY, cmd.y3); }
                });
                
                // Set current shape bounds for gradient calculation
                this.currentShapeBounds = { x, y, w: width, h: height };
                
                this.drawGeometryPathDirectly(path, fillColor, strokeColor, lineWidth, x, y, width, height, coordWidth, coordHeight, pathMinX, pathMinY);
            });
        } catch (error) {
            const ctx = this.context;
            ctx.restore(); // Ensure context is restored even on error
        }
    }

    /**
     * Draw a single geometry path
     */
    drawGeometryPath(path, fillColor, strokeColor, lineWidth) {
        if (!path || !path.commands || path.commands.length === 0) {
            return;
        }

        const ctx = this.context;
        ctx.beginPath();


        // Process path commands
        let hasDrawnPath = false;
        path.commands.forEach((command, index) => {
            switch (command.type) {
                case 'moveTo':
                    ctx.moveTo(command.x, command.y);
                    break;
                case 'lineTo':
                    ctx.lineTo(command.x, command.y);
                    hasDrawnPath = true;
                    break;
                case 'curveTo':
                case 'cubicBezTo':
                    const endX = command.x !== undefined ? command.x : command.x3;
                    const endY = command.y !== undefined ? command.y : command.y3;
                    
                    if (endX !== undefined && endY !== undefined) {
                        ctx.bezierCurveTo(command.x1, command.y1, command.x2, command.y2, endX, endY);
                        hasDrawnPath = true;
                    }
                    break;
                case 'quadTo':
                case 'quadBezTo':
                    ctx.quadraticCurveTo(command.x1, command.y1, command.x, command.y);
                    hasDrawnPath = true;
                    break;
                case 'arcTo':
                    // Simple arc implementation
                    ctx.arcTo(command.x1, command.y1, command.x, command.y, command.radius || 0);
                    hasDrawnPath = true;
                    break;
                case 'close':
                    ctx.closePath();
                    break;
                default:
            }
        });
        

        // Apply stroke first for better visual appearance (stroke-first rendering)
        if (strokeColor && lineWidth > 0) {
            this.setStrokeStyle(strokeColor, lineWidth);
            ctx.stroke();
        } else {
        }

        // Apply fill only if explicitly requested and no stroke, or if both are specified
        if (fillColor && (!strokeColor || lineWidth === 0)) {
            this.setFillStyle(fillColor);
            ctx.fill();
        } else {
        }
        
        ctx.restore();
    }

    /**
     * Set stroke style for graphics
     */
    setStrokeStyle(strokeColor, lineWidth) {
        const ctx = this.context;
        ctx.strokeStyle = strokeColor;
        ctx.lineWidth = CoordinateTransform.mmToPixels(lineWidth || 1);
        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';
    }

    /**
     * Set fill style for graphics  
     */
    setFillStyle(fillColor) {
        const ctx = this.context;
        ctx.fillStyle = fillColor;
    }

    /**
     * Draw geometry path directly in screen coordinates
     */
    drawGeometryPathDirectly(path, fillColor, strokeColor, lineWidth, screenX, screenY, screenWidth, screenHeight, coordWidth, coordHeight, pathMinX, pathMinY) {
        
        if (!path || !path.commands || path.commands.length === 0) {
            return;
        }

        const ctx = this.context;
        
        
        // Calculate coordinate transformation factors
        const scaleX = screenWidth / coordWidth;
        const scaleY = screenHeight / coordHeight;
        

        ctx.save();
        ctx.beginPath();
        
        // Helper function to transform coordinates
        const transformX = (x) => screenX + (x - pathMinX) * scaleX;
        const transformY = (y) => screenY + (y - pathMinY) * scaleY;

        // Process path commands with direct coordinate transformation
        let hasDrawnPath = false;
        let currentX = 0, currentY = 0; // Track current position
        
        path.commands.forEach((command, index) => {
            switch (command.type) {
                case 'moveTo':
                    const moveX = transformX(command.x);
                    const moveY = transformY(command.y);
                    ctx.moveTo(moveX, moveY);
                    currentX = moveX;
                    currentY = moveY;
                    break;
                case 'lineTo':
                    const lineX = transformX(command.x);
                    const lineY = transformY(command.y);
                    ctx.lineTo(lineX, lineY);
                    currentX = lineX;
                    currentY = lineY;
                    hasDrawnPath = true;
                    break;
                case 'curveTo':
                case 'cubicBezTo':
                    // PowerPoint uses x3,y3 for the end point of cubic bezier curves
                    const endX = command.x3 !== undefined ? command.x3 : command.x;
                    const endY = command.y3 !== undefined ? command.y3 : command.y;
                    
                    if (endX !== undefined && endY !== undefined) {
                        const cp1X = transformX(command.x1);
                        const cp1Y = transformY(command.y1);
                        const cp2X = transformX(command.x2);
                        const cp2Y = transformY(command.y2);
                        const bezEndX = transformX(endX);
                        const bezEndY = transformY(endY);
                        
                        // Check if this is actually a straight line disguised as a curve
                        if (this.isStraightLine(currentX, currentY, cp1X, cp1Y, cp2X, cp2Y, bezEndX, bezEndY)) {
                            // Draw as straight line to avoid unnecessary bending
                            ctx.lineTo(bezEndX, bezEndY);
                        } else {
                            // Use native cubic bezier for actual curves
                            ctx.bezierCurveTo(cp1X, cp1Y, cp2X, cp2Y, bezEndX, bezEndY);
                        }
                        
                        currentX = bezEndX;
                        currentY = bezEndY;
                        hasDrawnPath = true;
                    }
                    break;
                case 'quadTo':
                case 'quadBezTo': {
                    const qcpX = transformX(command.x1);
                    const qcpY = transformY(command.y1);
                    const qendX = transformX(command.x);
                    const qendY = transformY(command.y);
                    ctx.quadraticCurveTo(qcpX, qcpY, qendX, qendY);
                    currentX = qendX;
                    currentY = qendY;
                    hasDrawnPath = true;
                    break;
                }
                case 'close':
                    ctx.closePath();
                    break;
            }
        });


        // Check if this is a closed shape that should have fill (not a line path)
        const hasCloseCommand = path.commands.some(cmd => cmd.type === 'close');
        const shouldHaveFill = hasCloseCommand && fillColor && fillColor !== 'none';
        
        // Apply fill and stroke to the completed path

        // Apply fill for closed shapes 
        if (shouldHaveFill) {
            if (hasDrawnPath) {
                if (typeof fillColor === 'object' && fillColor.r !== undefined) {
                    ctx.fillStyle = `rgba(${fillColor.r}, ${fillColor.g}, ${fillColor.b}, ${fillColor.a / 255})`;
                } else if (typeof fillColor === 'object' && fillColor.type) {
                    // Handle gradient fills
                    if (fillColor.type === 'linear' && fillColor.stops) {
                        // Create proper linear gradient
                        const bounds = this.getCanvasBounds(ctx);
                        const gradient = this.createCanvasGradient(ctx, bounds, fillColor);
                        if (gradient) {
                            ctx.fillStyle = gradient;
                        } else {
                            // Use first stop color or return if no stops
                            const firstColor = fillColor.stops[0]?.color;
                            if (firstColor) {
                                ctx.fillStyle = this.colorToRgb(firstColor);
                            } else {
                                return; // No gradient stops - skip fill
                            }
                        }
                    } else {
                        return; // No gradient data - skip fill
                    }
                } else if (typeof fillColor === 'string' && fillColor !== 'none' && fillColor !== 'null') {
                    ctx.fillStyle = fillColor;
                }
                ctx.fill();
            } else {
                // For shapes with only moveTo+close, fill as rectangle using bounds
                if (typeof fillColor === 'object' && fillColor.r !== undefined) {
                    ctx.fillStyle = `rgba(${fillColor.r}, ${fillColor.g}, ${fillColor.b}, ${fillColor.a / 255})`;
                } else if (typeof fillColor === 'string' && fillColor !== 'none' && fillColor !== 'null') {
                    ctx.fillStyle = fillColor;
                }
                ctx.fillRect(screenX, screenY, screenWidth, screenHeight);
            }
        }

        // Apply stroke with correct line width from DOM
        if (strokeColor && strokeColor !== 'none' && hasDrawnPath) {
            // Use a simple approach: scale line width based on screen size ratio
            const baseLineWidthPixels = CoordinateTransform.mmToPixels(lineWidth > 0 ? lineWidth : 1);
            
            // Calculate scale based on screen dimensions relative to a standard size
            // Using 1280x720 as reference for full slide rendering
            const referenceArea = 1280 * 720;
            const currentArea = screenWidth * screenHeight;
            const areaScale = Math.sqrt(currentArea / referenceArea);
            
            const finalLineWidth = Math.max(baseLineWidthPixels * areaScale, 0.5); // Minimum 0.5px line width
            
            if (typeof strokeColor === 'object' && strokeColor.r !== undefined) {
                ctx.strokeStyle = `rgba(${strokeColor.r}, ${strokeColor.g}, ${strokeColor.b}, ${strokeColor.a / 255})`;
            } else {
                ctx.strokeStyle = strokeColor;
            }
            ctx.lineWidth = finalLineWidth;
            ctx.lineCap = 'square'; // Changed from 'round' to 'square' for sharp edges
            ctx.lineJoin = 'miter'; // Changed from 'round' to 'miter' for sharp corners
            ctx.stroke();
        } else {
        }

        ctx.restore();
    }

    /**
     * Check if a cubic bezier curve is actually a straight line
     */
    isStraightLine(startX, startY, cp1X, cp1Y, cp2X, cp2Y, endX, endY) {
        // Calculate the expected line between start and end points
        const deltaX = endX - startX;
        const deltaY = endY - startY;
        
        // If start and end are the same point, it's degenerate
        if (Math.abs(deltaX) < 0.1 && Math.abs(deltaY) < 0.1) {
            return true;
        }
        
        // Check if control points lie on or very close to the straight line
        // between start and end points
        const tolerance = 2; // pixels
        
        // Calculate expected positions of control points if this were a straight line
        const expectedCp1X = startX + deltaX * 0.33; // 1/3 along the line
        const expectedCp1Y = startY + deltaY * 0.33;
        const expectedCp2X = startX + deltaX * 0.67; // 2/3 along the line  
        const expectedCp2Y = startY + deltaY * 0.67;
        
        // Check if actual control points are close to expected straight line positions
        const cp1Distance = Math.sqrt(Math.pow(cp1X - expectedCp1X, 2) + Math.pow(cp1Y - expectedCp1Y, 2));
        const cp2Distance = Math.sqrt(Math.pow(cp2X - expectedCp2X, 2) + Math.pow(cp2Y - expectedCp2Y, 2));
        
        return cp1Distance < tolerance && cp2Distance < tolerance;
    }

    // ========================================
    // Graphics Methods
    // ========================================

    /**
     * Start path
     */
    _s() {
        this._context.beginPath();
    }

    /**
     * Move to point
     */
    _m(x, y) {
        this._context.moveTo(x, y);
    }

    /**
     * Line to point
     */
    _l(x, y) {
        this._context.lineTo(x, y);
    }

    /**
     * Cubic Bezier curve
     */
    _c(x1, y1, x2, y2, x, y) {
        this._context.bezierCurveTo(x1, y1, x2, y2, x, y);
    }

    /**
     * Close path
     */
    _z() {
        this._context.closePath();
    }

    /**
     * Draw stroke
     */
    ds() {
        this._context.stroke();
    }

    /**
     * Draw fill
     */
    df() {
        this._context.fill();
    }

    /**
     * Fill text with style properties (for DOM compatibility)
     * Expected interface: fillText(text, x, y, styleProperties)
     * @param {string} text - Text to render
     * @param {number} x - X coordinate
     * @param {number} y - Y coordinate
     * @param {Object} styleProperties - Style properties including fontSize, color, fontFamily, etc.
     */
    fillText(text, x, y, styleProperties = {}) {
        if (!this._context || !text) {
            return;
        }

        this._context.save();

        try {
            // Extract style properties with defaults and apply scaling
            const baseFontSize = styleProperties.fontSize || 12;
            const scaleFactor = this.getTextScaleFactor();
            const fontSize = baseFontSize * scaleFactor;
            const fontFamily = styleProperties.fontFamily || 'Arial';
            const fontWeight = styleProperties.fontWeight || (styleProperties.bold ? 'bold' : 'normal');
            const fontStyle = styleProperties.fontStyle || (styleProperties.italic ? 'italic' : 'normal');
            
            // Set font
            this._context.font = `${fontStyle} ${fontWeight} ${fontSize}px "${fontFamily}"`;
            
            // Set color
            if (styleProperties.color) {
                if (typeof styleProperties.color === 'object' && styleProperties.color.r !== undefined) {
                    // RGBA object format: {r: 255, g: 0, b: 0, a: 255} or {r: 255, g: 0, b: 0}
                    const alpha = styleProperties.color.a !== undefined ? styleProperties.color.a / 255 : 1;
                    const colorString = `rgba(${styleProperties.color.r}, ${styleProperties.color.g}, ${styleProperties.color.b}, ${alpha})`;
                    this._context.fillStyle = colorString;
                } else if (typeof styleProperties.color === 'string') {
                    // String format: "#FF0000" or "red"
                    this._context.fillStyle = styleProperties.color;
                } else {
                    this._context.fillStyle = '#000000';
                }
            } else {
                this._context.fillStyle = '#000000';
            }

            // Set text baseline and alignment - use consistent baseline
            this._context.textBaseline = styleProperties.textBaseline || 'alphabetic';
            this._context.textAlign = styleProperties.textAlign || 'left';

            // Render the text
            this._context.fillText(text, x, y);
            
        } catch (error) {
        } finally {
            this._context.restore();
        }
    }

    /**
     * Measure text with style properties (for DOM compatibility)
     * Expected interface: measureText(text, styleProperties)
     * @param {string} text - Text to measure
     * @param {Object} styleProperties - Style properties including fontSize, fontFamily, etc.
     * @returns {Object} Text metrics with width and height properties
     */
    measureText(text, styleProperties = {}) {
        if (!this._context || !text) {
            return { width: 0, height: 0 };
        }

        this._context.save();

        try {
            // Extract style properties with defaults and apply scaling
            const baseFontSize = styleProperties.fontSize || 12;
            const scaleFactor = this.getTextScaleFactor();
            const fontSize = baseFontSize * scaleFactor;
            const fontFamily = styleProperties.fontFamily || 'Arial';
            const fontWeight = styleProperties.fontWeight || (styleProperties.bold ? 'bold' : 'normal');
            const fontStyle = styleProperties.fontStyle || (styleProperties.italic ? 'italic' : 'normal');
            
            // Set font for measurement
            this._context.font = `${fontStyle} ${fontWeight} ${fontSize}px "${fontFamily}"`;
            
            // Measure the text
            const metrics = this._context.measureText(text);
            
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
            this._context.restore();
        }
    }

    /**
     * Get current canvas bounds for gradient calculation
     */
    getCanvasBounds(ctx) {
        const transform = ctx.getTransform();
        // Use the current shape bounds if available, otherwise use canvas dimensions
        // Use logical dimensions for bounds calculation
        let logicalWidth, logicalHeight;
        if (ctx.canvas.style.width && ctx.canvas.style.height) {
            logicalWidth = parseFloat(ctx.canvas.style.width);
            logicalHeight = parseFloat(ctx.canvas.style.height);
        } else {
            logicalWidth = ctx.canvas.width / (window.devicePixelRatio || 1);
            logicalHeight = ctx.canvas.height / (window.devicePixelRatio || 1);
        }
        return this.currentShapeBounds || { x: 0, y: 0, w: logicalWidth, h: logicalHeight };
    }

    /**
     * Create canvas gradient from gradient definition
     */
    createCanvasGradient(ctx, bounds, gradientDef) {
        if (!gradientDef || !gradientDef.stops || gradientDef.stops.length === 0) {
            return null;
        }

        let gradient;
        if (gradientDef.type === 'linear') {
            // Linear gradient
            const angle = gradientDef.angle || 0;
            const radians = (angle * Math.PI) / 180;

            const x1 = bounds.x + bounds.w / 2 - (Math.cos(radians) * bounds.w) / 2;
            const y1 = bounds.y + bounds.h / 2 - (Math.sin(radians) * bounds.h) / 2;
            const x2 = bounds.x + bounds.w / 2 + (Math.cos(radians) * bounds.w) / 2;
            const y2 = bounds.y + bounds.h / 2 + (Math.sin(radians) * bounds.h) / 2;

            gradient = ctx.createLinearGradient(x1, y1, x2, y2);
        } else if (gradientDef.type === 'radial') {
            // Radial gradient
            const centerX = bounds.x + bounds.w / 2;
            const centerY = bounds.y + bounds.h / 2;
            const radius = Math.max(bounds.w, bounds.h) / 2;

            gradient = ctx.createRadialGradient(centerX, centerY, 0, centerX, centerY, radius);
        } else {
            return null;
        }

        // Add color stops
        for (const stop of gradientDef.stops) {
            const colorString = this.colorToRgb(stop.color);
            gradient.addColorStop(stop.position, colorString);
        }

        return gradient;
    }

    /**
     * Convert color object/string to CSS rgb(a) string, delegating to graphics if possible
     */
    colorToRgb(color) {
        if (this.graphics && typeof this.graphics.colorToRgb === 'function') {
            return this.graphics.colorToRgb(color);
        }
        if (this.graphicsEngine && typeof this.graphicsEngine.colorToRgb === 'function') {
            return this.graphicsEngine.colorToRgb(color);
        }
        if (typeof color === 'string') {
            return color;
        }
        if (color && typeof color === 'object') {
            const r = color.r !== undefined ? color.r : (color.R !== undefined ? color.R : 0);
            const g = color.g !== undefined ? color.g : (color.G !== undefined ? color.G : 0);
            const b = color.b !== undefined ? color.b : (color.B !== undefined ? color.B : 0);
            const a = color.a !== undefined ? color.a : (color.A !== undefined ? color.A : 255);
            if (a !== undefined && a !== 255) {
                return `rgba(${r}, ${g}, ${b}, ${a / 255})`;
            }
            return `rgb(${r}, ${g}, ${b})`;
        }
        return 'rgb(0, 0, 0)';
    }

}
/**
 * Drawing Document Implementation
 */
class CDrawingDocument {
    constructor() {
        this.graphics = null;
        this.canvas = null;
        this.processor = null; // Add processor reference
        this.logger = new Logger('CDrawingDocument');
        // Track which layer we are rendering (slide | layout | master)
        this.currentLayer = null;
    }

    /**
     * Initialize with canvas and processor
     */
    init(canvas, processor = null) {
        this.canvas = canvas;
        this.processor = processor; // Store processor reference
        
        // Reset existing graphics state if reusing the same instance
        if (this.graphics) {
            this.graphics.resetState();
        } else {
            // Always create the adapter as primary graphics
            this.graphics = new CanvasGraphicsAdapter();
            // Create engine in parallel if available, and bridge arrow-capable methods
            try {
                const EngineCtor = (typeof globalThis !== 'undefined' && globalThis.CGraphics)
                    ? globalThis.CGraphics
                    : (typeof window !== 'undefined' && window.CGraphics ? window.CGraphics : null);
                if (typeof EngineCtor === 'function') {
                    this.graphicsEngine = new EngineCtor(canvas);
                    // Bridge engine arrow rendering through adapter surface
                    this.graphics.drawLineWithArrows = (x1, y1, x2, y2, strokeColor, lineWidth = 1, strokeInfo = null) => {
                        this.graphicsEngine.drawLineWithArrows(x1, y1, x2, y2, strokeColor, lineWidth, strokeInfo);
                    };
                    // Optional: allow applying stroke info via engine normalizer if needed
                    this.graphics.applyStrokeInfo = (strokeInfo) => {
                        if (typeof this.graphicsEngine.applyStrokeInfo === 'function') {
                            this.graphicsEngine.applyStrokeInfo(strokeInfo);
                        }
                    };
                }
            } catch (_e) {}
        }

        // Use logical display dimensions for coordinate system, not scaled canvas dimensions
        // This ensures proper positioning with high-DPI scaling
        let logicalWidth;
        let logicalHeight;

        const rect = typeof canvas.getBoundingClientRect === 'function'
            ? canvas.getBoundingClientRect()
            : { width: 0, height: 0 };

        const styleW = parseFloat(canvas.style.width);
        const styleH = parseFloat(canvas.style.height);
        // Explicit render resolution: when the host sets data-pvjs-render-width/height
        // (logical CSS px), use it verbatim and ignore the canvas's live CSS/box size.
        // This decouples render resolution from on-screen size, so the host can let
        // CSS control display (e.g. width:100%) while still rasterising at a chosen
        // resolution × devicePixelRatio. It also keeps a delayed re-render crisp,
        // because the attribute persists on the element regardless of CSS changes.
        const dataW = parseFloat(canvas.dataset && canvas.dataset.pvjsRenderWidth);
        const dataH = parseFloat(canvas.dataset && canvas.dataset.pvjsRenderHeight);
        if (dataW > 0 && dataH > 0) {
            logicalWidth = dataW;
            logicalHeight = dataH;
        } else if (rect.width > 0 && rect.height > 0) {
            // If style dimensions are explicitly set and LARGER than the bounding rect
            // (e.g. container constrained by page layout), prefer style dimensions.
            // This prevents canvas from being shrunk back by layout constraints during batch export.
            if (styleW > rect.width || styleH > rect.height) {
                logicalWidth = styleW || rect.width;
                logicalHeight = styleH || rect.height;
            } else {
                logicalWidth = rect.width;
                logicalHeight = rect.height;
            }
        } else if (canvas.style.width && canvas.style.height) {
            logicalWidth = styleW;
            logicalHeight = styleH;
        } else {
            const pixelRatio = typeof window !== 'undefined' && window.devicePixelRatio ? window.devicePixelRatio : 1;
            logicalWidth = canvas.width / pixelRatio;
            logicalHeight = canvas.height / pixelRatio;
        }

        
        // Initialize adapter surface (engine configures itself in its constructor)
        const pixelRatio = (this.processor && this.processor.renderContext && this.processor.renderContext.pixelRatio)
            || (typeof window !== 'undefined' && window.devicePixelRatio) || 1;

        const targetWidth = Math.max(1, Math.round(logicalWidth * pixelRatio));
        const targetHeight = Math.max(1, Math.round(logicalHeight * pixelRatio));

        if (canvas.width !== targetWidth || canvas.height !== targetHeight) {
            canvas.width = targetWidth;
            canvas.height = targetHeight;
        }

        // Only pin an inline display size when the host hasn't opted into the
        // explicit render-resolution mode (data-pvjs-render-*). In that mode the
        // host controls display via CSS (e.g. width:100%), so writing an inline
        // pixel size here would override it.
        if (!(dataW > 0 && dataH > 0)) {
            if (!canvas.style.width) {
                canvas.style.width = `${logicalWidth}px`;
            }
            if (!canvas.style.height) {
                canvas.style.height = `${logicalHeight}px`;
            }
        }

        const ctx = canvas.getContext('2d');
        if (ctx) {
            ctx.setTransform(1, 0, 0, 1, 0, 0);
            ctx.scale(pixelRatio, pixelRatio);
        }

        if (this.graphics && typeof this.graphics.init === 'function' && ctx) {
            this.graphics.init(
                ctx,
                logicalWidth,
                logicalHeight,
                logicalWidth * 25.4 / 96, // Convert px to mm for compatibility
                logicalHeight * 25.4 / 96
            );
        }

    }

    /**
     * Convert color object/string to CSS rgb(a) string, delegating to graphics if possible
     */
    colorToRgb(color) {
        if (this.graphics && typeof this.graphics.colorToRgb === 'function') {
            return this.graphics.colorToRgb(color);
        }
        if (this.graphicsEngine && typeof this.graphicsEngine.colorToRgb === 'function') {
            return this.graphicsEngine.colorToRgb(color);
        }
        if (typeof color === 'string') {
            return color;
        }
        if (color && typeof color === 'object') {
            const r = color.r !== undefined ? color.r : (color.R !== undefined ? color.R : 0);
            const g = color.g !== undefined ? color.g : (color.G !== undefined ? color.G : 0);
            const b = color.b !== undefined ? color.b : (color.B !== undefined ? color.B : 0);
            const a = color.a !== undefined ? color.a : (color.A !== undefined ? color.A : 255);
            if (a !== undefined && a !== 255) {
                return `rgba(${r}, ${g}, ${b}, ${a / 255})`;
            }
            return `rgb(${r}, ${g}, ${b})`;
        }
        return 'rgb(0, 0, 0)';
    }

    /**
     * Draw slide
     */
    async drawSlide(slide, slideIndex) {
        if (!this.graphics) {
            return;
        }

        // Store current slide for context resolution
        this.currentSlide = slide;
        this.currentSlideIndex = slideIndex;

        // Clear canvas
        this.graphics.clear();

        // Draw slide content
        await this.drawSlideContent(slide);
    }

    /**
     * Draw slide content
     * Rendering order: Background -> Master -> Layout -> Slide
     */
    async drawSlideContent(slide) {

        if (slide.layout) {
        }
        
        // Set up coordinate system for slide
        this.setupSlideCoordinateSystem();

        // Clip all rendering strictly to the slide rectangle to prevent spillover
        try {
            const ctx = this.graphics?.context;
            const coords = this.coordinateSystem;
            if (ctx && coords) {
                const clipX = Math.round(coords.offsetX);
                const clipY = Math.round(coords.offsetY);
                const clipW = Math.round(coords.scaledWidth);
                const clipH = Math.round(coords.scaledHeight);
                ctx.save();
                ctx.beginPath();
                ctx.rect(clipX, clipY, clipW, clipH);
                ctx.clip();
                this._slideClipApplied = true;
            }
        } catch (e) {
            this._slideClipApplied = false;
        }

        // Layer 1: Draw slide background first
        this.drawSlideBackground(slide);

        // Layer 2: Draw master elements if enabled
        if (slide.showMasterShapes !== false && slide.layout && slide.layout.master) {
            await this.drawMasterElements(slide.layout.master);
        } else {
            
        }

        // Layer 3: Draw layout elements
        if (slide.layout) {
            await this.drawLayoutElements(slide.layout);
        } else {
        }

        // Layer 4: Draw slide shapes (highest priority)
        if (slide.commonSlideData && slide.commonSlideData.shapeTree) {
            const spTree = slide.commonSlideData.shapeTree;

            // Sort shapes by rendering order (z-order) if available
            const sortedShapes = this.sortShapesByRenderOrder(spTree);

            // Render each shape in order
            const __prevLayer = this.currentLayer;
            this.currentLayer = 'slide';
            for (let i = 0; i < sortedShapes.length; i++) {
                const shape = sortedShapes[i];

                if (this.isHiddenShape(shape)) {
                    continue;
                }

                // Skip empty placeholder shapes on slide layer too - they should be invisible
                if (this.isPlaceholderShape(shape) && !this.placeholderHasContent(shape)) {
                    continue;
                }

                // Enhanced shape drawing - await for async chart rendering
                await this.drawShapeEnhanced(shape, i);
            }
            this.currentLayer = __prevLayer;
        }

        // Restore clipping region if applied
        try {
            if (this._slideClipApplied && this.graphics?.context) {
                this.graphics.context.restore();
            }
        } catch (e) {}
    }

    /**
     * Sort shapes by rendering order (z-order)
     * In PowerPoint, shapes are typically rendered in the order they appear in the XML,
     * but we should respect any explicit ordering properties
     */
    sortShapesByRenderOrder(shapes) {
        if (!shapes || shapes.length === 0) {
            return [];
        }

        // DEBUG: Log shape information to understand z-order properties

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

    /**
     * Draw slide background
     */
    drawSlideBackground(slide) {
        if (!this.graphics || !this.graphics.context) {return;}

        const ctx = this.graphics.context;

        // Check for slide background from DOM data
        let backgroundColor = null;

        // Check for slide background first
        if (slide.backgroundFill) {
            backgroundColor = this.getBackgroundColor(slide.backgroundFill);
        }
        // Check layout background
        else if (slide.layout && slide.layout.commonSlideData && slide.layout.commonSlideData.backgroundFill) {
            backgroundColor = this.getBackgroundColor(slide.layout.commonSlideData.backgroundFill);
        }
        // Check layout bg directly
        else if (slide.layout && slide.layout.cSld && slide.layout.cSld.bg) {
            backgroundColor = this.getBackgroundColor(slide.layout.cSld.bg);
        }
        // Check master background
        else if (slide.layout && slide.layout.master && slide.layout.master.commonSlideData && slide.layout.master.commonSlideData.backgroundFill) {
            backgroundColor = this.getBackgroundColor(slide.layout.master.commonSlideData.backgroundFill);
        }
        // Check for theme-based background colors
        else if (slide.layout && slide.layout.master && slide.layout.master.theme) {
            // Try to get background color from theme
            const theme = slide.layout.master.theme;
            if (theme.colors && theme.colors.bg1) {
                backgroundColor = theme.colors.bg1;
            } else if (theme.colors && theme.colors.bg2) {
                backgroundColor = theme.colors.bg2;
            }
        }

        // Fill the slide area with background color using SAME coordinate system as shapes
        if (this.coordinateSystem) {
            const { scale, offsetX, offsetY } = this.coordinateSystem;
            const slideWidthPx = this.coordinateSystem.slideWidthPx;
            const slideHeightPx = this.coordinateSystem.slideHeightPx;

            // Snap to integer pixels to avoid 1px gaps
            const actualSlideWidth = Math.round(slideWidthPx * scale);
            const actualSlideHeight = Math.round(slideHeightPx * scale);
            const snappedOffsetX = Math.round(offsetX);
            const snappedOffsetY = Math.round(offsetY);

            ctx.save();
            if (backgroundColor && backgroundColor.type === 'image' && backgroundColor.imageData?.relationshipId) {
                // Render image background - prefer resolvedCacheKey to avoid relId collisions
                const cacheKey = backgroundColor.imageData.resolvedCacheKey || backgroundColor.imageData.relationshipId;
                const imgData = this.processor?.imageCache?.get(cacheKey);
                if (imgData && imgData.image) {
                    ctx.drawImage(imgData.image, snappedOffsetX, snappedOffsetY, actualSlideWidth, actualSlideHeight);
                    ctx.restore();
                    return;
                }
                // Image not loaded yet - fall through to white background
                ctx.fillStyle = '#ffffff';
            } else if (backgroundColor && backgroundColor.type === 'gradient' && backgroundColor.gradient) {
                // Render gradient background
                const grad = backgroundColor.gradient;
                const stops = grad.stops || [];
                let canvasGrad;
                if (grad.type === 'radial') {
                    const cx = snappedOffsetX + actualSlideWidth / 2;
                    const cy = snappedOffsetY + actualSlideHeight / 2;
                    const radius = Math.max(actualSlideWidth, actualSlideHeight) / 2;
                    canvasGrad = ctx.createRadialGradient(cx, cy, 0, cx, cy, radius);
                } else {
                    // Linear gradient - use angle if available
                    const angle = (grad.angle || 0) * Math.PI / 180;
                    const cos = Math.cos(angle), sin = Math.sin(angle);
                    const x0 = snappedOffsetX + actualSlideWidth / 2 - cos * actualSlideWidth / 2;
                    const y0 = snappedOffsetY + actualSlideHeight / 2 - sin * actualSlideHeight / 2;
                    const x1 = snappedOffsetX + actualSlideWidth / 2 + cos * actualSlideWidth / 2;
                    const y1 = snappedOffsetY + actualSlideHeight / 2 + sin * actualSlideHeight / 2;
                    canvasGrad = ctx.createLinearGradient(x0, y0, x1, y1);
                }
                for (const stop of stops) {
                    const pos = (stop.position !== undefined) ? stop.position / 100000 : 0;
                    const color = this.parseColorToHex(stop.color) || '#ffffff';
                    canvasGrad.addColorStop(Math.min(1, Math.max(0, pos)), color);
                }
                ctx.fillStyle = canvasGrad;
            } else {
                ctx.fillStyle = (typeof backgroundColor === 'string' ? backgroundColor : null) || '#ffffff';
            }
            ctx.fillRect(snappedOffsetX, snappedOffsetY, actualSlideWidth, actualSlideHeight);
            ctx.restore();
        } else {
            // Fill background if color is specified in DOM
            ctx.save();
            if (backgroundColor && backgroundColor.type === 'image' && backgroundColor.imageData?.relationshipId) {
                // Render image background - prefer resolvedCacheKey to avoid relId collisions
                const cacheKey = backgroundColor.imageData.resolvedCacheKey || backgroundColor.imageData.relationshipId;
                const imgData = this.processor?.imageCache?.get(cacheKey);
                if (imgData && imgData.image) {
                    ctx.drawImage(imgData.image, 0, 0, this.canvas.width, this.canvas.height);
                    ctx.restore();
                    return;
                }
                ctx.fillStyle = '#ffffff';
            } else if (backgroundColor && backgroundColor.type === 'gradient' && backgroundColor.gradient) {
                const grad = backgroundColor.gradient;
                const stops = grad.stops || [];
                const w = this.canvas.width, h = this.canvas.height;
                let canvasGrad;
                if (grad.type === 'radial') {
                    canvasGrad = ctx.createRadialGradient(w/2, h/2, 0, w/2, h/2, Math.max(w, h)/2);
                } else {
                    const angle = (grad.angle || 0) * Math.PI / 180;
                    const cos = Math.cos(angle), sin = Math.sin(angle);
                    canvasGrad = ctx.createLinearGradient(w/2 - cos*w/2, h/2 - sin*h/2, w/2 + cos*w/2, h/2 + sin*h/2);
                }
                for (const stop of stops) {
                    const pos = (stop.position !== undefined) ? stop.position / 100000 : 0;
                    const color = this.parseColorToHex(stop.color) || '#ffffff';
                    canvasGrad.addColorStop(Math.min(1, Math.max(0, pos)), color);
                }
                ctx.fillStyle = canvasGrad;
            } else {
                ctx.fillStyle = (typeof backgroundColor === 'string' ? backgroundColor : null) || '#ffffff';
            }
            ctx.fillRect(0, 0, this.canvas.width, this.canvas.height);
            ctx.restore();
        }
    }

    /**
     * Draw master elements
     */
    async drawMasterElements(master) {
        if (!master || !master.commonSlideData || !master.commonSlideData.shapeTree) {
            return;
        }

        const masterShapes = master.commonSlideData.shapeTree;

        // Sort master shapes by rendering order
        const sortedMasterShapes = this.sortShapesByRenderOrder(masterShapes);
        const previousLayer = this.currentLayer;
        this.currentLayer = 'master';
        try {
            for (let i = 0; i < sortedMasterShapes.length; i++) {
                const shape = sortedMasterShapes[i];

                if (this.isHiddenShape(shape)) {
                    continue;
                }

                // Skip empty placeholder shapes in master layer - they should be invisible
                if (this.isPlaceholderShape(shape) && !this.placeholderHasContent(shape)) {
                    continue;
                }

                await this.drawShapeEnhanced(shape, i);
            }
        } finally {
            this.currentLayer = previousLayer;
        }
    }

    /**
     * Draw layout elements
     */
    async drawLayoutElements(layout) {
        
        
        if (!layout) {
            return;
        }
        
        if (!layout.commonSlideData) {
            return;
        }
        
        if (!layout.commonSlideData.shapeTree) {
            return;
        }

        const layoutShapes = layout.commonSlideData.shapeTree;

        // Print detailed information about each shape BEFORE rendering
        layoutShapes.forEach((shape, i) => {
        });

        // Sort layout shapes by rendering order
        const sortedLayoutShapes = this.sortShapesByRenderOrder(layoutShapes);

        const previousLayer = this.currentLayer;
        this.currentLayer = 'layout';
        try {
            for (let i = 0; i < sortedLayoutShapes.length; i++) {
                const shape = sortedLayoutShapes[i];

                if (this.isHiddenShape(shape)) {
                    continue;
                }

                // Skip empty placeholder shapes in layout layer - they should be invisible
                if (this.isPlaceholderShape(shape) && !this.placeholderHasContent(shape)) {
                    continue;
                }

                await this.drawShapeEnhanced(shape, i);
            }
        } finally {
            this.currentLayer = previousLayer;
        }
        
    }

    /**
     * Check if a shape is a placeholder
     */
    isPlaceholderShape(shape) {
        // Check for placeholder properties
        if (shape.placeholder || shape.phType || shape.isPlaceholder) {
            return true;
        }

        // Check for placeholder in shape properties
        if (shape.properties && shape.properties.placeholder) {
            return true;
        }

        // Check for placeholder in nvSpPr (non-visual shape properties)
        if (shape.nvSpPr && shape.nvSpPr.nvPr && shape.nvSpPr.nvPr.ph) {
            return true;
        }

        return false;
    }

    /**
     * Check if a placeholder shape has actual content that should be rendered.
     * Placeholders with fills, images, charts, or text should render.
     * Empty placeholders should not render their bounding boxes.
     */
    placeholderHasContent(shape) {
        // Has explicit fill (background color)
        if (shape.fill && shape.fill.type !== 'none' && shape.fill.type !== 'noFill') {
            return true;
        }
        // Has image
        if (shape.imageRelId || shape.type === 'pic') {
            return true;
        }
        // Has chart data
        if (shape.chartData) {
            return true;
        }
        // Has real text content (not just empty paragraphs) or field runs (e.g. slidenum)
        if (shape.textBody && shape.textBody.paragraphs) {
            const hasText = shape.textBody.paragraphs.some(p =>
                p.runs && p.runs.some(r => (r.text || '').trim().length > 0 || r.fieldType)
            );
            if (hasText) return true;
        }
        // Has table data
        if (shape.graphicData && shape.graphicData.tableXml) {
            return true;
        }
        return false;
    }

    /**
     * Process placeholder shapes and apply styling from master/layout
     */
    processPlaceholderShape(shape, slide) {
        if (!shape.isPlaceholder || !shape.placeholder) {
            return shape;
        }

        const placeholder = shape.placeholder;

        // Find matching placeholder in layout or master
        let masterPlaceholder = null;
        let layoutPlaceholder = null;

        // Look for matching placeholder in layout
        if (slide.layout && slide.layout.commonSlideData && slide.layout.commonSlideData.shapeTree) {
            layoutPlaceholder = this.findMatchingPlaceholder(slide.layout.commonSlideData.shapeTree, placeholder);
        }

        // Look for matching placeholder in master
        if (slide.layout && slide.layout.master && slide.layout.master.commonSlideData && slide.layout.master.commonSlideData.shapeTree) {
            masterPlaceholder = this.findMatchingPlaceholder(slide.layout.master.commonSlideData.shapeTree, placeholder);
        }

        // Apply styling from layout or master placeholder
        if (layoutPlaceholder) {
            this.applyPlaceholderStyling(shape, layoutPlaceholder);
        } else if (masterPlaceholder) {
            this.applyPlaceholderStyling(shape, masterPlaceholder);
        }

        return shape;
    }

    /**
     * Find matching placeholder in shape tree
     */
    findMatchingPlaceholder(shapeTree, targetPlaceholder) {
        for (const shape of shapeTree) {
            if (shape.isPlaceholder && shape.placeholder) {
                const ph = shape.placeholder;

                // Match by type and index
                if (ph.type === targetPlaceholder.type && ph.idx === targetPlaceholder.idx) {
                    return shape;
                }

                // Match by type if no index match
                if (ph.type === targetPlaceholder.type && !ph.idx) {
                    return shape;
                }
            }
        }
        return null;
    }

    /**
     * Apply placeholder styling from master/layout to slide placeholder
     */
    applyPlaceholderStyling(slideShape, templateShape) {
        // Apply text styling if present, but DO NOT copy placeholder text content
        if (templateShape.textBody) {
            if (!slideShape.textBody) {
                slideShape.textBody = { paragraphs: [] };
            }
            // Inherit only body/list styles; leave paragraphs empty unless slide provides its own
            if (templateShape.textBody.bodyProperties && !slideShape.textBody.bodyProperties) {
                slideShape.textBody.bodyProperties = JSON.parse(JSON.stringify(templateShape.textBody.bodyProperties));
            }
            if (templateShape.textBody.lstStyle && !slideShape.textBody.lstStyle) {
                slideShape.textBody.lstStyle = JSON.parse(JSON.stringify(templateShape.textBody.lstStyle));
            }
        }

        // Apply fill styling if not specified
        if (templateShape.fill && !slideShape.fill) {
            slideShape.fill = templateShape.fill;
        }

        // Apply stroke styling if not specified
        if (templateShape.stroke && !slideShape.stroke) {
            slideShape.stroke = templateShape.stroke;
        }

        // Apply geometry if not specified
        if (templateShape.geometry && !slideShape.geometry) {
            slideShape.geometry = templateShape.geometry;
        }

        // Apply style properties
        if (templateShape.style && !slideShape.style) {
            slideShape.style = templateShape.style;
        }
    }

    /**
     * Decide whether text should be rendered for a shape in the current context
     */
    shouldRenderText(shape) {
        try {
            // Suppress placeholder text in master/layout layers entirely
            if (this.currentLayer && this.currentLayer !== 'slide' && this.isPlaceholderShape(shape)) {
                return false;
            }

            // For slide layer: allow placeholder shapes only if they contain real text content
            if (this.isPlaceholderShape(shape)) {
                const hasRealText = !!(shape.textBody && Array.isArray(shape.textBody.paragraphs) &&
                    shape.textBody.paragraphs.some(p => Array.isArray(p.runs) && p.runs.some(r => (r.text || '').trim().length > 0)));
                return hasRealText;
            }
        } catch (_e) {}

        // Default: render text
        return true;
    }

    /**
     * Get background color from background fill with inheritance support
     */
    getBackgroundColor(backgroundFill) {
        if (!backgroundFill) {
            // Try inheritance from layout and master if no background fill
            const inheritedBg = this.resolveInheritedBackgroundColor();
            if (inheritedBg) {
                return inheritedBg;
            }
            return null; // No background color specified
        }

        // Handle different background fill types
        if (backgroundFill.type === 'solid' && backgroundFill.color) {
            const color = this.parseColorToHex(backgroundFill.color);
            return color;
        }

        if (backgroundFill.fill && backgroundFill.fill.type === 'solid' && backgroundFill.fill.color) {
            const color = this.parseColorToHex(backgroundFill.fill.color);
            return color;
        }

        // Handle bgPr type background
        if (backgroundFill.type === 'bgPr' && backgroundFill.fill) {
            if (backgroundFill.fill.type === 'solid' && backgroundFill.fill.color) {
                const color = this.parseColorToHex(backgroundFill.fill.color);
                return color;
            }
            if (backgroundFill.fill.type === 'gradient') {
                if (backgroundFill.fill.gradient && backgroundFill.fill.gradient.stops && backgroundFill.fill.gradient.stops.length > 0) {
                    // Return gradient info so drawSlideBackground can render it properly
                    return {
                        type: 'gradient',
                        gradient: backgroundFill.fill.gradient
                    };
                }
            }
            if (backgroundFill.fill.type === 'image' && backgroundFill.fill.imageData?.relationshipId) {
                return {
                    type: 'image',
                    imageData: backgroundFill.fill.imageData
                };
            }
        }

        // Handle bgRef type background (theme reference)
        if (backgroundFill.type === 'bgRef') {
            // Try to resolve theme background colors
            const idx = backgroundFill.idx;
            if (idx !== undefined) {
                // Logic: index > 1000 means bgFillStyleLst[index - 1000]
                if (idx > 1000) {
                    const bgStyleIndex = idx - 1000; // 1001 -> 1

                    // For bgFillStyleLst[1], this typically corresponds to a light theme background
                    // Based on common themes, bgFillStyleLst[1] is often bg2 (light background)
                    if (this.currentSlide && this.currentSlide.theme && this.currentSlide.theme.colors) {
                        const theme = this.currentSlide.theme;
                        // Try bg2 first for bgFillStyleLst[1], then bg1 as fallback
                        const bgColor = this.parseColorToHex(theme.colors.bg2) ||
                                      this.parseColorToHex(theme.colors.bg1) ||
                                      this.parseColorToHex(theme.colors.lt1);
                        return bgColor;
                    }
                } else {
                    // Direct theme color references (< 1000)
                    const themeColorMap = {
                        '0': 'bg1',      // usually white
                        '1': 'tx1',      // usually black
                        '2': 'bg2',      // usually light variant
                        '3': 'tx2'       // usually dark variant
                    };

                    const themeColorName = themeColorMap[idx.toString()];
                    if (themeColorName && this.currentSlide && this.currentSlide.theme && this.currentSlide.theme.colors) {
                        const bgColor = this.parseColorToHex(this.currentSlide.theme.colors[themeColorName]);
                        return bgColor;
                    }
                }

                // No theme color available for this index
                return null;
            }

            // No background reference resolution
            return null;
        }

        // No background color specified
        return null;
    }

    /**
     * Resolve inherited background color from layout and master
     */
    resolveInheritedBackgroundColor() {
        try {
            // Get current slide context
            if (!this.currentSlide) {return null;}

            // 1. Try to get background from layout
            if (this.currentSlide.layout && this.currentSlide.layout.backgroundFill) {
                const layoutBg = this.getBackgroundColorFromFill(this.currentSlide.layout.backgroundFill);
                if (layoutBg) {
                    return layoutBg;
                }
            }

            // 2. Try to get background from master
            if (this.currentSlide.layout && this.currentSlide.layout.master && this.currentSlide.layout.master.backgroundFill) {
                const masterBg = this.getBackgroundColorFromFill(this.currentSlide.layout.master.backgroundFill);
                if (masterBg) {
                    return masterBg;
                }
            }

            // 3. Try theme defaults (prefer light background colors)
            if (this.currentSlide.theme && this.currentSlide.theme.colors) {
                const themeBg = this.parseColorToHex(this.currentSlide.theme.colors.bg1) ||
                                this.parseColorToHex(this.currentSlide.theme.colors.lt1) ||
                                this.parseColorToHex(this.currentSlide.theme.colors.bg2);
                if (themeBg) {
                    return themeBg;
                }
            }

            return null;

        } catch (error) {
            return null;
        }
    }

    /**
     * Get background color from a fill object (helper method)
     */
    getBackgroundColorFromFill(backgroundFill) {
        if (!backgroundFill) {return null;}

        // Handle different background fill types (avoid recursion by using simplified parsing)
        if (backgroundFill.type === 'solid' && backgroundFill.color) {
            return this.parseColorToHex(backgroundFill.color);
        }

        if (backgroundFill.fill && backgroundFill.fill.type === 'solid' && backgroundFill.fill.color) {
            return this.parseColorToHex(backgroundFill.fill.color);
        }

        // Handle bgPr type background
        if (backgroundFill.type === 'bgPr' && backgroundFill.fill) {
            if (backgroundFill.fill.type === 'solid' && backgroundFill.fill.color) {
                return this.parseColorToHex(backgroundFill.fill.color);
            }
        }

        // Handle bgRef type background (theme reference)
        if (backgroundFill.type === 'bgRef') {
            const idx = backgroundFill.idx;
            if (idx !== undefined) {
                        const themeBgColors = {
            '0': '#ffffff',  // bg1 - usually white
            '1': '#000000',  // tx1 - usually black
            '2': '#f5f5f5',  // bg2 - usually light variant
            '3': '#404040',  // tx2 - usually dark variant
            '1000': '#ffffff', // white default for complex themes
        };

                return themeBgColors[idx.toString()] || null;
            }
        }

        return null;
    }

    /**
     * Parse color to hex format with enhanced support for scheme colors
     */
    parseColorToHex(color) {
        if (!color) {return '#ffffff';}

        if (typeof color === 'string') {
            return color.startsWith('#') ? color : `#${color}`;
        }

        // Handle scheme colors with proper mapping
        if (color.scheme) {
            const schemeColors = {
                'bg1': '#ffffff',    // Background 1 - usually white
                'tx1': '#000000',    // Text 1 - usually black
                'bg2': '#f8f8f8',    // Background 2 - light gray
                'tx2': '#404040',    // Text 2 - dark gray
                'accent1': '#4f81bd', // Accent 1 - blue
                'accent2': '#c0504d', // Accent 2 - red
                'accent3': '#9bbb59', // Accent 3 - green
                'accent4': '#8064a2', // Accent 4 - purple
                'accent5': '#4bacc6', // Accent 5 - cyan
                'accent6': '#f79646', // Accent 6 - orange
                'lt1': '#ffffff',    // Light 1
                'dk1': '#000000',    // Dark 1
                'lt2': '#f0f0f0',    // Light 2 - light gray
                'dk2': '#404040'     // Dark 2
            };

            let baseColor = schemeColors[color.scheme];
            if (baseColor) {
                // Apply color modifications (tint, shade, etc.)
                baseColor = this.applyColorModifications(baseColor, color);
                return baseColor;
            }
        }

        if (color.r !== undefined && color.g !== undefined && color.b !== undefined) {
            if (color.a !== undefined && color.a < 255) {
                return `rgba(${color.r}, ${color.g}, ${color.b}, ${(color.a / 255).toFixed(4)})`;
            }
            const hex = `#${color.r.toString(16).padStart(2, '0')}${color.g.toString(16).padStart(2, '0')}${color.b.toString(16).padStart(2, '0')}`;
            return hex;
        }

        return '#ffffff';
    }
    /**
     * Apply color modifications like tint, shade, etc.
     */
    applyColorModificationsLegacy(hexColor, colorInfo) {
        if (!colorInfo || (!colorInfo.tint && !colorInfo.shade && !colorInfo.lumMod && !colorInfo.lumOff)) {
            return hexColor;
        }

        // Convert hex to RGB
        const r = parseInt(hexColor.substr(1, 2), 16);
        const g = parseInt(hexColor.substr(3, 2), 16);
        const b = parseInt(hexColor.substr(5, 2), 16);

        let newR = r, newG = g, newB = b;

        // Apply tint (lighten)
        if (colorInfo.tint !== undefined) {
            const tint = colorInfo.tint / 100000; // Convert from percentage
            newR = Math.round(r + (255 - r) * tint);
            newG = Math.round(g + (255 - g) * tint);
            newB = Math.round(b + (255 - b) * tint);
        }

        // Apply shade (darken)
        if (colorInfo.shade !== undefined) {
            const shade = colorInfo.shade / 100000; // Convert from percentage
            newR = Math.round(r * (1 - shade));
            newG = Math.round(g * (1 - shade));
            newB = Math.round(b * (1 - shade));
        }

        // Apply luminance modulation
        if (colorInfo.lumMod !== undefined) {
            const lumMod = colorInfo.lumMod / 100000; // Convert from percentage
            newR = Math.round(r * lumMod);
            newG = Math.round(g * lumMod);
            newB = Math.round(b * lumMod);
        }

        // Apply luminance offset
        if (colorInfo.lumOff !== undefined) {
            const lumOff = colorInfo.lumOff / 100000 * 255; // Convert from percentage to 0-255
            newR = Math.round(Math.min(255, r + lumOff));
            newG = Math.round(Math.min(255, g + lumOff));
            newB = Math.round(Math.min(255, b + lumOff));
        }

        // Ensure values are within valid range
        newR = Math.max(0, Math.min(255, newR));
        newG = Math.max(0, Math.min(255, newG));
        newB = Math.max(0, Math.min(255, newB));

        const result = `#${newR.toString(16).padStart(2, '0')}${newG.toString(16).padStart(2, '0')}${newB.toString(16).padStart(2, '0')}`;
        return result;
    }
    /**
     * Setup coordinate system for slide to canvas mapping using slide renderer's calculations
     */
    setupSlideCoordinateSystem() {
        // Use the slide renderer's canvas rect calculation for consistency
        if (this.processor && this.processor.calculateCanvasRect) {
            const slideSize = this.processor.getSlideDimensions();
            const canvasRect = this.processor.calculateCanvasRect(this.canvas, slideSize);
            
            
            // Store coordinate system using slide renderer's calculations
            this.coordinateSystem = {
                slideWidthEMU: slideSize.cx,
                slideHeightEMU: slideSize.cy,
                slideWidthPx: canvasRect.slideWidthPx,
                slideHeightPx: canvasRect.slideHeightPx,
                scale: canvasRect.scale,
                offsetX: canvasRect.offsetX,
                offsetY: canvasRect.offsetY,
                scaledWidth: canvasRect.scaledWidth,
                scaledHeight: canvasRect.scaledHeight
            };
        } else {
            // Fallback to original calculation if slide renderer not available
            const slideSize = this.processor.getSlideDimensions();
            const slideWidthEMU = slideSize.cx;
            const slideHeightEMU = slideSize.cy;

            // Convert EMU to pixels
            const slideWidthPx = slideWidthEMU / 914400 * 96;
            const slideHeightPx = slideHeightEMU / 914400 * 96;

            // Calculate scale to fit canvas
            let logicalWidth;
            let logicalHeight;
            const rect = typeof this.canvas.getBoundingClientRect === 'function'
                ? this.canvas.getBoundingClientRect()
                : { width: 0, height: 0 };
            if (rect.width > 0 && rect.height > 0) {
                logicalWidth = rect.width;
                logicalHeight = rect.height;
            } else if (this.canvas.style.width && this.canvas.style.height) {
                logicalWidth = parseFloat(this.canvas.style.width);
                logicalHeight = parseFloat(this.canvas.style.height);
            } else {
                const pixelRatio = typeof window !== 'undefined' && window.devicePixelRatio ? window.devicePixelRatio : 1;
                logicalWidth = this.canvas.width / pixelRatio;
                logicalHeight = this.canvas.height / pixelRatio;
            }

            const scaleX = logicalWidth / slideWidthPx;
            const scaleY = logicalHeight / slideHeightPx;
            const scale = Math.min(scaleX, scaleY);

            // Center the slide
            let scaledWidth = slideWidthPx * scale;
            let scaledHeight = slideHeightPx * scale;
            let offsetX = (logicalWidth - scaledWidth) / 2;
            let offsetY = (logicalHeight - scaledHeight) / 2;

            // Snap to integer pixels to avoid 1px gutters
            scaledWidth = Math.round(scaledWidth);
            scaledHeight = Math.round(scaledHeight);
            offsetX = Math.round(offsetX);
            offsetY = Math.round(offsetY);

            // Store coordinate system
            this.coordinateSystem = {
                slideWidthEMU,
                slideHeightEMU,
                slideWidthPx,
                slideHeightPx,
                scale,
                offsetX,
                offsetY,
                scaledWidth,
                scaledHeight,
                canvasWidth: logicalWidth,
                canvasHeight: logicalHeight
            };
        }
    }


    /**
     * Draw individual shape with enhanced positioning
     */
    async drawShapeEnhanced(shape, shapeIndex) {
        // Ensure inheritance is applied before computing bounds so placeholder positions are available
        try {
            if (this.processor && typeof this.processor.applyPropertyInheritance === 'function') {
                this.processor.applyPropertyInheritance(shape, this.currentSlide);
            }
        } catch (_e) {}

        // Get bounds after inheritance
        const bounds = this.getShapeBounds(shape);

        // If this is a text shape with zero/invalid height, estimate auto height from content
        // Skip for line-like shapes: their cy=0 is intentional; text is positioned separately.
        if (bounds && (bounds.h === 0 || !isFinite(bounds.h) || bounds.h < 1) && shape && shape.textBody && shape.textBody.paragraphs && shape.textBody.paragraphs.length > 0 && !this.isLineLikeShape(shape)) {
            try {
                const estimated = this.estimateTextBoxHeightPx(shape, bounds.w);
                if (estimated && estimated > 0) {
                    bounds.h = Math.ceil(estimated);
                }
            } catch (_e) {}
        }

        if (!this.graphics || !shape) {
            return;
        }

        try {
            // Allow line-like shapes even if one dimension is zero
            const isLineLike = this.isLineLikeShape(shape);
            if (!bounds || (bounds.w <= 0 && bounds.h <= 0) || (!isLineLike && (bounds.w <= 0 || bounds.h <= 0))) {
                return;
            }

            // Draw shape based on type
            switch (shape.type) {
                case 'sp':
                    this.drawRegularShape(shape, bounds);
                    break;
                case 'pic':
                    this.drawPictureShape(shape, bounds);
                    break;
                case 'grpSp':
                    this.drawGroupShape(shape, bounds);
                    break;
                case 'cxnSp':
                    this.drawConnectorShape(shape, bounds);
                    break;
                case 'graphicFrame':
                    await this.drawGraphicFrame(shape, bounds);
                    break;
                default:
                    this.drawDefaultShape(shape, bounds);
                    break;
            }

        } catch (error) {
        }
    }

    /**
     * Draw shape geometry
     */
    drawShapeGeometry(shape, bounds) {
        // Skip shapes with zero or invalid dimensions
        if (!bounds || bounds.w <= 0 || bounds.h <= 0) {
            return;
        }
        
        const fillColor = this.getShapeFillColor(shape);
        const strokeColor = this.getShapeStrokeColor(shape);
        const lineWidth = this.getShapeLineWidth(shape);

        // Set colors
        if (fillColor) {
            const rgb = this.parseColor(fillColor);
            this.graphics.b_color1(rgb.r, rgb.g, rgb.b, rgb.a);
        }

        if (strokeColor) {
            const rgb = this.parseColor(strokeColor);
            this.graphics.p_color(rgb.r, rgb.g, rgb.b, rgb.a);
            this.graphics.p_width(lineWidth);
        }

        // Draw based on shape type
        switch (shape.type) {
            case 'sp':
            case 'shape':
                this.drawRegularShape(shape, bounds);
                break;
            case 'pic':
                this.drawPictureShape(shape, bounds);
                break;
            case 'grpSp':
                this.drawGroupShape(shape, bounds);
                break;
            default:
                this.drawDefaultShape(shape, bounds);
                break;
        }
    }

    /**
     * Draw regular shape
     */
    drawRegularShape(shape, bounds) {
        // Skip shapes with zero or invalid dimensions
        // Allow line-like sp shapes (preset=line) with one zero dimension
        if (!bounds) return;
        const isLinePreset = this.isLineLikeShape(shape);
        if (isLinePreset) {
            if (bounds.w <= 0 && bounds.h <= 0) return;
        } else {
            if (bounds.w <= 0 || bounds.h <= 0) return;
        }

        if (!this.graphics || !this.graphics.context) {
            return;
        }

        const ctx = this.graphics.context;
        ctx.save();

        // Apply rotation if present
        const transform = shape.properties?.transform;
        if (transform && transform.rotation && transform.rotation !== 0) {
            const centerX = bounds.x + bounds.w / 2;
            const centerY = bounds.y + bounds.h / 2;
            ctx.translate(centerX, centerY);
            ctx.rotate(transform.rotation * Math.PI / 180);
            ctx.translate(-centerX, -centerY);
        }

        // Transparency is handled in drawPresetGeometry via fillColor.a

        // Get enhanced fill and stroke colors
        const fillColor = this.getShapeFillColor(shape);
        const strokeColor = this.getShapeStrokeColor(shape);
        const lineWidth = this.getShapeLineWidth(shape);
        const strokeInfo = this.getShapeStrokeInfo(shape);

        // Apply effects (shadow, glow, etc.) before drawing
        if (shape.properties && shape.properties.effectLst) {
            this.applyEffectsToCanvas(ctx, shape.properties.effectLst);
        }
        
        

        // Get shape geometry with layout inheritance
        let preset = null;
        if (this.processor && this.processor.getShapePresetGeometry) {
            // Use slide-renderer's enhanced geometry method with layout inheritance
            preset = this.processor.getShapePresetGeometry(shape, this.processor.currentSlide);
        } else {
            // Fallback to local method
            preset = this.getShapePreset(shape);
        }

        // Special handling for line preset or connector type
        if (preset === 'line' || this.isLineLikeShape(shape)) {
            this.drawEngineLine(bounds, shape);
            // Reset effects after drawing
            this.resetEffectsOnContext(ctx);
            // Draw text if present (e.g., "LINE size=1" labels).
            // PowerPoint anchors line-shape text at the line's start point (shape.x),
            // extending in the direction set by the paragraph alignment:
            //   'l'/default → text starts at shape.x (extends right)
            //   'ctr'       → text centered at shape.x
            //   'r'         → text ends at shape.x (extends left)
            if (shape.textBody && shape.textBody.paragraphs && this.shouldRenderText(shape)) {
                const estimatedH = this.estimateTextBoxHeightPx(shape, bounds.w) || 24;
                const firstPara = shape.textBody.paragraphs[0];
                // Alignment is stored under paragraph.properties.align (set by slide-renderer parseParagraphProperties).
                // Fallback chain: properties.align → properties.alignment → pPr.algn → 'l'
                const rawAlgn = firstPara?.properties?.align ||
                                firstPara?.properties?.alignment ||
                                firstPara?.paragraphProperties?.align ||
                                firstPara?.pPr?.algn || 'l';
                let textX;
                if (rawAlgn === 'r' || rawAlgn === 'right') {
                    textX = bounds.x - bounds.w;
                } else if (rawAlgn === 'ctr' || rawAlgn === 'center') {
                    textX = bounds.x - bounds.w / 2;
                } else {
                    textX = bounds.x; // 'l' or default
                }
                const lineBounds = {
                    x: textX,
                    y: bounds.y - estimatedH / 2,
                    w: bounds.w,
                    h: estimatedH
                };
                this.setCurrentRenderingShape(shape);
                this.drawShapeText(shape, lineBounds);
            }
            ctx.restore();
            return;
        }


        // Apply flip for geometry only (not for text rendering) via a nested save/restore
        const needsFlip = transform && (transform.flipH || transform.flipV);
        if (needsFlip) {
            ctx.save();
            const flipCx = bounds.x + bounds.w / 2;
            const flipCy = bounds.y + bounds.h / 2;
            ctx.translate(flipCx, flipCy);
            ctx.scale(transform.flipH ? -1 : 1, transform.flipV ? -1 : 1);
            ctx.translate(-flipCx, -flipCy);
        }

        // Check for custom geometry first
        if (shape.geometry && shape.geometry.type === 'custom') {
            try {
                // Set current shape bounds for gradient calculation
                this.currentShapeBounds = bounds;

                // Apply effects before calling graphics engine
                if (shape.properties && shape.properties.effectLst) {
                    this.applyEffectsToCanvas(ctx, shape.properties.effectLst);
                }

                // Use the graphics drawCustomGeometry method directly
                this.graphics.drawCustomGeometry(shape.geometry, bounds.x, bounds.y, bounds.w, bounds.h, fillColor, strokeColor, lineWidth);

                // Reset effects after drawing
                this.resetEffectsOnContext(ctx);
            } catch (error) {
                // Fallback to rectangle if custom geometry fails
                this.graphics.drawPresetGeometry('rect', bounds.x, bounds.y, bounds.w, bounds.h, fillColor, strokeColor, lineWidth, strokeInfo);
            }
        }
        // Use the graphics engine for proper shape rendering
        else if (this.graphics.drawPresetGeometry && preset) {
            // Apply effects before calling graphics engine
            if (shape.properties && shape.properties.effectLst) {
                this.applyEffectsToCanvas(ctx, shape.properties.effectLst);
            }
            const adjustments = shape.geometry?.adjustments || shape.properties?.geometry?.adjustments || {};
            this.graphics.drawPresetGeometry(preset, bounds.x, bounds.y, bounds.w, bounds.h, fillColor, strokeColor, lineWidth, strokeInfo, adjustments);
            // Reset effects after drawing
            this.resetEffectsOnContext(ctx);
        } else {
            // For shapes with significant stroke width, prioritize stroke over fill
            const hasSignificantStroke = strokeColor && lineWidth > 2;
            const shouldPrioritizeStroke = hasSignificantStroke || (strokeColor && !fillColor);
            

            // Handle stroke first for path-like shapes
            if (strokeColor && shouldPrioritizeStroke) {
                ctx.strokeStyle = this.graphics.colorToRgb(strokeColor);
                ctx.lineWidth = CoordinateTransform.mmToPixels(lineWidth);

                // Apply advanced stroke styling if available
                if (strokeInfo) {
                    this.applyStrokeStyle(ctx, strokeInfo, CoordinateTransform.mmToPixels(lineWidth));
                }

                ctx.strokeRect(bounds.x, bounds.y, bounds.w, bounds.h);
            }
            
            // Only apply fill if not prioritizing stroke
            if (!shouldPrioritizeStroke) {
                // Handle gradient fills
                if (fillColor && fillColor.type === 'linear') {
                    this.drawGradientFill(ctx, bounds, fillColor);
                } else if (fillColor) {
                    ctx.fillStyle = this.graphics.colorToRgb(fillColor);
                    ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
                }

                // Handle stroke after fill for regular shapes
                if (strokeColor) {
                    ctx.strokeStyle = this.graphics.colorToRgb(strokeColor);
                    ctx.lineWidth = CoordinateTransform.mmToPixels(lineWidth);

                    // Apply advanced stroke styling if available
                    if (strokeInfo) {
                        this.applyStrokeStyle(ctx, strokeInfo, CoordinateTransform.mmToPixels(lineWidth));
                    }

                    ctx.strokeRect(bounds.x, bounds.y, bounds.w, bounds.h);
                }
            }
        }

        // Restore flip transform so text is drawn unflipped
        if (needsFlip) {
            ctx.restore();
        }

        // Draw text content for all rendering paths using adapter's own method
        if (shape.textBody && shape.textBody.paragraphs && this.shouldRenderText(shape)) {
            // Set current shape context for inheritance
            this.setCurrentRenderingShape(shape);
            this.drawShapeText(shape, bounds);
        }

        // Reset effects after drawing
        this.resetEffectsOnContext(ctx);

        ctx.restore();
    }

    /**
     * Determine if shape is a line or connector
     */
    isLineLikeShape(shape) {
        const preset = (shape && (shape.geometry?.preset || shape.properties?.geometry?.preset || this.getShapePreset(shape))) || null;
        const name = (typeof preset === 'string') ? preset.toLowerCase() : '';
        if (name === 'line' || name.includes('connector')) {return true;}
        if (shape && (shape.type === 'cxnSp' || shape.type === 'connector')) {return true;}
        return false;
    }

    /**
     * Draw a line using the engine with arrow/dash support if available
     */
    drawEngineLine(bounds, shape) {
        const strokeColor = this.getShapeStrokeColor(shape);
        const lineWidth = this.getShapeLineWidth(shape) || 1;
        const strokeInfo = this.getShapeStrokeInfo(shape);
        // Compute endpoints from XML transform (rotation, flips) and bounds.
        const { x1, y1, x2, y2 } = this._computeConnectorEndpoints(bounds, shape);
        if (this.graphics && typeof this.graphics.drawLineWithArrows === 'function') {
            // Ensure stroke color/width defaults if missing so arrowheads render
            const safeColor = strokeColor || { r: 0, g: 0, b: 0, a: 255 };
            const safeWidth = (lineWidth && lineWidth > 0) ? lineWidth : 0.75; // ~0.75mm default
            this.graphics.drawLineWithArrows(x1, y1, x2, y2, safeColor, safeWidth, strokeInfo);
        } else if (this.graphicsEngine && typeof this.graphicsEngine.drawLineWithArrows === 'function') {
            // Use bridged engine directly if adapter surface lacks the method
            const safeColor = strokeColor || { r: 0, g: 0, b: 0, a: 255 };
            const safeWidth = (lineWidth && lineWidth > 0) ? lineWidth : 0.75;
            this.graphicsEngine.drawLineWithArrows(x1, y1, x2, y2, safeColor, safeWidth, strokeInfo);
        } else if (this.graphics && typeof this.graphics.drawLine === 'function') {
            const safeColor = strokeColor || { r: 0, g: 0, b: 0, a: 255 };
            const safeWidth = (lineWidth && lineWidth > 0) ? lineWidth : 0.75;
            // Draw line via engine fallback
            this.graphics.drawLine(x1, y1, x2, y2, safeColor, safeWidth);
        } else if (this.graphics && this.graphics.context) {
            const ctx = this.graphics.context;
            ctx.save();
            ctx.beginPath();
            ctx.moveTo(x1, y1);
            ctx.lineTo(x2, y2);
            if (strokeColor) {
                ctx.strokeStyle = this.graphics.colorToRgb(strokeColor);
                // Convert mm to px for proper thickness
                const pxWidth = this.graphics && this.graphics.CoordinateTransform ? this.graphics.CoordinateTransform.mmToPixels(lineWidth) : CoordinateTransform.mmToPixels(lineWidth);
                ctx.lineWidth = Math.max(pxWidth, 1);
                // Apply stroke styling if present (caps, joins, dashes)
                try { this.applyStrokeStyle(ctx, strokeInfo, ctx.lineWidth); } catch(_e) {}
                ctx.stroke();
            }
            ctx.restore();
        }
    }

    /**
     * Compute connector endpoints honoring XML transform (flipH/flipV/rotation) and using
     * the shape's `spPr/ln` arrow info to keep start/end consistent with tailEnd/headEnd.
     * We derive endpoints from bounds as: start at the tail side, end at the head side.
     */
    _computeConnectorEndpoints(bounds, shape) {
        // Prefer geometry path order from XML if available (start = first moveTo, end = last point).
        // Otherwise, derive from bbox diagonal (off/ext) and apply flips/rotation.
        const width = Math.max(0, bounds.w);
        const height = Math.max(0, bounds.h);
        const centerX = bounds.x + width / 2;
        const centerY = bounds.y + height / 2;

        let x1, y1, x2, y2;

        // 1) Try geometry path-based endpoints
        const geom = (shape && (shape.geometry || shape.properties?.geometry)) || null;
        const pathList = geom && (geom.pathList || geom.pathLst);
        if (Array.isArray(pathList) && pathList.length > 0) {
            const path = pathList[0];
            const cmds = path.commands || [];
            let startPt = null;
            let endPt = null;
            for (let i = 0; i < cmds.length; i++) {
                const c = cmds[i];
                if (!startPt && c.type === 'moveTo') {
                    startPt = { x: c.x, y: c.y };
                }
                // Update endPt for each drawable command
                if (c.type === 'lineTo') {
                    endPt = { x: c.x, y: c.y };
                } else if (c.type === 'cubicBezTo') {
                    endPt = { x: c.x3, y: c.y3 };
                } else if (c.type === 'moveTo' && !endPt && startPt) {
                    // If a moveTo appears after start and no drawable command yet, use it as end candidate
                    endPt = { x: c.x, y: c.y };
                }
            }
            if (!startPt && cmds.length > 0) {
                const c0 = cmds[0];
                if (c0.x != null && c0.y != null) { startPt = { x: c0.x, y: c0.y }; }
            }
            if (!endPt && startPt) { endPt = { ...startPt }; }

            if (startPt && endPt) {
                const pathW = path.w || geom.pathW || 100;
                const pathH = path.h || geom.pathH || 100;
                const sx = (pathW === 0) ? 1 : (width / pathW);
                const sy = (pathH === 0) ? 1 : (height / pathH);
                x1 = bounds.x + startPt.x * sx;
                y1 = bounds.y + startPt.y * sy;
                x2 = bounds.x + endPt.x * sx;
                y2 = bounds.y + endPt.y * sy;
                
            }
        }

        // 2) Fallback to bbox diagonal if geometry is not available
        if (x1 === undefined) {
            x1 = bounds.x; y1 = bounds.y;
            x2 = bounds.x + width; y2 = bounds.y + height;
            // If the shape's original OOXML cy was 0 (horizontal line), snap to horizontal.
            // The bounds height may have been inflated for text layout, but the line itself
            // should run at the vertical center of the (possibly inflated) bounding box.
            const xfrm = shape && (shape.spPr?.xfrm || shape.transform);
            const origCy = xfrm && (xfrm.cy !== undefined ? xfrm.cy : xfrm.height);
            if (origCy === 0) {
                y1 = bounds.y + height / 2;
                y2 = y1;
            }
        }

        // Apply flips around the bbox center
        const flipH = !!(shape && shape.transform && shape.transform.flipH);
        const flipV = !!(shape && shape.transform && shape.transform.flipV);
        if (flipH) {
            x1 = centerX - (x1 - centerX);
            x2 = centerX - (x2 - centerX);
        }
        if (flipV) {
            y1 = centerY - (y1 - centerY);
            y2 = centerY - (y2 - centerY);
        }

        // Apply rotation (degrees) around center
        const rot = (shape && shape.transform && (shape.transform.rot || shape.transform.rotation)) || 0;
        if (rot) {
            const rad = (rot * Math.PI) / 180;
            const cos = Math.cos(rad);
            const sin = Math.sin(rad);
            const rx = (x, y) => ({
                x: centerX + (x - centerX) * cos - (y - centerY) * sin,
                y: centerY + (x - centerX) * sin + (y - centerY) * cos
            });
            const p1 = rx(x1, y1);
            const p2 = rx(x2, y2);
            x1 = p1.x; y1 = p1.y;
            x2 = p2.x; y2 = p2.y;
        }

        // Return start→end consistent with XML path order; arrow placement uses these directly
        
        return { x1, y1, x2, y2 };
    }

    /**
     * Override for line preset to use engine arrow drawing when possible
     */
    drawLinePath(ctx, x, y, width, height) {
        // If we have access to graphics engine with arrow support, draw via engine
        if (this.graphics && typeof this.graphics.drawLineWithArrows === 'function') {
            const strokeColor = this.getShapeStrokeColor(this.getCurrentShape());
            const lineWidth = this.getShapeLineWidth(this.getCurrentShape()) || 1;
            const strokeInfo = this.getShapeStrokeInfo(this.getCurrentShape());

            // Line from top-left to bottom-right of bounds
            const x1 = x;
            const y1 = y;
            const x2 = x + width;
            const y2 = y + height;
            this.graphics.drawLineWithArrows(x1, y1, x2, y2, strokeColor, lineWidth, strokeInfo);
            return;
        }
        // Fallback path-based line
        ctx.moveTo(x, y + height / 2);
        ctx.lineTo(x + width, y + height / 2);
    }

    /**
     * Draw a line with optional arrowheads using the adapter's canvas context.
     * Provides arrow rendering without requiring the engine.
     */
    drawLineWithArrows(x1, y1, x2, y2, strokeColor, lineWidth = 1, strokeInfo = null) {
        const ctx = this._context;
        if (!ctx) {return;}

        ctx.save();
        ctx.beginPath();
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);

        const color = this.colorToRgb(strokeColor) || 'rgba(0,0,0,1)';
        const pxWidth = Math.max(CoordinateTransform.mmToPixels(lineWidth || 1), 1);
        ctx.strokeStyle = color;
        ctx.lineWidth = pxWidth;
        if (strokeInfo) { try { this.applyStrokeStyle(ctx, strokeInfo, pxWidth); } catch(_e) {} }
        ctx.stroke();

        const mapSize = (s) => {
            if (typeof s === 'number') {return s;}
            const v = (s || '').toString();
            if (v === 'lg' || v === 'large') {return 2.0;}
            if (v === 'med' || v === 'medium') {return 1.5;}
            return 1.0;
        };
        const startDef = strokeInfo && strokeInfo.tailEnd;
        const endDef = strokeInfo && strokeInfo.headEnd;

        if (startDef || endDef) {
            const angle = Math.atan2(y2 - y1, x2 - x1);
            // Get current transformation scale to make arrow heads proportional to display size
            const transform = ctx.getTransform();
            const currentScale = Math.sqrt(transform.a * transform.a + transform.b * transform.b);
            // Scale arrow head base sizes with canvas scale, with reasonable minimums
            const scaledMinLen = Math.max(4, 10 * currentScale);
            const scaledMinWid = Math.max(2.5, 6 * currentScale);
            const headLen = Math.max(scaledMinLen, pxWidth * 8);
            const headWid = Math.max(scaledMinWid, pxWidth * 3);
            

            const drawHead = (x, y, ang, def) => {
                if (!def) {return;}
                const length = headLen * (def.lengthScale || mapSize(def.len));
                const width = headWid * (def.widthScale || mapSize(def.w));
                const type = (def.type || def.val || 'arrow').toString();
                const halfW = width / 2;
                ctx.save();
                ctx.translate(x, y);
                ctx.rotate(ang);
                ctx.beginPath();
                if (type === 'open') {
                    ctx.moveTo(0, 0);
                    ctx.lineTo(-length, -halfW);
                    ctx.moveTo(0, 0);
                    ctx.lineTo(-length, halfW);
                    ctx.strokeStyle = color;
                    ctx.lineWidth = Math.max(1, pxWidth);
                    ctx.stroke();
                } else if (type === 'stealth') {
                    ctx.moveTo(0, 0);
                    ctx.lineTo(-length, -(halfW * 0.6));
                    ctx.lineTo(-length, (halfW * 0.6));
                    ctx.closePath();
                    ctx.fillStyle = color;
                    ctx.fill();
                    ctx.strokeStyle = color;
                    ctx.lineWidth = Math.max(1, pxWidth * 0.5);
                    ctx.stroke();
                } else if (type === 'diamond') {
                    ctx.moveTo(0, 0);
                    ctx.lineTo(-length / 2, -halfW);
                    ctx.lineTo(-length, 0);
                    ctx.lineTo(-length / 2, halfW);
                    ctx.closePath();
                    ctx.fillStyle = color;
                    ctx.fill();
                    ctx.strokeStyle = color;
                    ctx.lineWidth = Math.max(1, pxWidth * 0.5);
                    ctx.stroke();
                } else if (type === 'oval') {
                    ctx.ellipse(-length / 2, 0, length / 2, halfW, 0, 0, Math.PI * 2);
                    ctx.fillStyle = color;
                    ctx.fill();
                    ctx.strokeStyle = color;
                    ctx.lineWidth = Math.max(1, pxWidth * 0.5);
                    ctx.stroke();
                } else {
                    // triangle/arrow
                    ctx.moveTo(0, 0);
                    ctx.lineTo(-length, -halfW);
                    ctx.lineTo(-length, halfW);
                    ctx.closePath();
                    ctx.fillStyle = color;
                    ctx.fill();
                    ctx.strokeStyle = color;
                    ctx.lineWidth = Math.max(1, pxWidth * 0.5);
                    ctx.stroke();
                }
                ctx.restore();
            };

            if (hasStart) { drawHead(x1, y1, angle + Math.PI, startDef); }
            if (hasEnd) { drawHead(x2, y2, angle, endDef); }
        } else {
            
        }

        ctx.restore();
    }

    /**
     * Draw gradient fill on canvas
     */
    drawGradientFill(ctx, bounds, gradient) {
        if (!gradient || !gradient.stops || gradient.stops.length === 0) {return;}

        let grad;
        if (gradient.type === 'linear') {
            // Linear gradient
            const angle = gradient.angle || 0;
            const radians = (angle * Math.PI) / 180;

            const x1 = bounds.x + bounds.w / 2 - (Math.cos(radians) * bounds.w) / 2;
            const y1 = bounds.y + bounds.h / 2 - (Math.sin(radians) * bounds.h) / 2;
            const x2 = bounds.x + bounds.w / 2 + (Math.cos(radians) * bounds.w) / 2;
            const y2 = bounds.y + bounds.h / 2 + (Math.sin(radians) * bounds.h) / 2;

            grad = ctx.createLinearGradient(x1, y1, x2, y2);
        } else {
            // Radial gradient
            const centerX = bounds.x + bounds.w / 2;
            const centerY = bounds.y + bounds.h / 2;
            const radius = Math.max(bounds.w, bounds.h) / 2;

            grad = ctx.createRadialGradient(centerX, centerY, 0, centerX, centerY, radius);
        }

        // Add color stops
        for (const stop of gradient.stops) {
            grad.addColorStop(stop.position, this.graphics.colorToRgb(stop.color));
        }

        ctx.fillStyle = grad;
        ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
    }

    /**
     * Apply advanced stroke styling to canvas context
     */
    applyStrokeStyle(ctx, strokeInfo, lineWidth) {
        if (!ctx || !strokeInfo) {return;}

        // Set line cap
        if (strokeInfo.cap) {
            switch (strokeInfo.cap) {
                case 'rnd':
                    ctx.lineCap = 'round';
                    break;
                case 'sq':
                    ctx.lineCap = 'square';
                    break;
                case 'flat':
                default:
                    ctx.lineCap = 'butt';
                    break;
            }
        }

        // Set line join
        if (strokeInfo.join) {
            switch (strokeInfo.join) {
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

        // Set dash pattern
        if (strokeInfo.dashArray && strokeInfo.dashArray.length > 0) {
            // Scale dash pattern by line width
            const scaledDashArray = strokeInfo.dashArray.map(dash => dash * lineWidth);
            ctx.setLineDash(scaledDashArray);
        } else {
            ctx.setLineDash([]);
        }
    }
    /**
     * Get shape preset geometry
     */
    getShapePreset(shape) {
        // Check direct geometry preset (primary location)
        if (shape.geometry && shape.geometry.preset) {
            return shape.geometry.preset;
        }

        // Check properties geometry preset
        if (shape.properties && shape.properties.geometry && shape.properties.geometry.preset) {
            return shape.properties.geometry.preset;
        }

        // Check spPr geometry preset
        if (shape.spPr && shape.spPr.geometry && shape.spPr.geometry.preset) {
            return shape.spPr.geometry.preset;
        }

        // Check spPr direct prstGeom structure (legacy compatibility)
        if (shape.spPr && shape.spPr.prstGeom && shape.spPr.prstGeom.prst) {
            return shape.spPr.prstGeom.prst;
        }

        // Default preset for placeholder shapes - they should typically be rectangles
        if (this.isPlaceholderShape(shape)) {
            return 'rect';
        }

        // Default preset for standard shapes without explicit geometry
        if (shape.type === 'sp' || shape.type === 'shape') {
            return 'rect';
        }

        return null;
    }
    /**
     * Draw picture shape
     */
    drawPictureShape(shape, bounds) {
        
        
        // Skip shapes with zero or invalid dimensions
        if (!bounds || bounds.w <= 0 || bounds.h <= 0) {
            return;
        }
        
        if (!this.graphics || !this.graphics.context) {return;}

        const ctx = this.graphics.context;
        ctx.save();

        // Apply rotation if present
        const transform = shape.properties?.transform;
        if (transform && transform.rotation && transform.rotation !== 0) {
            const centerX = bounds.x + bounds.w / 2;
            const centerY = bounds.y + bounds.h / 2;
            ctx.translate(centerX, centerY);
            ctx.rotate(transform.rotation * Math.PI / 180);
            ctx.translate(-centerX, -centerY);
        }

        // Apply geometry clip (e.g. ellipse for circular/rounded images)
        const prstGeomPreset = this.getShapePreset(shape);
        if (prstGeomPreset && prstGeomPreset !== 'rect') {
            if (prstGeomPreset === 'ellipse') {
                ctx.beginPath();
                ctx.ellipse(bounds.x + bounds.w / 2, bounds.y + bounds.h / 2,
                    bounds.w / 2, bounds.h / 2, 0, 0, 2 * Math.PI);
                ctx.clip();
            }
        }

        // Check if we have an image for this shape
        if (shape.imageRelId && this.processor && this.processor.imageCache) {
            const imageData = this.processor.imageCache.get(shape.imageRelId);
            
            if (imageData) {
                
            }

            // Check if this is SVG content and we have an SVG renderer
            if (imageData && imageData.type === 'svg' && this.processor.svgRenderer) {
                
                try {
                    // Use async SVG rendering - handle context properly
                    const svgRenderPromise = this.processor.svgRenderer.renderSVG(
                        imageData.content,
                        bounds.x,
                        bounds.y,
                        bounds.w,
                        bounds.h,
                        { preserveAspectRatio: true }
                    );
                    
                    svgRenderPromise.then(() => {}).catch(() => {});
                    
                    // Restore context immediately and return - SVG renderer handles its own rendering context
                    ctx.restore();
                    return; // Exit early for SVG handling
                } catch (error) {
                    // Fall through to regular image handling or placeholder
                }
            }

            if (imageData && imageData.image) {
                // Draw the actual image
                try {
                    const image = imageData.image;
                    const { width: naturalWidth, height: naturalHeight } = imageData;

                    // Calculate image positioning based on fill mode
                    let drawX = bounds.x, drawY = bounds.y, drawW = bounds.w, drawH = bounds.h;

                    if (shape.imageFillMode === 'stretch') {
                        if (shape.imageSrcRect) {
                            // srcRect stretches a (possibly virtual) portion of the source to fill the shape.
                            // Positive values crop inward; negative values add virtual padding (letterbox).
                            // Formula: image destination = shape_bounds scaled so that srcRect maps to (0,0,w,h).
                            const sr = shape.imageSrcRect;
                            const sW = 100000 - sr.l - sr.r;   // source width fraction (1/100000 units)
                            const sH = 100000 - sr.t - sr.b;   // source height fraction
                            if (sW > 0 && sH > 0) {
                                drawX = bounds.x + bounds.w * (-sr.l) / sW;
                                drawY = bounds.y + bounds.h * (-sr.t) / sH;
                                drawW = bounds.w * 100000 / sW;
                                drawH = bounds.h * 100000 / sH;
                                // Clip to shape rect so the image doesn't overflow outside bounds
                                ctx.beginPath();
                                ctx.rect(bounds.x, bounds.y, bounds.w, bounds.h);
                                ctx.clip();
                            }
                        } else if (shape.imageFillRect) {
                            // Apply fill rectangle if specified
                            const fillRect = shape.imageFillRect;
                            drawX = bounds.x + (bounds.w * fillRect.l / 100000);
                            drawY = bounds.y + (bounds.h * fillRect.t / 100000);
                            drawW = bounds.w * (100000 - fillRect.l - fillRect.r) / 100000;
                            drawH = bounds.h * (100000 - fillRect.t - fillRect.b) / 100000;
                        }
                    } else if (shape.imageFillMode === 'tile') {
                        // Tile the image - simplified implementation
                        // For now, just draw once
                        drawX = bounds.x;
                        drawY = bounds.y;
                        drawW = bounds.w;
                        drawH = bounds.h;
                    } else {
                        // Maintain aspect ratio (fit or fill)
                        const imageAspect = naturalWidth / naturalHeight;
                        const shapeAspect = bounds.w / bounds.h;

                        if (imageAspect > shapeAspect) {
                            // Image is wider - fit to width
                            drawW = bounds.w;
                            drawH = bounds.w / imageAspect;
                            drawX = bounds.x;
                            drawY = bounds.y + (bounds.h - drawH) / 2;
                        } else {
                            // Image is taller - fit to height
                            drawH = bounds.h;
                            drawW = bounds.h * imageAspect;
                            drawX = bounds.x + (bounds.w - drawW) / 2;
                            drawY = bounds.y;
                        }
                    }

                    // Draw the image
                    ctx.drawImage(image, drawX, drawY, drawW, drawH);

                    ctx.restore();
                    return; // Successfully drew image, exit early

                } catch (error) {
                    // Fall through to placeholder
                }
            }
        }

        // No fallback text or labels for missing images

        ctx.restore();

        // Draw text content if present (text overlay on image)
        if (shape.textBody && shape.textBody.paragraphs && this.shouldRenderText(shape)) {
            // Set current shape context for inheritance
            this.setCurrentRenderingShape(shape);
            this.drawShapeText(shape, bounds);
        }
    }

    /**
     * Draw image placeholder with custom message
     */
    drawImagePlaceholder(shape, x, y, width, height, message = 'IMAGE') {
        const ctx = this.m_oContext;
        if (!ctx) {return;}
        
        ctx.save();
        
        // Draw placeholder rectangle with subtle styling
        ctx.fillStyle = '#f8f9fa';
        ctx.fillRect(x, y, width, height);
        
        // Draw border
        ctx.strokeStyle = '#dee2e6';
        ctx.lineWidth = 1;
        ctx.strokeRect(x, y, width, height);
        
        ctx.restore();
    }

    /**
     * Draw group shape
     */
    drawGroupShape(shape, bounds) {
        // Skip shapes with zero or invalid dimensions
        if (!bounds || bounds.w <= 0 || bounds.h <= 0) {
            return;
        }
        
        if (!this.graphics || !this.graphics.context) {return;}

        this.graphics.SaveGrState();

        try {
            // Apply group-level transformations if available
            if (shape.transform) {
                const ctx = this.graphics.context;

                // Apply rotation
                if (shape.transform.rotation && shape.transform.rotation !== 0) {
                    const centerX = bounds.x + bounds.w / 2;
                    const centerY = bounds.y + bounds.h / 2;
                    ctx.save();
                    ctx.translate(centerX, centerY);
                    ctx.rotate(shape.transform.rotation * Math.PI / 180);
                    ctx.translate(-centerX, -centerY);
                }

                // Apply flip transformations
                if (shape.transform.flipH || shape.transform.flipV) {
                    const scaleX = shape.transform.flipH ? -1 : 1;
                    const scaleY = shape.transform.flipV ? -1 : 1;
                    const centerX = bounds.x + bounds.w / 2;
                    const centerY = bounds.y + bounds.h / 2;

                    ctx.save();
                    ctx.translate(centerX, centerY);
                    ctx.scale(scaleX, scaleY);
                    ctx.translate(-centerX, -centerY);
                }
            }

            // DON'T draw group background or border - let child shapes be visible
            // Group shapes are logical containers and should not render visual boundaries

            // Draw child shapes
            if (shape.shapeTree && shape.shapeTree.length > 0) {
                // Sort child shapes by rendering order for proper z-index
                const sortedChildShapes = this.sortShapesByRenderOrder(shape.shapeTree);

                for (let i = 0; i < sortedChildShapes.length; i++) {
                    const childShape = sortedChildShapes[i];

                    try {
                        // Get child shape bounds directly from the child shape
                        const childBounds = this.getShapeBounds(childShape);

                        if (childBounds && childBounds.w > 0 && childBounds.h > 0) {
                            // Draw child shape directly using its own bounds
                            this.drawShapeGeometry(childShape, childBounds);
                        }

                    } catch (error) {
                        // Continue with other child shapes
                    }
                }
            }

        } catch (error) {
        } finally {
            // Always restore graphics state
            this.graphics.RestoreGrState();
        }
    }

    /**
     * Calculate child shape bounds within group - fallback method
     */
    calculateChildShapeBounds(childShape, groupBounds) {
        if (!childShape.bounds) {
            return { x: groupBounds.x, y: groupBounds.y, w: 50, h: 50 };
        }

        // Use transformed bounds if available
        const bounds = childShape.bounds;

        // Get slide dimensions for scaling
        const slideSize = this.processor ? this.processor.getSlideDimensions() : { cx: 9144000, cy: 6858000 };
        const canvasRect = this.processor ? this.processor.calculateCanvasRect(this.canvas, slideSize) : {
            scale: 1,
            offsetX: 0,
            offsetY: 0,
            slideWidthPx: this.canvas.width,
            slideHeightPx: this.canvas.height
        };

        // Convert EMU bounds to canvas pixels
        const slideWidthPx = canvasRect.slideWidthPx;
        const slideHeightPx = canvasRect.slideHeightPx;

        const shapeSlidePx = {
            x: ((bounds.l || 0) / slideSize.cx) * slideWidthPx,
            y: ((bounds.t || 0) / slideSize.cy) * slideHeightPx,
            w: (((bounds.r || 0) - (bounds.l || 0)) / slideSize.cx) * slideWidthPx,
            h: (((bounds.b || 0) - (bounds.t || 0)) / slideSize.cy) * slideHeightPx
        };

        // Apply canvas scaling and centering
        const x = canvasRect.offsetX + shapeSlidePx.x * canvasRect.scale;
        const y = canvasRect.offsetY + shapeSlidePx.y * canvasRect.scale;
        const w = shapeSlidePx.w * canvasRect.scale;
        const h = shapeSlidePx.h * canvasRect.scale;

        return { x, y, w, h };
    }

    /**
     * Draw default shape
     */
    drawDefaultShape(shape, bounds) {
        // Skip shapes with zero or invalid dimensions
        if (!bounds || bounds.w <= 0 || bounds.h <= 0) {
            return;
        }
        
        if (!this.graphics || !this.graphics.context) {return;}

        const ctx = this.graphics.context;
        ctx.save();

        // Apply rotation if present
        const transform = shape.properties?.transform;
        if (transform && transform.rotation && transform.rotation !== 0) {
            const centerX = bounds.x + bounds.w / 2;
            const centerY = bounds.y + bounds.h / 2;
            ctx.translate(centerX, centerY);
            ctx.rotate(transform.rotation * Math.PI / 180);
            ctx.translate(-centerX, -centerY);
        }

        // Draw default shape background
        ctx.fillStyle = '#e0e0e0';
        ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);

        // Draw border
        ctx.strokeStyle = '#999';
        ctx.lineWidth = 1;
        ctx.strokeRect(bounds.x, bounds.y, bounds.w, bounds.h);

        ctx.restore();

        // Draw text content if present
        if (shape.textBody && shape.textBody.paragraphs) {
            // Set current shape context for inheritance
            this.setCurrentRenderingShape(shape);
            this.drawShapeText(shape, bounds);
        }
    }

    /**
     * Draw connector shape
     */
    drawConnectorShape(shape, bounds) {
        // Skip shapes with zero or invalid dimensions
        if (!bounds || (bounds.w <= 0 && bounds.h <= 0)) {return;}

        // Draw connector as line with stroke/arrow styling
        this.drawEngineLine(bounds, shape);
    }

    /**
     * Draw graphic frame
     */
    async drawGraphicFrame(shape, bounds) {
        if (shape.graphicData) {
        }
        
        // IMPORTANT: Do not replace pixel bounds with EMU bounds for tables.
        // Pixel-space `bounds` is already computed via the slide coordinate system
        // in `getShapeBounds`. Using `shape.bounds` (EMU) here shifts tables to (0,0).
        
        // Skip shapes with zero or invalid dimensions
        if (!bounds || bounds.w <= 0 || bounds.h <= 0) {
            return;
        }
        
        // Handle tables specifically
        if (shape.graphicData && shape.graphicData.uri === 'http://schemas.openxmlformats.org/drawingml/2006/table') {
            this.drawTable(shape, bounds);
        } 
        // Handle charts specifically
        else if (shape.graphicData && shape.graphicData.uri === 'http://schemas.openxmlformats.org/drawingml/2006/chart') {
            await this.drawChart(shape, bounds);
        } 
        else {
            // For other graphic frames, use default shape method
            this.drawDefaultShape(shape, bounds);
        }
    }

    /**
     * Draw table in graphics adapter
     */
    drawTable(shape, bounds) {
        // Try to parse and render the table
        if (shape.graphicData && shape.graphicData.tableXml) {
            try {
                // Use the comprehensive table processor instead of simple parser
                if (!this.tableProcessor) {
                    this.tableProcessor = new TableProcessor();
                }
                const table = this.tableProcessor.parseTableFromXML(shape.graphicData.tableXml);
                if (table) {
                    // Compute natural table size from grid/rows and layout properties (autofit support)
                    let x = bounds.x;
                    let y = bounds.y;
                    let width = bounds.w;
                    let height = bounds.h;

                    try {
                        const props = (typeof table.getTableProperties === 'function') ? (table.getTableProperties() || {}) : (table.tableProperties || {});
                        const layoutType = (props && props.layout) ? String(props.layout).toLowerCase() : 'autofit';

                        // Natural width from tblGrid (EMU → px) - use actual column totals
                        const csScale = (this.coordinateSystem && typeof this.coordinateSystem.scale === 'number') ? this.coordinateSystem.scale : 1;
                        let naturalWidthPx = 0;
                        const grid = (typeof table.getTableGrid === 'function') ? table.getTableGrid() : (table.tableGrid || []);
                        if (Array.isArray(grid) && grid.length > 0) {
                            const totalEmu = grid.reduce((s, c) => s + (parseInt(c.width) || 0), 0);
                            // Convert to base px, then scale to canvas space using slide scale
                            naturalWidthPx = ((totalEmu / 914400) * 96) * csScale;
                        }

                        // Natural height from row heights (EMU → px) when available
                        let naturalHeightPx = 0;
                        const rows = table.rows || [];
                        if (rows.length > 0 && rows.some(r => r && r.height)) {
                            const totalEmuH = rows.reduce((s, r) => s + (parseInt(r && r.height) || 0), 0);
                            naturalHeightPx = ((totalEmuH / 914400) * 96) * csScale;
                        }

                        // Preferred width from tblW if provided (pct/dxa/auto)
                        if (!naturalWidthPx && props && props.preferredWidth) {
                            const pw = props.preferredWidth;
                            if (pw.type === 'pct') {
                                const pct = (parseInt(pw.value) || 0) / 5000; // 5000 = 100%
                                naturalWidthPx = Math.max(0, bounds.w * pct);
                            } else if (pw.type === 'dxa') {
                                // dxa (twips): 1 twip = 1/1440 inch
                                naturalWidthPx = (((parseInt(pw.value) || 0) / 1440) * 96) * csScale;
                            }
                        }

                        // Use natural width unless the frame width is within 10% of natural
                        // (indicating the user intentionally resized the table to fill the frame).
                        // When natural >> frame the frame is a placeholder, not a layout constraint.
                        const fillRatio = naturalWidthPx > 0 ? naturalWidthPx / bounds.w : 0;
                        if (naturalWidthPx > bounds.w || fillRatio < 0.9) {
                            width = naturalWidthPx;
                        }
                        if (naturalHeightPx > 0) {
                            height = naturalHeightPx;
                        }
                    } catch(_e) {}
                    
                    // PPTX SPECIFICATION: Apply proper clipping to table content
                    // This ensures text doesn't render outside the defined table bounds
                    const tableCtx = this.canvas.getContext('2d');

                    let layoutMetrics = null;
                    try {
                        layoutMetrics = this.computeTableLayout(table, width, height, tableCtx);
                        if (layoutMetrics && typeof layoutMetrics.height === 'number' && layoutMetrics.height > 0) {
                            height = layoutMetrics.height;
                        }
                    } catch (_layoutError) {
                        layoutMetrics = null;
                    }

                    tableCtx.save();
                    // Use computed layout height (sum of actual row heights) not the graphicFrame bounds.h,
                    // because PowerPoint renders all rows even when they overflow the frame.
                    tableCtx.rect(x, y, width, height);
                    tableCtx.clip();
                    
                    
                    // Always render the table using the graphicFrame bounds
                    // The table should scale to fit the entire graphicFrame container
                    
                    // Get canvas context for direct drawing
                    const ctx = this.graphics.context;
                    if (!ctx) {
                        return;
                    }
                    
                    // Create a simple graphics wrapper for table drawing
                    const self = this;
                    const tableGraphics = {
                        strokeRect: (x, y, w, h, color, lineWidth) => {
                            if (self.graphics && typeof self.graphics.p_color === 'function') {
                                self.graphics.p_color(color.r, color.g, color.b, color.a || 255);
                                self.graphics.p_width(lineWidth || 1);
                                if (typeof self.graphics._s === 'function') {self.graphics._s();}
                                if (typeof self.graphics._m === 'function') {self.graphics._m(x, y);}
                                if (typeof self.graphics._l === 'function') {
                                    self.graphics._l(x + w, y);
                                    self.graphics._l(x + w, y + h);
                                    self.graphics._l(x, y + h);
                                }
                                if (typeof self.graphics._z === 'function') {self.graphics._z();}
                                if (typeof self.graphics.ds === 'function') {self.graphics.ds();}
                            }
                        },
                        fillRect: (x, y, w, h, color) => {
                            if (self.graphics && typeof self.graphics.b_color1 === 'function') {
                                self.graphics.b_color1(color.r, color.g, color.b, color.a || 255);
                                if (typeof self.graphics._s === 'function') {self.graphics._s();}
                                if (typeof self.graphics._m === 'function') {self.graphics._m(x, y);}
                                if (typeof self.graphics._l === 'function') {
                                    self.graphics._l(x + w, y);
                                    self.graphics._l(x + w, y + h);
                                    self.graphics._l(x, y + h);
                                }
                                if (typeof self.graphics._z === 'function') {self.graphics._z();}
                                if (typeof self.graphics.df === 'function') {self.graphics.df();}
                            }
                        },
                        drawLine: (x1, y1, x2, y2, color, lineWidth) => {
                            if (self.graphics && typeof self.graphics.p_color === 'function') {
                                self.graphics.p_color(color.r, color.g, color.b, color.a || 255);
                                self.graphics.p_width(lineWidth || 1);
                                if (typeof self.graphics._s === 'function') {self.graphics._s();}
                                if (typeof self.graphics._m === 'function') {self.graphics._m(x1, y1);}
                                if (typeof self.graphics._l === 'function') {self.graphics._l(x2, y2);}
                                if (typeof self.graphics.ds === 'function') {self.graphics.ds();}
                            }
                        },
                        fillText: (text, x, y, properties) => {
                            if (self.graphics && typeof self.graphics.FillText === 'function') {
                                if (typeof self.graphics.font === 'function') {
                                    self.graphics.font(properties.fontFamily || 'Arial', properties.fontSize || 12, 
                                                     properties.bold ? 'bold' : 'normal', 
                                                     properties.italic ? 'italic' : 'normal');
                                }
                                self.graphics.b_color1(properties.color.r, properties.color.g, properties.color.b, properties.color.a || 255);
                                self.graphics.FillText(x, y, text);
                            }
                        }
                    };
                    
                    // Try direct canvas drawing with the full graphicFrame bounds
                    this.drawTableDirect(table, tableCtx, x, y, width, height, layoutMetrics);
                    
                    // PPTX SPECIFICATION: Restore clipping after table is rendered
                    tableCtx.restore();
                } else {
                    this.drawTablePlaceholder(bounds, 'Parse Failed');
                    tableCtx.restore(); // Restore clipping for error case
                }
            } catch (error) {
                this.drawTablePlaceholder(bounds, 'Error');
                tableCtx.restore(); // Restore clipping for error case
            }
        } else {
            this.drawTablePlaceholder(bounds, 'No Data');
        }
    }

    /**
     * Compute table layout metrics (column widths, row heights, final block height).
     */
    computeTableLayout(table, width, height, ctx) {
        const layout = { cellWidths: [], cellHeights: [], height };
        if (!table) {return layout;}

        const rowCount = (typeof table.getRowCount === 'function') ? table.getRowCount() : (table.rows ? table.rows.length : 0);
        const baseColCount = (typeof table.getColumnCount === 'function') ? table.getColumnCount() : 0;
        if (rowCount <= 0) {return layout;}

        // csScale factor for consistent coordinate system conversion (EMU → canvas pixels)
        const csScale = (this.coordinateSystem && typeof this.coordinateSystem.scale === 'number') ? this.coordinateSystem.scale : 1;

        let cellWidths;
        try {
            const grid = (typeof table.getTableGrid === 'function') ? table.getTableGrid() : (table.tableGrid || []);
            if (Array.isArray(grid) && grid.length > 0) {
                cellWidths = grid.map(col => {
                    const wEmu = parseInt(col.width, 10) || 0;
                    return ((wEmu / 914400) * 96) * csScale;
                });

                const naturalTotal = cellWidths.reduce((sum, value) => sum + value, 0);
                if (naturalTotal > 0 && isFinite(width) && width > 0 && Math.abs(naturalTotal - width) > 1) {
                    // Scale columns proportionally to fit within the available width
                    const scale = width / naturalTotal;
                    cellWidths = cellWidths.map(value => value * scale);
                }

                const colCount = Math.max(baseColCount, cellWidths.length);
                if (cellWidths.length < colCount && colCount > 0) {
                    const remaining = colCount - cellWidths.length;
                    const used = cellWidths.reduce((sum, value) => sum + value, 0);
                    const pad = remaining > 0 ? (width - used) / remaining : 0;
                    for (let i = 0; i < remaining; i++) {cellWidths.push(pad);}
                }
            }
        } catch (_e) {
            cellWidths = undefined;
        }

        const resolvedColCount = Math.max(baseColCount, Array.isArray(cellWidths) ? cellWidths.length : 0);
        if (!Array.isArray(cellWidths) || cellWidths.length === 0) {
            const equalWidth = width / Math.max(resolvedColCount || 1, 1);
            cellWidths = Array.from({ length: Math.max(resolvedColCount, 1) }, () => equalWidth);
        }

        const cellHeights = new Array(rowCount);
        const autoRows = new Array(rowCount).fill(false);
        const fallbackRowHeight = (isFinite(height) && height > 0)
            ? (height / Math.max(rowCount, 1))
            : 0;

        for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            const row = table.rows ? table.rows[rowIndex] : null;
            let resolvedHeight = null;
            if (row && row.height !== undefined && row.height !== null) {
                const parsed = parseInt(row.height, 10);
                if (!Number.isNaN(parsed) && parsed > 0) {
                    resolvedHeight = ((parsed / 914400) * 96) * csScale;
                }
            }

            if (resolvedHeight && resolvedHeight > 0) {
                cellHeights[rowIndex] = resolvedHeight;
            } else {
                cellHeights[rowIndex] = 0;
                autoRows[rowIndex] = true;
            }
        }

        const sumSegment = (arr, start, count) => {
            let total = 0;
            for (let i = 0; i < count; i++) {total += arr[start + i] || 0;}
            return total;
        };

        if (ctx && typeof ctx.measureText === 'function') {
            const rowAdjustments = new Array(rowCount).fill(0);
            const measurementCtx = ctx;
            if (typeof measurementCtx.save === 'function') {measurementCtx.save();}
            try {
                measurementCtx.textBaseline = 'top';
                for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                    const row = table.rows ? table.rows[rowIndex] : null;
                    if (!row || !Array.isArray(row.cells)) {continue;}

                    let logicalColIndex = 0;
                    for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex++) {
                        const cell = row.cells[cellIndex];
                        const gridSpan = Math.max(1, cell?.gridSpan || 1);
                        const rowSpan = Math.max(1, cell?.rowSpan || 1);

                        if (cell && cell.isMergedContinue) {
                            if (cell.vMerge === 'continue') {logicalColIndex += 1;}
                            continue;
                        }

                        const cellWidth = sumSegment(cellWidths, logicalColIndex, gridSpan);
                        const margins = this.getCellMarginsPx(cell);
                        const contentWidth = Math.max(0, cellWidth - margins.left - margins.right);
                        if (contentWidth <= 0) {
                            logicalColIndex += gridSpan;
                            continue;
                        }

                        const layoutInfo = this.computeCellTextLayout(measurementCtx, cell, contentWidth, Number.POSITIVE_INFINITY, { contentX: 0 });
                        const requiredHeight = layoutInfo.totalHeight + margins.top + margins.bottom;
                        if (requiredHeight <= 0) {
                            logicalColIndex += gridSpan;
                            continue;
                        }

                        const currentSpanHeight = sumSegment(cellHeights, rowIndex, rowSpan);
                        if (requiredHeight > currentSpanHeight) {
                            const extra = requiredHeight - currentSpanHeight;
                            if (rowSpan === 1) {
                                rowAdjustments[rowIndex] = Math.max(rowAdjustments[rowIndex], extra);
                            } else {
                                const perRow = extra / rowSpan;
                                for (let spanOffset = 0; spanOffset < rowSpan; spanOffset++) {
                                    const targetRow = rowIndex + spanOffset;
                                    if (targetRow >= rowCount) {break;}
                                    rowAdjustments[targetRow] = Math.max(rowAdjustments[targetRow], perRow);
                                }
                            }
                        }

                        logicalColIndex += gridSpan;
                    }
                }
            } finally {
                if (typeof measurementCtx.restore === 'function') {measurementCtx.restore();}
            }

            for (let i = 0; i < rowCount; i++) {
                if (rowAdjustments[i] > 0) {
                    cellHeights[i] += rowAdjustments[i];
                }
            }
        }

        for (let i = 0; i < rowCount; i++) {
            if ((cellHeights[i] === undefined || cellHeights[i] <= 0) && autoRows[i]) {
                cellHeights[i] = fallbackRowHeight || 0;
            }
            if (cellHeights[i] <= 0) {
                cellHeights[i] = 0;
            }
        }

        const totalHeight = cellHeights.reduce((sum, value) => sum + (value || 0), 0);
        layout.cellWidths = cellWidths;
        layout.cellHeights = cellHeights;
        if (totalHeight > 0) {
            layout.height = totalHeight;
        } else if (isFinite(height) && height > 0) {
            layout.height = height;
        } else {
            layout.height = 0;
        }

        return layout;
    }

    /**
     * Draw table directly to canvas context
     */
    drawTableDirect(table, ctx, x, y, width, height, layoutMetrics = null) {
        if (!table || !ctx) {return;}

        ctx.save();

        const rowCount = (typeof table.getRowCount === 'function') ? table.getRowCount() : (table.rows ? table.rows.length : 0);
        const colCount = (typeof table.getColumnCount === 'function') ? table.getColumnCount() : 0;
        if (rowCount === 0 || (colCount === 0 && !(table.rows && table.rows[0] && table.rows[0].cells))) {
            ctx.restore();
            return;
        }

        let metrics = layoutMetrics;
        if (!metrics) {
            metrics = this.computeTableLayout(table, width, height, ctx);
        }

        let cellWidths = Array.isArray(metrics?.cellWidths) ? metrics.cellWidths.slice() : null;
        const effectiveColCount = cellWidths && cellWidths.length > 0 ? cellWidths.length : Math.max(colCount, 1);
        if (!cellWidths || cellWidths.length === 0) {
            cellWidths = Array.from({ length: effectiveColCount }, () => width / Math.max(effectiveColCount, 1));
        }

        let cellHeights = Array.isArray(metrics?.cellHeights) ? metrics.cellHeights.slice() : null;
        if (!cellHeights || cellHeights.length === 0) {
            cellHeights = Array.from({ length: rowCount }, () => height / Math.max(rowCount, 1));
        }

        if (metrics && typeof metrics.height === 'number' && metrics.height > 0) {
            height = metrics.height;
        }

        const sum = (arr, start, count) => {
            let total = 0;
            for (let i = 0; i < count; i++) {total += arr[start + i] || 0;}
            return total;
        };
        const offset = (arr, count) => {
            let total = 0;
            for (let i = 0; i < count; i++) {total += arr[i] || 0;}
            return total;
        };

        // FIXED: Don't draw default outer border - let table borders be defined by PPTX
        // ctx.strokeStyle = '#000000';
        // ctx.lineWidth = 1;
        // ctx.strokeRect(x, y, width, height);

        // Retrieve table-level borders (insideH/insideV)
        let tableBorders = null;
        try {
            const props = (typeof table.getTableProperties === 'function') ? table.getTableProperties() : (table.tableProperties || {});
            tableBorders = props && props.borders ? props.borders : null;
        } catch(_e) { tableBorders = null; }

        // Iterate rows/cells respecting merged spans
        let currentY = y;
        const rowBottomYs = [];
        for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            const rowHeight = cellHeights[rowIndex];
            const row = table.rows[rowIndex];
            if (!row) {continue;}

            let logicalColIndex = 0;
            for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex++) {
                const cell = row.cells[cellIndex];
                const gridSpan = Math.max(1, cell?.gridSpan || 1);
                const rowSpan = Math.max(1, cell?.rowSpan || 1);

                // Skip merged continuation cells
                if (cell && cell.isMergedContinue) {
                    // Only vertical continuation occupies a column in this row
                    if (cell.vMerge === 'continue') {
                        logicalColIndex += 1;
                    }
                    continue;
                }

                const cellX = x + offset(cellWidths, logicalColIndex);
                const cellY = currentY;
                const cellW = sum(cellWidths, logicalColIndex, gridSpan);
                const cellH = sum(cellHeights, rowIndex, rowSpan);

                // Background - only draw if explicitly specified (no default white)
                // Pass rowIndex and table for table style support
                const cellStyle = this.extractCellStyle(cell, rowIndex, table);
                if (cellStyle.backgroundColor) {
                    ctx.fillStyle = cellStyle.backgroundColor;
                    ctx.fillRect(cellX, cellY, cellW, cellH);
                }
                // No background drawn if not specified - keeps transparent cells transparent

                // FIXED: Always use PPTX-specific border drawing - no default borders
                this.drawCellBordersFromPPTX(ctx, cell, cellX, cellY, cellW, cellH);

                // Text
                if (cell && cell.textBody) {
                    this.drawCellText(ctx, cell, cellX, cellY, cellW, cellH, cellStyle);
                }

                logicalColIndex += gridSpan;
            }
            currentY += rowHeight;
            rowBottomYs.push(currentY);
        }

        // Scale table-level border widths from 96dpi to actual canvas scale
        const tblCsScale = (this.coordinateSystem && typeof this.coordinateSystem.scale === 'number') ? this.coordinateSystem.scale : 1;
        const scaledBorderWidth = (b) => Math.max(1, Math.round((b.width || 1) * tblCsScale));

        // Draw table-level outer borders and interior gridlines
        if (tableBorders) {
            // Draw outer top border
            if (tableBorders.top && tableBorders.top.color) {
                const tb = tableBorders.top;
                ctx.strokeStyle = `rgb(${tb.color.r}, ${tb.color.g}, ${tb.color.b})`;
                ctx.lineWidth = scaledBorderWidth(tb);
                ctx.beginPath();
                ctx.moveTo(x, y);
                ctx.lineTo(x + width, y);
                ctx.stroke();
            }

            // Draw outer left border
            if (tableBorders.left && tableBorders.left.color) {
                const lb = tableBorders.left;
                ctx.strokeStyle = `rgb(${lb.color.r}, ${lb.color.g}, ${lb.color.b})`;
                ctx.lineWidth = scaledBorderWidth(lb);
                ctx.beginPath();
                ctx.moveTo(x, y);
                ctx.lineTo(x, y + height);
                ctx.stroke();
            }

            // Draw outer right border
            if (tableBorders.right && tableBorders.right.color) {
                const rb = tableBorders.right;
                ctx.strokeStyle = `rgb(${rb.color.r}, ${rb.color.g}, ${rb.color.b})`;
                ctx.lineWidth = scaledBorderWidth(rb);
                ctx.beginPath();
                ctx.moveTo(x + width, y);
                ctx.lineTo(x + width, y + height);
                ctx.stroke();
            }

            // Draw outer bottom border
            if (tableBorders.bottom && tableBorders.bottom.color) {
                const bb = tableBorders.bottom;
                ctx.strokeStyle = `rgb(${bb.color.r}, ${bb.color.g}, ${bb.color.b})`;
                ctx.lineWidth = scaledBorderWidth(bb);
                ctx.beginPath();
                ctx.moveTo(x, y + height);
                ctx.lineTo(x + width, y + height);
                ctx.stroke();
            }

            // Draw internal horizontal gridlines (insideH)
            if (tableBorders.insideH && tableBorders.insideH.color) {
                const hb = tableBorders.insideH;
                ctx.strokeStyle = `rgb(${hb.color.r}, ${hb.color.g}, ${hb.color.b})`;
                ctx.lineWidth = scaledBorderWidth(hb);
                for (let i = 0; i < rowBottomYs.length - 1; i++) {
                    const yLine = rowBottomYs[i];
                    ctx.beginPath();
                    ctx.moveTo(x, yLine);
                    ctx.lineTo(x + width, yLine);
                    ctx.stroke();
                }
            }

            // Draw internal vertical gridlines (insideV)
            if (tableBorders.insideV && tableBorders.insideV.color) {
                const vb = tableBorders.insideV;
                ctx.strokeStyle = `rgb(${vb.color.r}, ${vb.color.g}, ${vb.color.b})`;
                ctx.lineWidth = scaledBorderWidth(vb);
                let colX = x;
                for (let c = 0; c < cellWidths.length - 1; c++) {
                    colX += cellWidths[c];
                    ctx.beginPath();
                    ctx.moveTo(colX, y);
                    ctx.lineTo(colX, y + height);
                    ctx.stroke();
                }
            }
        }

        ctx.restore();
    }

    /**
     * Draw cell borders based on PPTX border definitions
     */
    drawCellBordersFromPPTX(ctx, cell, x, y, w, h) {
        if (!cell || !cell.borders) {return;}

        const borders = cell.borders;
        // Scale border widths from 96dpi to actual canvas scale
        const csScale = (this.coordinateSystem && typeof this.coordinateSystem.scale === 'number') ? this.coordinateSystem.scale : 1;
        const borderSides = [
            { name: 'top', x1: x, y1: y, x2: x + w, y2: y },
            { name: 'right', x1: x + w, y1: y, x2: x + w, y2: y + h },
            { name: 'bottom', x1: x, y1: y + h, x2: x + w, y2: y + h },
            { name: 'left', x1: x, y1: y, x2: x, y2: y + h }
        ];

        borderSides.forEach(side => {
            const border = borders[side.name];
            // Only draw border if it exists, has a color, and width > 0
            if (border && border.color && border.width && border.width > 0) {
                ctx.strokeStyle = `rgb(${border.color.r}, ${border.color.g}, ${border.color.b})`;
                ctx.lineWidth = Math.max(1, Math.round(border.width * csScale));
                ctx.beginPath();
                ctx.moveTo(side.x1, side.y1);
                ctx.lineTo(side.x2, side.y2);
                ctx.stroke();
            }
        });
    }
    /**
     * Extract cell styling information - Enhanced with table style support
     */
    extractCellStyle(cell, rowIndex, table) {
        const style = {};
        
        if (!cell) {return style;}
        
        
        // Extract background color - only if explicitly specified
        if (cell.shading && cell.shading.fillColor && 
            cell.shading.fillColor.r !== undefined && 
            cell.shading.fillColor.g !== undefined && 
            cell.shading.fillColor.b !== undefined) {
            const color = cell.shading.fillColor;
            style.backgroundColor = `rgba(${color.r}, ${color.g}, ${color.b}, ${(color.a || 255) / 255})`;
        } else if (cell.shading && cell.shading.fill && cell.shading.fill !== 'none') {
            // Only accept explicit hex color; ignore scheme or unknown tokens to avoid black boxes
            const hex = cell.shading.fill.startsWith('#') ? cell.shading.fill.slice(1) : null;
            if (hex && hex.length === 6) {
                const r = parseInt(hex.substring(0, 2), 16);
                const g = parseInt(hex.substring(2, 4), 16);
                const b = parseInt(hex.substring(4, 6), 16);
                style.backgroundColor = `rgb(${r}, ${g}, ${b})`;
            }
        }
        
        // ENHANCED: Apply table style formatting if no explicit cell background and table is provided
        if (!style.backgroundColor && table && rowIndex !== undefined) {
            const tableProps = (typeof table.getTableProperties === 'function') ? 
                table.getTableProperties() : (table.tableProperties || {});
            
            if (tableProps) {
                const colIndex = cell.logicalColIndex || 0;
                
                // firstRow - header row gets blue background
                if (tableProps.firstRow && rowIndex === 0) {
                    style.backgroundColor = 'rgba(68, 114, 196, 1)'; // Default blue header
                }
                // bandRow - alternating row colors
                else if (tableProps.bandRow && rowIndex > 0) {
                    const effectiveRow = tableProps.firstRow ? rowIndex - 1 : rowIndex;
                    if (effectiveRow % 2 === 1) {
                        style.backgroundColor = 'rgba(242, 242, 242, 1)'; // Light gray
                    }
                }
                // firstCol - first column gets special formatting
                else if (tableProps.firstCol && colIndex === 0 && rowIndex > 0) {
                    style.backgroundColor = 'rgba(221, 235, 247, 1)'; // Light blue
                }
            }
        }
        // No default background color - transparent cells remain transparent
        
        // FIXED: Don't extract general border styling - borders are handled individually by drawCellBordersFromPPTX
        
        // Extract text color from first run if available, apply table style defaults
        if (cell.textBody && cell.textBody.paragraphs && cell.textBody.paragraphs.length > 0) {
            const firstParagraph = cell.textBody.paragraphs[0];
            if (firstParagraph.runs && firstParagraph.runs.length > 0) {
                const firstRun = firstParagraph.runs[0];
                if (firstRun.properties && firstRun.properties.color) {
                    const color = firstRun.properties.color;
                    style.textColor = `rgba(${color.r}, ${color.g}, ${color.b}, ${(color.a || 255) / 255})`;
                } else {
                    // Apply table style-based text color when no explicit color
                    style.textColor = this.getTableStyleTextColor(rowIndex, table);
                }
            } else {
                style.textColor = this.getTableStyleTextColor(rowIndex, table);
            }
        } else {
            style.textColor = this.getTableStyleTextColor(rowIndex, table);
        }
        
        return style;
    }

    /**
     * Get table style-based text color
     * Returns appropriate text color based on table style and row position
     */
    getTableStyleTextColor(rowIndex, table) {
        // Get table properties
        const tableProps = (table && typeof table.getTableProperties === 'function') ? 
            table.getTableProperties() : (table && table.tableProperties ? table.tableProperties : {});
        
        if (tableProps && rowIndex !== undefined) {
            // firstRow - header row gets white text
            if (tableProps.firstRow && rowIndex === 0) {
                return 'rgb(255, 255, 255)'; // White text for header
            }
        }
        
        // Default: black text for body rows (standard table text)
        return 'rgb(0, 0, 0)';
    }

    /**
     * Calculate effective cell margins in pixels (defaults to 1 mm each side).
     */
    getCellMarginsPx(cell) {
        const csScale = (this.coordinateSystem && typeof this.coordinateSystem.scale === 'number') ? this.coordinateSystem.scale : 1;
        const defaultMargin = 3.78 * csScale;
        const defaults = { left: defaultMargin, top: defaultMargin, right: defaultMargin, bottom: defaultMargin };
        if (!cell) {return defaults;}

        const result = { ...defaults };
        const convert = value => {
            const numeric = Number(value);
            if (!Number.isFinite(numeric)) {return null;}
            return ((numeric / 914400) * 96) * csScale;
        };

        const applyLegacy = source => {
            if (!source) {return;}
            if (source.marL !== undefined) {
                const converted = convert(source.marL);
                if (converted !== null) {result.left = converted;}
            }
            if (source.marT !== undefined) {
                const converted = convert(source.marT);
                if (converted !== null) {result.top = converted;}
            }
            if (source.marR !== undefined) {
                const converted = convert(source.marR);
                if (converted !== null) {result.right = converted;}
            }
            if (source.marB !== undefined) {
                const converted = convert(source.marB);
                if (converted !== null) {result.bottom = converted;}
            }
        };

        const applyAlt = source => {
            if (!source) {return;}
            if (source.left !== undefined) {
                const converted = convert(source.left);
                if (converted !== null) {result.left = converted;}
            }
            if (source.top !== undefined) {
                const converted = convert(source.top);
                if (converted !== null) {result.top = converted;}
            }
            if (source.right !== undefined) {
                const converted = convert(source.right);
                if (converted !== null) {result.right = converted;}
            }
            if (source.bottom !== undefined) {
                const converted = convert(source.bottom);
                if (converted !== null) {result.bottom = converted;}
            }
        };

        applyLegacy(cell.cellProperties);
        applyAlt(cell.margins);

        return result;
    }

    /**
     * Build wrapped text layout for a table cell so we can measure height and render consistently.
     */
    computeCellTextLayout(ctx, cell, contentWidth, heightLimit = Number.POSITIVE_INFINITY, options = {}) {
        const layout = { lines: [], totalHeight: 0 };
        if (!cell || !cell.textBody || !Array.isArray(cell.textBody.paragraphs) || contentWidth <= 0) {
            return layout;
        }

        const effectiveLimit = (typeof heightLimit === 'number' && heightLimit > 0)
            ? heightLimit : Number.POSITIVE_INFINITY;
        const baseX = (options && typeof options.contentX === 'number') ? options.contentX : 0;
        let usedHeight = 0;

        // Scale font sizes to match canvas coordinate system (same as column widths)
        const csScale = (this.coordinateSystem && typeof this.coordinateSystem.scale === 'number') ? this.coordinateSystem.scale : 1;

        const wrapDisabled = Boolean(cell.textBody && cell.textBody.bodyProperties && cell.textBody.bodyProperties.wrap === false);
        const originalFont = ctx ? ctx.font : null;

        for (const paragraph of cell.textBody.paragraphs) {
            if (!paragraph || !Array.isArray(paragraph.runs) || paragraph.runs.length === 0) {continue;}
            if (usedHeight >= effectiveLimit) {break;}

            let textAlign = 'left';
            let textX = baseX;
            if (paragraph.properties && paragraph.properties.alignment) {
                const alignRaw = paragraph.properties.alignment;
                const align = typeof alignRaw === 'string' ? alignRaw.toLowerCase() : alignRaw;
                if (align === 'center' || align === 'ctr') {
                    textAlign = 'center';
                    textX = baseX + (contentWidth / 2);
                } else if (align === 'right' || align === 'r' || align === 'end') {
                    textAlign = 'right';
                    textX = baseX + contentWidth;
                } else if (align === 'start' || align === 'left' || align === 'l') {
                    textAlign = 'left';
                    textX = baseX;
                } else if (align === 'justify') {
                    textAlign = 'left';
                    textX = baseX;
                }
            }

            const firstRun = paragraph.runs.find(r => (r.text || '').trim().length > 0) || paragraph.runs[0];
            let fontSize = 12;
            if (firstRun?.properties?.fontSize) {
                fontSize = firstRun.properties.fontSize;
            } else if (firstRun?.properties?.sz) {
                const parsedSz = parseInt(firstRun.properties.sz, 10);
                if (!Number.isNaN(parsedSz)) {fontSize = parsedSz / 100;}
            }
            const fontSizePixels = Math.max(1, Math.round(fontSize * (96 / 72) * csScale));
            const fontFamily = firstRun?.properties?.fontFamily || 'Arial';
            const isBold = Boolean(firstRun?.properties?.bold);
            const isItalic = Boolean(firstRun?.properties?.italic);
            const isUnderline = Boolean(firstRun?.properties?.underline);
            
            // Use color from options.defaultTextColor if available (table style colors), otherwise from run properties
            let colorStr = options.defaultTextColor || 'rgb(159, 159, 159)';
            if (firstRun?.properties?.color) {
                const c = firstRun.properties.color;
                colorStr = `rgba(${c.r}, ${c.g}, ${c.b}, ${(c.a || 255) / 255})`;
            }

            // Default line spacing: 1.2x matches PowerPoint's default for table cells
            // (OpenXML "normal" line spacing is implementation-defined; PowerPoint uses ~1.2x)
            let lineSpacingFactor = 1.2;
            if (paragraph.properties) {
                if (typeof paragraph.properties.lineSpacingFactor === 'number') {
                    lineSpacingFactor = paragraph.properties.lineSpacingFactor;
                } else if (typeof paragraph.properties.lineSpacingPct === 'number') {
                    lineSpacingFactor = Math.max(0.5, Math.min(4, paragraph.properties.lineSpacingPct / 100));
                }
            }
            const lineHeight = Math.max(1, Math.round(fontSizePixels * lineSpacingFactor));
            const fullText = paragraph.runs.map(r => r.text || '').join('');
            if (!fullText) {continue;}

            const wrapEnabled = !wrapDisabled;
            const measurementFont = `${isItalic ? 'italic ' : ''}${isBold ? 'bold ' : ''}${fontSizePixels}px ${fontFamily}`;
            if (ctx && typeof ctx.font === 'string') {
                ctx.font = measurementFont;
            }

            // Get highlight from first run if present
            let highlightStr = null;
            if (firstRun?.properties?.highlight) {
                const hl = firstRun.properties.highlight;
                if (typeof hl === 'object' && hl.r !== undefined) {
                    highlightStr = `rgba(${hl.r}, ${hl.g}, ${hl.b}, ${(hl.a || 255) / 255})`;
                } else if (typeof hl === 'string') {
                    highlightStr = hl;
                }
            }

            const pushLine = textValue => {
                if (!textValue) {return;}
                layout.lines.push({
                    text: textValue,
                    textX,
                    textAlign,
                    fontSizePixels,
                    fontFamily,
                    isBold,
                    isItalic,
                    isUnderline,
                    colorStr,
                    lineHeight,
                    highlightStr
                });
                usedHeight += lineHeight;
            };

            if (wrapEnabled) {
                let remaining = fullText.trim();
                while (remaining.length > 0) {
                    if (usedHeight + lineHeight > effectiveLimit) {break;}
                    let low = 1;
                    let high = remaining.length;
                    let fit = 1;
                    while (low <= high) {
                        const mid = (low + high) >> 1;
                        const substr = remaining.slice(0, mid);
                        const metrics = ctx && typeof ctx.measureText === 'function'
                            ? ctx.measureText(substr)
                            : { width: substr.length * fontSizePixels * 0.5 };
                        if (metrics.width <= contentWidth || mid === 1) {
                            fit = mid;
                            low = mid + 1;
                        } else {
                            high = mid - 1;
                        }
                    }
                    let breakPos = fit;
                    if (fit < remaining.length) {
                        const spacePos = remaining.slice(0, fit).lastIndexOf(' ');
                        if (spacePos > 0) {breakPos = spacePos;}
                    }
                    let sliceEnd = Math.max(1, breakPos);
                    let lineText = remaining.slice(0, sliceEnd).trimEnd();
                    if (!lineText) {
                        lineText = remaining.charAt(0);
                        sliceEnd = 1;
                    }
                    pushLine(lineText);
                    remaining = remaining.slice(sliceEnd).trimStart();
                }
            } else {
                if (usedHeight + lineHeight > effectiveLimit) {break;}
                pushLine(fullText.trimEnd());
            }
        }

        if (ctx && originalFont !== null) {
            ctx.font = originalFont;
        }

        layout.totalHeight = usedHeight;
        return layout;
    }

    /**
     * Draw cell text content
     */
    drawCellText(ctx, cell, x, y, width, height, cellStyle) {
        if (!cell || !cell.textBody || !cell.textBody.paragraphs) {return;}

        ctx.save();
        ctx.textBaseline = 'top';

        const margins = this.getCellMarginsPx(cell);
        const contentWidth = Math.max(0, width - margins.left - margins.right);
        const contentHeight = Math.max(0, height - margins.top - margins.bottom);
        if (contentWidth <= 0 || contentHeight <= 0) {
            ctx.restore();
            return;
        }

        const contentX = x + margins.left;
        const contentY = y + margins.top;

        ctx.beginPath();
        ctx.rect(contentX, contentY, contentWidth, contentHeight);
        ctx.clip();

        const layout = this.computeCellTextLayout(ctx, cell, contentWidth, contentHeight, { 
            contentX,
            defaultTextColor: cellStyle.textColor 
        });

        if (layout.lines.length > 0) {
            const totalH = layout.totalHeight;
            let vAlign = 'top';
            if (cell && typeof cell.verticalAlignment === 'string') {
                vAlign = cell.verticalAlignment.toLowerCase();
            } else if (cell && cell.textBody && cell.textBody.bodyProperties && cell.textBody.bodyProperties.anchor) {
                const a = String(cell.textBody.bodyProperties.anchor).toLowerCase();
                if (a === 'ctr' || a === 'center' || a === 'middle') {vAlign = 'middle';}
                else if (a === 'b' || a === 'bottom') {vAlign = 'bottom';}
            }

            // For center/bottom alignment, center the visual em box rather than the line height box.
            // textBaseline='top' draws from em box top; trailing leading (lineHeight-fontSize) would
            // otherwise push text above the true visual center of the cell.
            const lastLine = layout.lines[layout.lines.length - 1];
            const trailingLeading = lastLine ? Math.max(0, lastLine.lineHeight - lastLine.fontSizePixels) : 0;

            let cursorY = contentY;
            if (vAlign === 'middle' || vAlign === 'center' || vAlign === 'ctr') {
                cursorY = contentY + Math.max(0, (contentHeight - totalH) / 2) + trailingLeading / 2;
            } else if (vAlign === 'b' || vAlign === 'bottom') {
                cursorY = contentY + Math.max(0, contentHeight - totalH) + trailingLeading;
            }

            for (const ln of layout.lines) {
                let fontStyle = '';
                if (ln.isItalic) {fontStyle += 'italic ';}
                if (ln.isBold) {fontStyle += 'bold ';}
                ctx.font = `${fontStyle}${ln.fontSizePixels}px ${ln.fontFamily}`;
                ctx.textAlign = ln.textAlign;
                // Draw highlight background if present
                if (ln.highlightStr) {
                    const textMetrics = ctx.measureText(ln.text);
                    ctx.save();
                    ctx.fillStyle = ln.highlightStr;
                    ctx.fillRect(Math.round(ln.textX), Math.round(cursorY - ln.fontSizePixels), textMetrics.width, ln.fontSizePixels * 1.2);
                    ctx.restore();
                }
                ctx.fillStyle = ln.colorStr;
                ctx.fillText(ln.text, Math.round(ln.textX), Math.round(cursorY));
                if (ln.isUnderline) {
                    const textW = ctx.measureText(ln.text).width;
                    const ulY = Math.round(cursorY + ln.fontSizePixels * 1.1);
                    const ulX = ln.textAlign === 'center' ? Math.round(ln.textX - textW / 2)
                              : ln.textAlign === 'right'  ? Math.round(ln.textX - textW)
                              : Math.round(ln.textX);
                    ctx.save();
                    ctx.strokeStyle = ln.colorStr;
                    ctx.lineWidth = Math.max(1, ln.fontSizePixels * 0.08);
                    ctx.beginPath();
                    ctx.moveTo(ulX, ulY);
                    ctx.lineTo(ulX + textW, ulY);
                    ctx.stroke();
                    ctx.restore();
                }
                cursorY += ln.lineHeight;
                if (cursorY > contentY + contentHeight) {break;}
            }
        }

        ctx.restore();
    }
    
    /**
     * Get scaled font size based on coordinate system
     */
    getScaledFontSize(pointSize) {
        // Base font size scaling - adjust based on coordinate system scale
        let scaledSize = pointSize;
        
        if (this.coordinateSystem && this.coordinateSystem.scale) {
            // Scale font size with the coordinate system
            scaledSize = pointSize * this.coordinateSystem.scale;
            
            // Ensure minimum readable size
            scaledSize = Math.max(scaledSize, 8);
            
            // Ensure maximum reasonable size  
            scaledSize = Math.min(scaledSize, 72);
        }
        
        
        return Math.round(scaledSize);
    }
    /**
     * Draw table placeholder
     */
    drawTablePlaceholder(bounds, label) {
        const color = { r: 100, g: 100, b: 200, a: 255 };
        
        // Draw border
        this.graphics.p_color(color.r, color.g, color.b, color.a);
        this.graphics.p_width(2);
        this.graphics._s();
        this.graphics._m(bounds.x, bounds.y);
        this.graphics._l(bounds.x + bounds.w, bounds.y);
        this.graphics._l(bounds.x + bounds.w, bounds.y + bounds.h);
        this.graphics._l(bounds.x, bounds.y + bounds.h);
        this.graphics._z();
        if (typeof this.graphics.ds === 'function') {
            this.graphics.ds();
        }
        
        // No placeholder text
    }

    /**
     * Draw chart in graphics adapter with enhanced processing
     */
    async drawChart(shape, bounds) {
        // Capture slide context BEFORE any awaits to prevent corruption
        // by concurrent drawSlide calls that modify this.currentSlideIndex
        const capturedSlideIndex = this.currentSlideIndex;
        const capturedSlide = this.currentSlide;

        // Attach the slide coordinate scale to chart data so the Chart.js renderer can
        // rasterise the chart at its native (zoom-independent) size and downscale into the
        // on-slide area. ChartRenderer receives `this.graphics` (the engine), which has no
        // coordinateSystem, so this scale must be supplied here where it is available.
        const _csScale = (this.coordinateSystem && typeof this.coordinateSystem.scale === 'number' && this.coordinateSystem.scale > 0)
            ? this.coordinateSystem.scale
            : null;
        const _csDpr = (typeof window !== 'undefined' && window.devicePixelRatio) ? window.devicePixelRatio : 1;
        const withChartScale = (cd) => {
            if (cd && _csScale !== null) {
                cd._scalingInfo = Object.assign({}, cd._scalingInfo, {
                    displayScale: _csScale,
                    devicePixelRatio: _csDpr
                });
            }
            return cd;
        };



        // Try to render chart data if available
        if (shape.chartData) {
            // Handle deferred charts from XML parsing phase
            if (shape.chartData.type === 'DEFERRED_CHART') {

                if (window.ChartProcessor) {
                    const chartProcessor = new ChartProcessor(this.context);

                    if (capturedSlideIndex !== undefined && capturedSlide) {
                        const slideContext = {
                            slideIndex: capturedSlideIndex,
                            slideName: capturedSlide?.name || `slide${capturedSlideIndex + 1}`,
                            slide: capturedSlide
                        };
                        chartProcessor.setSlideContext(slideContext);

                        try {
                            const chartData = await chartProcessor.loadChartFromRelationship(shape.chartData.relationshipId, slideContext);
                            if (chartData && chartData.type !== 'DEFERRED_CHART') {
                                shape.chartData = chartData;
                                // Continue with regular chart rendering below
                            } else {
                                this.drawChartPlaceholder(bounds, 'Load Failed');
                                return;
                            }
                        } catch (error) {
                            this.drawChartPlaceholder(bounds, 'Load Error');
                            return;
                        }
                    } else {
                        this.drawChartPlaceholder(bounds, 'No Context');
                        return;
                    }
                } else {
                    this.drawChartPlaceholder(bounds, 'No Processor');
                    return;
                }
            }

            try {
                const chartRenderer = new ChartRenderer(this.graphics);
                await chartRenderer.renderChart(withChartScale(shape.chartData), bounds.x, bounds.y, bounds.w, bounds.h);
            } catch (error) {
                this.drawChartPlaceholder(bounds, 'Render Failed');
            }
        }
        // Try to process chart data if not already processed
        else if (shape.graphicData && shape.graphicData.element && window.ChartProcessor) {
            try {
                const chartProcessor = new ChartProcessor(this.context);

                if (capturedSlideIndex !== undefined && capturedSlide) {
                    const slideContext = {
                        slideIndex: capturedSlideIndex,
                        slideName: capturedSlide?.name || `slide${capturedSlideIndex + 1}`,
                        slide: capturedSlide
                    };
                    chartProcessor.setSlideContext(slideContext);
                }

                // Try embedded chart processing first
                const embeddedData = chartProcessor.parseEmbeddedChartData(shape.graphicData.element);
                if (embeddedData) {
                    shape.chartData = embeddedData;
                    const chartRenderer2 = new ChartRenderer(this.graphics);
                    await chartRenderer2.renderChart(withChartScale(embeddedData), bounds.x, bounds.y, bounds.w, bounds.h);
                    return;
                }

                // Try relationship-based processing if embedded failed
                if (shape.graphicData.chartRef) {
                    const slideContext = {
                        slideIndex: capturedSlideIndex,
                        slideName: capturedSlide?.name || `slide${capturedSlideIndex + 1}`,
                        slide: capturedSlide
                    };
                    try {
                        const chartData = await chartProcessor.loadChartFromRelationship(shape.graphicData.chartRef, slideContext);
                        if (chartData) {
                            shape.chartData = chartData;
                            const chartRenderer3 = new ChartRenderer(this.graphics);
                            await chartRenderer3.renderChart(withChartScale(chartData), bounds.x, bounds.y, bounds.w, bounds.h);
                        } else {
                            this.drawChartPlaceholder(bounds, 'Processing Failed');
                        }
                    } catch (error) {
                        this.drawChartPlaceholder(bounds, 'Processing Error');
                    }
                    return;
                }

                this.drawChartPlaceholder(bounds, 'No Chart Reference');

            } catch (error) {
                this.drawChartPlaceholder(bounds, 'Processing Failed');
            }
        } else {
            this.drawChartPlaceholder(bounds, 'No Chart Data');
        }
    }

    /**
     * Draw chart styling (borders, rounded corners) based on PPTX settings
     */
    drawChartStyling(chartData, x, y, width, height) {
        if (!chartData) {return;}
        
        // Get the correct canvas context - try multiple approaches
        let ctx = null;
        if (this.graphics && this.graphics.ctx) {
            ctx = this.graphics.ctx;
        } else if (this.graphics && this.graphics._context) {
            ctx = this.graphics._context;
        } else if (this.graphics && this.graphics.context) {
            ctx = this.graphics.context;
        } else if (this.graphics && this.graphics.m_oContext) {
            ctx = this.graphics.m_oContext;
        }
        
        if (!ctx) {
            return;
        }
        
        // Use parsed plotArea shape properties to determine border
        const plotProps = chartData.plotArea?.shapeProperties || chartData.shapeProperties || {};
        const line = plotProps.line;
        if (!line || !line.color) {
            // No visible border defined in PPTX
            return;
        }
        
        // Convert EMU (ln@w) to pixels ~ 12700 EMU per px at 96dpi
        const emuPerPx = 12700;
        const borderWidth = Math.max(1, Math.round((line.width || emuPerPx) / emuPerPx));
        const borderColor = (typeof line.color === 'string') ? line.color : this.colorToRgb(line.color);
        const cornerRadius = chartData.roundedCorners ? Math.min(20, Math.floor(Math.min(width, height) * 0.1)) : 0;
        
        ctx.save();
        
        // Set border style
        ctx.strokeStyle = borderColor;
        ctx.lineWidth = borderWidth;
        ctx.lineCap = 'round';
        ctx.lineJoin = 'round';
        // Ensure border is drawn on top
        ctx.globalCompositeOperation = 'source-over';
        
        // Draw border (rounded or square based on roundedCorners setting)
        ctx.beginPath();
        if (chartData.roundedCorners && cornerRadius > 0) {
            // Draw rounded rectangle border - adjust for border width
            const adjustedX = x + borderWidth / 2;
            const adjustedY = y + borderWidth / 2; 
            const adjustedWidth = width - borderWidth;
            const adjustedHeight = height - borderWidth;
            
            // Use the CanvasGraphicsAdapter helper for rounded rectangles
            if (this.graphics && typeof this.graphics.drawRoundRectPath === 'function') {
                this.graphics.drawRoundRectPath(ctx, adjustedX, adjustedY, adjustedWidth, adjustedHeight, cornerRadius);
            } else if (typeof this.drawRoundRectPath === 'function') {
                // Fallback if method is available on this document (older builds)
                this.drawRoundRectPath(ctx, adjustedX, adjustedY, adjustedWidth, adjustedHeight, cornerRadius);
            } else {
                ctx.rect(adjustedX, adjustedY, adjustedWidth, adjustedHeight);
            }
        } else {
            // Draw square rectangle border - adjust for border width
            const adjustedX = x + borderWidth / 2;
            const adjustedY = y + borderWidth / 2;
            const adjustedWidth = width - borderWidth;
            const adjustedHeight = height - borderWidth;
            ctx.rect(adjustedX, adjustedY, adjustedWidth, adjustedHeight);
        }
        ctx.stroke();
        
        ctx.restore();
    }

    /**
     * Draw chart placeholder
     */
    drawChartPlaceholder(bounds, label) {
        const color = { r: 200, g: 100, b: 100, a: 255 };
        
        // Draw border
        this.graphics.p_color(color.r, color.g, color.b, color.a);
        this.graphics.p_width(2);
        this.graphics._s();
        this.graphics._m(bounds.x, bounds.y);
        this.graphics._l(bounds.x + bounds.w, bounds.y);
        this.graphics._l(bounds.x + bounds.w, bounds.y + bounds.h);
        this.graphics._l(bounds.x, bounds.y + bounds.h);
        this.graphics._z();
        if (typeof this.graphics.ds === 'function') {
            this.graphics.ds();
        }
        
        // No placeholder text
    }

    /**
     * Draw text for a shape using text rendering patterns
     */        
    drawShapeText(shape, bounds) {        
        if (!shape || !shape.textBody || !shape.textBody.paragraphs || shape.textBody.paragraphs.length === 0) {
            return;
        }

        // Set current shape for inheritance resolution
        this.setCurrentRenderingShape(shape);

        // Save graphics state
        this.graphics.SaveGrState();

        try {
            // Get text bounds with proper coordinate transformation
            const textBounds = this.calculateTextBounds(shape, bounds);

            // Draw text content using enhanced rendering
            this.drawTextContent(shape.textBody, textBounds);

        } catch (error) {
            // Error handling - draw fallback text
            this.drawFallbackText(shape, bounds);
        } finally {
            // Always restore graphics state
            this.graphics.RestoreGrState();

            // Clear current shape
            this.setCurrentRenderingShape(null);
        }
    }

    /**
     * Calculate text bounds with proper coordinate transformation
     */
    calculateTextBounds(shape, bounds) {
        let { x, y, w, h } = bounds;

        // Get text body properties for proper margin handling (accept bodyPr or bodyProperties)
        const bodyProps = shape.textBody?.bodyProperties || shape.textBody?.bodyPr || {};

        // For rtTriangle shapes, PowerPoint places text in the bottom half of the bounding box
        // (the triangle's area is concentrated there). Shift the text region down accordingly.
        const shapePreset = this.getShapePreset(shape);
        if (shapePreset === 'rtTriangle') {
            y = y + h / 2;
            h = h / 2;
        }

        // Convert EMU margins to pixels with proper scaling
        // Use proportional defaults based on shape size for better layout
        const defaultLeftMargin = Math.max(22860, w * 0.02 * 914400); // 2% of width or min 0.625mm
        const defaultTopMargin = Math.max(11430, h * 0.01 * 914400);  // 1% of height or min 0.3mm

        const leftMargin = this.emuToPixels(bodyProps.leftMargin || defaultLeftMargin);
        const rightMargin = this.emuToPixels(bodyProps.rightMargin || defaultLeftMargin);
        const topMargin = this.emuToPixels(bodyProps.topMargin || defaultTopMargin);
        const bottomMargin = this.emuToPixels(bodyProps.bottomMargin || defaultTopMargin);

        // Apply coordinate system transformation if available
        if (this.coordinateSystem) {
            const { scale } = this.coordinateSystem;

            // Apply scaling more precisely to prevent cumulative errors
            const scaledLeftMargin = Math.round(leftMargin * scale);
            const scaledRightMargin = Math.round(rightMargin * scale);
            const scaledTopMargin = Math.round(topMargin * scale);
            const scaledBottomMargin = Math.round(bottomMargin * scale);

            return {
                x: x + scaledLeftMargin,
                y: y + scaledTopMargin,
                w: Math.max(w - (scaledLeftMargin + scaledRightMargin), 0),
                h: Math.max(h - (scaledTopMargin + scaledBottomMargin), 0),
                scale: scale,
                margins: {
                    left: scaledLeftMargin,
                    right: scaledRightMargin,
                    top: scaledTopMargin,
                    bottom: scaledBottomMargin
                }
            };
        }

        // Fallback without coordinate system
        return {
            x: x + leftMargin,
            y: y + topMargin,
            w: Math.max(w - (leftMargin + rightMargin), 0),
            h: Math.max(h - (topMargin + bottomMargin), 0),
            scale: 1,
            margins: {
                left: leftMargin,
                right: rightMargin,
                top: topMargin,
                bottom: bottomMargin
            }
        };
    }

    /**
     * Draw fallback text when main rendering fails
     */
    drawFallbackText(shape, bounds) {
        const ctx = this.graphics.context;
        if (!ctx) {return;}

        const { x, y, w, h } = bounds;

        // Simple fallback text rendering
        ctx.save();
        const scaleFactor = this.getTextScaleFactor();
        const scaledFontSize = 12 * scaleFactor;
        ctx.font = `${scaledFontSize}px Arial`;
        ctx.fillStyle = '#000000';
        ctx.textAlign = 'left';
        ctx.textBaseline = 'top';

        let textY = y + 5;
        const lineHeight = 16;

        for (const paragraph of shape.textBody.paragraphs) {
            if (textY > y + h - lineHeight) {break;}

            for (const run of paragraph.runs) {
                if (run.text && textY <= y + h - lineHeight) {
                    ctx.fillText(run.text, x + 5, textY);
                    textY += lineHeight;
                }
            }
        }

        ctx.restore();
    }


    /**
     * Draw text content using text rendering patterns
     * Following the approach from sdkjs/word/Drawing/Graphics.js and sdkjs/common/libfont/textmeasurer.js
     */
    drawTextContent(textBody, bounds) {
        if (!textBody || !textBody.paragraphs || textBody.paragraphs.length === 0) {
            return;
        }

        if (!this.graphics || !this.graphics.context) {
            return;
        }

        const ctx = this.graphics.context;
        const { x, y, w, h } = bounds;

        // Standard text rendering approach
        // Following patterns from sdkjs/word/Drawing/Graphics.js and sdkjs/common/libfont/textmeasurer.js

        // If height is zero/very small, re-estimate height and adjust a temporary bounds for rendering
        if (!h || h < 1) {
            const estimatedHeight = this.estimateTextHeightPx(textBody, w);
            if (estimatedHeight && estimatedHeight > 0) {
                const tmpBounds = { ...bounds, h: estimatedHeight };
                this.renderTextBodyStandard(textBody, tmpBounds);
                return;
            }
        }
        this.renderTextBodyStandard(textBody, bounds);
    }

    /**
     * Estimate content height for a text body in pixels given a width
     */
    estimateTextHeightPx(textBody, widthPx) {
        try {
            if (!textBody || !textBody.paragraphs || textBody.paragraphs.length === 0) {return 0;}

            // Use body properties margins if present
            const bodyProps = textBody.bodyProperties || textBody.bodyPr || {};
            const leftMargin = this.emuToPixels(bodyProps.leftMargin || 0);
            const rightMargin = this.emuToPixels(bodyProps.rightMargin || 0);
            const contentWidth = Math.max(0, widthPx - leftMargin - rightMargin);

            let total = 0;
            for (const paragraph of textBody.paragraphs) {
                const paraProps = this.parseParagraphProperties(paragraph, this.getCurrentShape());
                const wrappedLines = this.calculateWrappedLines(paragraph, paraProps, contentWidth, this.getCurrentShape());
                const lineHeight = this.calculateStandardLineHeight(paraProps, wrappedLines);
                const spaceBeforePx = this.emuToPixels(paraProps.spaceBefore || 0);
                const spaceAfterPx = this.emuToPixels(paraProps.spaceAfter || 0);
                total += spaceBeforePx + wrappedLines.length * lineHeight + spaceAfterPx;
            }

            // Add top/bottom margins
            const topMargin = this.emuToPixels(bodyProps.topMargin || 0);
            const bottomMargin = this.emuToPixels(bodyProps.bottomMargin || 0);
            total += topMargin + bottomMargin;

            return Math.ceil(total);
        } catch (_e) {
            return 0;
        }
    }

    /**
     * Estimate overall textbox height, including margins, from a shape and target width
     */
    estimateTextBoxHeightPx(shape, widthPx) {
        const textBody = shape && shape.textBody;
        if (!textBody) {return 0;}
        return this.estimateTextHeightPx(textBody, widthPx);
    }
    /**
     * Render text body using standard patterns
     * Based on sdkjs/word/Drawing/Graphics.js text rendering
     */
    renderTextBodyStandard(textBody, bounds) {
        const { x, y, w, h, margins, scale } = bounds;

        // Use the text bounds directly - margins are already calculated in calculateTextBounds
        const textAreaX = x;
        const textAreaY = y;
        const textAreaWidth = w;
        const textAreaHeight = h;
        
        // Get body properties for alignment and text wrapping (accept bodyPr or bodyProperties)
        const bodyProps = textBody.bodyProperties || textBody.bodyPr || {};

        // Apply "shrink text on overflow" auto-fit. PowerPoint pre-computes a font scale and
        // line-space reduction (normAutofit) so the text fits the shape; honoring them keeps
        // text from rendering oversized (too wide / overflowing the box). These instance
        // fields are read by setupStandardFont and calculateStandardLineHeight for the
        // duration of this text body's layout + rendering, then reset below.
        const _afScale = bodyProps.fontScale;
        this._textAutofitScale = (typeof _afScale === 'number' && _afScale > 0 && _afScale <= 1) ? _afScale : 1;
        const _lnReduction = bodyProps.lineSpaceReduction;
        this._textLineReduction = (typeof _lnReduction === 'number' && _lnReduction > 0) ? Math.min(0.9, _lnReduction) : 0;

        // Handle vertical alignment from body properties
        // Map PPTX anchor values to our internal values
        let verticalAlign = bodyProps.anchor || bodyProps.verticalAlign || 't';
        
        // Normalize PPTX anchor values
        switch (verticalAlign) {
            case 'top':
                verticalAlign = 't';
                break;
            case 'middle':
            case 'center':
                verticalAlign = 'ctr';
                break;
            case 'bottom':
                verticalAlign = 'b';
                break;
            // 'ctr', 'just', 'dist', 't', 'b' are already correct PPTX values
        }
        let startY = textAreaY;

        // Pre-process all paragraphs to calculate layout
        const paragraphLayouts = [];
        let totalTextHeight = 0;

        for (let paraIndex = 0; paraIndex < textBody.paragraphs.length; paraIndex++) {
            const paragraph = textBody.paragraphs[paraIndex];
            const currentShape = this.getCurrentShape();
            const paraProps = this.parseParagraphProperties(paragraph, currentShape);

            // Calculate wrapped lines for this paragraph
            const wrappedLines = this.calculateWrappedLines(paragraph, paraProps, textAreaWidth, currentShape);
            const lineHeight = this.calculateStandardLineHeight(paraProps, wrappedLines);
            // Convert paragraph spacing (EMU) to canvas pixels; apply once per paragraph
            let spaceBeforePx = this.emuToPixels(paraProps.spaceBefore || 0) * scale;
            let spaceAfterPx = this.emuToPixels(paraProps.spaceAfter || 0) * scale;
            // Apply percentage-based spacing from master/layout styles (e.g. spcPct val="20000" = 20%)
            if (paraProps.spaceBeforePct && !paraProps.spaceBefore) {
                spaceBeforePx = paraProps.spaceBeforePct * lineHeight;
            }
            if (paraProps.spaceAfterPct && !paraProps.spaceAfter) {
                spaceAfterPx = paraProps.spaceAfterPct * lineHeight;
            }
            const paragraphHeight = spaceBeforePx + wrappedLines.length * lineHeight + spaceAfterPx;

            paragraphLayouts.push({
                paragraph: paragraph,
                paraProps: paraProps,
                wrappedLines: wrappedLines,
                lineHeight: lineHeight,
                spaceBefore: spaceBeforePx,
                spaceAfter: spaceAfterPx,
                height: paragraphHeight
            });

            totalTextHeight += paragraphHeight;
        }

        // Adjust starting Y position based on vertical alignment (PPTX standard values)
        switch (verticalAlign) {
            case 'ctr':
            case 'middle':
                // Center alignment - allow overflow above/below center point (PPTX standard)
                startY = textAreaY + (textAreaHeight - totalTextHeight) / 2;
                break;
            case 'b':
            case 'bottom':
                // Bottom alignment - allow overflow above bottom point (PPTX standard)
                startY = textAreaY + textAreaHeight - totalTextHeight;
                break;
            case 'just':
                // Justified - distribute evenly with spacing between paragraphs
                if (textBody.paragraphs.length > 1) {
                    startY = textAreaY;
                    // Will be handled in the rendering loop below
                } else {
                    startY = textAreaY; // Single paragraph acts like top align
                }
                break;
            case 'dist':
                // Distributed - similar to justified but includes space at top/bottom
                startY = textAreaY;
                // Will be handled in the rendering loop below
                break;
            case 't':
            case 'top':
            default:
                // Top alignment (default)
                startY = textAreaY;
                break;
        }

        let currentY = startY;
        const maxY = textAreaY + textAreaHeight;

        // Calculate spacing for justified and distributed alignment
        let paragraphSpacing = 0;
        if ((verticalAlign === 'just' || verticalAlign === 'dist') && paragraphLayouts.length > 1) {
            const remainingSpace = Math.max(0, textAreaHeight - totalTextHeight);
            if (remainingSpace > 0) {
                if (verticalAlign === 'just') {
                    // Justified: space between paragraphs only
                    paragraphSpacing = remainingSpace / (paragraphLayouts.length - 1);
                } else if (verticalAlign === 'dist') {
                    // Distributed: space at top, between, and bottom
                    paragraphSpacing = remainingSpace / (paragraphLayouts.length + 1);
                    currentY += paragraphSpacing; // Add space at top
                }
            }
        }

        // Track bullet numbering per level for proper hierarchical numbering
        const levelCounters = {}; // Track counters for each level
        let lastLevel = -1;
        
        // Pre-calculate correct bullet numbers for each paragraph
        const bulletNumbers = {};
        const levelSequence = {};
        let previousLevel = -1;

        // Derive indentation-based levels as a fallback when explicit levels are missing
        const indentValuesSet = new Set();
        for (let i = 0; i < paragraphLayouts.length; i++) {
            const { paraProps } = paragraphLayouts[i];
            if (paraProps && paraProps.bullet && (paraProps.bullet.type === 'number' || paraProps.bullet.type === 'autoNumber')) {
                const indentLeftEmu = (paraProps.indent && typeof paraProps.indent.left === 'number') ? paraProps.indent.left : 0;
                indentValuesSet.add(indentLeftEmu);
            }
        }
        const indentValues = Array.from(indentValuesSet).sort((a, b) => a - b);
        const minIndent = indentValues.length > 0 ? indentValues[0] : 0;
        
        const resolveLevel = (paraProps) => {
            const explicitLevel = (paraProps && typeof paraProps.level === 'number') ? paraProps.level : 0;
            // If explicit level is non-zero, respect it
            if (explicitLevel > 0) {return explicitLevel;}
            // Otherwise, derive from indentation if it indicates deeper nesting
            const indentLeftEmu = (paraProps.indent && typeof paraProps.indent.left === 'number') ? paraProps.indent.left : 0;
            if (indentValues.length <= 1) {return explicitLevel;}
            if (indentLeftEmu <= minIndent) {return 0;}
            let idx = indentValues.indexOf(indentLeftEmu);
            if (idx === -1) {
                // Find closest greater-or-equal indentation level
                idx = indentValues.findIndex(v => v >= indentLeftEmu);
                if (idx === -1) {idx = indentValues.length - 1;}
            }
            return Math.max(0, idx);
        };
        
        for (let i = 0; i < paragraphLayouts.length; i++) {
            const layout = paragraphLayouts[i];
            const paraProps = layout.paraProps;
            
            if (paraProps.bullet && (paraProps.bullet.type === 'number' || paraProps.bullet.type === 'autoNumber')) {
                const currentLevel = resolveLevel(paraProps);
                
                // Check if this bullet has an explicit startAt value
                const hasExplicitStart = paraProps.bullet.startAt !== undefined;
                const startValue = hasExplicitStart ? paraProps.bullet.startAt : 1;
                
                // Handle level transitions for proper hierarchical numbering
                if (currentLevel !== lastLevel) {
                    if (currentLevel < lastLevel) {
                        // Going back to a parent level: clear deeper level counters but keep this level's counter
                        for (let lvl = currentLevel + 1; lvl <= 9; lvl++) {
                            delete levelCounters[lvl];
                        }
                    } else if (currentLevel > lastLevel) {
                        // Going to a deeper level: always start at 1 for new sub-level
                        levelCounters[currentLevel] = 0; // Will be incremented to 1 below
                    }
                }

                // Initialize counter for this level if it doesn't exist yet
                if (!levelCounters[currentLevel] && levelCounters[currentLevel] !== 0) {
                    const startValue2 = hasExplicitStart ? startValue - 1 : 0;
                    levelCounters[currentLevel] = startValue2;
                }
                
                // Initialize or increment counter for current level
                if (!levelCounters[currentLevel]) {
                    levelCounters[currentLevel] = 0;
                }
                levelCounters[currentLevel]++;
                
                // Store the bullet number for this paragraph
                bulletNumbers[i] = levelCounters[currentLevel];
                
                // Update last level for next iteration
                lastLevel = currentLevel;
            }
        }
        
        // Render each paragraph with proper wrapping
        for (let layoutIndex = 0; layoutIndex < paragraphLayouts.length; layoutIndex++) {
            const layout = paragraphLayouts[layoutIndex];
            if (currentY >= maxY) {
                break;
            }

            const { paraProps, wrappedLines, lineHeight, spaceBefore, spaceAfter } = layout;

            // Apply spaceBefore once at the start of the paragraph
            currentY += spaceBefore || 0;
            
            // Handle bullet numbering per level
            const currentLevel = resolveLevel(paraProps);
            if (paraProps.bullet && (paraProps.bullet.type === 'number' || paraProps.bullet.type === 'autoNumber')) {
                // Handle level transitions for proper hierarchical numbering
                if (currentLevel !== lastLevel) {
                    if (currentLevel < lastLevel) {
                        // Going back to a higher level: clear all deeper level counters
                        for (let lvl = currentLevel + 1; lvl <= 9; lvl++) {
                            delete levelCounters[lvl];
                        }
                    } else if (currentLevel > lastLevel) {
                        // Going to a deeper level: always start at 1 for new sub-level
                        levelCounters[currentLevel] = 0; // Will be incremented to 1 below
                    }
                }
                
                // Initialize or increment counter for current level
                if (!levelCounters[currentLevel]) {
                    levelCounters[currentLevel] = 0;
                }
                levelCounters[currentLevel]++;
                
                // Update last level for next iteration
                lastLevel = currentLevel;
            }

            // Render each wrapped line
            for (let lineIndex = 0; lineIndex < wrappedLines.length; lineIndex++) {
                const line = wrappedLines[lineIndex];
                
                // Remove line height check to allow text to render even if it extends beyond bounds
                // This prevents text from disappearing when font sizes are larger
                // if (currentY + lineHeight > maxY) break;

                // Apply indentation for this paragraph (scale to canvas pixels)
                const indentLeft = this.emuToPixels(paraProps.indent?.left || 0) * scale;
                const hangingIndent = this.emuToPixels(paraProps.indent?.hanging || 0) * scale;
                
                // Calculate effective left margin
                // For bullets: bullet at (indentLeft + hangingIndent), text at (indentLeft)
                // For non-bullets: first line at (indentLeft + hangingIndent), subsequent lines at (indentLeft)
                let effectiveLeftMargin = indentLeft;
                if (lineIndex === 0 && hangingIndent !== 0) {
                    // First line uses hanging indent (typically negative for bullet points)
                    effectiveLeftMargin += hangingIndent;
                }

                // Calculate line positioning based on alignment
                const alignment = paraProps.align || 'left';
                let lineStartX = textAreaX + Math.max(0, effectiveLeftMargin);

                if (alignment === 'center' || alignment === 'ctr') {
                    lineStartX = textAreaX + effectiveLeftMargin + (textAreaWidth - line.width - effectiveLeftMargin) / 2;
                } else if (alignment === 'right' || alignment === 'r') {
                    lineStartX = textAreaX + textAreaWidth - line.width;
                }

                // Calculate proper baseline position for text using scaled font size
                const baseFontSize = paraProps.fontSize || 12;
                const scaleFactor = this.getTextScaleFactor();
                const scaledFontSize = baseFontSize * scaleFactor * (this._textAutofitScale || 1);
                const baselineY = currentY + scaledFontSize * 0.8;

                // Render bullet for first line of paragraph
                if (lineIndex === 0) {
                    if (paraProps.bullet && paraProps.bullet.type !== 'none') {
                        // Bullet position: at the hanging indent position (typically negative, so pulls bullet left)
                        const bulletX = textAreaX + indentLeft + hangingIndent;
                        
                        // Use level-specific counter for numbered lists
                        const bulletIndex = (paraProps.bullet.type === 'number' || paraProps.bullet.type === 'autoNumber') 
                            ? (levelCounters[currentLevel] || 1) // Use the current level's counter directly
                            : layoutIndex;
                        
                        // Store pre-calculated bullet number for direct access in renderBullet
                        paraProps._bulletNumber = bulletNumbers[layoutIndex];
                        
                        // Get the actual fontSize from the first run of the first line (this is what text uses)
                        const firstRunFontSize = (line.runs && line.runs[0] && line.runs[0].runProps && line.runs[0].runProps.fontSize) 
                            ? line.runs[0].runProps.fontSize 
                            : paraProps.fontSize || 12;
                        
                        // Create adjusted paraProps with the actual text fontSize for bullet rendering
                        const bulletParaProps = { ...paraProps, fontSize: firstRunFontSize };
                        
                        // Get bullet width to calculate proper spacing
                        const bulletWidth = this.getBulletWidth(paraProps.bullet, bulletParaProps, bulletIndex, levelCounters);
                        
                        this.renderBullet(paraProps.bullet, bulletX, baselineY, bulletParaProps, bulletIndex, levelCounters);
                        
                        // Text starts at indentLeft (the full left margin - this is where continuation lines start)
                        // This respects the list level hierarchy: higher levels have larger indentLeft values
                        // Add small spacing after bullet to prevent overlap
                        const actualScaledFontSize = firstRunFontSize * scaleFactor;
                        const minSpacing = actualScaledFontSize * 0.3; // 30% of font size spacing
                        const bulletEndX = bulletX + bulletWidth;
                        const expectedTextStartX = textAreaX + indentLeft;
                        
                        // If bullet extends past where text should start, push text right
                        if (bulletEndX + minSpacing > expectedTextStartX) {
                            lineStartX = bulletEndX + minSpacing;
                        } else {
                            lineStartX = expectedTextStartX;
                        }
                    }
                }

                // Render all runs in this line
                let currentX = lineStartX;
                for (const runData of line.runs) {
                    const { text, runProps } = runData;

                    // Apply text capitalization
                    let textToRender = text;
                    if (runProps.cap && runProps.cap !== 'none') {
                        switch (runProps.cap) {
                            case 'all':
                            case 'small':
                                textToRender = textToRender.toUpperCase();
                                break;
                            case 'words':
                                textToRender = textToRender.replace(/\b\w/g, char => char.toUpperCase());
                                break;
                        }
                    }

                    // Set up font and render text
                    this.setupStandardFont(runProps);

                    // Draw highlight background if present (before text so it appears behind)
                    if (runProps.highlight) {
                        const hlCtx = this.graphics.context;
                        if (hlCtx) {
                            const hl = runProps.highlight;
                            let hlColor;
                            if (typeof hl === 'object' && hl.r !== undefined) {
                                hlColor = `rgba(${hl.r}, ${hl.g}, ${hl.b}, ${(hl.a || 255) / 255})`;
                            } else if (typeof hl === 'string') {
                                hlColor = hl;
                            }
                            if (hlColor) {
                                const runWidth = hlCtx.measureText(textToRender).width;
                                const sFontSize = runProps.scaledFontSize || ((runProps.fontSize || 12) * this.getTextScaleFactor());
                                hlCtx.save();
                                hlCtx.fillStyle = hlColor;
                                hlCtx.fillRect(Math.round(currentX), Math.round(baselineY - sFontSize * 0.85), runWidth, sFontSize * 1.15);
                                hlCtx.restore();
                                // Restore fill color for text rendering
                                hlCtx.fillStyle = this.graphics.colorToRgb(runProps.color);
                            }
                        }
                    }

                    // Apply text-level effects (shadow, glow) before rendering
                    const effectCtx = this.graphics.context;
                    let hasTextEffect = false;
                    if (runProps.effectLst && effectCtx) {
                        const effects = runProps.effectLst;
                        if (effects.outerShadow) {
                            const shadow = effects.outerShadow;
                            effectCtx.save();
                            hasTextEffect = true;
                            const blur = (shadow.blurRadius || 0) / 12700; // EMU to approx pixels
                            const dist = (shadow.distance || 0) / 12700;
                            const dir = ((shadow.direction || 0) / 60000) * Math.PI / 180;
                            effectCtx.shadowOffsetX = Math.cos(dir) * dist * this.getTextScaleFactor();
                            effectCtx.shadowOffsetY = Math.sin(dir) * dist * this.getTextScaleFactor();
                            effectCtx.shadowBlur = blur * this.getTextScaleFactor();
                            if (shadow.color) {
                                const sc = shadow.color;
                                let r = 0, g = 0, b = 0, alpha = 0.5;
                                if (sc.r !== undefined) {
                                    // Already in {r,g,b,a} format
                                    r = sc.r; g = sc.g; b = sc.b;
                                    alpha = sc.a !== undefined ? sc.a / 255 : 0.5;
                                } else if (sc.value) {
                                    // {type, value, alpha} format from parseColor
                                    const hex = sc.value.replace('#', '');
                                    r = parseInt(hex.slice(0,2), 16) || 0;
                                    g = parseInt(hex.slice(2,4), 16) || 0;
                                    b = parseInt(hex.slice(4,6), 16) || 0;
                                    alpha = sc.alpha !== undefined ? sc.alpha : 0.5;
                                }
                                effectCtx.shadowColor = `rgba(${r}, ${g}, ${b}, ${alpha})`;
                            } else {
                                effectCtx.shadowColor = 'rgba(0, 0, 0, 0.4)';
                            }
                        } else if (effects.glow) {
                            const glow = effects.glow;
                            effectCtx.save();
                            hasTextEffect = true;
                            const radius = (glow.radius || 0) / 12700;
                            effectCtx.shadowOffsetX = 0;
                            effectCtx.shadowOffsetY = 0;
                            effectCtx.shadowBlur = radius * this.getTextScaleFactor();
                            if (glow.color) {
                                const gc = glow.color;
                                effectCtx.shadowColor = `rgba(${gc.r || 0}, ${gc.g || 0}, ${gc.b || 0}, ${(gc.a || 255) / 255})`;
                            } else {
                                effectCtx.shadowColor = 'rgba(255, 0, 0, 0.6)';
                            }
                        }
                    }

                    // Adjust Y position for superscript/subscript
                    let renderY = baselineY;
                    if (runProps.verticalAlign && runProps.verticalAlign !== 'normal' && runProps.verticalAlign !== 'baseline') {
                        const sFontSize = runProps.scaledFontSize || ((runProps.fontSize || 12) * this.getTextScaleFactor());
                        if (runProps.verticalAlign === 'superscript' || runProps.verticalAlign === 'super') {
                            renderY = baselineY - sFontSize * 0.35;
                        } else if (runProps.verticalAlign === 'subscript' || runProps.verticalAlign === 'sub') {
                            renderY = baselineY + sFontSize * 0.2;
                        }
                    }

                    currentX = this.renderRunTextStandard(textToRender, currentX, renderY, runProps, textAreaX + textAreaWidth);

                    // Restore context after text effects
                    if (hasTextEffect && effectCtx) {
                        effectCtx.restore();
                        // Re-apply fill style after restore
                        effectCtx.fillStyle = this.graphics.colorToRgb(runProps.color);
                    }
                }

                currentY += lineHeight;
            }

            // Apply spaceAfter once after the paragraph
            currentY += spaceAfter || 0;

            // Add paragraph spacing for justified/distributed alignment
            if (layoutIndex < paragraphLayouts.length - 1) { // Don't add spacing after last paragraph
                currentY += paragraphSpacing;
            }
        }

        // Reset auto-fit scaling so it doesn't leak into the next shape's text.
        this._textAutofitScale = 1;
        this._textLineReduction = 0;
    }

    /**
     * Calculate wrapped lines for a paragraph
     */
    calculateWrappedLines(paragraph, paraProps, maxWidth, currentShape) {
        const lines = [];
        let currentLine = { runs: [], width: 0 };
        let isFirstRun = true;

        for (const run of paragraph.runs) {
            if (!run.text) {continue;}

            const wasFirstRun = isFirstRun;
            isFirstRun = false;

            const runProps = this.parseRunProperties(run, paraProps, currentShape);

            let runText = run.text;
            // PptxGenJS includes literal "1. " in the text run AND sets buAutoNum in pPr —
            // strip the leading number prefix from the first run to avoid double-rendering.
            if (wasFirstRun && paraProps.bullet &&
                (paraProps.bullet.type === 'autoNumber' || paraProps.bullet.type === 'number')) {
                runText = runText.replace(/^\s*\d+[.)]\s*/, '');
            }
            // Resolve dynamic field values at render time
            if (run.fieldType) {
                const fType = (run.fieldType || '').toLowerCase();
                if (fType === 'slidenum') {
                    // currentSlideIndex is 0-based; slide numbers are 1-based
                    const slideNumber = (this.currentSlideIndex !== undefined && this.currentSlideIndex !== null)
                        ? (this.currentSlideIndex + 1)
                        : parseInt(runText, 10) || 1;
                    runText = String(slideNumber);
                }
            }
            if (!runText) {continue;}

            // Split text on explicit newlines first, then word-wrap each segment
            const segments = runText.split('\n');
            for (let segIdx = 0; segIdx < segments.length; segIdx++) {
                // Force a line break between segments (except before the first)
                if (segIdx > 0) {
                    if (currentLine.runs.length > 0) {
                        lines.push(currentLine);
                    }
                    currentLine = { runs: [], width: 0 };
                }

                const segment = segments[segIdx];
                if (!segment) {continue;}

                const words = segment.split(/(\s+)/); // Keep whitespace

                for (const word of words) {
                    if (!word) {continue;}

                    // Measure word width
                    const wordWidth = this.measureRunText(word, runProps);

                    // Check if word fits on current line
                    if (currentLine.width + wordWidth <= maxWidth) {
                        // Word fits — add to current line
                        currentLine.runs.push({ text: word, runProps: runProps });
                        currentLine.width += wordWidth;
                    } else if (wordWidth <= maxWidth) {
                        // Word doesn't fit on current line but fits on its own — start new line
                        if (currentLine.runs.length > 0) {
                            lines.push(currentLine);
                        }
                        currentLine = {
                            runs: [{ text: word, runProps: runProps }],
                            width: wordWidth
                        };
                    } else {
                        // Word is longer than maxWidth (e.g. a URL with no spaces) — break at character level
                        if (currentLine.runs.length > 0) {
                            lines.push(currentLine);
                            currentLine = { runs: [], width: 0 };
                        }
                        let charBuf = '';
                        let charBufWidth = 0;
                        for (const ch of word) {
                            const chWidth = this.measureRunText(ch, runProps);
                            if (charBufWidth + chWidth > maxWidth && charBuf.length > 0) {
                                currentLine.runs.push({ text: charBuf, runProps: runProps });
                                currentLine.width += charBufWidth;
                                lines.push(currentLine);
                                currentLine = { runs: [], width: 0 };
                                charBuf = '';
                                charBufWidth = 0;
                            }
                            charBuf += ch;
                            charBufWidth += chWidth;
                        }
                        if (charBuf) {
                            currentLine.runs.push({ text: charBuf, runProps: runProps });
                            currentLine.width += charBufWidth;
                        }
                    }
                }
            }
        }

        // Add final line
        if (currentLine.runs.length > 0) {
            lines.push(currentLine);
        }

        // If no lines, create empty line
        if (lines.length === 0) {
            lines.push({ runs: [], width: 0 });
        }

        return lines;
    }

    /**
     * Set up font following standard patterns
     * Fixed: Apply proper font scaling to match geometry scaling
     */
    setupStandardFont(runProps) {
        const ctx = this.graphics.context;
        if (!ctx) {return;}

        // Scale font size to match slide-to-canvas scaling, including any active
        // auto-fit ("shrink text on overflow") scale for the current text body.
        const scaleFactor = this.getTextScaleFactor();
        const autofitScale = this._textAutofitScale || 1;
        const scaledFontSize = (runProps.fontSize || 12) * scaleFactor * autofitScale;

        // Build font string
        const fontStyle = runProps.italic ? 'italic' : 'normal';
        const fontWeight = runProps.bold ? 'bold' : 'normal';
        const fontFamily = runProps.fontFamily || 'Arial';
        
        // Include emoji-capable fonts in font stack for Unicode character support
        // Use system fonts that are more likely to be available
        const fontStack = `"${fontFamily}", system-ui, -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial, "Noto Color Emoji", "Apple Color Emoji", "Segoe UI Emoji", "Segoe UI Symbol", sans-serif`;

        ctx.font = `${fontStyle} ${fontWeight} ${scaledFontSize}px ${fontStack}`;
        ctx.fillStyle = this.graphics.colorToRgb(runProps.color);
        ctx.textBaseline = 'alphabetic';

        // Store scaled font size for line height calculations
        runProps.scaledFontSize = scaledFontSize;
    }

    /**
     * Get bullet width for spacing calculations
     */
    getBulletWidth(bulletProps, paraProps, paragraphIndex = 0, levelCounters = null) {
        const ctx = this.graphics.context;
        if (!ctx || !bulletProps) {
            return 0;
        }

        ctx.save();
        
        try {
            // Set up font for bullet measurement (same as renderBullet)
            const baseFontSize = paraProps.fontSize || 12;
            const scaleFactor = this.getTextScaleFactor();
            const scaledFontSize = baseFontSize * scaleFactor;
            const fontFamily = paraProps.fontFamily || 'Arial';
            const fontWeight = paraProps.bold ? 'bold' : 'normal';
            const fontStyle = paraProps.italic ? 'italic' : 'normal';
            
            ctx.font = `${fontStyle} ${fontWeight} ${scaledFontSize}px ${fontFamily}`;
            
            let bulletText = '';
            
            switch (bulletProps.type) {
                case 'character':
                    bulletText = bulletProps.char || '•';
                    break;
                    
                case 'number':
                    {
                        const base = (bulletProps.startAt || 1) - 1;
                        const indexForLevel = (paraProps && typeof paraProps._bulletNumber === 'number')
                            ? paraProps._bulletNumber
                            : paragraphIndex;
                        bulletText = `${base + indexForLevel}.`;
                    }
                    break;
                    
                case 'autoNumber':
                    {
                        const base = (bulletProps.startAt || 1) - 1;
                        const levelNumber = (paraProps && typeof paraProps._bulletNumber === 'number')
                            ? paraProps._bulletNumber
                            : paragraphIndex;
                        bulletText = this.getAutoNumberBullet(bulletProps.subType, base + levelNumber);
                    }
                    break;
                    
                default:
                    bulletText = '•';
                    break;
            }

            const width = ctx.measureText(bulletText).width;
            ctx.restore();
            return width;
            
        } catch (error) {
            ctx.restore();
            return 0;
        }
    }

    /**
     * Calculate bullet information including width and spacing
     */
    calculateBulletInfo(bulletProps, paraProps, paragraphIndex = 0) {
        const ctx = this.graphics.getContext();
        ctx.save();
        
        // Set up font for bullet measurement
        const baseFontSize = paraProps.fontSize || 12;
        const scaleFactor = this.getTextScaleFactor();
        const scaledFontSize = Math.max(baseFontSize * scaleFactor, 14);
        const fontFamily = paraProps.fontFamily || 'Arial';
        const fontWeight = paraProps.bold ? 'bold' : 'normal';
        const fontStyle = paraProps.italic ? 'italic' : 'normal';
        
        ctx.font = `${fontStyle} ${fontWeight} ${scaledFontSize}px ${fontFamily}`;
        
        let bulletText = '';
        switch (bulletProps.type) {
            case 'character':
                bulletText = bulletProps.char || '•';
                break;
            case 'number':
                {
                    const base = (bulletProps.startAt || 1) - 1;
                    const indexForLevel = (paraProps && typeof paraProps._bulletNumber === 'number')
                        ? paraProps._bulletNumber
                        : paragraphIndex;
                    bulletText = `${base + indexForLevel}.`;
                }
                break;
            case 'autoNumber':
                {
                    const base = (bulletProps.startAt || 1) - 1;
                    const indexForLevel = (paraProps && typeof paraProps._bulletNumber === 'number')
                        ? paraProps._bulletNumber
                        : paragraphIndex;
                    const bulletNumber = base + indexForLevel;
                    bulletText = this.getAutoNumberBullet(bulletProps.subType, bulletNumber);
                }
                break;
            default:
                bulletText = '•';
                break;
        }
        
        const bulletWidth = ctx.measureText(bulletText).width;
        const bulletSpacing = scaledFontSize * 0.5; // Space between bullet and text
        
        ctx.restore();
        
        return {
            bulletText: bulletText,
            bulletWidth: bulletWidth,
            bulletSpacing: bulletSpacing,
            totalWidth: bulletWidth + bulletSpacing
        };
    }
    
    /**
     * Generate hierarchical number for auto-numbered bullets
     * Sub-levels restart numbering (e.g., 1., 2., 3. then 1., 2. for sub-level)
     */
    generateHierarchicalNumber(bulletProps, paraProps, paragraphIndex, levelCounters) {
        const currentLevel = paraProps.level || 0;
        
        // For all levels, use the current level's counter
        const bulletNumber = levelCounters && levelCounters[currentLevel] 
            ? levelCounters[currentLevel] 
            : (bulletProps.startAt || 1) + paragraphIndex;
            
        return this.getAutoNumberBullet(bulletProps.subType, bulletNumber);
    }

    /**
     * Render bullet for a paragraph
     */
    renderBullet(bulletProps, x, y, paraProps, paragraphIndex, levelCounters = null) {
        const ctx = this.graphics.context;
        if (!ctx || !bulletProps) {
            return;
        }

        ctx.save();

        try {
            // Set up font for bullet (use paragraph font properties)
            const baseFontSize = paraProps.fontSize || 12;
            const scaleFactor = this.getTextScaleFactor();
            const scaledFontSize = baseFontSize * scaleFactor; // Use same size as text - no additional zoom scaling
            const fontFamily = paraProps.fontFamily || 'Arial';
            const fontWeight = paraProps.bold ? 'bold' : 'normal';
            const fontStyle = paraProps.italic ? 'italic' : 'normal';
            
            // Build font string - match text rendering exactly (no currentZoom multiplication)
            ctx.font = `${fontStyle} ${fontWeight} ${scaledFontSize}px ${fontFamily}`;
            
            // Use the paragraph text color, but ensure it's visible (not white/transparent)
            const bulletColor = paraProps.color || { r: 0, g: 0, b: 0, a: 255 };
            const bulletColorRgb = this.graphics.colorToRgb(bulletColor);
            ctx.fillStyle = bulletColorRgb;
            ctx.textBaseline = 'alphabetic';

            let bulletText = '';
            
            switch (bulletProps.type) {
                case 'character':
                    bulletText = bulletProps.char || '•';
                    break;
                    
                case 'number':
                    {
                        const base = (bulletProps.startAt || 1) - 1;
                        const indexForLevel = (paraProps && typeof paraProps._bulletNumber === 'number')
                            ? paraProps._bulletNumber
                            : paragraphIndex;
                        bulletText = `${base + indexForLevel}.`;
                    }
                    break;
                    
                case 'autoNumber':
                    {
                        // Use the pre-calculated bullet number for correct sub-level numbering
                        const base = (bulletProps.startAt || 1) - 1;
                        const levelNumber = (paraProps && typeof paraProps._bulletNumber === 'number')
                            ? paraProps._bulletNumber
                            : paragraphIndex;
                        bulletText = this.getAutoNumberBullet(bulletProps.subType, base + levelNumber);
                    }
                    break;
                    
                default:
                    bulletText = '•'; // Default bullet
                    break;
            }

            // Render the bullet at the provided x position (already calculated correctly in caller)
            // The x position passed in is already the correct bullet position (indentLeft + hangingIndent)
            ctx.fillText(bulletText, x, y);

        } catch (error) {
        } finally {
            ctx.restore();
        }
    }

    /**
     * Get automatic numbering bullet text based on subType
     */
    getAutoNumberBullet(subType, number) {
        switch (subType) {
            case 'arabicPeriod':
            case 'arabic1Minus':
            case 'arabic2Minus':
                return `${number}.`;
            case 'arabicParenR':
                return `${number})`;
            case 'arabicParenBoth':
                return `(${number})`;
            case 'romanUcPeriod':
                return `${this.toRoman(number, true)}.`;
            case 'romanLcPeriod':
                return `${this.toRoman(number, false)}.`;
            case 'alphaUcPeriod':
                return `${this.toAlpha(number, true)}.`;
            case 'alphaLcPeriod':
                return `${this.toAlpha(number, false)}.`;
            case 'alphaUcParenR':
                return `${this.toAlpha(number, true)})`;
            case 'alphaLcParenR':
                return `${this.toAlpha(number, false)})`;
            case 'alphaUcParenBoth':
                return `(${this.toAlpha(number, true)})`;
            case 'alphaLcParenBoth':
                return `(${this.toAlpha(number, false)})`;
            default:
                return `${number}.`; // Default to arabic period
        }
    }

    /**
     * Convert number to Roman numerals
     */
    toRoman(num, uppercase = true) {
        const values = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
        const symbols = uppercase 
            ? ['M', 'CM', 'D', 'CD', 'C', 'XC', 'L', 'XL', 'X', 'IX', 'V', 'IV', 'I']
            : ['m', 'cm', 'd', 'cd', 'c', 'xc', 'l', 'xl', 'x', 'ix', 'v', 'iv', 'i'];
        
        let result = '';
        for (let i = 0; i < values.length; i++) {
            while (num >= values[i]) {
                result += symbols[i];
                num -= values[i];
            }
        }
        return result;
    }

    /**
     * Convert number to alphabetic (A, B, C, ... Z, AA, BB, etc.)
     */
    toAlpha(num, uppercase = true) {
        let result = '';
        const base = 26;
        const offset = uppercase ? 65 : 97; // ASCII 'A' or 'a'
        
        while (num > 0) {
            num--; // Make it 0-based
            result = String.fromCharCode(offset + (num % base)) + result;
            num = Math.floor(num / base);
        }
        
        return result || (uppercase ? 'A' : 'a');
    }

    /**
     * Render run text character-by-character
     */
    renderRunTextStandard(text, startX, y, runProps, maxX) {
        const ctx = this.graphics.context;
        if (!ctx || !text) {return startX;}

        let currentX = startX;

        // Apply text effects from current rendering shape before rendering
        if (this.currentRenderingShape && this.currentRenderingShape.effects) {
            this.applyTextEffectsToContext(ctx, this.currentRenderingShape.effects);
        }

        // Handle emoji and multi-byte Unicode characters properly
        const textArray = Array.from(text); // This properly handles Unicode code points
        
        for (let i = 0; i < textArray.length; i++) {
            const char = textArray[i];

            // Check if character fits within bounds - but allow some overflow for larger fonts
            const charWidth = ctx.measureText(char).width;

            // Render character directly using canvas API (keep original behavior)
            ctx.fillText(char, currentX, y);

            // Advance position with proper character spacing
            currentX += charWidth + (runProps.letterSpacing || 0);
        }

        // Reset text effects after rendering
        this.resetTextEffectsOnContext(ctx);

        // Draw text decorations
        this.drawTextDecorations(startX, y, currentX - startX, runProps);

        return currentX;
    }

    /**
     * Apply effects to canvas context for shapes
     */
    applyEffectsToCanvas(ctx, effects) {
        if (!effects || !ctx) {
            return;
        }

        // Apply outer shadow
        if (effects.outerShadow) {
            this.applyOuterShadowToContext(ctx, effects.outerShadow);
        }

        // Apply glow effect
        if (effects.glow) {
            this.applyGlowEffectToContext(ctx, effects.glow);
        }

        // Apply inner shadow
        if (effects.innerShadow) {
            // Inner shadow is not directly supported by canvas
        }
    }

    /**
     * Apply text effects directly to canvas context
     */
    applyTextEffectsToContext(ctx, effects) {
        if (!effects || !ctx) {
            return;
        }

        // Apply outer shadow
        if (effects.outerShadow) {
            this.applyOuterShadowToContext(ctx, effects.outerShadow);
        }

        // Apply glow effect
        if (effects.glow) {
            this.applyGlowEffectToContext(ctx, effects.glow);
        }

        // Apply inner shadow
        if (effects.innerShadow) {
            // Inner shadow is not directly supported by canvas
        }
    }

    /**
     * Apply outer shadow effect to canvas context
     */
    applyOuterShadowToContext(ctx, shadow) {
        if (!shadow || !ctx) {
            return;
        }

        // Convert EMU values to pixels at 96 DPI base
        const blurRadius = (shadow.blurRadius || 0) / 9525;
        const distance = (shadow.distance || 0) / 9525;
        const direction = (shadow.direction || 0) / 60000; // Convert to degrees

        // Canvas shadowBlur is in physical pixels (not affected by ctx transform).
        // Scale it by the coordinate system scale so it matches the rendered text size.
        const scale = (this.coordinateSystem && this.coordinateSystem.scale) || 1;

        // Calculate offset from distance and direction (in canvas coordinate units)
        const angleRad = (direction * Math.PI) / 180;
        const offsetX = Math.cos(angleRad) * distance;
        const offsetY = Math.sin(angleRad) * distance;

        // Set shadow properties
        ctx.shadowOffsetX = offsetX;
        ctx.shadowOffsetY = offsetY;
        ctx.shadowBlur = blurRadius * scale;

        // Set shadow color
        if (shadow.color) {
            const colorStr = this.convertColorToString(shadow.color);
            ctx.shadowColor = colorStr;
        } else {
            ctx.shadowColor = 'rgba(0, 0, 0, 0.5)';
        }
    }

    /**
     * Apply glow effect to canvas context (simulated with shadow)
     */
    applyGlowEffectToContext(ctx, glow) {
        if (!glow || !ctx) {
            return;
        }

        const radius = (glow.radius || 0) / 9525; // EMU to pixels

        // Set glow properties (no offset for glow)
        ctx.shadowOffsetX = 0;
        ctx.shadowOffsetY = 0;
        ctx.shadowBlur = radius;

        // Set glow color
        if (glow.color) {
            const colorStr = this.convertColorToString(glow.color);
            ctx.shadowColor = colorStr;
        } else {
            ctx.shadowColor = 'rgba(255, 255, 255, 0.8)';
        }
    }

    /**
     * Reset text effects on canvas context
     */
    /**
     * Reset effects on canvas context
     */
    resetEffectsOnContext(ctx) {
        if (!ctx) {
            return;
        }

        ctx.shadowOffsetX = 0;
        ctx.shadowOffsetY = 0;
        ctx.shadowBlur = 0;
        ctx.shadowColor = 'transparent';
    }
    resetTextEffectsOnContext(ctx) {
        if (!ctx) {
            return;
        }

        ctx.shadowOffsetX = 0;
        ctx.shadowOffsetY = 0;
        ctx.shadowBlur = 0;
        ctx.shadowColor = 'transparent';
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
     * Calculate text scale factor based on slide dimensions
     * Fixed: Apply proper scaling to match PowerPoint's text rendering
     */
    getTextScaleFactor() {
        if (!this.coordinateSystem) {return 1.0;}

        const { scale } = this.coordinateSystem;
        
        // PowerPoint renders text at a higher effective resolution than geometry
        // Text should be scaled more than geometry to match PowerPoint's appearance
        // PowerPoint uses approximately 1.33x scaling for text relative to geometry
        // This compensates for the difference between 72 DPI (points) and 96 DPI (screen)
        const textScaleMultiplier = 1.33; // 96/72 = 1.333...
        
        // Apply both the coordinate scale and the text-specific multiplier
        return scale * textScaleMultiplier;
    }


    /**
     * Calculate line height with proper scaling
     * Fixed: Apply proper scaling to match font scaling
     */
    calculateStandardLineHeight(paraProps, wrappedLines = null) {
        // Auto-fit ("shrink text on overflow") line-space reduction for the current text body.
        const lnReductionFactor = 1 - (this._textLineReduction || 0);
        // If absolute line spacing in points is provided, use it exactly
        if (paraProps.lineHeightPoints) {
            // Convert points to pixels: pts * (96/72) gives base pixels, then scale for canvas
            // Use coordinateSystem.scale directly (NOT scaleFactor which already includes 96/72)
            const pixelsPerPoint = 96 / 72;
            const csScale = (this.coordinateSystem && this.coordinateSystem.scale) || 1;
            return paraProps.lineHeightPoints * pixelsPerPoint * csScale * lnReductionFactor;
        }

        let baseFontSizePt = paraProps.fontSize || 12; // points

        // When runs use a smaller font than the inherited paragraph default (e.g., 12pt runs
        // inside an 18pt otherStyle paragraph), PowerPoint/LibreOffice both compute "single"
        // line spacing as actualRunFontSize × 1.2 (the "font-independent" factor per
        // LibreOffice ImplCalculateFontIndependentLineSpacing and CSS line-height:normal).
        // Only apply when runs are SMALLER than inherited — when runs are larger (e.g., 30pt
        // title with 18pt inherited), keep the inherited value to preserve correct rendering.
        if (wrappedLines && wrappedLines.length > 0) {
            let maxRunFontSize = 0;
            for (const line of wrappedLines) {
                for (const runData of (line.runs || [])) {
                    if (runData.runProps && runData.runProps.fontSize > 0) {
                        maxRunFontSize = Math.max(maxRunFontSize, runData.runProps.fontSize);
                    }
                }
            }
            if (maxRunFontSize > 0 && maxRunFontSize < baseFontSizePt) {
                baseFontSizePt = maxRunFontSize * 1.2;
            }
        }

        const scaleFactor = this.getTextScaleFactor();
        const autofitScale = this._textAutofitScale || 1;
        const scaledFontSizePx = baseFontSizePt * scaleFactor * autofitScale; // px

        // Otherwise, use percent of font size (default 100%)
        const lineHeightPercent = paraProps.lineHeight || 100;
        const lineHeight = (scaledFontSizePx * lineHeightPercent) / 100 * lnReductionFactor;

        // Return line height only; paragraph spacing is applied once per paragraph
        return lineHeight;
    }

    /**
     * Measure run text with proper scaling
     */
    measureRunText(text, runProps) {
        const ctx = this.graphics.context;
        if (!ctx || !text) {return 0;}

        // Set up font for measurement
        ctx.save();
        this.setupStandardFont(runProps);

        // Measure text character by character for accuracy
        let totalWidth = 0;
        for (let i = 0; i < text.length; i++) {
            const char = text.charAt(i);
            const metrics = ctx.measureText(char);
            totalWidth += metrics.width + (runProps.letterSpacing || 0);
        }

        ctx.restore();
        return totalWidth;
    }

    /**
     * Draw text decorations
     */
    drawTextDecorations(x, y, width, runProps) {
        const ctx = this.graphics.context;
        // Use scaled font size for consistent decoration thickness and positioning
        const fontSize = runProps.scaledFontSize || runProps.fontSize || 12;

                if (runProps.underline) {
                    ctx.save();
                    ctx.strokeStyle = ctx.fillStyle;
                    ctx.lineWidth = Math.max(1, fontSize * 0.05);
                    ctx.beginPath();
            ctx.moveTo(x, y + 2);
            ctx.lineTo(x + width, y + 2);
                    ctx.stroke();
                    ctx.restore();
                }

                if (runProps.strike) {
                    ctx.save();
                    ctx.strokeStyle = ctx.fillStyle;
                    ctx.lineWidth = Math.max(1, fontSize * 0.05);
                    ctx.beginPath();
            const strikeY = y - (fontSize * 0.3);
            ctx.moveTo(x, strikeY);
            ctx.lineTo(x + width, strikeY);
                    ctx.stroke();
                    ctx.restore();
                }
    }

    /**
     * Measure run text following measurement patterns
     * Based on sdkjs/common/libfont/textmeasurer.js Measure method
     * @deprecated Use measureRunText instead
     */
    measureRunTextAlternative(text, runProps) {
        const ctx = this.graphics.context;

        // Set up font for measurement
        ctx.save();
        this.setupStandardFont(runProps);

        // Measure text character by character
        let totalWidth = 0;
        for (let i = 0; i < text.length; i++) {
            const char = text.charAt(i);
            const metrics = ctx.measureText(char);
            totalWidth += metrics.width + (runProps.letterSpacing || 0);
        }

                ctx.restore();
        return totalWidth;
            }

    /**
     * Calculate line height - alternative
     * Based on sdkjs/common/libfont/textmeasurer.js GetHeight method
     * @deprecated Use calculateStandardLineHeight instead
     */
    calculateStandardLineHeightAlternative(paraProps) {
        const fontSize = paraProps.fontSize || 12;
        const lineHeight = paraProps.lineHeight || 120; // Default 120%

        // Calculate line height
            const actualLineHeight = (fontSize * lineHeight) / 100;

        // Ensure minimum line height
        return Math.max(actualLineHeight, fontSize * 1.2);
    }

    /**
     * Render a line of text with proper alignment
     */
    renderTextLine(ctx, lineContent, x, y, maxWidth, paraProps) {
        if (lineContent.length === 0) {return;}

        // Calculate total line width
        const totalWidth = lineContent.reduce((sum, item) => sum + item.width, 0);

        // Determine starting X position based on alignment
        let startX = x;
        const alignment = paraProps.align || 'left';

        // Get coordinate system scale factor for accurate positioning
        const scaleFactor = this.coordinateSystem?.scale || 1;
        const effectiveMaxWidth = maxWidth * scaleFactor;
        const effectiveTotalWidth = totalWidth * scaleFactor;

        if (alignment === 'center' || alignment === 'ctr') {
            // Improved centering calculation accounting for coordinate transformation
            const remainingSpace = Math.max(0, maxWidth - totalWidth);
            startX = x + (remainingSpace / 2);
        } else if (alignment === 'right' || alignment === 'r') {
            startX = x + Math.max(0, maxWidth - totalWidth);
        } else if (alignment === 'justify' && lineContent.length > 1) {
            // For justify, we'll space out the content (simplified)
            startX = x;
        }

        // Render each run in the line
        let currentX = startX;
        for (const item of lineContent) {
            // Set font and color for this run
            ctx.font = item.font;
            ctx.fillStyle = this.graphics.colorToRgb(item.props.color || { r: 0, g: 0, b: 0, a: 255 });
            
            // Use consistent text baseline for proper vertical alignment
            ctx.textBaseline = 'alphabetic';
            ctx.textAlign = 'left'; // Always use left align, we calculate position manually

            // Draw highlight background if present
            if (item.props.highlight) {
                const hl = item.props.highlight;
                let hlColor;
                if (typeof hl === 'object' && hl.r !== undefined) {
                    hlColor = `rgba(${hl.r}, ${hl.g}, ${hl.b}, ${(hl.a || 255) / 255})`;
                } else if (typeof hl === 'string') {
                    hlColor = hl;
                }
                if (hlColor) {
                    const scaleFactor = this.getTextScaleFactor();
                    const scaledFontSize = (item.props.fontSize || 12) * scaleFactor;
                    ctx.save();
                    ctx.fillStyle = hlColor;
                    ctx.fillRect(Math.round(currentX), Math.round(y - scaledFontSize), item.width, scaledFontSize * 1.2);
                    ctx.restore();
                    ctx.fillStyle = this.graphics.colorToRgb(item.props.color || { r: 0, g: 0, b: 0, a: 255 });
                }
            }

            // Draw the text with precise positioning
            ctx.fillText(item.text, Math.round(currentX), Math.round(y));

            // Draw text decorations
            if (item.props.underline) {
                ctx.save();
                ctx.strokeStyle = ctx.fillStyle;
                ctx.lineWidth = 1;
                ctx.beginPath();
                ctx.moveTo(currentX, y + 2);
                ctx.lineTo(currentX + item.width, y + 2);
                ctx.stroke();
                ctx.restore();
            }

            if (item.props.strike) {
                ctx.save();
                ctx.strokeStyle = ctx.fillStyle;
                ctx.lineWidth = 1;
                ctx.beginPath();
                // Use scaled font size for proper strikethrough positioning
                const scaleFactor = this.getTextScaleFactor();
                const scaledFontSize = (item.props.fontSize || 12) * scaleFactor;
                const strikeY = y - scaledFontSize * 0.3;
                ctx.moveTo(currentX, strikeY);
                ctx.lineTo(currentX + item.width, strikeY);
                ctx.stroke();
                ctx.restore();
            }

            currentX += item.width;
        }
    }

    /**
     * Parse paragraph properties from paragraph data with enhanced property handling and inheritance
     */
    parseParagraphProperties(paragraph, shape = null) {
        // Start with default properties
        const props = {
            align: 'left',
            lineHeight: 120, // 120% default
            spaceBefore: 0,
            spaceAfter: 0,
            fontSize: 18,
            fontFamily: 'Arial',
            color: { r: 0, g: 0, b: 0, a: 255 },
            bold: false,
            italic: false,
            bullet: null, // Bullet properties
            indent: {
                left: 0,   // marL - margin left
                hanging: 0 // indent - hanging indent (negative indent)
            }
        };

        // Get paragraph level (default to 0)
        const paragraphLevel = (paragraph.properties && typeof paragraph.properties.level === 'number') 
            ? paragraph.properties.level 
            : 0;

        // Apply inheritance from layout/master styles with level awareness
        const inheritedProps = this.resolveInheritedTextProperties(shape, 'paragraph', paragraphLevel);
        
        // Apply inherited properties, but handle indent specially
        for (const key in inheritedProps) {
            if (key === 'leftMargin') {
                props.indent.left = inheritedProps.leftMargin;
            } else if (key === 'hangingIndent') {
                props.indent.hanging = inheritedProps.hangingIndent;
            } else if (key !== 'leftMargin' && key !== 'hangingIndent') {
                props[key] = inheritedProps[key];
            }
        }

        // Map inherited alignment values to CSS values and only override when paragraph explicitly sets alignment
        if (inheritedProps.align) {
            const alignmentMap = {
                'l': 'left',
                'left': 'left',
                'ctr': 'center',
                'center': 'center',
                'r': 'right',
                'right': 'right',
                'just': 'justify',
                'justify': 'justify'
            };
            props.align = alignmentMap[inheritedProps.align] || inheritedProps.align;
        }

        // Apply placeholder-specific defaults - only formatting, not sizes or alignment
        // Let the layout/master define the font sizes and alignment properly
        if (shape && shape.placeholder) {
            switch (shape.placeholder.type) {
                case 'title':
                case 'ctrTitle':
                    // Don't override bold — inherit from master txStyles (OOXML default is false)
                    break;
                case 'subTitle':
                    // Don't override alignment - let layout define it
                    break;
                case 'body':
                    // Don't override alignment - let layout define it
                    break;
                case 'obj':
                    // No specific overrides for obj placeholders
                    break;
            }
        }

        if (paragraph.properties) {
            const pProps = paragraph.properties;

            // Alignment - override inherited only if explicitly set on paragraph
            if (pProps.align) {
                // Map PPTX alignment values to CSS values
                const alignmentMap = {
                    'l': 'left',
                    'left': 'left',
                    'ctr': 'center',
                    'center': 'center',
                    'r': 'right',
                    'right': 'right',
                    'just': 'justify',
                    'justify': 'justify'
                };
                props.align = alignmentMap[pProps.align] || pProps.align;
            }
            // If explicit properties have 'left' alignment, preserve inherited alignment
            // This handles cases where parsing creates default 'left' alignment that shouldn't override layout inheritance

            // Line height
            if (pProps.lineHeight !== undefined) {
                props.lineHeight = pProps.lineHeight;
            }
            if (pProps.lineHeightPoints !== undefined) {
                props.lineHeightPoints = pProps.lineHeightPoints;
            }

            // Spacing — only override inherited when slide explicitly sets a non-zero value
            if (pProps.spacing) {
                if (pProps.spacing.before) {
                    props.spaceBefore = pProps.spacing.before;
                }
                if (pProps.spacing.after) {
                    props.spaceAfter = pProps.spacing.after;
                }
            }

            // Default font properties from paragraph
            if (pProps.fontSize !== undefined) {
                props.fontSize = pProps.fontSize;
            }

            if (pProps.fontFamily) {
                props.fontFamily = pProps.fontFamily;
            }

            if (pProps.color) {
                props.color = pProps.color;
            }

            if (pProps.bold !== undefined) {
                props.bold = pProps.bold;
            }

            if (pProps.italic !== undefined) {
                props.italic = pProps.italic;
            }

            // Bullet properties
            if (pProps.bullet) {
                props.bullet = pProps.bullet;
            }

            // Indentation properties - only override if explicitly set in paragraph
            if (pProps.indent) {
                props.indent.left = pProps.indent.left;
                props.indent.hanging = pProps.indent.hanging;
            }
        }

        return props;
    }

    /**
     * Parse run properties from run data with enhanced property handling and inheritance
     */
    parseRunProperties(run, paraProps, shape = null) {
        // Start with default properties - use theme-aware defaults for layout slides
        const isLayoutSlide = this.currentSlide?.type === 'layout';
        const defaultColor = isLayoutSlide ? 
            { r: 68, g: 68, b: 68, a: 255 } : // Dark gray for layout slides
            { r: 0, g: 0, b: 0, a: 255 }; // Black for regular slides
            
        const props = {
            fontSize: 18, // Default font size matching PPTX standards (18pt)
            fontFamily: 'Arial',
            color: defaultColor,
            bold: false,
            italic: false,
            underline: false,
            strike: false,
            verticalAlign: 'normal',
            letterSpacing: 0,
            cap: 'none' // Text capitalization
        };

        // Apply inheritance from layout/master styles
        const inheritedProps = this.resolveInheritedTextProperties(shape, 'run');
        Object.assign(props, inheritedProps);

        // Apply paragraph properties (override inherited)
        if (paraProps) {
            if (paraProps.fontSize !== undefined) {props.fontSize = paraProps.fontSize;}
            if (paraProps.fontFamily) {props.fontFamily = paraProps.fontFamily;}
            if (paraProps.color) {props.color = paraProps.color;}
            if (paraProps.bold !== undefined) {props.bold = paraProps.bold;}
            if (paraProps.italic !== undefined) {props.italic = paraProps.italic;}
        }

        // Apply run-specific properties (highest priority)
        if (run.properties) {
            const rProps = run.properties;

            // Font size - ensure proper scaling
            if (rProps.fontSize !== undefined) {
                props.fontSize = rProps.fontSize;
            }

            // Font family
            if (rProps.fontFamily) {
                props.fontFamily = rProps.fontFamily;
            }

            // Color
            if (rProps.color) {
                props.color = rProps.color;
            }

            // Bold
            if (rProps.bold !== undefined) {
                props.bold = rProps.bold;
            }

            // Italic
            if (rProps.italic !== undefined) {
                props.italic = rProps.italic;
            }

            // Underline
            if (rProps.underline !== undefined) {
                props.underline = rProps.underline;
            }

            // Strike
            if (rProps.strike !== undefined) {
                props.strike = rProps.strike;
            }

            // Vertical alignment (superscript/subscript)
            if (rProps.verticalAlign) {
                props.verticalAlign = rProps.verticalAlign;
            }

            // Letter spacing
            if (rProps.letterSpacing !== undefined) {
                props.letterSpacing = rProps.letterSpacing;
            }

            // Text capitalization
            if (rProps.cap) {
                props.cap = rProps.cap;
            }

            // Highlight
            if (rProps.highlight) {
                props.highlight = rProps.highlight;
            }

            // Text effects (shadow, glow)
            if (rProps.effectLst) {
                props.effectLst = rProps.effectLst;
            }
        }

        return props;
    }

    /**
     * Convert EMU to pixels with proper DPI handling for layout slides
     * Fixed: Use consistent 96 DPI to match CoordinateTransform utility
     */
    emuToPixels(emu) {
        // Use CoordinateTransform's standard conversion for consistency
        return CoordinateTransform.emuToPixels(emu, 96);
    }


    /**
     * Convert color object to CSS string with enhanced color handling
     */
    colorToRgbAlternative(color) {
        if (typeof color === 'string') {
            // Handle hex colors
            if (color.startsWith('#')) {return color;}
            // Handle named colors
            if (color.match(/^[a-z]+$/i)) {return color;}
            return color;
        }

        if (!color) {return 'rgb(0, 0, 0)';}

        // Handle different color object formats
        let r = 0, g = 0, b = 0, a = 255;

        if (typeof color === 'object') {
            // Handle {r, g, b, a} format
            r = color.r !== undefined ? color.r : (color.R !== undefined ? color.R : 0);
            g = color.g !== undefined ? color.g : (color.G !== undefined ? color.G : 0);
            b = color.b !== undefined ? color.b : (color.B !== undefined ? color.B : 0);
            a = color.a !== undefined ? color.a : (color.A !== undefined ? color.A : 255);

            // Ensure values are in valid range
            r = Math.max(0, Math.min(255, Math.round(r)));
            g = Math.max(0, Math.min(255, Math.round(g)));
            b = Math.max(0, Math.min(255, Math.round(b)));
            a = Math.max(0, Math.min(255, Math.round(a)));
        }

        // Return appropriate format
        if (a !== undefined && a !== 255) {
            return `rgba(${r}, ${g}, ${b}, ${a / 255})`;
        }
        return `rgb(${r}, ${g}, ${b})`;
    }

    /**
     * Utility methods
     */
    isHiddenShape(shape) {
        return shape.hidden === true;
    }

    getShapeTransform(shape) {
        // Return null for now - would implement transform parsing
        return null;
    }

    getShapeBounds(shape) {
        // PRIORITY for graphicFrame: use parsed position (xfrm) on the frame itself
        // Tables/images often live in a graphicFrame whose transform is parsed into `position`
        // Prefer this over inherited spPr.xfrm to avoid picking up placeholder/layout transforms
        if (shape && shape.type === 'graphicFrame' && shape.position) {
            const pos = shape.position;
            const x = pos.x || 0;
            const y = pos.y || 0;
            const width = pos.width || 914400; // default 1 inch in EMU
            const height = pos.height || 914400;

            if (this.coordinateSystem) {
                const { slideWidthEMU, slideHeightEMU } = this.coordinateSystem;

                const validSlideWidth = slideWidthEMU || 9144000;
                const validSlideHeight = slideHeightEMU || 6858000;

                const normalizedBounds = {
                    x: x / validSlideWidth,
                    y: y / validSlideHeight,
                    w: width / validSlideWidth,
                    h: height / validSlideHeight
                };

                const canvas = this.graphics.context.canvas;
                let canvasWidth, canvasHeight;
                if (canvas.style.width && canvas.style.height) {
                    canvasWidth = parseFloat(canvas.style.width);
                    canvasHeight = parseFloat(canvas.style.height);
                } else {
                    canvasWidth = canvas.width / (window.devicePixelRatio || 1);
                    canvasHeight = canvas.height / (window.devicePixelRatio || 1);
                }

                const slideAspectRatio = validSlideWidth / validSlideHeight;
                const canvasAspectRatio = canvasWidth / canvasHeight;

                let slideCanvasWidth, slideCanvasHeight, slideOffsetX, slideOffsetY;
                if (slideAspectRatio > canvasAspectRatio) {
                    slideCanvasWidth = canvasWidth;
                    slideCanvasHeight = canvasWidth / slideAspectRatio;
                    slideOffsetX = 0;
                    slideOffsetY = (canvasHeight - slideCanvasHeight) / 2;
                } else {
                    slideCanvasWidth = canvasHeight * slideAspectRatio;
                    slideCanvasHeight = canvasHeight;
                    slideOffsetX = (canvasWidth - slideCanvasWidth) / 2;
                    slideOffsetY = 0;
                }

                return {
                    x: Math.round(slideOffsetX + (normalizedBounds.x * slideCanvasWidth)),
                    y: Math.round(slideOffsetY + (normalizedBounds.y * slideCanvasHeight)),
                    w: Math.round(normalizedBounds.w * slideCanvasWidth),
                    h: Math.round(normalizedBounds.h * slideCanvasHeight)
                };
            }

            const emuToPx = emu => emu / 914400 * 96;
            return {
                x: emuToPx(x),
                y: emuToPx(y),
                w: emuToPx(width),
                h: emuToPx(height)
            };
        }

        // First: honor spPr.xfrm (including inherited from layout/master)
        if (shape.spPr && shape.spPr.xfrm) {
            const xfrm = shape.spPr.xfrm;
            // Support multiple possible shapes of xfrm
            const x = (xfrm.x !== undefined ? xfrm.x : (xfrm.off && xfrm.off.x) !== undefined ? xfrm.off.x : 0);
            const y = (xfrm.y !== undefined ? xfrm.y : (xfrm.off && xfrm.off.y) !== undefined ? xfrm.off.y : 0);
            const width = (xfrm.width !== undefined ? xfrm.width : (xfrm.cx !== undefined ? xfrm.cx : (xfrm.ext && xfrm.ext.cx) !== undefined ? xfrm.ext.cx : 914400));
            const height = (xfrm.height !== undefined ? xfrm.height : (xfrm.cy !== undefined ? xfrm.cy : (xfrm.ext && xfrm.ext.cy) !== undefined ? xfrm.ext.cy : 914400));

            // Apply coordinate system transformation if available
            if (this.coordinateSystem) {
                const { slideWidthEMU, slideHeightEMU } = this.coordinateSystem;

                // Ensure we have valid slide dimensions
                const validSlideWidth = slideWidthEMU || 9144000; // Default slide width in EMU
                const validSlideHeight = slideHeightEMU || 6858000; // Default slide height in EMU

                // Convert to normalized coords
                const normalizedBounds = {
                    x: x / validSlideWidth,
                    y: y / validSlideHeight,
                    w: width / validSlideWidth,
                    h: height / validSlideHeight
                };

                // Get logical canvas size
                const canvas = this.graphics.context.canvas;
                let canvasWidth, canvasHeight;
                if (canvas.style.width && canvas.style.height) {
                    canvasWidth = parseFloat(canvas.style.width);
                    canvasHeight = parseFloat(canvas.style.height);
                } else {
                    canvasWidth = canvas.width / (window.devicePixelRatio || 1);
                    canvasHeight = canvas.height / (window.devicePixelRatio || 1);
                }

                // Slide area within canvas (maintain aspect ratio)
                const slideAspectRatio = validSlideWidth / validSlideHeight;
                const canvasAspectRatio = canvasWidth / canvasHeight;

                let slideCanvasWidth, slideCanvasHeight, slideOffsetX, slideOffsetY;
                if (slideAspectRatio > canvasAspectRatio) {
                    slideCanvasWidth = canvasWidth;
                    slideCanvasHeight = canvasWidth / slideAspectRatio;
                    slideOffsetX = 0;
                    slideOffsetY = (canvasHeight - slideCanvasHeight) / 2;
                } else {
                    slideCanvasWidth = canvasHeight * slideAspectRatio;
                    slideCanvasHeight = canvasHeight;
                    slideOffsetX = (canvasWidth - slideCanvasWidth) / 2;
                    slideOffsetY = 0;
                }

                return {
                    x: Math.round(slideOffsetX + (normalizedBounds.x * slideCanvasWidth)),
                    y: Math.round(slideOffsetY + (normalizedBounds.y * slideCanvasHeight)),
                    w: Math.round(normalizedBounds.w * slideCanvasWidth),
                    h: Math.round(normalizedBounds.h * slideCanvasHeight)
                };
            }

            // Fallback to 96 DPI conversion
            const emuToPx = emu => emu / 914400 * 96;
            return {
                x: emuToPx(x),
                y: emuToPx(y),
                w: emuToPx(width),
                h: emuToPx(height)
            };
        }

        // Check for shape.properties.transform first (test page format)
        if (shape.properties && shape.properties.transform) {
            const transform = shape.properties.transform;
            const x = transform.x || 0;
            const y = transform.y || 0;
            const width = transform.width || 914400; // Default 1 inch in EMU
            const height = transform.height || 914400;


            // Apply coordinate system transformation if available
            if (this.coordinateSystem) {
                const { scale, offsetX, offsetY, slideWidthEMU, slideHeightEMU } = this.coordinateSystem;


                // Ensure we have valid slide dimensions
                const validSlideWidth = slideWidthEMU || 9144000; // Default slide width in EMU
                const validSlideHeight = slideHeightEMU || 6858000; // Default slide height in EMU


                // Convert shape bounds from EMU to normalized coordinates (0-1)
                const normalizedBounds = {
                    x: x / validSlideWidth,
                    y: y / validSlideHeight,
                    w: width / validSlideWidth,
                    h: height / validSlideHeight
                };


                // Get logical canvas dimensions (not high-DPI scaled dimensions)
                const canvas = this.graphics.context.canvas;
                let canvasWidth, canvasHeight;
                if (canvas.style.width && canvas.style.height) {
                    canvasWidth = parseFloat(canvas.style.width);
                    canvasHeight = parseFloat(canvas.style.height);
                } else {
                    canvasWidth = canvas.width / (window.devicePixelRatio || 1);
                    canvasHeight = canvas.height / (window.devicePixelRatio || 1);
                }

                // Calculate slide area within canvas (maintaining aspect ratio)
                const slideAspectRatio = validSlideWidth / validSlideHeight;
                const canvasAspectRatio = canvasWidth / canvasHeight;

                let slideCanvasWidth, slideCanvasHeight, slideOffsetX, slideOffsetY;

                if (slideAspectRatio > canvasAspectRatio) {
                    // Slide is wider than canvas - fit to width
                    slideCanvasWidth = canvasWidth;
                    slideCanvasHeight = canvasWidth / slideAspectRatio;
                    slideOffsetX = 0;
                    slideOffsetY = (canvasHeight - slideCanvasHeight) / 2;
                } else {
                    // Slide is taller than canvas - fit to height
                    slideCanvasWidth = canvasHeight * slideAspectRatio;
                    slideCanvasHeight = canvasHeight;
                    slideOffsetX = (canvasWidth - slideCanvasWidth) / 2;
                    slideOffsetY = 0;
                }

                // Apply final transformation to canvas coordinates with precision
                const finalBounds = {
                    x: Math.round(slideOffsetX + (normalizedBounds.x * slideCanvasWidth)),
                    y: Math.round(slideOffsetY + (normalizedBounds.y * slideCanvasHeight)),
                    w: Math.round(normalizedBounds.w * slideCanvasWidth),
                    h: Math.round(normalizedBounds.h * slideCanvasHeight)
                };


                return finalBounds;
            }

            // Fallback: direct EMU to pixels conversion (96 DPI)
            const emuToPx = emu => emu / 914400 * 96;
            const fallbackBounds = {
                x: emuToPx(x),
                y: emuToPx(y),
                w: emuToPx(width),
                h: emuToPx(height)
            };

            return fallbackBounds;
        }

        // Fallback: some shapes (e.g., graphicFrame) store transform as `position` or `transform` directly on the shape
        // These values are in EMU and should be mapped through the same coordinate system as above
        const pos = shape.position || shape.transform;
        if (pos && (pos.x !== undefined || pos.y !== undefined || pos.width !== undefined || pos.height !== undefined)) {
            const x = pos.x || 0;
            const y = pos.y || 0;
            const width = pos.width || 914400; // default 1 inch
            const height = pos.height || 914400;

            if (this.coordinateSystem) {
                const { slideWidthEMU, slideHeightEMU } = this.coordinateSystem;

                const validSlideWidth = slideWidthEMU || 9144000;
                const validSlideHeight = slideHeightEMU || 6858000;

                const normalizedBounds = {
                    x: x / validSlideWidth,
                    y: y / validSlideHeight,
                    w: width / validSlideWidth,
                    h: height / validSlideHeight
                };

                const canvas = this.graphics.context.canvas;
                let canvasWidth, canvasHeight;
                if (canvas.style.width && canvas.style.height) {
                    canvasWidth = parseFloat(canvas.style.width);
                    canvasHeight = parseFloat(canvas.style.height);
                } else {
                    canvasWidth = canvas.width / (window.devicePixelRatio || 1);
                    canvasHeight = canvas.height / (window.devicePixelRatio || 1);
                }

                const slideAspectRatio = validSlideWidth / validSlideHeight;
                const canvasAspectRatio = canvasWidth / canvasHeight;

                let slideCanvasWidth, slideCanvasHeight, slideOffsetX, slideOffsetY;
                if (slideAspectRatio > canvasAspectRatio) {
                    slideCanvasWidth = canvasWidth;
                    slideCanvasHeight = canvasWidth / slideAspectRatio;
                    slideOffsetX = 0;
                    slideOffsetY = (canvasHeight - slideCanvasHeight) / 2;
                } else {
                    slideCanvasWidth = canvasHeight * slideAspectRatio;
                    slideCanvasHeight = canvasHeight;
                    slideOffsetX = (canvasWidth - slideCanvasWidth) / 2;
                    slideOffsetY = 0;
                }

                return {
                    x: Math.round(slideOffsetX + (normalizedBounds.x * slideCanvasWidth)),
                    y: Math.round(slideOffsetY + (normalizedBounds.y * slideCanvasHeight)),
                    w: Math.round(normalizedBounds.w * slideCanvasWidth),
                    h: Math.round(normalizedBounds.h * slideCanvasHeight)
                };
            }

            const emuToPx = emu => emu / 914400 * 96;
            return {
                x: emuToPx(x),
                y: emuToPx(y),
                w: emuToPx(width),
                h: emuToPx(height)
            };
        }

        // Use bounds if available (original format)
        if (shape.bounds) {
            const l = shape.bounds.l || 0;
            const t = shape.bounds.t || 0;
            const r = (typeof shape.bounds.r === 'number') ? shape.bounds.r : (l + 1000000);
            const b = (typeof shape.bounds.b === 'number') ? shape.bounds.b : (t + 1000000);

            // Apply coordinate system transformation if available
            if (this.coordinateSystem) {
                const { scale, offsetX, offsetY, slideWidthEMU, slideHeightEMU } = this.coordinateSystem;

                // Ensure we have valid slide dimensions
                const validSlideWidth = slideWidthEMU || 9144000; // Default slide width in EMU
                const validSlideHeight = slideHeightEMU || 6858000; // Default slide height in EMU

                // Convert shape bounds from EMU to normalized coordinates (0-1)
                const normalizedBounds = {
                    x: l / validSlideWidth,
                    y: t / validSlideHeight,
                    w: (r - l) / validSlideWidth,
                    h: (b - t) / validSlideHeight
                };

                // Get logical canvas dimensions (not high-DPI scaled dimensions)
                const canvas = this.graphics.context.canvas;
                let canvasWidth, canvasHeight;
                if (canvas.style.width && canvas.style.height) {
                    canvasWidth = parseFloat(canvas.style.width);
                    canvasHeight = parseFloat(canvas.style.height);
                } else {
                    canvasWidth = canvas.width / (window.devicePixelRatio || 1);
                    canvasHeight = canvas.height / (window.devicePixelRatio || 1);
                }

                // Calculate slide area within canvas (maintaining aspect ratio)
                const slideAspectRatio = validSlideWidth / validSlideHeight;
                const canvasAspectRatio = canvasWidth / canvasHeight;

                let slideCanvasWidth, slideCanvasHeight, slideOffsetX, slideOffsetY;

                if (slideAspectRatio > canvasAspectRatio) {
                    // Slide is wider than canvas - fit to width
                    slideCanvasWidth = canvasWidth;
                    slideCanvasHeight = canvasWidth / slideAspectRatio;
                    slideOffsetX = 0;
                    slideOffsetY = (canvasHeight - slideCanvasHeight) / 2;
                } else {
                    // Slide is taller than canvas - fit to height
                    slideCanvasWidth = canvasHeight * slideAspectRatio;
                    slideCanvasHeight = canvasHeight;
                    slideOffsetX = (canvasWidth - slideCanvasWidth) / 2;
                    slideOffsetY = 0;
                }

                // Apply final transformation to canvas coordinates with precision
                const finalBounds = {
                    x: Math.round(slideOffsetX + (normalizedBounds.x * slideCanvasWidth)),
                    y: Math.round(slideOffsetY + (normalizedBounds.y * slideCanvasHeight)),
                    w: Math.round(normalizedBounds.w * slideCanvasWidth),
                    h: Math.round(normalizedBounds.h * slideCanvasHeight)
                };


                return finalBounds;
            }

            // Fallback: direct EMU to pixels conversion (96 DPI)
            const emuToPx = emu => emu / 914400 * 96;
            const fallbackBounds = {
                x: emuToPx(l),
                y: emuToPx(t),
                w: emuToPx(r - l),
                h: emuToPx(b - t)
            };

            return fallbackBounds;
        }

        // No bounds available - return default
        const defaultBounds = { x: 50, y: 50, w: 100, h: 50 };
        return defaultBounds;
    }
    getShapeFillColor(shape) {
        // CRITICAL FIX: Check for explicit no fill first before any processing
        if (shape.fill && shape.fill.type === 'none') {
            return null; // Explicit no fill
        }
        if (shape.properties && shape.properties.fill && shape.properties.fill.type === 'none') {
            return null; // Explicit no fill from properties
        }
        
        // CRITICAL FIX: Ensure inheritance has run before getting fill color
        if (this.processor && this.processor.applyPropertyInheritance && 
            (!shape.fill || shape.fill === null) && 
            (!shape.properties?.fill || shape.properties?.fill === null)) {
            // Pass current slide context for inheritance
            this.processor.applyPropertyInheritance(shape, this.currentSlide);
        }

        // Check again for no fill after inheritance
        if (shape.fill && shape.fill.type === 'none') {
            return null; // No fill after inheritance
        }
        if (shape.properties && shape.properties.fill && shape.properties.fill.type === 'none') {
            return null; // No fill from properties after inheritance
        }

        // IMPORTANT: Only return null for text placeholders that explicitly have no fill
        // Allow text placeholders with actual fill properties to be rendered
        if (shape.textBody && shape.placeholder && !shape.fill && !shape.properties?.fill && !shape.fillColor) {
            return null; // No fill for text placeholders with explicit null fill
        }

        // Use the processor's enhanced color processing if available
        if (this.processor && this.processor.getShapeFillColor) {
            const processorResult = this.processor.getShapeFillColor(shape);
            // Return processor result even if null (null means explicit no fill)
            if (processorResult !== undefined) {
                return processorResult;
            }
        }

        // CRITICAL FIX: Check preserved style first for fillRef
        const styleToUse = shape.preservedStyle || shape.style;
        if (styleToUse && styleToUse.fillRef) {
            // IMPORTANT: Skip fillRef processing only for text placeholders with truly no fill
            if (shape.textBody && shape.placeholder && !shape.fill && !shape.properties?.fill && !shape.fillColor) {
                return null;
            }
            
            // Try to resolve style-based color using processor if available
            if (this.processor && this.processor.resolveStyleFillColor) {
                const styleColor = this.processor.resolveStyleFillColor(styleToUse.fillRef);
                if (styleColor) {
                    return styleColor;
                }
            }
            // Direct color from fillRef if available
            if (styleToUse.fillRef.color) {
                return styleToUse.fillRef.color;
            }
            
            // For layout slides, use no fill for text boxes only if they don't have explicit fill properties
            if (this.currentSlide?.type === 'layout' && shape.textBody && !shape.fill && !shape.properties?.fill) {
                return null;
            }
            
            // Default fill color is solid white for unresolved fillRef
            return '#FFFFFF';
        }

        // IMPORTANT: For text-only shapes that appear to be summary boxes or standalone text,
        // check if PowerPoint would render the background. Light colored backgrounds on text 
        // shapes are often ignored by PowerPoint to match its rendering behavior.
        if (shape.textBody && !shape.placeholder && shape.fillColor) {
            // Check if this is a light background color that PowerPoint typically ignores on text shapes
            const fillColor = shape.fillColor;
            if (typeof fillColor === 'string') {
                // Parse hex color to check if it's a light background
                const hex = fillColor.replace('#', '');
                if (hex.length === 6) {
                    const r = parseInt(hex.substr(0, 2), 16);
                    const g = parseInt(hex.substr(2, 2), 16);
                    const b = parseInt(hex.substr(4, 2), 16);
                    
                    // Check if this is a very light green background (like #E8F8F5)
                    // PowerPoint often doesn't render such subtle backgrounds on text shapes
                    if (r > 200 && g > 240 && b > 240) {
                        return null;
                    }
                }
            }
            // Also check object-format colors
            else if (typeof fillColor === 'object' && fillColor.r !== undefined) {
                // Check if this is a very light green background
                if (fillColor.r > 200 && fillColor.g > 240 && fillColor.b > 240) {
                    return null;
                }
            }
        }

        // Fallback to basic color extraction
        if (shape.fill && shape.fill.type === 'solid' && shape.fill.color) {
            return shape.fill.color;
        }
        if (shape.properties && shape.properties.fill && shape.properties.fill.color) {
            return shape.properties.fill.color;
        }
        return shape.fillColor || null;
    }

    getShapeStrokeColor(shape) {
        // Only return a stroke when it explicitly exists on the shape
        const hasExplicitStroke = !!(
            (shape && shape.spPr && shape.spPr.ln) ||
            (shape && shape.properties && shape.properties.stroke) ||
            (shape && shape.stroke)
        );

        if (!hasExplicitStroke) {
            return null;
        }

        // Use the processor's enhanced color processing if available
        if (this.processor && this.processor.getShapeStrokeColor) {
            const result = this.processor.getShapeStrokeColor(shape);
            return result;
        }

        // CRITICAL FIX: Check preserved style first for lnRef
        const styleToUse = shape.preservedStyle || shape.style;
        if (styleToUse && styleToUse.lnRef) {
            // Try to resolve style-based color using processor if available
            if (this.processor && this.processor.resolveStyleStrokeColor) {
                const styleColor = this.processor.resolveStyleStrokeColor(styleToUse.lnRef);
                if (styleColor) {
                    return styleColor;
                }
            }
            // Direct color from lnRef if available
            if (styleToUse.lnRef.color) {
                return styleToUse.lnRef.color;
            }
        }

        // Fallback to basic color extraction
        if (shape.stroke && shape.stroke.color) {
            return shape.stroke.color;
        }
        if (shape.properties && shape.properties.stroke && shape.properties.stroke.color) {
            return shape.properties.stroke.color;
        }
        const fallbackResult = shape.strokeColor || null;
        return fallbackResult;
    }

    getShapeLineWidth(shape) {
        // Use the processor's enhanced line width processing if available
        if (this.processor && this.processor.getShapeLineWidth) {
            return this.processor.getShapeLineWidth(shape);
        }

        // Only return a width when stroke is explicitly defined
        const hasExplicitStroke = !!(
            (shape && shape.spPr && shape.spPr.ln) ||
            (shape && shape.properties && shape.properties.stroke) ||
            (shape && shape.stroke)
        );
        if (!hasExplicitStroke) {return 0;}

        // Fallback to basic line width extraction
        if (shape.stroke && shape.stroke.width) {
            return shape.stroke.width / 914400 * 25.4; // Convert EMU to MM
        }
        if (shape.properties && shape.properties.stroke && shape.properties.stroke.width) {
            return shape.properties.stroke.width / 914400 * 25.4;
        }
        // Extract from spPr.ln if available
        if (shape.spPr && shape.spPr.ln && shape.spPr.ln.w) {
            return shape.spPr.ln.w / 914400 * 25.4;
        }
        return 0;
    }

    getShapeStrokeInfo(shape) {
        // Use the processor's enhanced stroke info processing if available
        if (this.processor && this.processor.getShapeStrokeInfo) {
            return this.processor.getShapeStrokeInfo(shape);
        }

        // Fallback to basic stroke info extraction
        if (shape.stroke) {
            return shape.stroke;
        }
        if (shape.properties && shape.properties.stroke) {
            return shape.properties.stroke;
        }
        return null;
    }

    getShapeStrokeInfoLegacy(shape) {
        // Use the processor's enhanced stroke info processing if available
        if (this.processor && this.processor.getShapeStrokeInfo) {
            return this.processor.getShapeStrokeInfo(shape);
        }

        // Fallback to basic stroke info extraction
        if (shape.stroke) {
            return shape.stroke;
        }
        if (shape.properties && shape.properties.stroke) {
            return shape.properties.stroke;
        }
        return null;
    }

    getParagraphProperties(paragraph) {
        return paragraph.properties || {};
    }

    getRunProperties(run, paraProps) {
        return { ...paraProps, ...(run.properties || {}) };
    }

    parseColor(color) {
        if (typeof color === 'string') {
            if (color.startsWith('#')) {
                const hex = color.slice(1);
                return {
                    r: parseInt(hex.slice(0, 2), 16),
                    g: parseInt(hex.slice(2, 4), 16),
                    b: parseInt(hex.slice(4, 6), 16),
                    a: 255
                };
            }
        }
        if (color && typeof color === 'object') {
            return color;
        }
        return null;
    }

    clipTextRect(bounds) {
        this.graphics.AddClipRect(bounds.x, bounds.y, bounds.w, bounds.h);
    }

    /**
     * Get current canvas bounds for gradient calculation
     */
    getCanvasBounds(ctx) {
        const transform = ctx.getTransform();
        // Use the current shape bounds if available, otherwise use canvas dimensions
        // Use logical dimensions for bounds calculation
        let logicalWidth, logicalHeight;
        if (ctx.canvas.style.width && ctx.canvas.style.height) {
            logicalWidth = parseFloat(ctx.canvas.style.width);
            logicalHeight = parseFloat(ctx.canvas.style.height);
        } else {
            logicalWidth = ctx.canvas.width / (window.devicePixelRatio || 1);
            logicalHeight = ctx.canvas.height / (window.devicePixelRatio || 1);
        }
        return this.currentShapeBounds || { x: 0, y: 0, w: logicalWidth, h: logicalHeight };
    }

    /**
     * Create canvas gradient from gradient definition
     */
    createCanvasGradient(ctx, bounds, gradientDef) {
        if (!gradientDef || !gradientDef.stops || gradientDef.stops.length === 0) {
            return null;
        }

        let gradient;
        if (gradientDef.type === 'linear') {
            // Linear gradient
            const angle = gradientDef.angle || 0;
            const radians = (angle * Math.PI) / 180;

            const x1 = bounds.x + bounds.w / 2 - (Math.cos(radians) * bounds.w) / 2;
            const y1 = bounds.y + bounds.h / 2 - (Math.sin(radians) * bounds.h) / 2;
            const x2 = bounds.x + bounds.w / 2 + (Math.cos(radians) * bounds.w) / 2;
            const y2 = bounds.y + bounds.h / 2 + (Math.sin(radians) * bounds.h) / 2;

            gradient = ctx.createLinearGradient(x1, y1, x2, y2);
        } else if (gradientDef.type === 'radial') {
            // Radial gradient
            const centerX = bounds.x + bounds.w / 2;
            const centerY = bounds.y + bounds.h / 2;
            const radius = Math.max(bounds.w, bounds.h) / 2;

            gradient = ctx.createRadialGradient(centerX, centerY, 0, centerX, centerY, radius);
        } else {
            return null;
        }

        // Add color stops
        for (const stop of gradientDef.stops) {
            const colorString = this.colorToRgb(stop.color);
            gradient.addColorStop(stop.position, colorString);
        }

        return gradient;
    }

    /**
     * Convert color object to RGB string
     */
    colorToRgbAlternative2(color) {
        if (typeof color === 'string') {return color;}
        if (color && typeof color === 'object') {
            const r = color.r || color.R || 0;
            const g = color.g || color.G || 0;
            const b = color.b || color.B || 0;
            const a = (color.a !== undefined ? color.a : (color.A !== undefined ? color.A : 255)) / 255;
            return `rgba(${r}, ${g}, ${b}, ${a})`;
        }
        return '#000000';
    }

    /**
     * Find layout shape that matches a placeholder
     */
    findLayoutShapeForPlaceholder(layoutShapes, shape) {
        if (!shape || !shape.placeholder || !layoutShapes) {
            return null;
        }

        const placeholder = shape.placeholder;

        // Find shape with matching placeholder type and index
        for (const layoutShape of layoutShapes) {
            if (layoutShape.placeholder) {
                const layoutPh = layoutShape.placeholder;

                // Normalize idx values to numbers if possible
                const slideIdx = (placeholder.idx === undefined || placeholder.idx === null || isNaN(Number(placeholder.idx))) ? undefined : Number(placeholder.idx);
                const layoutIdx = (layoutPh.idx === undefined || layoutPh.idx === null || isNaN(Number(layoutPh.idx))) ? undefined : Number(layoutPh.idx);

                // Helper to compare types including title/ctrTitle equivalence
                const typesMatch = (a, b) => {
                    if (a === b) {return true;}
                    const pair = new Set([a, b]);
                    // Treat title/ctrTitle as equivalent and body/obj as equivalent
                    if (pair.has('title') && pair.has('ctrTitle')) {return true;}
                    if (pair.has('body') && pair.has('obj')) {return true;}
                    return false;
                };

                // Primary: match by type and idx when both defined
                if (typesMatch(layoutPh.type, placeholder.type) && slideIdx !== undefined && layoutIdx !== undefined && slideIdx === layoutIdx) {
                    return layoutShape;
                }

                // Secondary: if either idx is missing, match on type only
                if (typesMatch(layoutPh.type, placeholder.type) && (slideIdx === undefined || layoutIdx === undefined)) {
                    return layoutShape;
                }
            }
        }

        return null;
    }

    /**
     * Resolve inherited text properties from layout and master styles
     */
    resolveInheritedTextProperties(shape, level = 'run', paragraphLevel = 0) {
        const inheritedProps = {};

        try {
            const slideContext = this.buildSlideContext();

            // Determine text style type based on shape placeholder type
            let styleType = 'otherStyle'; // Default
            if (shape && shape.placeholder) {
                switch (shape.placeholder.type) {
                    case 'title':
                    case 'ctrTitle':
                        styleType = 'titleStyle';
                        break;
                    case 'body':
                    case 'subTitle':
                    case 'obj':
                        styleType = 'bodyStyle';
                        break;
                    default:
                        styleType = 'otherStyle';
                }
            }

            // Apply theme-based defaults first
            if (slideContext.theme) {
                if (slideContext.theme.fonts) {
                    // Use minor font for body text, major font for titles
                    const isTitle = styleType === 'titleStyle';
                    inheritedProps.fontFamily = isTitle ?
                        slideContext.theme.fonts.major :
                        slideContext.theme.fonts.minor;
                }

                // Apply theme colors if available
                if (slideContext.theme.colors && slideContext.theme.colors.text1) {
                    inheritedProps.color = this.parseColorFromTheme(slideContext.theme.colors.text1);
                }
            }

            // Try to get style from master first (highest precedence for inheritance)
            if (slideContext.master && slideContext.master.txStyles) {
                const masterStyle = slideContext.master.txStyles[styleType];
                if (masterStyle) {
                    const masterProps = this.extractTextPropertiesFromStyle(masterStyle, level, paragraphLevel);
                    Object.assign(inheritedProps, masterProps);
                }
            }

            // Try to get style from layout (can override master)
            if (slideContext.layout && slideContext.layout.txStyles) {
                const layoutStyle = slideContext.layout.txStyles[styleType];
                if (layoutStyle) {
                    const layoutProps = this.extractTextPropertiesFromStyle(layoutStyle, level, paragraphLevel);
                    Object.assign(inheritedProps, layoutProps);
                }
            }

            // Try to get properties from layout shape's lstStyle (highest precedence)
            let layoutHasTitleStyling = false;
            if (slideContext.layout && slideContext.layout.shapes) {
                const layoutShape = this.findLayoutShapeForPlaceholder(slideContext.layout.shapes, shape);
                if (layoutShape && layoutShape.textBody && layoutShape.textBody.lstStyle) {
                    layoutHasTitleStyling = true;
                    const lstStyle = layoutShape.textBody.lstStyle;
                    // Determine paragraph level from first paragraph if available
                    let paragraphLevel = 0;
                    try {
                        if (shape && shape.textBody && Array.isArray(shape.textBody.paragraphs) && shape.textBody.paragraphs.length > 0) {
                            const lvl = shape.textBody.paragraphs[0]?.properties?.level;
                            if (typeof lvl === 'number' && lvl >= 0 && lvl <= 8) {paragraphLevel = lvl;}
                        }
                    } catch(_e) {}

                    const levelKey = `lvl${paragraphLevel + 1}pPr`;
                    const levelProps = lstStyle[levelKey] || lstStyle.lvl1pPr;
                    if (levelProps) {
                        if (levelProps.fontSize) {inheritedProps.fontSize = levelProps.fontSize;}
                        if (levelProps.fontFamily) {inheritedProps.fontFamily = levelProps.fontFamily;}
                        if (levelProps.color) {inheritedProps.color = levelProps.color;}
                        if (levelProps.bold !== undefined) {inheritedProps.bold = levelProps.bold;}
                        if (levelProps.italic !== undefined) {inheritedProps.italic = levelProps.italic;}
                        if (levelProps.align) {inheritedProps.align = levelProps.align;}
                        if (levelProps.cap) {inheritedProps.cap = levelProps.cap;}
                    }
                }
            }

            // Fallback: Try to get properties from master shape's lstStyle when layout doesn't provide
            if ((!inheritedProps.fontSize || !inheritedProps.align || !inheritedProps.fontFamily) && slideContext.master && slideContext.master.cSld && slideContext.master.cSld.spTree) {
                const masterShapes = slideContext.master.cSld.spTree;
                const masterShape = this.findLayoutShapeForPlaceholder(masterShapes, shape);
                if (masterShape && masterShape.textBody && masterShape.textBody.lstStyle) {
                    const lstStyle = masterShape.textBody.lstStyle;
                    if (lstStyle.lvl1pPr) {
                        const lvl1Props = lstStyle.lvl1pPr;
                        if (!inheritedProps.fontSize && lvl1Props.fontSize) {inheritedProps.fontSize = lvl1Props.fontSize;}
                        if (!inheritedProps.fontFamily && lvl1Props.fontFamily) {inheritedProps.fontFamily = lvl1Props.fontFamily;}
                        if (!inheritedProps.color && lvl1Props.color) {inheritedProps.color = lvl1Props.color;}
                        if (lvl1Props.bold !== undefined && inheritedProps.bold === undefined) {inheritedProps.bold = lvl1Props.bold;}
                        if (lvl1Props.italic !== undefined && inheritedProps.italic === undefined) {inheritedProps.italic = lvl1Props.italic;}
                        if (!inheritedProps.align && lvl1Props.align) {inheritedProps.align = lvl1Props.align;}
                        if (!inheritedProps.cap && lvl1Props.cap) {inheritedProps.cap = lvl1Props.cap;}
                    }
                }
            }

            // Apply placeholder-specific style overrides - only formatting, not sizes or alignment
            // Let the layout/master define the font sizes, alignment, and capitalization properly
            if (shape && shape.placeholder) {
                switch (shape.placeholder.type) {
                    case 'title':
                    case 'ctrTitle':
                        // PowerPoint built-in: titles are bold unless the layout provides explicit style context.
                        // If the layout has a title shape with lstStyle, it owns the style (bold defaults to false).
                        // If the layout has no title shape/lstStyle, apply PowerPoint's built-in bold=true default.
                        if (inheritedProps.bold === undefined && !layoutHasTitleStyling) {
                            inheritedProps.bold = true;
                        }
                        break;
                    case 'subTitle':
                        // Don't override align - let layout define alignment
                        break;
                    case 'body':
                        // Don't override align - let layout define alignment
                        break;
                    case 'obj':
                        // No specific overrides for obj placeholders
                        break;
                }
            }

            // Final fallback: sensible defaults when still missing
            if ((shape?.placeholder?.type === 'title' || shape?.placeholder?.type === 'ctrTitle')) {
                if (!inheritedProps.align) {inheritedProps.align = 'ctr';}
                if (!inheritedProps.fontSize) {inheritedProps.fontSize = 44;} // typical title size
            }

        } catch (error) {
            // Silent error handling - use defaults
        }

        return inheritedProps;
    }

    /**
     * Extract text properties from a text style definition
     */
    extractTextPropertiesFromStyle(style, level = 'run', paragraphLevel = 0) {
        const props = {};

        if (!style || !style.element) {
            return props;
        }

        try {
            // Parse the style element if not already parsed
            if (!style.parsed) {
                style.parsedData = this.parseTextStyleElement(style.element);
                style.parsed = true;
            }

            const styleData = style.parsedData;
            if (!styleData) {return props;}

            // Extract properties based on level
            if (level === 'paragraph' || level === 'run') {
                // Look for default paragraph properties
                const defPPr = styleData.defPPr;
                if (defPPr) {
                    if (defPPr.align) {props.align = defPPr.align;}
                    if (defPPr.lineHeight) {props.lineHeight = defPPr.lineHeight;}
                    if (defPPr.spaceBefore) {props.spaceBefore = defPPr.spaceBefore;}
                    if (defPPr.spaceAfter) {props.spaceAfter = defPPr.spaceAfter;}
                    if (defPPr.bullet) {props.bullet = defPPr.bullet;}
                }

                // Look for default run properties
                const defRPr = styleData.defRPr;
                if (defRPr) {
                    if (defRPr.fontSize) {props.fontSize = defRPr.fontSize;}
                    if (defRPr.fontFamily) {props.fontFamily = defRPr.fontFamily;}
                    if (defRPr.color) {props.color = defRPr.color;}
                    if (defRPr.bold !== undefined) {props.bold = defRPr.bold;}
                    if (defRPr.italic !== undefined) {props.italic = defRPr.italic;}
                    if (defRPr.underline !== undefined) {props.underline = defRPr.underline;}
                    if (defRPr.strike !== undefined) {props.strike = defRPr.strike;}
                    if (defRPr.cap) {props.cap = defRPr.cap;}
                }

                // Look for level-specific properties based on paragraphLevel
                // In PowerPoint, lvl1pPr corresponds to level 0, lvl2pPr to level 1, etc.
                const levelIndex = paragraphLevel + 1;
                const levelPPr = styleData[`lvl${levelIndex}pPr`];
                if (levelPPr) {
                    if (levelPPr.align) {props.align = levelPPr.align;}
                    if (levelPPr.lineHeight) {props.lineHeight = levelPPr.lineHeight;}
                    if (levelPPr.bullet) {props.bullet = levelPPr.bullet;}
                    if (levelPPr.leftMargin !== undefined) {props.leftMargin = levelPPr.leftMargin;}
                    if (levelPPr.indent !== undefined) {props.hangingIndent = levelPPr.indent;}
                    if (levelPPr.spaceBefore !== undefined) {props.spaceBefore = levelPPr.spaceBefore;}
                    if (levelPPr.spaceBeforePct !== undefined) {props.spaceBeforePct = levelPPr.spaceBeforePct;}
                    if (levelPPr.spaceAfter !== undefined) {props.spaceAfter = levelPPr.spaceAfter;}
                    if (levelPPr.spaceAfterPct !== undefined) {props.spaceAfterPct = levelPPr.spaceAfterPct;}

                    // Check for run properties within level
                    if (levelPPr.defRPr) {
                        const levelRPr = levelPPr.defRPr;
                        if (levelRPr.fontSize) {props.fontSize = levelRPr.fontSize;}
                        if (levelRPr.fontFamily) {props.fontFamily = levelRPr.fontFamily;}
                        if (levelRPr.color) {props.color = levelRPr.color;}
                        if (levelRPr.bold !== undefined) {props.bold = levelRPr.bold;}
                        if (levelRPr.italic !== undefined) {props.italic = levelRPr.italic;}
                        if (levelRPr.cap) {props.cap = levelRPr.cap;}
                    }
                }
            }

        } catch (error) {
            // Silent error handling
        }

        return props;
    }

    /**
     * Parse color from theme color reference
     */
    parseColorFromTheme(colorRef) {
        if (typeof colorRef === 'string') {
            // Handle hex colors
            if (colorRef.startsWith('#')) {
                const hex = colorRef.substring(1);
                const r = parseInt(hex.substring(0, 2), 16);
                const g = parseInt(hex.substring(2, 4), 16);
                const b = parseInt(hex.substring(4, 6), 16);
                return { r, g, b, a: 255 };
            }
        }

        // Return null if parsing fails - no color specified
        return null;
    }

    /**
     * Parse text style element to extract properties
     */
    parseTextStyleElement(element) {
        const styleData = {};

        if (!element) {return styleData;}

        try {
            // Parse default paragraph properties
            const defPPrElement = element.querySelector('defPPr, a\\:defPPr');
            if (defPPrElement) {
                styleData.defPPr = this.parseStyleParagraphProperties(defPPrElement);
            }

            // Parse default run properties
            const defRPrElement = element.querySelector('defRPr, a\\:defRPr') ||
                                 element.querySelector('defPPr > defRPr, a\\:defPPr > a\\:defRPr');
            if (defRPrElement) {
                styleData.defRPr = this.parseStyleRunProperties(defRPrElement);
            }

            // Parse level-specific properties (lvl1pPr, lvl2pPr, etc.)
            for (let i = 1; i <= 9; i++) {
                const levelElement = element.querySelector(`lvl${i}pPr, a\\:lvl${i}pPr`);
                if (levelElement) {
                    const levelData = this.parseStyleParagraphProperties(levelElement);

                    // Check for run properties within level
                    const levelRPrElement = levelElement.querySelector('defRPr, a\\:defRPr');
                    if (levelRPrElement) {
                        levelData.defRPr = this.parseStyleRunProperties(levelRPrElement);
                    }

                    styleData[`lvl${i}pPr`] = levelData;
                }
            }

        } catch (error) {
            // Silent error handling
        }

        return styleData;
    }

    /**
     * Parse paragraph properties from style element
     */
    parseStyleParagraphProperties(pPrElement) {
        const props = {};

        if (!pPrElement) {return props;}

        try {
            // Alignment
            const align = pPrElement.getAttribute('algn') || pPrElement.getAttribute('align');
            if (align) {
                const alignmentMap = {
                    'l': 'left',
                    'left': 'left',
                    'ctr': 'center',
                    'center': 'center',
                    'r': 'right',
                    'right': 'right',
                    'just': 'justify',
                    'justify': 'justify'
                };
                props.align = alignmentMap[align] || align;
            }

            // Line height
            const lineHeight = pPrElement.getAttribute('lnSpc') || pPrElement.getAttribute('lineHeight');
            if (lineHeight) {
                props.lineHeight = parseInt(lineHeight) || 120;
            }

            // Spacing before — child element form: <a:spcBef><a:spcPct val="20000"/></a:spcBef>
            const spcBefEl = pPrElement.querySelector('spcBef, a\\:spcBef');
            if (spcBefEl) {
                const spcPct = spcBefEl.querySelector('spcPct, a\\:spcPct');
                const spcPts = spcBefEl.querySelector('spcPts, a\\:spcPts');
                if (spcPct) {
                    props.spaceBeforePct = parseInt(spcPct.getAttribute('val')) / 100000;
                } else if (spcPts) {
                    props.spaceBefore = parseInt(spcPts.getAttribute('val')) * 127;
                }
            }

            // Spacing after — child element form: <a:spcAft><a:spcPct val="0"/></a:spcAft>
            const spcAftEl = pPrElement.querySelector('spcAft, a\\:spcAft');
            if (spcAftEl) {
                const spcPct = spcAftEl.querySelector('spcPct, a\\:spcPct');
                const spcPts = spcAftEl.querySelector('spcPts, a\\:spcPts');
                if (spcPct) {
                    props.spaceAfterPct = parseInt(spcPct.getAttribute('val')) / 100000;
                } else if (spcPts) {
                    props.spaceAfter = parseInt(spcPts.getAttribute('val')) * 127;
                }
            }

            // Margins
            const leftMargin = pPrElement.getAttribute('marL') || pPrElement.getAttribute('leftMargin');
            if (leftMargin) {
                props.leftMargin = parseInt(leftMargin) || 0;
            }

            const rightMargin = pPrElement.getAttribute('marR') || pPrElement.getAttribute('rightMargin');
            if (rightMargin) {
                props.rightMargin = parseInt(rightMargin) || 0;
            }

            // Indentation
            const indent = pPrElement.getAttribute('indent');
            if (indent) {
                props.indent = parseInt(indent) || 0;
            }

            // Bullet properties
            // Check for bullet character (buChar)
            const buCharElement = pPrElement.querySelector('buChar, a\\:buChar');
            if (buCharElement) {
                const char = buCharElement.getAttribute('char');
                if (char) {
                    props.bullet = {
                        type: 'character',
                        char: char
                    };
                }
            }

            // Check for bullet numbering (buAutoNum)
            const buAutoNumElement = pPrElement.querySelector('buAutoNum, a\\:buAutoNum');
            if (buAutoNumElement) {
                const type = buAutoNumElement.getAttribute('type') || 'arabicPeriod';
                const startAt = buAutoNumElement.getAttribute('startAt') || '1';
                props.bullet = {
                    type: 'autoNumber',
                    subType: type,
                    startAt: parseInt(startAt)
                };
            }

            // Check for no bullet specified (buNone)
            const buNoneElement = pPrElement.querySelector('buNone, a\\:buNone');
            if (buNoneElement) {
                props.bullet = {
                    type: 'none'
                };
            }

        } catch (error) {
            // Silent error handling
        }

        return props;
    }

    /**
     * Parse run properties from style element
     */
    parseStyleRunProperties(rPrElement) {
        const props = {};

        if (!rPrElement) {return props;}

        try {
            // Font size
            const fontSize = rPrElement.getAttribute('sz') || rPrElement.getAttribute('fontSize');
            if (fontSize) {
                // Convert from half-points to points if needed
                const sizeValue = parseInt(fontSize);
                props.fontSize = sizeValue > 100 ? sizeValue / 100 : sizeValue;
            }

            // Font family
            const fontFamily = rPrElement.getAttribute('typeface') || rPrElement.getAttribute('fontFamily');
            if (fontFamily) {
                props.fontFamily = fontFamily;
            }

            // Bold
            const bold = rPrElement.getAttribute('b') || rPrElement.getAttribute('bold');
            if (bold !== null) {
                props.bold = bold === '1' || bold === 'true' || bold === true;
            }

            // Italic
            const italic = rPrElement.getAttribute('i') || rPrElement.getAttribute('italic');
            if (italic !== null) {
                props.italic = italic === '1' || italic === 'true' || italic === true;
            }

            // Underline — OOXML uses u="sng"|"dbl"|"heavy"|"dotted"|etc.; any value except "none" is underline
            const underline = rPrElement.getAttribute('u') || rPrElement.getAttribute('underline');
            if (underline !== null) {
                props.underline = underline !== 'none' && underline !== '' && underline !== '0' && underline !== 'false';
            }

            // Strike - OOXML uses 'sngStrike', 'dblStrike', 'noStrike'
            const strike = rPrElement.getAttribute('strike') || rPrElement.getAttribute('strikethrough');
            if (strike !== null) {
                props.strike = strike !== 'noStrike' && strike !== '0' && strike !== 'false' && strike !== false;
            }

            // Capitalization
            const cap = rPrElement.getAttribute('cap') || rPrElement.getAttribute('capitalization');
            if (cap) {
                props.cap = cap;
            }

            // Color
            const colorElement = rPrElement.querySelector('solidFill, a\\:solidFill') ||
                               rPrElement.querySelector('schemeClr, a\\:schemeClr') ||
                               rPrElement.querySelector('srgbClr, a\\:srgbClr');
            if (colorElement) {
                props.color = this.parseColorFromElement(colorElement);
            }

            // Letter spacing
            const letterSpacing = rPrElement.getAttribute('spc') || rPrElement.getAttribute('letterSpacing');
            if (letterSpacing) {
                props.letterSpacing = parseInt(letterSpacing) || 0;
            }

        } catch (error) {
            // Silent error handling
        }

        return props;
    }
    /**
     * Parse color from color element
     */
    parseColorFromElement(colorElement) {
        try {
            // Check for sRGB color
            const srgbClr = colorElement.querySelector('srgbClr, a\\:srgbClr');
            if (srgbClr) {
                const val = srgbClr.getAttribute('val');
                if (val && val.length === 6) {
                    const r = parseInt(val.substr(0, 2), 16);
                    const g = parseInt(val.substr(2, 2), 16);
                    const b = parseInt(val.substr(4, 2), 16);
                    return { r, g, b, a: 255 };
                }
            }

            // Check for scheme color
            const schemeClr = colorElement.querySelector('schemeClr, a\\:schemeClr');
            if (schemeClr) {
                const val = schemeClr.getAttribute('val');
                // Map common scheme colors - PowerPoint standard theme colors
                const schemeColors = {
                    // No hardcoded scheme colors - only use colors from DOM
                };
                return schemeColors[val] || null;
            }

        } catch (error) {
            // Silent error handling
        }

        return null;
    }

    /**
     * Build slide context for placeholder style resolution
     */
    buildSlideContext() {
        const context = {};

        // Get current slide if available
        if (this.currentSlide) {
            context.slide = this.currentSlide;

            // Get layout if available
            if (this.currentSlide.layout || this.currentSlide.Layout) {
                context.layout = this.currentSlide.layout || this.currentSlide.Layout;

                // Include layout shapes for text style inheritance
                if (context.layout.commonSlideData && context.layout.commonSlideData.shapeTree) {
                    context.layout.shapes = context.layout.commonSlideData.shapeTree;
                } else if (context.layout.cSld && context.layout.cSld.spTree) {
                    context.layout.shapes = context.layout.cSld.spTree;
                }

                // Get master if available
                if (context.layout.master || context.layout.Master) {
                    context.master = context.layout.master || context.layout.Master;
                }
            }

            // Get theme if available
            if (this.currentSlide.theme || this.currentSlide.Theme) {
                context.theme = this.currentSlide.theme || this.currentSlide.Theme;
            }
        }

        return context;
    }

    /**
     * Get the current shape being rendered
     */
    getCurrentShape() {
        // Return the shape currently being processed
        // This would be set by drawShapeEnhanced when processing each shape
        return this.currentRenderingShape || null;
    }

    /**
     * Helper function to apply precise coordinate transformations
     * Reduces cumulative floating-point errors in coordinate calculations
     */
    applyPreciseTransform(value, scale = 1, offset = 0) {
        return Math.round(value * scale + offset);
    }

    /**
     * Set the current shape being rendered (called from drawShapeEnhanced)
     */
    setCurrentRenderingShape(shape) {
        this.currentRenderingShape = shape;
    }
}

/**
 * Enhanced Custom Geometry Processor
 * Handles parsing and processing of complex custom geometries
 */
class CustomGeometryProcessor {
    constructor() {
        this.cache = new Map();
    }

    /**
     * Process shape data format like the provided example
     */
    processShapeGeometry(shapeData) {
        const cacheKey = this.generateCacheKey(shapeData);
        
        if (this.cache.has(cacheKey)) {
            return this.cache.get(cacheKey);
        }

        let geometry = null;

        // Check for custom geometry in the shape data
        if (shapeData.geometry && shapeData.geometry.type === 'custom') {
            geometry = this.processCustomGeometry(shapeData.geometry);
        } else if (shapeData.properties && shapeData.properties.geometry && shapeData.properties.geometry.type === 'custom') {
            geometry = this.processCustomGeometry(shapeData.properties.geometry);
        }

        if (geometry) {
            this.cache.set(cacheKey, geometry);
        }

        return geometry;
    }

    /**
     * Process custom geometry with pathList
     */
    processCustomGeometry(geometryData) {
        if (!geometryData.pathLst || geometryData.pathLst.length === 0) {
            return null;
        }

        const processedGeometry = {
            type: 'custom',
            pathLst: []
        };

        // Process each path in the pathLst
        for (const path of geometryData.pathLst) {
            const processedPath = this.processPath(path);
            if (processedPath) {
                processedGeometry.pathLst.push(processedPath);
            }
        }

        return processedGeometry;
    }

    /**
     * Process individual path data
     */
    processPath(pathData) {
        if (!pathData.commands || pathData.commands.length === 0) {
            return null;
        }

        return {
            w: pathData.w || 100,
            h: pathData.h || 100,
            commands: pathData.commands.map(cmd => this.normalizeCommand(cmd)),
            fill: pathData.fill !== false,
            stroke: pathData.stroke !== false
        };
    }

    /**
     * Normalize command format to ensure consistency
     */
    normalizeCommand(command) {
        const normalized = { ...command };

        // Ensure consistent command type naming
        switch (command.type) {
            case 'M':
            case 'moveTo':
                normalized.type = 'moveTo';
                break;
            case 'L':
            case 'lineTo':
                normalized.type = 'lineTo';
                break;
            case 'C':
            case 'curveTo':
            case 'cubicBezTo':
                normalized.type = 'cubicBezTo';
                break;
            case 'Q':
            case 'quadTo':
            case 'quadBezTo':
                normalized.type = 'quadBezTo';
                break;
            case 'A':
            case 'arcTo':
                normalized.type = 'arcTo';
                break;
            case 'Z':
            case 'close':
                normalized.type = 'close';
                break;
        }

        return normalized;
    }

    /**
     * Generate cache key for geometry data
     */
    generateCacheKey(shapeData) {
        const geometry = shapeData.geometry || shapeData.properties?.geometry;
        if (!geometry) {return 'no-geometry';}

        const pathData = geometry.pathLst || [];
        return `custom-${JSON.stringify(pathData).substring(0, 100)}-${pathData.length}`;
    }
}

// Export classes for use in slide editor (maintain backward compatibility)
if (typeof window !== 'undefined') {
    window.CanvasGraphicsAdapter = CanvasGraphicsAdapter;
    window.CDrawingDocument = CDrawingDocument;
    window.GraphicsMatrix = GraphicsMatrix;
    try { __augmentAdapterWithArrows__(); } catch (_e) {}
}

// Intentionally no ES module exports to support classic <script> usage in root demo

// Logger class for compatibility
if (typeof Logger === 'undefined' && typeof window !== 'undefined') {
    window.Logger = class Logger {
        constructor(name) {
            this.name = name;
        }

        log(level, message, data) {
            // Silent logging
        }

        logError(module, message) {
            // Silent error logging
        }
    };
}
