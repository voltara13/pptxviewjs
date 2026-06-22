/**
 * Enhanced Presentation DOM Module
 * Based on presentation object model with enhanced shape support
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
 * Enhanced Presentation class - based on CPresentation
 */
class CPresentation {
    constructor() {
        this.type = 'CPresentation';
        this.id = Math.random().toString(36).substr(2, 9);
        this.slides = [];
        this.slideMasters = [];
        this.notesMasters = [];
        this.slideSize = { cx: 9144000, cy: 6858000 }; // Default size in EMU
        this.currentSlide = 0;
        this.app = null;
        this.core = null;
        this.customProperties = null;
        this.masterIds = []; // rId values for slide masters
        this.theme = null; // Presentation theme
    }

    /**
     * Get slide count
     */
    getSlidesCount() {
        return this.slides.length;
    }

    /**
     * Get current slide
     */
    getCurrentSlide() {
        return this.slides[this.currentSlide] || null;
    }

    /**
     * Get slide by index
     */
    getSlide(index) {
        return this.slides[index] || null;
    }

    /**
     * Add slide
     */
    addSlide(slide) {
        this.slides.push(slide);
        slide.setParent(this);
    }

    /**
     * Set current slide index
     */
    setCurrentSlide(index) {
        if (index >= 0 && index < this.slides.length) {
            this.currentSlide = index;
        }
    }

    /**
     * Get slide dimensions in millimeters
     */
    getSlideDimensions() {
        return {
            width: CoordinateTransform.emuToMM(this.slideSize.cx),
            height: CoordinateTransform.emuToMM(this.slideSize.cy)
        };
    }

    addSlideMaster(master) {
        this.slideMasters.push(master);
        master.setParent(this);
    }

    /**
     * Enhanced drawing method for presentation
     */
    draw(graphics) {
        const currentSlide = this.getCurrentSlide();
        if (currentSlide) {
            currentSlide.draw(graphics);
        }
    }
}

/**
 * Enhanced Slide class - based on standard CSlide
 */
class CSlide {
    constructor() {

        this.id = Math.random().toString(36).substr(2, 9);
        this.commonSlideData = new CSld();
        this.layout = null;
        this.notes = null;
        this.timing = null;
        this.transition = null;
        this.showMasterShapes = true;
        this.width = 254;  // Default width in mm
        this.height = 190.5; // Default height in mm
        this.parent = null;

        // Enhanced properties
        this.bounds = { l: 0, t: 0, r: 254, b: 190.5 };
        this.backgroundFill = null;

    }

    /**
     * Set parent presentation
     */
    setParent(parent) {
        this.parent = parent;
    }

    /**
     * Get slide name
     */
    getName() {
        return this.commonSlideData ? this.commonSlideData.name : '';
    }

    /**
     * Get shape tree
     */
    getShapeTree() {
        return this.commonSlideData ? this.commonSlideData.shapeTree : [];
    }

    /**
     * Add shape to slide
     */
    addToSpTree(pos, shape) {

        if (!this.commonSlideData) {

            this.commonSlideData = new CSld();
        }

        if (pos === undefined || pos === null) {
            pos = this.commonSlideData.shapeTree.length;
        }

        this.commonSlideData.shapeTree.splice(pos, 0, shape);
        shape.setParent(this);

    }

    /**
     * Remove shape from slide
     */
    removeFromSpTree(pos) {
        if (pos >= 0 && pos < this.commonSlideData.shapeTree.length) {
            return this.commonSlideData.shapeTree.splice(pos, 1)[0];
        }
        return null;
    }

    /**
     * Get background
     */
    getBackground() {
        return this.backgroundFill || this.commonSlideData?.background;
    }

    /**
     * Enhanced draw method - standard style
     */
    draw(graphics) {
        if (!graphics) {return;}


        // Draw background
        this.drawBackground(graphics);

        // Draw master slide elements if enabled
        if (this.showMasterShapes && this.layout && this.layout.master) {
            this.drawMasterElements(graphics);
        }

        // Draw layout elements
        if (this.layout) {
            this.drawLayoutElements(graphics);
        }

        // Draw slide shapes
        this.drawShapeTree(graphics);
    }

    /**
     * Draw background with enhanced support
     */
    drawBackground(_graphics) {
        // DISABLED: This method conflicts with standard adapter background drawing
        // The standard adapter handles background drawing
        // with the correct coordinate system. This method uses MM coordinates while the
        // adapter uses pixel coordinates, which causes coordinate system mismatches.

        // Original code disabled:
        // const bg = this.getBackground();
        // if (bg) {
        //     if (bg.type === 'solid' && bg.color) {
        //         graphics.fillRect(0, 0, this.width, this.height, bg.color);
        //     } else if (bg.type === 'gradient') {
        //         this.drawGradientBackground(graphics, bg);
        //     } else if (bg.type === 'image') {
        //         this.drawImageBackground(graphics, bg);
        //     }
        // } else {
        //     // Default white background
        //     graphics.fillRect(0, 0, this.width, this.height, { r: 255, g: 255, b: 255 });
        // }
    }

    /**
     * Draw gradient background
     */
    drawGradientBackground(graphics, bg) {
        // Simplified gradient - just use first color
        const color = bg.colors && bg.colors[0] ? bg.colors[0] : { r: 255, g: 255, b: 255 };
        graphics.fillRect(0, 0, this.width, this.height, color);
    }

    /**
     * Draw image background
     */
    drawImageBackground(graphics, _bg) {
        // Placeholder for image background
        graphics.fillRect(0, 0, this.width, this.height, { r: 240, g: 240, b: 240 });
    }

    /**
     * Draw master slide elements
     */
    drawMasterElements(graphics) {
        if (this.layout && this.layout.master && this.layout.master.commonSlideData) {
            const masterShapes = this.layout.master.commonSlideData.shapeTree;
            for (const shape of masterShapes) {
                if (!shape.isPlaceholder) {
                    this.drawShape(graphics, shape);
                }
            }
        }
    }

    /**
     * Draw layout elements
     */
    drawLayoutElements(graphics) {
        if (this.layout && this.layout.commonSlideData) {
            const layoutShapes = this.layout.commonSlideData.shapeTree;
            for (const shape of layoutShapes) {
                if (!shape.isPlaceholder) {
                    this.drawShape(graphics, shape);
                }
            }
        }
    }

    /**
     * Draw shape tree with enhanced support
     */
    drawShapeTree(graphics) {
        const shapes = this.getShapeTree();

        for (let i = 0; i < shapes.length; i++) {
            const shape = shapes[i];
            
            if (shape && !shape.isHidden) {
                this.drawShape(graphics, shape);
            }
        }
    }

    /**
     * Enhanced shape drawing with standard-style support
     */
    drawShape(graphics, shape) {
        if (!shape || !shape.properties) {
            return;
        }

        // Save graphics state
        graphics.SaveGrState();

        try {
            // Apply shape transform if available
            if (shape.properties.transform) {
                const matrix = this.createShapeTransformMatrix(shape.properties.transform);
                graphics.transform3(matrix);
            }

            // Check for group shapes
            if (shape.type === 'grpSp' || shape.isGroup?.()) {
                this.drawGroupShape(graphics, shape);
            } else {
                this.drawSingleShape(graphics, shape);
            }

        } finally {
            graphics.RestoreGrState();
        }
    }

    /**
     * Draw group shape - Enhanced with standard patterns
     */
    drawGroupShape(graphics, shape) {
        if (!shape.shapeTree || shape.shapeTree.length === 0) {
            return;
        }

        graphics.SaveGrState();

        try {
            // Apply group-level transformations if available
            if (shape.transform) {
                // Apply rotation
                if (shape.transform.rotation && shape.transform.rotation !== 0) {
                    const rotationMatrix = this.createRotationMatrix(shape.transform.rotation);
                    graphics.transform3(rotationMatrix);
                }

                // Apply flip transformations
                if (shape.transform.flipH || shape.transform.flipV) {
                    const scaleX = shape.transform.flipH ? -1 : 1;
                    const scaleY = shape.transform.flipV ? -1 : 1;
                    const flipMatrix = this.createScaleMatrix(scaleX, scaleY);
                    graphics.transform3(flipMatrix);
                }
            }

            // Apply group coordinate system transformation if available
            if (shape.groupCoordSystem) {
                const coordMatrix = this.createGroupCoordinateMatrix(shape.groupCoordSystem);
                if (coordMatrix) {
                    graphics.transform3(coordMatrix);
                }
            }

            // Draw all child shapes in the group
            for (const childShape of shape.shapeTree) {
                this.drawShape(graphics, childShape);
            }

        } finally {
            graphics.RestoreGrState();
        }
    }

    /**
     * Create group coordinate system transformation matrix
     */
    createGroupCoordinateMatrix(groupCoordSystem) {
        if (!groupCoordSystem) {return null;}

        // Calculate scale factors from group coordinate system
        let scaleX = 1, scaleY = 1;

        if (groupCoordSystem.chExt.cx > 0 && groupCoordSystem.groupExt.cx > 0) {
            scaleX = groupCoordSystem.groupExt.cx / groupCoordSystem.chExt.cx;
        }
        if (groupCoordSystem.chExt.cy > 0 && groupCoordSystem.groupExt.cy > 0) {
            scaleY = groupCoordSystem.groupExt.cy / groupCoordSystem.chExt.cy;
        }

        // Calculate translation offsets
        const offsetX = groupCoordSystem.groupOff.x - groupCoordSystem.chOff.x * scaleX;
        const offsetY = groupCoordSystem.groupOff.y - groupCoordSystem.chOff.y * scaleY;

        // Create transformation matrix
        return {
            m11: scaleX,
            m12: 0,
            m21: 0,
            m22: scaleY,
            dx: offsetX,
            dy: offsetY
        };
    }

    /**
     * Create rotation matrix
     */
    createRotationMatrix(degrees) {
        const radians = degrees * Math.PI / 180;
        const cos = Math.cos(radians);
        const sin = Math.sin(radians);

        return {
            m11: cos,
            m12: sin,
            m21: -sin,
            m22: cos,
            dx: 0,
            dy: 0
        };
    }

    /**
     * Create scale matrix
     */
    createScaleMatrix(scaleX, scaleY) {
        return {
            m11: scaleX,
            m12: 0,
            m21: 0,
            m22: scaleY,
            dx: 0,
            dy: 0
        };
    }

    /**
     * Draw single shape with enhanced geometry support
     */
    drawSingleShape(graphics, shape) {

        const transform = shape.properties.transform;
        if (!transform) {
            return;
        }

        // Convert EMU to mm
        const x = CoordinateTransform.emuToMM(transform.x || 0);
        const y = CoordinateTransform.emuToMM(transform.y || 0);
        const width = CoordinateTransform.emuToMM(transform.width || 0);
        const height = CoordinateTransform.emuToMM(transform.height || 0);

        // Get shape styling
        const fillColor = this.getShapeFillColor(shape);
        const strokeColor = this.getShapeStrokeColor(shape);
        const lineWidth = this.getShapeLineWidth(shape);


        // Draw based on shape type and geometry
        switch (shape.type) {
            case 'sp':
                this.drawTextShape(graphics, shape, x, y, width, height, fillColor, strokeColor, lineWidth);
                break;
            case 'pic':
                this.drawImageShape(graphics, shape, x, y, width, height);
                break;
            case 'cxnSp':
                this.drawConnectorShape(graphics, shape, x, y, width, height, strokeColor, lineWidth);
                break;
            case 'graphicFrame':
                this.drawGraphicFrame(graphics, shape, x, y, width, height);
                break;
            default:
                this.drawDefaultShape(graphics, shape, x, y, width, height, fillColor, strokeColor, lineWidth);
                break;
        }
    }

    /**
     * Draw text shape with enhanced geometry support
     */
    drawTextShape(graphics, shape, x, y, width, height, fillColor, strokeColor, lineWidth) {
        // Check for preset geometry
        const geometry = shape.geometry || shape.properties?.geometry;


        if (geometry && geometry.preset) {
            // Draw preset shape
            graphics.drawPresetGeometry(
                geometry.preset, x, y, width, height,
                fillColor, strokeColor, lineWidth
            );
        } else if (geometry && geometry.pathLst) {
            // Draw custom geometry
            this.drawCustomGeometry(graphics, geometry, x, y, width, height, fillColor, strokeColor, lineWidth);
        } else {
            // Default rectangle
            if (fillColor) {
                graphics.fillRect(x, y, width, height, fillColor);
            }
            if (strokeColor) {
                graphics.strokeRect(x, y, width, height, strokeColor, lineWidth);
            }
        }

        // Draw text content
        this.drawTextContent(graphics, shape, x, y, width, height);
    }

    /**
     * Draw custom geometry
     */
    drawCustomGeometry(graphics, geometry, x, y, width, height, fillColor, strokeColor, lineWidth) {

        graphics.SaveGrState();

        // Set up coordinate transformation for geometry
        const scaleX = width / (geometry.pathW || width);
        const scaleY = height / (geometry.pathH || height);

        // Apply geometry transform
        graphics.context.translate(x, y);
        graphics.context.scale(scaleX, scaleY);

        // Draw each path in the geometry
        if (geometry.pathLst) {
            for (let i = 0; i < geometry.pathLst.length; i++) {
                const path = geometry.pathLst[i];
                
                // For custom geometry paths, prioritize stroke over fill if both are present
                // Check if path has explicit fill configuration
                let pathFillColor = fillColor;
                if (path.fill === 'none' || path.fill === false) {
                    pathFillColor = null;
                } else if (!Object.prototype.hasOwnProperty.call(path, 'fill') && strokeColor) {
                    // If no explicit fill and we have stroke, prefer stroke-only rendering
                    pathFillColor = null;
                }
                
                
                this.drawGeometryPath(graphics, path, pathFillColor, strokeColor, lineWidth);
            }
        }

        graphics.RestoreGrState();
    }

    /**
     * Draw geometry path
     */
    drawGeometryPath(graphics, path, fillColor, strokeColor, lineWidth) {
        if (!path.commands || path.commands.length === 0) {return;}


        graphics._s();

        for (let i = 0; i < path.commands.length; i++) {
            const command = path.commands[i];
            switch (command.type) {
                case 'M': // MoveTo
                case 'moveTo':
                    graphics._m(command.x, command.y);
                    break;
                case 'L': // LineTo
                case 'lineTo':
                    graphics._l(command.x, command.y);
                    break;
                case 'C': // CurveTo
                case 'curveTo':
                    graphics._c(command.x1, command.y1, command.x2, command.y2, command.x, command.y);
                    break;
                case 'Z': // Close
                case 'close':
                    // Only close path if we're doing fill, not stroke-only
                    if (fillColor && path.fill !== 'none' && path.fill !== false) {
                        graphics._z();
                    } else {
                        // Path not closed - no action needed
                    }
                    break;
                case 'Q': // Quadratic curve
                case 'quadTo': {
                    // Convert quadratic to cubic bezier
                    const cp1x = command.cpx + (command.x - command.cpx) * 2/3;
                    const cp1y = command.cpy + (command.y - command.cpy) * 2/3;
                    graphics._c(command.cpx, command.cpy, cp1x, cp1y, command.x, command.y);
                    break;
                }
                case 'A': // Arc
                case 'arcTo':
                    // Better arc handling - convert to bezier curves for smoother rendering
                    if (command.rx && command.ry && command.x && command.y) {
                        // Use quadratic curve as approximation for arcs
                        const midX = (graphics.context.lastX || 0 + command.x) / 2;
                        const midY = (graphics.context.lastY || 0 + command.y) / 2;
                        graphics._c(midX, midY, midX, midY, command.x, command.y);
                    } else {
                        graphics._l(command.x, command.y);
                    }
                    break;
            }
        }

        // Apply stroke first (for paths with both fill and stroke, stroke should be visible)
        if (strokeColor && path.stroke !== false) {
            graphics.p_color(strokeColor.r, strokeColor.g, strokeColor.b, strokeColor.a || 255);
            graphics.p_width(lineWidth || 1);
            graphics.ds();
        } else {
            // No stroke needed
        }

        // Apply fill only if explicitly specified and not 'none'
        if (fillColor && path.fill !== 'none' && path.fill !== false) {
            graphics.b_color1(fillColor.r, fillColor.g, fillColor.b, fillColor.a || 255);
            graphics.df();
        } else {
            // No fill needed
        }

        graphics._e();
    }

    /**
     * Draw text content
     */
    drawTextContent(graphics, shape, x, y, width, height) {
        if (!shape.textBody || !shape.textBody.paragraphs) {return;}

        let textY = y + 5;
        const maxTextHeight = height - 10;
        const lineHeight = 14; // Default line height

            for (const paragraph of shape.textBody.paragraphs) {
            if (textY > y + maxTextHeight) {break;}

            let textX = x + 5; // Left padding

                for (const run of paragraph.runs) {
                if (run.text && textY <= y + maxTextHeight) {
                    const textProperties = {
                        fontSize: (run.properties?.fontSize) || 18, // Default to 18pt if not specified
                        fontFamily: (run.properties?.fontFamily) || 'Arial',
                        bold: run.properties?.bold || false,
                        italic: run.properties?.italic || false,
                        color: (run.properties?.color) || { r: 0, g: 0, b: 0 },
                        highlight: run.properties?.highlight || null
                    };

                    graphics.fillText(run.text, textX, textY, textProperties);

                    // Approximate text width for positioning
                    textX += run.text.length * (textProperties.fontSize * 0.6);
                }
            }

            textY += lineHeight;
        }
    }

    /**
     * Draw image shape
     */
    drawImageShape(graphics, shape, x, y, width, height) {
        // Draw neutral border only when no image is available (no placeholder text)
        graphics.strokeRect(x, y, width, height, { r: 160, g: 160, b: 160 }, 1);

        // If actual image data is available, draw it
        if (shape.imageData) {
            graphics.drawImage(shape.imageData, x, y, width, height);
        }
    }

    /**
     * Draw connector shape (lines)
     */
    drawConnectorShape(graphics, shape, x, y, width, height, strokeColor, lineWidth) {
        const startX = x;
        const startY = y;
        const endX = x + width;
        const endY = y + height;

        graphics.drawLine(startX, startY, endX, endY, strokeColor, lineWidth);
    }

    /**
     * Draw graphic frame (tables, charts)
     */
    drawGraphicFrame(graphics, shape, x, y, width, height) {
        if (shape.graphicData) {
            // Graphic data will be processed by specialized renderers
        }
        
        // CRITICAL FIX: Skip rendering if async chart processing is scheduled to prevent duplicate rendering
        if (shape.asyncChartProcessing) {
            // Draw a temporary neutral border while async processing is in progress
            graphics.strokeRect(x, y, width, height, { r: 200, g: 200, b: 200 }, 1);
            return;
        }
        
        // Check if this is a chart
        if (shape.chartData) {
            // Render chart using ChartRenderer
            if (window.ChartRenderer) {
                const chartRenderer = new ChartRenderer(graphics);
                chartRenderer.renderChart(shape.chartData, x, y, width, height);
            } else {
                // Fallback if ChartRenderer not available
                graphics.strokeRect(x, y, width, height, { r: 160, g: 160, b: 160 }, 1);
            }
        } else if (shape.graphicData && shape.graphicData.uri === 'http://schemas.openxmlformats.org/drawingml/2006/chart') {
            // Parse chart data if not already parsed
            if (window.ChartProcessor) {
                const chartProcessor = new ChartProcessor();
                const chartData = chartProcessor.parseEmbeddedChartData(shape.graphicData);
                if (chartData) {
                    shape.chartData = chartData;
                    if (window.ChartRenderer) {
                        const chartRenderer = new ChartRenderer(graphics);
                        chartRenderer.renderChart(chartData, x, y, width, height);
                    }
                } else {
                    graphics.strokeRect(x, y, width, height, { r: 160, g: 160, b: 160 }, 1);
                }
            } else {
                graphics.strokeRect(x, y, width, height, { r: 160, g: 160, b: 160 }, 1);
            }
        } else if (shape.table && shape.table instanceof CTable) {
            // Render table
            shape.table.draw(graphics, x, y, width, height);
        } else if (shape.graphicData && shape.graphicData.uri === 'http://schemas.openxmlformats.org/drawingml/2006/table') {
            // Try to parse and render the table
            try {
                // We need access to the parseTableFromXML method from slide-renderer
                // For now, create a simple table parser here
                const table = this.parseSimpleTable(shape.graphicData.tableXml);
                if (table) {
                    shape.table = table;
                    table.draw(graphics, x, y, width, height);
                } else {
                    graphics.strokeRect(x, y, width, height, { r: 160, g: 160, b: 160 }, 1);
                }
            } catch (error) {
                graphics.strokeRect(x, y, width, height, { r: 160, g: 160, b: 160 }, 1);
            }
        } else {
            // Draw neutral border for other graphic frames
            graphics.strokeRect(x, y, width, height, { r: 160, g: 160, b: 160 }, 1);
        }
    }

    /**
     * Simple table parser for DOM rendering
     */
    parseSimpleTable(tableXml) {
        if (!tableXml) {
            return null;
        }

        try {
            const parser = new DOMParser();
            const doc = parser.parseFromString(tableXml, 'text/xml');
            
            if (doc.documentElement.nodeName === 'parsererror') {
                return null;
            }

            const table = new CTable();
            
            // Find table grid (column definitions) - handle namespace
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
                    columns.push({ width: width ? parseInt(width) : 914400 }); // Default 1 inch
                });
                table.setTableGrid(columns);
            }

            // Parse table rows - handle namespace
            let tableRows = doc.querySelectorAll('tr');
            if (tableRows.length === 0) {
                tableRows = doc.querySelectorAll('a\\:tr');
            }
            
            tableRows.forEach(trElement => {
                const row = new CTableRow();
                
                // Get row height - handle namespace
                let trPr = trElement.querySelector('trPr');
                if (!trPr) {trPr = trElement.querySelector('a\\:trPr');}
                
                if (trPr) {
                    let trHeight = trPr.querySelector('trHeight');
                    if (!trHeight) {trHeight = trPr.querySelector('a\\:trHeight');}
                    
                    if (trHeight) {
                        row.height = parseInt(trHeight.getAttribute('val')) || null;
                    }
                }

                // Parse table cells - handle namespace
                let tableCells = trElement.querySelectorAll('tc');
                if (tableCells.length === 0) {
                    tableCells = trElement.querySelectorAll('a\\:tc');
                }
                
                tableCells.forEach(tcElement => {
                    const cell = new CTableCell();
                    
                    // Parse cell properties - handle namespace
                    let tcPr = tcElement.querySelector('tcPr');
                    if (!tcPr) {tcPr = tcElement.querySelector('a\\:tcPr');}
                    
                    if (tcPr) {
                        // Grid span
                        let gridSpan = tcPr.querySelector('gridSpan');
                        if (!gridSpan) {gridSpan = tcPr.querySelector('a\\:gridSpan');}
                        if (gridSpan) {
                            cell.gridSpan = parseInt(gridSpan.getAttribute('val')) || 1;
                        }

                        // Row span
                        let rowSpan = tcPr.querySelector('rowSpan');
                        if (!rowSpan) {rowSpan = tcPr.querySelector('a\\:rowSpan');}
                        if (rowSpan) {
                            cell.rowSpan = parseInt(rowSpan.getAttribute('val')) || 1;
                        }

                        // Cell fill
                        let solidFill = tcPr.querySelector('solidFill');
                        if (!solidFill) {solidFill = tcPr.querySelector('a\\:solidFill');}
                        if (solidFill) {
                            let colorElement = solidFill.querySelector('srgbClr, schemeClr, scrgbClr');
                            if (!colorElement) {colorElement = solidFill.querySelector('a\\:srgbClr, a\\:schemeClr, a\\:scrgbClr');}
                            if (colorElement) {
                                // Simple color parsing
                                const color = this.parseSimpleColor(colorElement);
                                if (color) {
                                    cell.fill = { color: color };
                                }
                            }
                        }
                    }

                    // Parse text content - simplified
                    const textContent = this.parseSimpleTextContent(tcElement);
                    if (textContent) {
                        cell.setTextBody(textContent);
                    }

                    row.addCell(cell);
                });

                table.addRow(row);
            });

            return table;

        } catch (error) {
            return null;
        }
    }

    /**
     * Simple color parser
     */
    parseSimpleColor(colorElement) {
        if (colorElement.tagName.includes('srgbClr')) {
            const val = colorElement.getAttribute('val');
            if (val) {
                const hex = val.replace('#', '');
                return {
                    r: parseInt(hex.substr(0, 2), 16),
                    g: parseInt(hex.substr(2, 2), 16),
                    b: parseInt(hex.substr(4, 2), 16),
                    a: 255
                };
            }
        }
        return { r: 200, g: 200, b: 200, a: 255 }; // Default gray
    }

    /**
     * Simple text content parser
     */
    parseSimpleTextContent(tcElement) {
        // Find run elements (text with properties)
        const runElements = tcElement.querySelectorAll('r, a\\:r');
        if (runElements.length === 0) {
            return null;
        }

        const textBody = {
            paragraphs: [],
            bodyProperties: {
                wrap: true,
                verticalAlign: 'top'
            }
        };

        const paragraph = {
            runs: [],
            properties: {}
        };

        runElements.forEach(runEl => {
            const textEl = runEl.querySelector('t, a\\:t');
            const text = textEl ? textEl.textContent : '';
            if (text) {
                // Parse run properties (rPr)
                const rPr = runEl.querySelector('rPr, a\\:rPr');
                const props = {
                    fontSize: 12,
                    fontFamily: 'Arial',
                    bold: false,
                    italic: false,
                    color: { r: 0, g: 0, b: 0, a: 255 } // Default black
                };

                if (rPr) {
                    // Font size
                    const sizeAttr = rPr.getAttribute('sz');
                    if (sizeAttr) {
                        props.fontSize = parseInt(sizeAttr) / 100; // Convert from 1/100 points to points
                    }

                    // Bold
                    const boldAttr = rPr.getAttribute('b');
                    if (boldAttr === '1') {
                        props.bold = true;
                    }

                    // Color parsing from solidFill > srgbClr
                    const solidFill = rPr.querySelector('solidFill, a\\:solidFill');
                    if (solidFill) {
                        const srgbClr = solidFill.querySelector('srgbClr, a\\:srgbClr');
                        if (srgbClr) {
                            const colorVal = srgbClr.getAttribute('val');
                            if (colorVal) {
                                // Convert hex color to RGB
                                const r = parseInt(colorVal.substr(0, 2), 16);
                                const g = parseInt(colorVal.substr(2, 2), 16);
                                const b = parseInt(colorVal.substr(4, 2), 16);
                                props.color = { r, g, b, a: 255 };
                            }
                        }
                    }
                }

                paragraph.runs.push({
                    text: text,
                    properties: props
                });
            }
        });

        if (paragraph.runs.length > 0) {
            textBody.paragraphs.push(paragraph);
        }

        return textBody.paragraphs.length > 0 ? textBody : null;
    }

    /**
     * Draw default shape
     */
    drawDefaultShape(graphics, shape, x, y, width, height, fillColor, strokeColor, lineWidth) {
        if (fillColor) {
            graphics.fillRect(x, y, width, height, fillColor);
        }
        if (strokeColor) {
            graphics.strokeRect(x, y, width, height, strokeColor, lineWidth);
        }
    }

    /**
     * Helper methods for shape properties
     */
    getShapeFillColor(shape) {
        if (shape.brush?.color) {return shape.brush.color;}
        if (shape.fill?.color) {return shape.fill.color;}
        if (shape.properties?.fill?.color) {return shape.properties.fill.color;}
        return { r: 200, g: 200, b: 200, a: 255 }; // Default light gray
    }

    getShapeStrokeColor(shape) {
        
        // Check for explicit no stroke indicators (standard patterns)
        if (shape.pen === null || shape.pen === 'none') {
            return null;
        }
        if (shape.stroke === null || shape.stroke === 'none') {
            return null;
        }
        if (shape.properties?.stroke === null || shape.properties?.stroke === 'none') {
            return null;
        }
        
        // Check for explicit stroke width of 0 (no stroke)
        if (shape.pen?.width === 0) {
            return null;
        }
        if (shape.stroke?.width === 0) {
            return null;
        }
        if (shape.properties?.stroke?.width === 0) {
            return null;
        }
        
        // Check for explicit stroke colors
        if (shape.pen?.color) {
            return shape.pen.color;
        }
        if (shape.stroke?.color) {
            return shape.stroke.color;
        }
        if (shape.properties?.stroke?.color) {
            return shape.properties.stroke.color;
        }
        
        // For text shapes (sp type), default to no stroke following standard pattern
        // Only apply default stroke for connector shapes or explicit shape borders
        return null; // No default stroke - let DOM define stroke behavior
    }

    getShapeLineWidth(shape) {
        // Check for explicit no stroke indicators (standard patterns)
        if (shape.pen === null || shape.pen === 'none') {return 0;}
        if (shape.stroke === null || shape.stroke === 'none') {return 0;}
        if (shape.properties?.stroke === null || shape.properties?.stroke === 'none') {return 0;}
        
        // Get explicit width values
        if (shape.pen?.width !== undefined) {return CoordinateTransform.emuToMM(shape.pen.width);}
        if (shape.stroke?.width !== undefined) {return CoordinateTransform.emuToMM(shape.stroke.width);}
        if (shape.properties?.stroke?.width !== undefined) {return CoordinateTransform.emuToMM(shape.properties.stroke.width);}
        
        // For text shapes, default to no line width (following standard pattern)
        return 0; // No default line width - let DOM define stroke behavior
    }

    /**
     * Create transformation matrix from shape transform
     */
    createShapeTransformMatrix(transform) {
        const matrix = new CMatrix();

        // Apply rotation if present
        if (transform.rotation) {
            const angle = (transform.rotation / 60000) * Math.PI / 180; // Convert from 1/60000th degrees
            const cos = Math.cos(angle);
            const sin = Math.sin(angle);
            matrix.sx = cos;
            matrix.sy = cos;
            matrix.shx = -sin;
            matrix.shy = sin;
        }

        // Apply translation
        matrix.tx = CoordinateTransform.emuToMM(transform.x || 0);
        matrix.ty = CoordinateTransform.emuToMM(transform.y || 0);

        return matrix;
    }
}

/**
 * Enhanced Common Slide Data class - based on standard CSld
 */
class CSld {
    constructor() {

        this.name = '';
        this.shapeTree = [];
        this.background = null;
        this.colorMap = null;
        this.parent = null;

    }

    /**
     * Set parent slide
     */
    setParent(parent) {
        this.parent = parent;
    }

    /**
     * Add shape to shape tree
     */
    addToSpTree(pos, shape) {
        if (pos === undefined || pos === null) {
            pos = this.shapeTree.length;
        }
        this.shapeTree.splice(pos, 0, shape);
        shape.setParent(this);
    }

    /**
     * Remove shape from shape tree
     */
    removeFromSpTree(pos) {
        if (pos >= 0 && pos < this.shapeTree.length) {
            return this.shapeTree.splice(pos, 1)[0];
        }
        return null;
    }
}

/**
 * Enhanced Shape class - based on standard CShape
 */
class CShape {
    constructor() {
        this.id = Math.random().toString(36).substr(2, 9);
        this.type = 'sp'; // Default shape type
        this.properties = {
            transform: {
                x: 0,
                y: 0,
                width: 100000, // 1 inch in EMU
                height: 100000,
                rotation: 0
            }
        };
        this.geometry = null;
        this.textBody = null;
        this.fill = null;
        this.stroke = null;
        this.brush = null;
        this.pen = null;
        this.parent = null;
        this.isHidden = false;
        this.isPlaceholder = false;

        // Effect properties
        this.effects = null; // Store parsed effect list

        // Enhanced properties
        this.bounds = { l: 0, t: 0, r: 0, b: 0 };
        this.transform = null; // Calculated transform matrix
        this.style = null; // Shape style information (fillRef, lnRef, etc.)
        this.preservedStyle = null; // Backup of original style data for DOM preservation

        // Additional style-related properties for standard compatibility
        this.nvSpPr = null; // Non-visual shape properties (name, id, etc.)
        this.spPr = null; // Shape properties (for standard adapter compatibility)
        this.name = null; // Shape name for identification

        // Chart data support for PptxGenJS compatibility
        this.chartData = null; // ChartData object for chart shapes
        this.mediaInfo = null; // Media information for video/audio
        this.svgContent = null; // SVG data for vector graphics
        this.graphicData = null; // Generic graphic data container
    }

    /**
     * Set parent container
     */
    setParent(parent) {
        this.parent = parent;
    }

    /**
     * Check if this is a group shape
     */
    isGroup() {
        return this.type === 'grpSp';
    }

    /**
     * Get object type for rendering decisions
     */
    getObjectType() {
        switch (this.type) {
            case 'sp': return 'shape';
            case 'pic': return 'image';
            case 'grpSp': return 'group';
            case 'cxnSp': return 'connector';
            case 'graphicFrame': 
                // Check graphic data to determine specific type
                if (this.chartData) {return 'chart';}
                if (this.graphicData?.uri === 'http://schemas.openxmlformats.org/drawingml/2006/table') {return 'table';}
                if (this.graphicData?.uri === 'http://schemas.openxmlformats.org/drawingml/2006/chart') {return 'chart';}
                return 'graphic';
            default: return 'unknown';
        }
    }

    /**
     * Set geometry for the shape
     */
    setGeometry(geometry) {
        this.geometry = geometry;
    }

    /**
     * Set text body for the shape
     */
    setTextBody(textBody) {
        this.textBody = textBody;
    }

    /**
     * Set chart data for chart shapes
     */
    setChartData(chartData) {
        this.chartData = chartData;
    }

    /**
     * Check if this shape contains a chart
     */
    isChart() {
        return this.chartData !== null || 
               (this.graphicData?.uri === 'http://schemas.openxmlformats.org/drawingml/2006/chart');
    }

    /**
     * Check if this shape contains a table
     */
    isTable() {
        return this.graphicData?.uri === 'http://schemas.openxmlformats.org/drawingml/2006/table';
    }

    /**
     * Set media information for video/audio shapes
     */
    setMediaInfo(mediaInfo) {
        this.mediaInfo = mediaInfo;
    }

    /**
     * Set SVG content for vector graphics
     */
    setSvgContent(svgContent) {
        this.svgContent = svgContent;
    }

    /**
     * Calculate bounds for the shape
     */
    recalculateBounds() {
        const transform = this.properties.transform;
        if (transform) {
            this.bounds.l = CoordinateTransform.emuToMM(transform.x);
            this.bounds.t = CoordinateTransform.emuToMM(transform.y);
            this.bounds.r = this.bounds.l + CoordinateTransform.emuToMM(transform.width);
            this.bounds.b = this.bounds.t + CoordinateTransform.emuToMM(transform.height);
        }
    }

    /**
     * Enhanced draw method with standard-style rendering
     */
    draw(graphics) {
        if (this.isHidden) {return;}

        // Update bounds
        this.recalculateBounds();

        // Draw based on object type
        switch (this.getObjectType()) {
            case 'group':
                this.drawGroupShape(graphics);
                break;
            case 'image':
                this.drawImageShape(graphics);
                break;
            case 'connector':
                this.drawConnectorShape(graphics);
                break;
            case 'graphic':
                this.drawGraphicFrame(graphics);
                break;
            default:
                this.drawShape(graphics);
                break;
        }
    }

    /**
     * Draw regular shape
     */
    drawShape(graphics) {
        const x = this.bounds.l;
        const y = this.bounds.t;
        const width = this.bounds.r - this.bounds.l;
        const height = this.bounds.b - this.bounds.t;

        // Draw shape based on geometry
        if (this.geometry && this.geometry.preset) {
            graphics.drawPresetGeometry(
                this.geometry.preset, x, y, width, height,
                this.getFillColor(), this.getStrokeColor(), this.getLineWidth()
            );
        } else {
            // Default rectangle
            if (this.getFillColor()) {
                graphics.fillRect(x, y, width, height, this.getFillColor());
            }
            if (this.getStrokeColor()) {
                graphics.strokeRect(x, y, width, height, this.getStrokeColor(), this.getLineWidth());
            }
        }

        // Draw text if present
        if (this.textBody) {
            this.drawTextContent(graphics, x, y, width, height);
        }
    }

    /**
     * Draw group shape
     */
    drawGroupShape(graphics) {
        if (this.shapeTree) {
            graphics.SaveGrState();

            // Apply group transform if needed
            if (this.transform) {
                graphics.transform3(this.transform);
            }

            // Draw all child shapes
            for (const childShape of this.shapeTree) {
                childShape.draw(graphics);
            }

            graphics.RestoreGrState();
        }
    }

    /**
     * Draw image shape
     */
    drawImageShape(graphics) {
        const x = this.bounds.l;
        const y = this.bounds.t;
        const width = this.bounds.r - this.bounds.l;
        const height = this.bounds.b - this.bounds.t;

        graphics.strokeRect(x, y, width, height, { r: 160, g: 160, b: 160 }, 1);
    }

    /**
     * Draw connector shape
     */
    drawConnectorShape(graphics) {
        const x = this.bounds.l;
        const y = this.bounds.t;
        const width = this.bounds.r - this.bounds.l;
        const height = this.bounds.b - this.bounds.t;

        graphics.drawLine(x, y, x + width, y + height, this.getStrokeColor(), this.getLineWidth());
    }

    /**
     * Draw graphic frame
     */
    drawGraphicFrame(graphics) {
        const x = this.bounds.l;
        const y = this.bounds.t;
        const width = this.bounds.r - this.bounds.l;
        const height = this.bounds.b - this.bounds.t;

        // CRITICAL FIX: Skip rendering if async chart processing is scheduled to prevent duplicate rendering
        if (this.asyncChartProcessing) {
            // Draw a temporary neutral border while async processing is in progress
            graphics.strokeRect(x, y, width, height, { r: 200, g: 200, b: 200 }, 1);
            return;
        }

        // Check if this is a chart
        if (this.chartData) {
            // Render chart using ChartRenderer
            if (window.ChartRenderer) {
                const chartRenderer = new ChartRenderer(graphics);
                chartRenderer.renderChart(this.chartData, x, y, width, height);
            } else {
                graphics.strokeRect(x, y, width, height, { r: 160, g: 160, b: 160 }, 1);
            }
        } else if (this.graphicData && this.graphicData.uri === 'http://schemas.openxmlformats.org/drawingml/2006/chart') {
            // Parse chart data if not already parsed
            if (window.ChartProcessor) {
                const chartProcessor = new ChartProcessor();
                const chartData = chartProcessor.parseEmbeddedChartData(this.graphicData);
                if (chartData) {
                    this.chartData = chartData;
                    if (window.ChartRenderer) {
                        const chartRenderer = new ChartRenderer(graphics);
                        chartRenderer.renderChart(chartData, x, y, width, height);
                    }
                } else {
                    graphics.strokeRect(x, y, width, height, { r: 160, g: 160, b: 160 }, 1);
                }
            } else {
                graphics.strokeRect(x, y, width, height, { r: 160, g: 160, b: 160 }, 1);
            }
        } else {
            // Default neutral border without placeholder text
            graphics.strokeRect(x, y, width, height, { r: 160, g: 160, b: 160 }, 1);
        }
    }

    /**
     * Draw text content
     */
    drawTextContent(graphics, x, y, width, height) {
        if (!this.textBody || !this.textBody.paragraphs) {return;}

        let textY = y + 5;
        for (const paragraph of this.textBody.paragraphs) {
            for (const run of paragraph.runs) {
                if (run.text) {
                    graphics.fillText(run.text, x + 5, textY, run.properties || {});
                    textY += 15;
                }
            }
        }
    }

    /**
     * Helper methods for styling
     */
    getFillColor() {
        // Return null if no fill is specified
        if (this.fill === null) {return null;}
        return this.fill?.color || this.brush?.color || null;
    }

    getStrokeColor() {
        // Return null if no stroke is specified
        if (this.stroke === null) {return null;}
        return this.stroke?.color || this.pen?.color || { r: 0, g: 0, b: 0 };
    }

    getLineWidth() {
        const width = this.stroke?.width || this.pen?.width || 12700; // Default 1pt in EMU
        return CoordinateTransform.emuToMM(width);
    }
}

/**
 * Group Shape class - Enhanced with standard patterns
 */
class CGroupShape extends CShape {
    constructor() {
        super();
        this.type = 'grpSp';
        this.shapeTree = [];

        // Group coordinate system properties (standard pattern)
        this.groupCoordSystem = {
            chOff: { x: 0, y: 0 },    // Child offset
            chExt: { cx: 0, cy: 0 },  // Child extent
            groupOff: { x: 0, y: 0 }, // Group offset
            groupExt: { cx: 0, cy: 0 } // Group extent
        };

        // Group transformation properties
        this.transform = {
            rotation: 0,
            flipH: false,
            flipV: false,
            groupOff: { x: 0, y: 0 },
            groupExt: { cx: 0, cy: 0 }
        };
    }

    /**
     * Add shape to group
     */
    addToSpTree(pos, shape) {
        if (pos === undefined || pos === null) {
            pos = this.shapeTree.length;
        }
        this.shapeTree.splice(pos, 0, shape);

        // Set parent reference
        if (shape.setParent) {
            shape.setParent(this);
        }

        // Mark shape as being in a group
        shape.inGroup = true;
        shape.parentGroup = this;

        // Recalculate bounds after adding child
        this.recalculateBounds();
    }

    /**
     * Remove shape from group
     */
    removeFromSpTree(pos) {
        if (pos >= 0 && pos < this.shapeTree.length) {
            const removedShape = this.shapeTree.splice(pos, 1)[0];

            // Clear parent reference
            if (removedShape.setParent) {
                removedShape.setParent(null);
            }

            // Clear group properties
            removedShape.inGroup = false;
            removedShape.parentGroup = null;

            // Recalculate bounds after removal
            this.recalculateBounds();

            return removedShape;
        }
        return null;
    }

    /**
     * Set group coordinate system
     */
    setGroupCoordSystem(chOff, chExt, groupOff, groupExt) {
        this.groupCoordSystem = {
            chOff: chOff || { x: 0, y: 0 },
            chExt: chExt || { cx: 0, cy: 0 },
            groupOff: groupOff || { x: 0, y: 0 },
            groupExt: groupExt || { cx: 0, cy: 0 }
        };

        // Update transform as well
        this.transform.groupOff = { ...this.groupCoordSystem.groupOff };
        this.transform.groupExt = { ...this.groupCoordSystem.groupExt };

        // Transform child shapes with new coordinate system
        this.transformChildShapes();
    }

    /**
     * Transform child shapes according to group coordinate system
     */
    transformChildShapes() {
        if (!this.shapeTree || this.shapeTree.length === 0) {
            return;
        }

        const groupCoords = this.groupCoordSystem;

        // Calculate scale factors
        let scaleX = 1, scaleY = 1;

        if (groupCoords.chExt.cx > 0 && groupCoords.groupExt.cx > 0) {
            scaleX = groupCoords.groupExt.cx / groupCoords.chExt.cx;
        }
        if (groupCoords.chExt.cy > 0 && groupCoords.groupExt.cy > 0) {
            scaleY = groupCoords.groupExt.cy / groupCoords.chExt.cy;
        }

        // Transform each child shape
        for (const childShape of this.shapeTree) {
            if (childShape.bounds) {
                const childBounds = childShape.bounds;

                // Store original bounds if not already stored
                if (!childShape.originalBounds) {
                    childShape.originalBounds = { ...childBounds };
                }

                // Transform coordinates from group space to slide space
                const transformedBounds = {
                    l: groupCoords.groupOff.x + ((childBounds.l - groupCoords.chOff.x) * scaleX),
                    t: groupCoords.groupOff.y + ((childBounds.t - groupCoords.chOff.y) * scaleY),
                    r: groupCoords.groupOff.x + ((childBounds.r - groupCoords.chOff.x) * scaleX),
                    b: groupCoords.groupOff.y + ((childBounds.b - groupCoords.chOff.y) * scaleY)
                };

                childShape.bounds = transformedBounds;

                // Store transformation info
                childShape.groupTransform = {
                    scaleX: scaleX,
                    scaleY: scaleY,
                    offsetX: groupCoords.groupOff.x - groupCoords.chOff.x * scaleX,
                    offsetY: groupCoords.groupOff.y - groupCoords.chOff.y * scaleY
                };
            }
        }
    }

    /**
     * Recalculate bounds based on child shapes - Enhanced with standard patterns
     */
    recalculateBounds() {
        if (this.shapeTree.length === 0) {
            super.recalculateBounds();
            return;
        }

        let minX = Infinity, minY = Infinity;
        let maxX = -Infinity, maxY = -Infinity;
        let hasValidBounds = false;

        for (const shape of this.shapeTree) {
            if (shape.recalculateBounds) {
                shape.recalculateBounds();
            }

            if (shape.bounds) {
                minX = Math.min(minX, shape.bounds.l);
                minY = Math.min(minY, shape.bounds.t);
                maxX = Math.max(maxX, shape.bounds.r);
                maxY = Math.max(maxY, shape.bounds.b);
                hasValidBounds = true;
            }
        }

        if (hasValidBounds) {
            this.bounds.l = minX;
            this.bounds.t = minY;
            this.bounds.r = maxX;
            this.bounds.b = maxY;
        } else {
            // Use fallback bounds
            this.bounds.l = 914400;
            this.bounds.t = 914400;
            this.bounds.r = 5486400;
            this.bounds.b = 2743200;
        }
    }

    /**
     * Enhanced draw method for group shapes
     */
    draw(graphics) {
        if (this.isHidden) {return;}

        // Update bounds
        this.recalculateBounds();

        // Save graphics state
        graphics.SaveGrState();

        try {
            // Apply group transformations
            if (this.transform) {
                // Apply rotation
                if (this.transform.rotation && this.transform.rotation !== 0) {
                    const rotationMatrix = this.createRotationMatrix(this.transform.rotation);
                    graphics.transform3(rotationMatrix);
                }

                // Apply flip transformations
                if (this.transform.flipH || this.transform.flipV) {
                    const scaleX = this.transform.flipH ? -1 : 1;
                    const scaleY = this.transform.flipV ? -1 : 1;
                    const flipMatrix = this.createScaleMatrix(scaleX, scaleY);
                    graphics.transform3(flipMatrix);
                }
            }

            // Apply group coordinate system transformation
            if (this.groupCoordSystem) {
                const coordMatrix = this.createGroupCoordinateMatrix(this.groupCoordSystem);
                if (coordMatrix) {
                    graphics.transform3(coordMatrix);
                }
            }

            // Draw all child shapes
            for (const childShape of this.shapeTree) {
                childShape.draw(graphics);
            }

        } finally {
            graphics.RestoreGrState();
        }
    }

    /**
     * Create group coordinate system transformation matrix
     */
    createGroupCoordinateMatrix(groupCoordSystem) {
        if (!groupCoordSystem) {return null;}

        // Calculate scale factors from group coordinate system
        let scaleX = 1, scaleY = 1;

        if (groupCoordSystem.chExt.cx > 0 && groupCoordSystem.groupExt.cx > 0) {
            scaleX = groupCoordSystem.groupExt.cx / groupCoordSystem.chExt.cx;
        }
        if (groupCoordSystem.chExt.cy > 0 && groupCoordSystem.groupExt.cy > 0) {
            scaleY = groupCoordSystem.groupExt.cy / groupCoordSystem.chExt.cy;
        }

        // Calculate translation offsets
        const offsetX = groupCoordSystem.groupOff.x - groupCoordSystem.chOff.x * scaleX;
        const offsetY = groupCoordSystem.groupOff.y - groupCoordSystem.chOff.y * scaleY;

        // Create transformation matrix
        return {
            m11: scaleX,
            m12: 0,
            m21: 0,
            m22: scaleY,
            dx: offsetX,
            dy: offsetY
        };
    }

    /**
     * Create rotation matrix
     */
    createRotationMatrix(degrees) {
        const radians = degrees * Math.PI / 180;
        const cos = Math.cos(radians);
        const sin = Math.sin(radians);

        return {
            m11: cos,
            m12: sin,
            m21: -sin,
            m22: cos,
            dx: 0,
            dy: 0
        };
    }

    /**
     * Create scale matrix
     */
    createScaleMatrix(scaleX, scaleY) {
        return {
            m11: scaleX,
            m12: 0,
            m21: 0,
            m22: scaleY,
            dx: 0,
            dy: 0
        };
    }

    /**
     * Get object type
     */
    getObjectType() {
        return 'group';
    }

    /**
     * Check if this is a group shape
     */
    isGroup() {
        return true;
    }
}

/**
 * Slide Master class - based on standard CSlideMaster
 */
class CSlideMaster {
    constructor() {
        this.id = Math.random().toString(36).substr(2, 9);
        this.commonSlideData = new CSld();
        this.theme = null;
        this.layouts = [];
        this.parent = null;
    }

    /**
     * Set parent presentation
     */
    setParent(parent) {
        this.parent = parent;
    }

    /**
     * Add layout to master
     */
    addLayout(layout) {
        this.layouts.push(layout);
        layout.setMaster(this);
    }
}

/**
 * Slide Layout class - based on standard CSlideLayout
 */
class CSlideLayout {
    constructor() {
        this.id = Math.random().toString(36).substr(2, 9);
        this.commonSlideData = new CSld();
        this.master = null;
        this.type = 'blank'; // Layout type
    }

    /**
     * Set master slide
     */
    setMaster(master) {
        this.master = master;
    }
}

/**
 * Table class - based on standard CTable for PPTX presentations
 */
class CTable {
    constructor() {
        this.id = Math.random().toString(36).substr(2, 9);
        this.type = 'table';
        this.rows = [];
        this.tableGrid = [];  // Column width definitions
        this.tableProperties = null;
        this.parent = null;
        
        // Table style properties
        this.firstRow = false;
        this.firstCol = false;
        this.lastRow = false;
        this.lastCol = false;
        this.bandRow = false;
        this.bandCol = false;
        this.tableStyleId = null;
        
        // Default dimensions
        this.bounds = { l: 0, t: 0, r: 0, b: 0 };
    }

    /**
     * Set parent container
     */
    setParent(parent) {
        this.parent = parent;
    }

    /**
     * Add row to table
     */
    addRow(row) {
        this.rows.push(row);
        row.setParent(this);
        row.rowIndex = this.rows.length - 1;
    }

    /**
     * Set table grid (column widths)
     */
    setTableGrid(gridColumns) {
        this.tableGrid = gridColumns || [];
    }

    /**
     * Get total table width
     */
    getTotalWidth() {
        return this.tableGrid.reduce((sum, col) => sum + (col.width || 0), 0);
    }

    /**
     * Get row count
     */
    getRowCount() {
        return this.rows.length;
    }

    /**
     * Get column count - returns the maximum of grid columns or actual cells in any row
     */
    getColumnCount() {
        // Get the defined grid column count
        const gridColumnCount = this.tableGrid.length;
        
        // Get the maximum number of cells in any row
        let maxCellCount = 0;
        for (const row of this.rows) {
            const cellCount = row.getCells().length;
            if (cellCount > maxCellCount) {
                maxCellCount = cellCount;
            }
        }
        
        // Return the maximum to handle malformed tables
        return Math.max(gridColumnCount, maxCellCount);
    }

    /**
     * Get table properties
     */
    getTableProperties() {
        return this.tableProperties || {};
    }

    /**
     * Get table borders
     */
    getTableBorders() {
        return this.tableProperties?.borders || {};
    }

    /**
     * Get table shading
     */
    getTableShading() {
        return this.tableProperties?.shading || {};
    }

    /**
     * Get table grid (column definitions)
     */
    getTableGrid() {
        return this.tableGrid || [];
    }

    /**
     * Get specific cell by row and column index
     */
    getCell(rowIndex, colIndex) {
        if (rowIndex >= 0 && rowIndex < this.rows.length) {
            const row = this.rows[rowIndex];
            if (colIndex >= 0 && colIndex < row.cells.length) {
                return row.cells[colIndex];
            }
        }
        return null;
    }

    /**
     * Get row by index
     */
    getRow(rowIndex) {
        if (rowIndex >= 0 && rowIndex < this.rows.length) {
            return this.rows[rowIndex];
        }
        return null;
    }

    /**
     * Get all rows
     */
    getRows() {
        return this.rows;
    }

    /**
     * Recalculate table bounds
     */
    recalculateBounds() {
        const totalWidth = this.getTotalWidth();
        const totalHeight = this.rows.reduce((sum, row) => sum + (row.height || 0), 0);
        
        this.bounds.r = this.bounds.l + CoordinateTransform.emuToMM(totalWidth);
        this.bounds.b = this.bounds.t + CoordinateTransform.emuToMM(totalHeight);
    }

    /**
     * Draw table
     */
    draw(graphics, x, y, width, height) {
        if (!this.rows || this.rows.length === 0) {
            // Neutral border only; no placeholder text
            graphics.strokeRect(x, y, width, height, { r: 160, g: 160, b: 160 }, 1);
            return;
        }

        // Calculate cell dimensions
        const colCount = this.getColumnCount();
        const rowCount = this.getRowCount();
        
        const cellWidth = colCount > 0 ? width / colCount : width;
        const cellHeight = rowCount > 0 ? height / rowCount : height;

        // FIXED: Don't draw default table border - let table borders be defined by PPTX
        // graphics.strokeRect(x, y, width, height, { r: 0, g: 0, b: 0 }, 1);

        // Draw each row
        let currentY = y;
        for (let rowIndex = 0; rowIndex < this.rows.length; rowIndex++) {
            const row = this.rows[rowIndex];
            const rowHeight = row.height ? CoordinateTransform.emuToMM(row.height) : cellHeight;
            
            row.draw(graphics, x, currentY, width, rowHeight, cellWidth, this.tableGrid);
            currentY += rowHeight;
        }
    }

    // getTableProperties method moved to line 1886

    // getTableBorders method moved to line 1893

    // getTableShading method moved to line 1900

    // getTableGrid method moved to line 1907

    // getRows method moved to line 1937

    // getRow and getCell methods moved to earlier in file (lines 1927 and 1914)
}

/**
 * Table Row class - represents a row in a table
 */
class CTableRow {
    constructor() {
        this.id = Math.random().toString(36).substr(2, 9);
        this.type = 'tableRow';
        this.cells = [];
        this.height = null;  // Row height in EMU
        this.parent = null;
        this.rowIndex = 0;
    }

    /**
     * Set parent table
     */
    setParent(parent) {
        this.parent = parent;
    }

    /**
     * Add cell to row
     */
    addCell(cell) {
        this.cells.push(cell);
        cell.setParent(this);
        cell.cellIndex = this.cells.length - 1;
    }

    /**
     * Get cell count
     */
    getCellCount() {
        return this.cells.length;
    }

    /**
     * Draw row
     */
    draw(graphics, x, y, width, height, defaultCellWidth, tableGrid) {
        if (!this.cells || this.cells.length === 0) {
            return;
        }

        // Draw row background if specified
        // (Row background styling would go here)

        // Draw each cell
        let currentX = x;
        for (let cellIndex = 0; cellIndex < this.cells.length; cellIndex++) {
            const cell = this.cells[cellIndex];
            
            // Calculate cell width from table grid or use default
            let cellWidth = defaultCellWidth;
            if (tableGrid && tableGrid[cellIndex] && tableGrid[cellIndex].width) {
                cellWidth = CoordinateTransform.emuToMM(tableGrid[cellIndex].width);
            }
            
            cell.draw(graphics, currentX, y, cellWidth, height);
            currentX += cellWidth;
        }

        // Draw horizontal row separator
        graphics.drawLine(x, y + height, x + width, y + height, { r: 200, g: 200, b: 200 }, 1);
    }

    /**
     * Get cell by index
     */
    getCell(index) {
        return this.cells[index] || null;
    }

    /**
     * Get cells array
     */
    getCells() {
        return this.cells;
    }
}

/**
 * Table Cell class - represents a cell in a table
 */
class CTableCell {
    constructor() {
        this.id = Math.random().toString(36).substr(2, 9);
        this.type = 'tableCell';
        this.textBody = null;  // Cell text content
        this.cellProperties = null;  // Cell formatting
        this.parent = null;
        this.cellIndex = 0;
        
        // Cell spanning
        this.gridSpan = 1;  // Column span
        this.rowSpan = 1;   // Row span
        
        // Cell borders and fill
        this.fill = null;
        this.borders = {
            top: null,
            right: null,
            bottom: null,
            left: null
        };
    }

    /**
     * Set parent row
     */
    setParent(parent) {
        this.parent = parent;
    }

    /**
     * Set text content
     */
    setTextBody(textBody) {
        this.textBody = textBody;
    }

    /**
     * Draw cell
     */
    draw(graphics, x, y, width, height) {
        // Draw cell background
        if (this.fill && this.fill.color) {
            graphics.fillRect(x, y, width, height, this.fill.color);
        }

        // FIXED: Don't draw default cell borders - let cell borders be defined by PPTX
        // graphics.strokeRect(x, y, width, height, { r: 200, g: 200, b: 200 }, 1);

        // Draw cell content
        if (this.textBody && this.textBody.paragraphs) {
            this.drawCellText(graphics, x, y, width, height);
        }

        // FIXED: Don't draw default vertical separator - let borders be defined by PPTX
        // graphics.drawLine(x + width, y, x + width, y + height, { r: 200, g: 200, b: 200 }, 1);
    }

    /**
     * Draw cell text content
     */
    drawCellText(graphics, x, y, width, height) {
        if (!this.textBody || !this.textBody.paragraphs) {
            return;
        }

        let textY = y + 5;
        const maxTextHeight = height - 10;
        const lineHeight = 14;

        for (const paragraph of this.textBody.paragraphs) {
            if (textY > y + maxTextHeight) {
                break;
            }

            let textX = x + 5; // Left padding
            
            if (paragraph.runs) {
                for (const run of paragraph.runs) {
                    if (run.text && textY <= y + maxTextHeight) {
                        const textProperties = {
                            fontSize: (run.properties?.fontSize) || 12,
                            fontFamily: (run.properties?.fontFamily) || 'Arial',
                            bold: run.properties?.bold || false,
                            italic: run.properties?.italic || false,
                            color: (run.properties?.color) || { r: 0, g: 0, b: 0 }
                        };

                        graphics.fillText(run.text, textX, textY, textProperties);
                        
                        // Approximate text width for positioning
                        textX += run.text.length * (textProperties.fontSize * 0.6);
                    }
                }
            }

            textY += lineHeight;
        }
    }

    /**
     * Get grid span
     */
    getGridSpan() {
        return this.gridSpan || 1;
    }

    /**
     * Get row span
     */
    getRowSpan() {
        return this.rowSpan || 1;
    }

    /**
     * Get cell shading/background
     */
    getCellShading() {
        return this.shading;
    }

    /**
     * Get text body
     */
    getTextBody() {
        return this.textBody;
    }

    /**
     * Get cell borders
     */
    getCellBorders() {
        return this.borders;
    }

    // setTextBody and setParent methods already defined earlier in this class
}

/**
 * Theme class - based on standard CTheme
 */
class CTheme {
    constructor() {
        this.id = Math.random().toString(36).substr(2, 9);
        this.name = 'Default Theme';
        this.colorScheme = null;
        this.fontScheme = null;
        this.formatScheme = null;
    }
}

// Export enhanced classes

// Export classes (maintain backward compatibility)
if (typeof window !== 'undefined') {
    window.CPresentation = CPresentation;
    window.CSlide = CSlide;
    window.CSld = CSld;
    window.CShape = CShape;
    window.CGroupShape = CGroupShape;
    window.CSlideMaster = CSlideMaster;
    window.CSlideLayout = CSlideLayout;
    window.CTheme = CTheme;
    window.CTable = CTable;
    window.CTableRow = CTableRow;
    window.CTableCell = CTableCell;
}
if (typeof globalThis !== 'undefined') {
    globalThis.CPresentation = CPresentation;
    globalThis.CSlide = CSlide;
    globalThis.CSld = CSld;
    globalThis.CShape = CShape;
    globalThis.CGroupShape = CGroupShape;
    globalThis.CSlideMaster = CSlideMaster;
    globalThis.CSlideLayout = CSlideLayout;
    globalThis.CTheme = CTheme;
    globalThis.CTable = CTable;
    globalThis.CTableRow = CTableRow;
    globalThis.CTableCell = CTableCell;
}

// Intentionally no ES module exports here to keep classic <script> compatibility in root demo
