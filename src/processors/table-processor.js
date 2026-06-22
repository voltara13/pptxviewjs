/**
 * Enhanced Table Processor for PowerPoint Presentations
 * Handles table parsing, rendering, and data extraction with comprehensive validation
 * Uses a general approach with enhanced error handling and input validation
 */

// import { Logger } from '../utils/utils.js';

class TableProcessor {
    constructor(options = {}) {
        // Initialize logger
        this.logger = new Logger('TableProcessor');
        
        // Initialize validation framework
        this.validator = this._getValidator();
        this.errorBoundary = new (this._getErrorBoundary())({
            context: 'TableProcessor',
            enableLogging: options.enableLogging !== false
        });
        
        // Core functionality
        this.tableCache = new Map();
        
        // Configuration with validation
        this.config = this._validateConfig({
            maxTableSize: options.maxTableSize || 1000, // Max cells
            maxCellTextLength: options.maxCellTextLength || 10000,
            enableCaching: options.enableCaching !== false,
            enableValidation: options.enableValidation !== false,
            renderTimeout: options.renderTimeout || 5000,
            ...options
        });
        
        // Performance tracking
        this.performanceMetrics = {
            tablesProcessed: 0,
            cacheHits: 0,
            cacheMisses: 0,
            averageRenderTime: 0
        };
    }

    /**
     * Enhanced table rendering with comprehensive validation and error handling
     * Uses a five-stage pipeline with defensive programming
     */
    async renderTableFrame(graphics, shape, x, y, w, h, options = {}) {
        // Input validation
        const validationResult = this._validateRenderInput(graphics, shape, x, y, w, h, options);
        
        if (!validationResult.valid) {
            this._handleValidationError('renderTableFrame', validationResult.errors);
            this.drawTablePlaceholder(graphics, x, y, w, h);
            return;
        }
        
        return this.errorBoundary.wrap(async () => {
            const startTime = performance.now();
            
            try {
                // Stage 1: Extract and validate table data with timeout
                const table = await this._getTableInstanceWithTimeout(shape, this.config.renderTimeout);
                
                if (!table) {
                    if (this.logger) {
                        this.logger.log("warn", this.constructor.name, 'TableProcessor', 'No valid table data found, drawing placeholder');
                    }
                    this.drawTablePlaceholder(graphics, x, y, w, h);
                    return;
                }
                
                // Validate table structure
                if (this.config.enableValidation) {
                    this._validateTableStructure(table);
                }
                
                // Save graphics state for table rendering
                graphics.SaveGrState();
                
                try {
                    // Stage 2: Draw table background and outer border
                    await this._safeDrawTableBackgroundAndOuterBorder(graphics, table, x, y, w, h);
                    
                    // Stage 3: Draw cell backgrounds
                    await this._safeDrawCellsBackground(graphics, table, x, y, w, h);
                    
                    // Stage 4: Draw cell content
                    await this._safeDrawCellsContent(graphics, table, x, y, w, h);
                    
                    // Stage 5: Draw cell borders
                    await this._safeDrawCellsBorders(graphics, table, x, y, w, h);
                    
                    // Update performance metrics
                    this._updatePerformanceMetrics(startTime);
                    
                } finally {
                    graphics.RestoreGrState();
                }
                
            } catch (error) {
                console.error('[TableProcessor] MAIN RENDER ERROR:', error.message);
                console.error('[TableProcessor] Error stack:', error.stack);
                console.error('[TableProcessor] Error name:', error.name);
                if (this.logger) {
                    this.logger.logError(this.constructor.name, 'TableProcessor', 'Error during table rendering:', error);
                }
                this.drawTablePlaceholder(graphics, x, y, w, h);
                throw this._enhanceError(error, 'table_rendering', { x, y, w, h });
            }
        }, { context: 'renderTableFrame' })();
    }

    /**
     * Get table instance with enhanced validation and caching
     */
    async getTableInstance(shape) {
        return this.errorBoundary.wrap(async () => {
            // Input validation
            if (!shape || typeof shape !== 'object') {
                throw this._createError('Invalid shape object provided', {
                    name: 'ValidationError',
                    code: 'INVALID_SHAPE'
                });
            }
            
            // Check for existing cached table instance
            if (shape.table && shape.table instanceof CTable) {
                this.performanceMetrics.cacheHits++;
                return shape.table;
            }
            
            this.performanceMetrics.cacheMisses++;
            
            // Parse table from XML if available
            if (shape.graphicData && shape.graphicData.tableXml) {
                try {
                    // Validate XML before parsing
                    if (this.config.enableValidation) {
                        this._validateTableXML(shape.graphicData.tableXml);
                    }
                    
                    const table = await this.parseTableFromXML(shape.graphicData.tableXml);
                    if (table) {
                        // Validate parsed table
                        if (this.config.enableValidation) {
                            this._validateParsedTable(table);
                        }
                        
                        // Cache the parsed table if caching is enabled
                        if (this.config.enableCaching) {
                            shape.table = table;
                        }
                        
                        return table;
                    }
                } catch (error) {
                    if (this.logger) {
                        this.logger.logError(this.constructor.name, 'TableProcessor', 'Failed to parse table XML:', error);
                    }
                    throw this._enhanceError(error, 'table_parsing');
                }
            }
            
            return null;
        }, { context: 'getTableInstance' })();
    }
    
    /**
     * Get table instance with timeout protection
     */
    async _getTableInstanceWithTimeout(shape, timeoutMs) {
        return new Promise(async (resolve, reject) => {
            const timeout = setTimeout(() => {
                reject(this._createError(`Table parsing timed out after ${timeoutMs}ms`, {
                    name: 'TimeoutError',
                    code: 'TABLE_PARSING_TIMEOUT'
                }));
            }, timeoutMs);
            
            try {
                const result = await this.getTableInstance(shape);
                clearTimeout(timeout);
                resolve(result);
            } catch (error) {
                clearTimeout(timeout);
                reject(error);
            }
        });
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

        // FIXED: Only draw outer table border if there are actual borders defined
        // Check if tableBorders has any actual border definitions (not just an empty object)
        if (tableBorders && Object.keys(tableBorders).length > 0) {
            this.drawTableOuterBorder(graphics, tableBorders, x, y, w, h);
        }
    }

    /**
     * Stage 2: Draw cell backgrounds - Fixed for merged cells + table style support
     */
    drawCellsBackground(graphics, table, x, y, w, h) {
        const rows = table.getRows();
        const cellWidths = this.calculateCellWidths(table, w);
        const cellHeights = this.calculateCellHeights(table, h);
        const tableProps = table.getTableProperties() || {};

        let currentY = y;
        
        for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
            const row = rows[rowIndex];
            const cells = row.getCells();
            const rowHeight = cellHeights[rowIndex];
            
            // Check if this is the total row (contains "TOTAL" in first cell)
            const isTotalRow = cells.length > 0 && 
                cells[0].textBody?.paragraphs?.[0]?.runs?.[0]?.text === 'TOTAL';
            
            // Check if this is the header row (contains "Product" in first cell)
            const isHeaderRow = cells.length > 0 && 
                cells[0].textBody?.paragraphs?.[0]?.runs?.[0]?.text === 'Product';
            
            for (let cellIndex = 0; cellIndex < cells.length; cellIndex++) {
                const cell = cells[cellIndex];
                const gridSpan = Math.max(1, cell.gridSpan || 1);
                const rowSpan = Math.max(1, cell.rowSpan || 1);
                
                // Skip cells that are continuations of merged cells
                if (!cell.isMergedContinue) {
                    // Use the stored logical column index from parsing
                    const logicalColIndex = cell.logicalColIndex || 0;
                    
                    // Calculate cell position and dimensions
                    const cellX = x + this.calculateCellXOffset(cellWidths, logicalColIndex);
                    const actualWidth = this.calculateMergedCellWidth(cellWidths, logicalColIndex, gridSpan);
                    const actualHeight = this.calculateMergedCellHeight(cellHeights, rowIndex, rowSpan);
                    
                    // Draw cell background with row context including table properties
                    this.drawCellBackground(graphics, cell, cellX, currentY, actualWidth, actualHeight, {
                        isHeaderRow: isHeaderRow,
                        isTotalRow: isTotalRow,
                        rowIndex: rowIndex,
                        tableProps: tableProps
                    });
                }
            }
            
            currentY += rowHeight;
        }
    }

    /**
     * Stage 3: Draw cell content - Fixed for merged cells
     */
    drawCellsContent(graphics, table, x, y, w, h) {
        const rows = table.getRows();
        const cellWidths = this.calculateCellWidths(table, w);
        const cellHeights = this.calculateCellHeights(table, h);
        let currentY = y;
        
        for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
            const row = rows[rowIndex];
            const cells = row.getCells();
            const rowHeight = cellHeights[rowIndex];
            
            for (let cellIndex = 0; cellIndex < cells.length; cellIndex++) {
                const cell = cells[cellIndex];
                const gridSpan = Math.max(1, cell.gridSpan || 1);
                const rowSpan = Math.max(1, cell.rowSpan || 1);
                
                // Skip cells that are continuations of merged cells - content only in parent
                if (!cell.isMergedContinue) {
                    // Use the stored logical column index from parsing
                    const logicalColIndex = cell.logicalColIndex || 0;

                    // Calculate cell position and dimensions
                    const cellX = x + this.calculateCellXOffset(cellWidths, logicalColIndex);
                    const actualWidth = this.calculateMergedCellWidth(cellWidths, logicalColIndex, gridSpan);
                    const actualHeight = this.calculateMergedCellHeight(cellHeights, rowIndex, rowSpan);

                    // Draw cell content
                    this.drawCellContent(graphics, cell, cellX, currentY, actualWidth, actualHeight);
                }
            }
            
            currentY += rowHeight;
        }
    }

    /**
     * Stage 4: Draw cell borders - Fixed for merged cells
     */
    drawCellsBorders(graphics, table, x, y, w, h) {
        const rows = table.getRows();
        const cellWidths = this.calculateCellWidths(table, w);
        const cellHeights = this.calculateCellHeights(table, h);

        let currentY = y;
        
        // Draw borders for each row
        for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
            const row = rows[rowIndex];
            const cells = row.getCells();
            const rowHeight = cellHeights[rowIndex];
            
            for (let cellIndex = 0; cellIndex < cells.length; cellIndex++) {
                const cell = cells[cellIndex];
                const gridSpan = Math.max(1, cell.gridSpan || 1);
                const rowSpan = Math.max(1, cell.rowSpan || 1);
                
                // Handle border drawing for merged cells
                if (!cell.isMergedContinue) {
                    // Use the stored logical column index from parsing
                    const logicalColIndex = cell.logicalColIndex || 0;
                    
                    // Calculate cell position and dimensions
                    const cellX = x + this.calculateCellXOffset(cellWidths, logicalColIndex);
                    const actualWidth = this.calculateMergedCellWidth(cellWidths, logicalColIndex, gridSpan);
                    const actualHeight = this.calculateMergedCellHeight(cellHeights, rowIndex, rowSpan);
                    
                    // Draw cell borders
                    this.drawCellBorders(graphics, cell, cellX, currentY, actualWidth, actualHeight, rowIndex, logicalColIndex, table);
                }
            }
            
            currentY += rowHeight;
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
                const width = Math.max(border.width || 1, 3); // Minimum 3px borders to match PowerPoint style
                
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
     * Draw cell background using actual PPTX cell shading data + table style formatting
     * Enhanced to apply table style rules (firstRow, bandRow, etc.)
     */
    drawCellBackground(graphics, cell, x, y, w, h, rowContext = {}) {
        // Parse cell background from PPTX data only - no hardcoded fallbacks
        let backgroundColor = null;
        
        // Method 1: Try cell.shading.fillColor (parsed during XML processing)
        if (cell.shading && cell.shading.fillColor) {
            backgroundColor = cell.shading.fillColor;
        }
        
        // Method 2: Try getCellShading() if available
        if (!backgroundColor && typeof cell.getCellShading === 'function') {
            try {
                const cellShading = cell.getCellShading();
                if (cellShading && cellShading.fill && cellShading.fill.color) {
                    backgroundColor = cellShading.fill.color;
                }
            } catch (e) {
                // getCellShading failed, continue with other methods
            }
        }
        
        // Method 3: Try cell.shading.color as backup
        if (!backgroundColor && cell.shading && cell.shading.color) {
            backgroundColor = cell.shading.color;
        }
        
        // Method 4: Try direct cell properties
        if (!backgroundColor) {
            const cellProps = cell.properties || cell.cellProperties;
            if (cellProps && cellProps.shading && cellProps.shading.fillColor) {
                backgroundColor = cellProps.shading.fillColor;
            }
        }
        
        // ENHANCED: Apply table style-based formatting if no explicit cell background
        if (!backgroundColor && rowContext.tableProps) {
            const props = rowContext.tableProps;
            const rowIndex = rowContext.rowIndex || 0;
            const colIndex = cell.logicalColIndex || 0;
            
            // firstRow - header row gets darker background
            if (props.firstRow && rowIndex === 0) {
                backgroundColor = { r: 68, g: 114, b: 196, a: 255 }; // Default blue header
            }
            // bandRow - alternating row colors
            else if (props.bandRow && rowIndex > 0) {
                // Odd rows get light gray background (skip header if firstRow is true)
                const effectiveRow = props.firstRow ? rowIndex - 1 : rowIndex;
                if (effectiveRow % 2 === 1) {
                    backgroundColor = { r: 242, g: 242, b: 242, a: 255 }; // Light gray
                }
            }
            // firstCol - first column gets special formatting
            else if (props.firstCol && colIndex === 0 && rowIndex > 0) {
                backgroundColor = { r: 221, g: 235, b: 247, a: 255 }; // Light blue
            }
        }
        
        // Draw background if we have a color (from cell or table style)
        if (backgroundColor && 
            backgroundColor.r !== undefined && 
            backgroundColor.g !== undefined && 
            backgroundColor.b !== undefined && 
            graphics.m_oContext) {
            
            graphics.m_oContext.save();
            graphics.m_oContext.fillStyle = `rgba(${backgroundColor.r}, ${backgroundColor.g}, ${backgroundColor.b}, ${(backgroundColor.a || 255) / 255})`;
            graphics.m_oContext.fillRect(x, y, w, h);
            graphics.m_oContext.restore();
        }
    }

    /**
     * Draw cell content with proper PPTX margins and alignment
     */
    drawCellContent(graphics, cell, x, y, w, h) {
        const textBody = cell.getTextBody();
        if (!textBody || !textBody.paragraphs) {return;}

        // FIXED: Apply both cell-level margins AND text body insets
        // Cell margins (tcMar) define the outer padding of the cell
        // Text body insets (lIns, rIns, tIns, bIns) define additional inner padding for text
        
        // Start with cell-level margins (default 1mm in PPTX)
        const defaultMarginEMU = 38100; // 1mm in EMU
        const defaultMarginPx = this.convertEMUToPixels(defaultMarginEMU);
        
        let leftMargin = defaultMarginPx;
        let rightMargin = defaultMarginPx;
        let topMargin = defaultMarginPx;
        let bottomMargin = defaultMarginPx;
        
        // Apply cell-specific margins if available
        if (cell.margins) {
            leftMargin = cell.margins.left ? this.convertEMUToPixels(cell.margins.left) : defaultMarginPx;
            rightMargin = cell.margins.right ? this.convertEMUToPixels(cell.margins.right) : defaultMarginPx;
            topMargin = cell.margins.top ? this.convertEMUToPixels(cell.margins.top) : defaultMarginPx;
            bottomMargin = cell.margins.bottom ? this.convertEMUToPixels(cell.margins.bottom) : defaultMarginPx;
        }
        
        // Text body insets override cell margins when present (not additive)
        if (textBody.bodyProperties) {
            const props = textBody.bodyProperties;
            if (props.leftMargin != null) {
                leftMargin = this.convertEMUToPixels(props.leftMargin);
            }
            if (props.rightMargin != null) {
                rightMargin = this.convertEMUToPixels(props.rightMargin);
            }
            if (props.topMargin != null) {
                topMargin = this.convertEMUToPixels(props.topMargin);
            }
            if (props.bottomMargin != null) {
                bottomMargin = this.convertEMUToPixels(props.bottomMargin);
            }
        }

        const contentX = x + leftMargin;
        const contentY = y + topMargin;
        const contentW = Math.max(0, w - (leftMargin + rightMargin));
        const contentH = Math.max(0, h - (topMargin + bottomMargin));

        // Render text content using enhanced text rendering with proper alignment
        this.renderCellTextContent(graphics, textBody, contentX, contentY, contentW, contentH, cell);
    }

    /**
     * Draw cell borders with conflict resolution - Fixed to only draw specified borders
     */
    drawCellBorders(graphics, cell, x, y, w, h, rowIndex, colIndex, table) {
        const cellBorders = cell.getCellBorders();
        if (!cellBorders) {
            return;
        }

        const borders = ['top', 'right', 'bottom', 'left'];
        
        borders.forEach(borderSide => {
            const border = cellBorders[borderSide];
            
            // FIXED: Only draw borders that are explicitly defined in PPTX
            // Don't draw default borders - only draw what's specified
            if (!border || border.style === 'none' || !border.color || border.width === 0) {
                return;
            }

            // Border conflict resolution - pass cell for merged cell logic
            if (this.shouldDrawBorder(borderSide, rowIndex, colIndex, table, cell)) {
                const color = this.resolveBorderColor(border);
                // FIXED: Use actual border width from PPTX - don't fallback to 1 if width is 0
                const width = border.width;
                
                if (graphics.m_oContext) {
                    graphics.m_oContext.save();
                    graphics.m_oContext.strokeStyle = `rgba(${color.r}, ${color.g}, ${color.b}, ${(color.a || 255) / 255})`;
                    graphics.m_oContext.lineWidth = width;
                    graphics.m_oContext.beginPath();
                    
                    switch (borderSide) {
                        case 'top':
                            graphics.m_oContext.moveTo(x, y);
                            graphics.m_oContext.lineTo(x + w, y);
                            break;
                        case 'right':
                            graphics.m_oContext.moveTo(x + w, y);
                            graphics.m_oContext.lineTo(x + w, y + h);
                            break;
                        case 'bottom':
                            graphics.m_oContext.moveTo(x, y + h);
                            graphics.m_oContext.lineTo(x + w, y + h);
                            break;
                        case 'left':
                            graphics.m_oContext.moveTo(x, y);
                            graphics.m_oContext.lineTo(x, y + h);
                            break;
                    }
                    
                    graphics.m_oContext.stroke();
                    graphics.m_oContext.restore();
                }
            }
        });
    }

    /**
     * Border conflict resolution - Enhanced for merged cells and internal borders
     */
    shouldDrawBorder(borderSide, rowIndex, colIndex, table, cell = null) {
        const rowCount = table.getRowCount();
        const colCount = table.getColumnCount();
        
        // Get cell span information if cell is provided
        const gridSpan = cell ? Math.max(1, cell.gridSpan || 1) : 1;
        const rowSpan = cell ? Math.max(1, cell.rowSpan || 1) : 1;
        
        switch (borderSide) {
            case 'top':
                // Draw top border for first row and for any row that's not a continuation of a vertical merge
                if (rowIndex === 0) {return true;}
                // Check if this cell is a continuation of a vertical merge
                if (cell && cell.isMergedContinue && cell.vMerge === 'continue') {return false;}
                return true;
                
            case 'right':
                // Draw right border for last column or when cell doesn't span to the next column
                const rightmostCol = colIndex + gridSpan - 1;
                if (rightmostCol === colCount - 1) {return true;}
                // Draw internal borders to show cell separation
                return true;
                
            case 'bottom':
                // Draw bottom border for last row or when cell doesn't span to the next row
                const bottommostRow = rowIndex + rowSpan - 1;
                if (bottommostRow === rowCount - 1) {return true;}
                // Draw internal borders to show cell separation
                return true;
                
            case 'left':
                // Draw left border for first column and internal borders
                if (colIndex === 0) {return true;}
                // For merged cells, don't draw internal left borders within the merge
                // This is determined by checking if the previous cell extends into this position
                return true;
                
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

        // Normalize EMU widths and scale to fit totalWidth proportionally
        const widths = tableGrid.map(col => Math.max(0, parseInt(col.width) || 0));
        const totalGridWidth = widths.reduce((sum, w) => sum + w, 0) || 1;
        const scaled = widths.map(w => totalWidth * (w / totalGridWidth));

        // If fewer columns than cells, pad remaining equally
        const colCount = table.getColumnCount();
        if (scaled.length < colCount) {
            const remaining = colCount - scaled.length;
            const used = scaled.reduce((s, v) => s + v, 0);
            const pad = (totalWidth - used) / Math.max(remaining, 1);
            for (let i = 0; i < remaining; i++) {scaled.push(pad);}
        }
        return scaled;
    }

    /**
     * Calculate cell heights
     */
    calculateCellHeights(table, totalHeight) {
        const rowCount = table.getRowCount();
        if (rowCount <= 0) {return [];}

        // Use row.height (EMU) if provided; otherwise distribute equally
        const unitHeights = [];
        for (let i = 0; i < rowCount; i++) {
            const row = table.getRow(i);
            const emu = row && row.height;
            const n = (emu !== undefined && emu !== null) ? Number(emu) : NaN;
            unitHeights.push(isNaN(n) ? 1 : Math.max(n, 1));
        }
        const totalUnits = unitHeights.reduce((s, v) => s + v, 0) || rowCount;
        return unitHeights.map(u => totalHeight * (u / totalUnits));
    }

    /**
     * Check if cell is the starting cell of a merged range - Removed for new merge handling
     */
    isStartingCellOfMerge(cell, rowIndex, colIndex) {
        // Always return true now since we handle merging in the coordinate calculation
        return true;
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
        if (tableShading && tableShading.fill) {
            const color = this.parseColorFromHex(tableShading.fill);
            if (color) {return color;}
        }
        // Default white background
        return { r: 255, g: 255, b: 255, a: 255 };
    }

    /**
     * Resolve cell background color
     */
    resolveCellBackgroundColor(cellShading) {
        if (cellShading && cellShading.fillColor) {
            return cellShading.fillColor;
        }
        if (cellShading && cellShading.fill) {
            const color = this.parseColorFromHex(cellShading.fill);
            if (color) {return color;}
        }
        if (cellShading && cellShading.color) {
            const color = this.parseColorFromHex(cellShading.color);
            if (color) {return color;}
        }
        // Default white background
        return { r: 255, g: 255, b: 255, a: 255 };
    }

    /**
     * Resolve border color
     */
    resolveBorderColor(border) {
        if (border && border.color) {
            return border.color;
        }
        // Default black border
        return { r: 0, g: 0, b: 0, a: 255 };
    }

    /**
     * Render cell text content with proper PPTX alignment, color, and font size
     * Enhanced to support per-run styling and proper vertical alignment
     */
    renderCellTextContent(graphics, textBody, x, y, w, h, cell = null) {
        if (!textBody.paragraphs || textBody.paragraphs.length === 0) {
            return;
        }

        // Build lines with per-run styling support
        const lines = [];
        for (let p = 0; p < textBody.paragraphs.length; p++) {
            const paragraph = textBody.paragraphs[p];
            if (!paragraph.runs || paragraph.runs.length === 0) {continue;}

            // Get paragraph-level alignment
            let textAlign = 'left';
            if (paragraph.properties && paragraph.properties.alignment) {
                const a = paragraph.properties.alignment;
                if (a === 'center' || a === 'ctr') { textAlign = 'center'; }
                else if (a === 'right' || a === 'r') { textAlign = 'right'; }
                else if (a === 'just' || a === 'justify') { textAlign = 'left'; } // Justify treated as left for now
            }

            // Process each run with its own styling
            const runSegments = [];
            for (const run of paragraph.runs) {
                const text = run.text || '';
                if (!text) continue;
                
                const runProps = run.properties || {};
                let baseFontSize = 12;
                if (runProps.fontSize) { baseFontSize = runProps.fontSize; }
                else if (runProps.sz) { baseFontSize = parseInt(runProps.sz) / 100; }
                
                const fontSizePx = Math.max(10, Math.round(baseFontSize * 1.33));
                const fontFamily = runProps.fontFamily || 'Arial';
                const fontWeight = runProps.bold ? 'bold' : 'normal';
                const fontStyle = runProps.italic ? 'italic' : 'normal';
                const underline = runProps.underline || false;
                const strikethrough = runProps.strikethrough || false;
                const colorObj = runProps.color || { r: 0, g: 0, b: 0 };
                const color = `rgb(${colorObj.r}, ${colorObj.g}, ${colorObj.b})`;
                
                runSegments.push({
                    text,
                    fontSizePx,
                    fontFamily,
                    fontWeight,
                    fontStyle,
                    underline,
                    strikethrough,
                    color
                });
            }
            
            if (runSegments.length === 0) continue;
            
            // Calculate line height from largest font in paragraph
            const maxFontSize = Math.max(...runSegments.map(r => r.fontSizePx));
            const lineHeight = Math.round(maxFontSize * 1.25);
            
            lines.push({ runSegments, textAlign, lineHeight });
        }

        if (lines.length === 0) { return; }

        // ENHANCED: Use vertical alignment from textBody.bodyProperties.anchor or cell.verticalAlignment
        const totalHeight = lines.reduce((s, ln) => s + ln.lineHeight, 0);
        let vAlign = 'ctr'; // Default to center (PPTX default)
        
        // Priority 1: Cell-level vertical alignment
        if (cell && cell.verticalAlignment) {
            vAlign = cell.verticalAlignment;
        }
        // Priority 2: TextBody anchor property
        else if (textBody.bodyProperties && textBody.bodyProperties.anchor) {
            vAlign = textBody.bodyProperties.anchor;
        }
        
        // Calculate starting Y position based on vertical alignment
        let cursorY = y;
        if (vAlign === 't' || vAlign === 'top') {
            cursorY = y;
        } else if (vAlign === 'ctr' || vAlign === 'center' || vAlign === 'middle') {
            cursorY = y + Math.max(0, (h - totalHeight) / 2);
        } else if (vAlign === 'b' || vAlign === 'bottom') {
            cursorY = y + Math.max(0, h - totalHeight);
        } else {
            // Default to center if unknown value
            cursorY = y + Math.max(0, (h - totalHeight) / 2);
        }
        
        if (!isFinite(cursorY)) { cursorY = y; }

        if (graphics.m_oContext) {
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i];
                const yMid = cursorY + line.lineHeight / 2;
                
                // Calculate total line width for alignment
                let totalLineWidth = 0;
                graphics.m_oContext.save();
                for (const seg of line.runSegments) {
                    graphics.m_oContext.font = `${seg.fontStyle} ${seg.fontWeight} ${seg.fontSizePx}px ${seg.fontFamily}`;
                    totalLineWidth += graphics.m_oContext.measureText(seg.text).width;
                }
                graphics.m_oContext.restore();
                
                // Calculate starting X based on alignment
                let currentX = x;
                if (line.textAlign === 'center') {
                    currentX = x + (w - totalLineWidth) / 2;
                } else if (line.textAlign === 'right') {
                    currentX = x + w - totalLineWidth;
                }
                
                // Render each run segment with its own styling
                for (const seg of line.runSegments) {
                    graphics.m_oContext.save();
                    graphics.m_oContext.fillStyle = seg.color;
                    graphics.m_oContext.font = `${seg.fontStyle} ${seg.fontWeight} ${seg.fontSizePx}px ${seg.fontFamily}`;
                    graphics.m_oContext.textAlign = 'left';
                    graphics.m_oContext.textBaseline = 'middle';
                    graphics.m_oContext.fillText(seg.text, currentX, yMid);
                    
                    // Add underline if specified
                    if (seg.underline) {
                        const textWidth = graphics.m_oContext.measureText(seg.text).width;
                        graphics.m_oContext.strokeStyle = seg.color;
                        graphics.m_oContext.lineWidth = 1;
                        graphics.m_oContext.beginPath();
                        graphics.m_oContext.moveTo(currentX, yMid + seg.fontSizePx * 0.1);
                        graphics.m_oContext.lineTo(currentX + textWidth, yMid + seg.fontSizePx * 0.1);
                        graphics.m_oContext.stroke();
                    }
                    
                    // Add strikethrough if specified
                    if (seg.strikethrough) {
                        const textWidth = graphics.m_oContext.measureText(seg.text).width;
                        graphics.m_oContext.strokeStyle = seg.color;
                        graphics.m_oContext.lineWidth = 1;
                        graphics.m_oContext.beginPath();
                        graphics.m_oContext.moveTo(currentX, yMid);
                        graphics.m_oContext.lineTo(currentX + textWidth, yMid);
                        graphics.m_oContext.stroke();
                    }
                    
                    currentX += graphics.m_oContext.measureText(seg.text).width;
                    graphics.m_oContext.restore();
                }
                
                cursorY += line.lineHeight;
            }
        }
    }

    /**
     * Resolve text color with proper PPTX defaults
     */
    resolveTextColor(paragraph) {
        if (paragraph && paragraph.runs && paragraph.runs.length > 0) {
            for (const run of paragraph.runs) {
                if (run.properties && run.properties.color) {
                    return run.properties.color;
                }
            }
        }
        // Default gray text for PPTX tables (#9F9F9F)
        return { r: 159, g: 159, b: 159, a: 255 };
    }

    /**
     * Draw centered text
     */
    drawCenteredText(graphics, text, x, y, w, h) {
        if (!graphics.m_oContext) {return;}

        graphics.m_oContext.save();
        graphics.m_oContext.fillStyle = 'rgba(128, 128, 128, 0.8)';
        graphics.m_oContext.font = '14px Arial';
        graphics.m_oContext.textAlign = 'center';
        graphics.m_oContext.textBaseline = 'middle';
        graphics.m_oContext.fillText(text, x + w / 2, y + h / 2);
        graphics.m_oContext.restore();
    }

    /**
     * Enhanced table parsing
     */
    parseTableFromXML(tableXml) {
        if (!tableXml) {
            return null;
        }

        try {
            const parser = new DOMParser();
            const doc = parser.parseFromString(tableXml, 'text/xml');
            
            if (doc.documentElement.nodeName === 'parsererror') {
                console.warn('XML parsing error in table XML');
                return null;
            }

            const table = new CTable();
            
            // Initialize row data tracking for merged cell processing
            this.tableRowData = [];
            
            // Stage 1: Parse table properties
            this.parseTableProperties(doc, table);
            
            // Stage 2: Parse table grid (column definitions)
            this.parseTableGrid(doc, table);
            
            // Stage 3: Parse table rows and cells
            this.parseTableRows(doc, table);
            
            // Stage 4: Process merged cells and relationships
            this.processTableMerges(table);
            
            // Stage 5: Validate table structure (non-fatal for merged layouts)
            if (!this.validateTableStructure(table)) {
                console.warn('Table structure validation failed');
                // Debug grid and a few rows to diagnose span issues
                try {
                    const gridCols = (table.getTableGrid() || []).length;
                    console.warn('[TableProcessor][DEBUG] Grid columns:', gridCols);
                    const rows = table.getRows();
                    for (let r = 0; r < Math.min(rows.length, 6); r++) {
                        const cells = rows[r].getCells();
                        const details = cells.map((c, i) => `c${i}{gs:${c.gridSpan||1},rs:${c.rowSpan||1},cont:${!!c.isMergedContinue},v:${c.vMerge||''},lc:${c.logicalColIndex||0}}`).join(' | ');
                        console.warn(`[TableProcessor][DEBUG] Row ${r} cells(${cells.length}): ${details}`);
                    }
                } catch (_e) {}
                // Continue rendering with best-effort table to ensure visibility
            }

            return table;

        } catch (error) {
            console.warn('Error parsing table XML:', error);
            return null;
        } finally {
            // Clean up table row data tracking
            this.tableRowData = null;
        }
    }

    /**
     * Parse table properties - Enhanced with table style flags
     */
    parseTableProperties(doc, table) {
        // Find table properties element
        let tblPr = doc.querySelector('tblPr');
        if (!tblPr) {tblPr = doc.querySelector('a\\:tblPr');}
        
        if (tblPr) {
            const tableProps = {};
            
            // ENHANCED: Parse table style flags
            // firstRow - apply special formatting to first row (header)
            const firstRow = tblPr.getAttribute('firstRow');
            if (firstRow === '1' || firstRow === 'true') {
                tableProps.firstRow = true;
            }
            
            // bandRow - apply alternating row colors
            const bandRow = tblPr.getAttribute('bandRow');
            if (bandRow === '1' || bandRow === 'true') {
                tableProps.bandRow = true;
            }
            
            // firstCol - apply special formatting to first column
            const firstCol = tblPr.getAttribute('firstCol');
            if (firstCol === '1' || firstCol === 'true') {
                tableProps.firstCol = true;
            }
            
            // lastRow - apply special formatting to last row (total row)
            const lastRow = tblPr.getAttribute('lastRow');
            if (lastRow === '1' || lastRow === 'true') {
                tableProps.lastRow = true;
            }
            
            // lastCol - apply special formatting to last column
            const lastCol = tblPr.getAttribute('lastCol');
            if (lastCol === '1' || lastCol === 'true') {
                tableProps.lastCol = true;
            }
            
            // bandCol - apply alternating column colors
            const bandCol = tblPr.getAttribute('bandCol');
            if (bandCol === '1' || bandCol === 'true') {
                tableProps.bandCol = true;
            }
            
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
            
            table.tableProperties = tableProps;
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

        // Fallback: many PPTX writers (incl. PptxGenJS) encode row height on the <a:tr> "h" attribute.
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
     * Parse row cells - Enhanced merged cell handling with proper column tracking
     */
    parseRowCells(trElement, row, rowIndex) {
        let tableCells = trElement.querySelectorAll('tc');
        if (tableCells.length === 0) {
            tableCells = trElement.querySelectorAll('a\\:tc');
        }
        
        // Get actual column count from table grid XML directly
        const tableElement = trElement.closest('tbl') || trElement.closest('a\\:tbl');
        let actualColumnCount = 10; // fallback
        if (tableElement) {
            let tblGrid = tableElement.querySelector('tblGrid');
            if (!tblGrid) {tblGrid = tableElement.querySelector('a\\:tblGrid');}
            if (tblGrid) {
                let gridCols = tblGrid.querySelectorAll('gridCol');
                if (gridCols.length === 0) {gridCols = tblGrid.querySelectorAll('a\\:gridCol');}
                actualColumnCount = gridCols.length;
            }
        }
        
        // Track which logical column positions are occupied by merged cells from this and previous rows
        const occupiedColumns = new Array(actualColumnCount).fill(false);
        
        // CRITICAL FIX: Mark columns occupied by vertical merges from previous rows
        if (rowIndex > 0 && this.tableRowData) {
            for (let prevRowIndex = 0; prevRowIndex < rowIndex; prevRowIndex++) {
                const prevRowData = this.tableRowData[prevRowIndex];
                if (prevRowData) {
                    prevRowData.forEach(prevCell => {
                        if (prevCell.rowSpan > 1) {
                            const spanEnd = prevRowIndex + prevCell.rowSpan - 1;
                            if (spanEnd >= rowIndex) {
                                // This cell spans into the current row
                                const startCol = prevCell.logicalColIndex;
                                const endCol = startCol + (prevCell.gridSpan || 1);
                                for (let col = startCol; col < endCol; col++) {
                                    if (col < actualColumnCount) {
                                        occupiedColumns[col] = true;
                                    }
                                }
                            }
                        }
                    });
                }
            }
        }
        
        // Initialize table row data tracking if not exists
        if (!this.tableRowData) {
            this.tableRowData = [];
        }
        this.tableRowData[rowIndex] = [];

        // Track how many cells remain as implicit horizontal span continuations
        // PptxGenJS does NOT set hMerge on continuation cells; they're just empty <a:tc> elements
        let remainingHSpanCells = 0;

        tableCells.forEach((tcElement, cellIndex) => {
            // Pre-parse merge attributes - check both as <a:tc> attribute (PptxGenJS) and inside <a:tcPr> (PowerPoint)
            const vMergeAttr = tcElement.getAttribute('vMerge');
            const gridSpanAttr = tcElement.getAttribute('gridSpan');
            const rowSpanAttr = tcElement.getAttribute('rowSpan');
            // Some producers (e.g., PowerPoint) emit hMerge="1" placeholders after a gridSpan cell
            // Treat them as horizontal continuation cells that should not contribute new span coverage
            const hMergeAttr = tcElement.getAttribute('hMerge');

            // Also check vMerge inside <a:tcPr> for PowerPoint-generated PPTX
            let tcPrEl = tcElement.querySelector('tcPr');
            if (!tcPrEl) tcPrEl = tcElement.querySelector('a\\:tcPr');
            let vMergeInTcPr = null;
            if (tcPrEl) {
                vMergeInTcPr = tcPrEl.querySelector('vMerge') || tcPrEl.querySelector('a\\:vMerge');
            }

            // In PPTX, horizontal merging is handled via gridSpan, not hMerge
            // hMerge="1" marks the placeholder cell that visually continues the previous cell.
            // PptxGenJS does NOT set hMerge — continuation cells are just empty <a:tc> elements.
            // We detect them via remainingHSpanCells counter set after processing a gridSpan cell.
            const isHorizontalContinuation = hMergeAttr === '1' || hMergeAttr === 'true' || remainingHSpanCells > 0;
            // Only check for vertical merge continuation (vMerge="1" or "true" or "continue")
            // - as <a:tc vMerge="1"> attribute (PptxGenJS style)
            // - OR as <a:tcPr><a:vMerge val="cont"/> (PowerPoint style, where val != "restart")
            const vMergeFromAttr = vMergeAttr === '1' || vMergeAttr === 'true' || vMergeAttr === 'continue';
            const vMergeFromTcPr = vMergeInTcPr && vMergeInTcPr.getAttribute('val') !== 'restart';
            const isVerticalContinuation = vMergeFromAttr || vMergeFromTcPr;

            // gridSpan and rowSpan may be on <a:tc> attributes (PptxGenJS) OR inside <a:tcPr> (PowerPoint)
            let cellGridSpan = gridSpanAttr ? parseInt(gridSpanAttr) : 1;
            let cellRowSpan = rowSpanAttr ? parseInt(rowSpanAttr) : 1;
            if (tcPrEl) {
                const gsEl = tcPrEl.querySelector('gridSpan') || tcPrEl.querySelector('a\\:gridSpan');
                if (gsEl) { const v = parseInt(gsEl.getAttribute('val')); if (v > 1) cellGridSpan = v; }
                const rsEl = tcPrEl.querySelector('rowSpan') || tcPrEl.querySelector('a\\:rowSpan');
                if (rsEl) { const v = parseInt(rsEl.getAttribute('val')); if (v > 1) cellRowSpan = v; }
            }
            
            // Find logical column position
            let logicalColumnIndex = 0;
            if (isVerticalContinuation) {
                // Vertical continuation cells must be placed at the parent's occupied column
                // (they inherit the column from their rowSpan parent above)
                while (logicalColumnIndex < actualColumnCount && !occupiedColumns[logicalColumnIndex]) {
                    logicalColumnIndex++;
                }
                if (logicalColumnIndex >= actualColumnCount) {logicalColumnIndex = 0;}
            } else {
                // Regular/horizontal cells: find next unoccupied column
                while (logicalColumnIndex < actualColumnCount && occupiedColumns[logicalColumnIndex]) {
                    logicalColumnIndex++;
                }
                if (logicalColumnIndex >= actualColumnCount) {
                    logicalColumnIndex = Math.min(cellIndex, actualColumnCount - 1);
                }
            }
            
            const cell = new CTableCell();
            
            // Set merge attributes
            if (isVerticalContinuation) {
                cell.isMergedContinue = true;
                cell.vMerge = 'continue';
            }
            
            if (cellRowSpan > 1) {
                cell.rowSpan = cellRowSpan;
            }
            
            if (cellGridSpan > 1) {
                cell.gridSpan = cellGridSpan;
            }
            
            // Mark columns as occupied
            // For regular cells: mark gridSpan columns
            // For vertical continuations: mark only 1 column (they don't span horizontally)
            // For horizontal continuation placeholders: don't mark anything (already covered by the starter)
            const columnsToMark = isHorizontalContinuation ? 0 : (isVerticalContinuation ? 1 : cellGridSpan);
            for (let i = 0; i < columnsToMark && (logicalColumnIndex + i) < actualColumnCount; i++) {
                occupiedColumns[logicalColumnIndex + i] = true;
            }

            // Update the implicit horizontal span counter
            if (isHorizontalContinuation) {
                if (remainingHSpanCells > 0) remainingHSpanCells--;
            } else if (!isVerticalContinuation && cellGridSpan > 1) {
                remainingHSpanCells = cellGridSpan - 1;
            } else {
                remainingHSpanCells = 0;
            }

            // Parse cell properties from tcPr element
            this.parseCellProperties(tcElement, cell, rowIndex, cellIndex);

            // Parse cell content
            this.parseCellContent(tcElement, cell);

            // Set flags for horizontal continuation so renderer skips it
            if (isHorizontalContinuation) {
                cell.isMergedContinue = true;
                cell.mergeParent = null; // parent resolved later if needed
            }

            // Set logical position for merged cells
            cell.logicalRowIndex = rowIndex;
            cell.logicalColIndex = logicalColumnIndex;

            // Store cell data for vertical merge tracking
            this.tableRowData[rowIndex].push({
                logicalColIndex: logicalColumnIndex,
                gridSpan: cellGridSpan,
                rowSpan: cellRowSpan,
                isVerticalContinuation: isVerticalContinuation
            });
            
            // Insert the cell  
            row.addCell(cell);
        });
    }

    /**
     * Parse cell properties - Enhanced to capture all cell-level styling
     */
    parseCellProperties(tcElement, cell, rowIndex, cellIndex) {
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

            // Vertical merge
            let vMerge = tcPr.querySelector('vMerge');
            if (!vMerge) {vMerge = tcPr.querySelector('a\\:vMerge');}
            if (vMerge) {
                const val = vMerge.getAttribute('val');
                cell.vMerge = val === 'restart' ? 'start' : 'continue';
            }

            // ENHANCED: Vertical alignment (anchor)
            // In PPTX, cell vertical alignment can be specified at cell level
            const anchor = tcPr.getAttribute('anchor');
            if (anchor) {
                // Values: 't' (top), 'ctr' (center), 'b' (bottom)
                cell.verticalAlignment = anchor;
            }
            
            // ENHANCED: Text direction
            const vert = tcPr.getAttribute('vert');
            if (vert) {
                // Values: 'horz' (horizontal), 'vert' (vertical), 'vert270' (rotated)
                cell.textDirection = vert;
            }

            // Cell borders - parse direct line elements (OnlyOffice table format)
            const borders = this.parseCellBordersFromTcPr(tcPr);
            if (borders) {
                cell.borders = borders;
            }

            // Cell shading - look for solid fill in tcPr (but NOT in border elements)
            // Find solidFill that is a direct child of tcPr, not nested in border elements
            const directSolidFill = Array.from(tcPr.children).find(child => 
                child.tagName === 'a:solidFill' || child.tagName === 'solidFill'
            );
            
            if (directSolidFill) {
                cell.shading = this.parseCellShading(directSolidFill);
            } else {
                // Look for solidFill but exclude those inside border elements (lnL, lnR, lnT, lnB)
                const allSolidFills = tcPr.querySelectorAll('solidFill, a\\:solidFill');
                let cellSolidFill = null;
                
                for (const fill of allSolidFills) {
                    // Check if this solidFill is inside a border element
                    let parent = fill.parentElement;
                    let isInBorder = false;
                    
                    while (parent && parent !== tcPr) {
                        if (parent.tagName === 'a:lnL' || parent.tagName === 'a:lnR' || 
                            parent.tagName === 'a:lnT' || parent.tagName === 'a:lnB' ||
                            parent.tagName === 'lnL' || parent.tagName === 'lnR' || 
                            parent.tagName === 'lnT' || parent.tagName === 'lnB') {
                            isInBorder = true;
                            break;
                        }
                        parent = parent.parentElement;
                    }
                    
                    if (!isInBorder) {
                        cellSolidFill = fill;
                        break;
                    }
                }
                
                if (cellSolidFill) {
                    cell.shading = this.parseCellShading(cellSolidFill);
                } else {
                }
            }

            // Cell margins - from <tcMar> child OR direct attributes on tcPr (PptxGenJS style)
            let tcMar = tcPr.querySelector('tcMar');
            if (!tcMar) {tcMar = tcPr.querySelector('a\\:tcMar');}
            if (tcMar) {
                cell.margins = this.parseCellMargins(tcMar);
            } else {
                // PptxGenJS stores margins as attributes: marL, marR, marT, marB (EMU)
                const marL = tcPr.getAttribute('marL');
                const marR = tcPr.getAttribute('marR');
                const marT = tcPr.getAttribute('marT');
                const marB = tcPr.getAttribute('marB');
                if (marL || marR || marT || marB) {
                    cell.margins = {
                        left: marL ? parseInt(marL) : null,
                        right: marR ? parseInt(marR) : null,
                        top: marT ? parseInt(marT) : null,
                        bottom: marB ? parseInt(marB) : null
                    };
                }
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
     * Append child's text body into parent cell's text body for merged cells
     * Ensures content from continuation cells is shown once in the parent
     */
    appendMergedCellText(parentCell, childCell) {
        if (!childCell || childCell.mergedContentTransferred) {return;}
        const childBody = childCell.getTextBody && childCell.getTextBody();
        if (!childBody || !childBody.paragraphs || childBody.paragraphs.length === 0) {return;}
        if (!parentCell.getTextBody || !parentCell.setTextBody) {return;}
        const parentBody = parentCell.getTextBody();
        if (!parentBody || !parentBody.paragraphs) {
            // If parent has no body, adopt child's body directly
            parentCell.setTextBody({ paragraphs: [...childBody.paragraphs] });
        } else {
            // Insert a blank paragraph separator if parent already has content
            if (parentBody.paragraphs.length > 0) {
                parentBody.paragraphs.push({ runs: [], properties: {} });
            }
            parentBody.paragraphs.push(...childBody.paragraphs);
        }
        // Mark as transferred and clear child's content to avoid accidental drawing
        childCell.mergedContentTransferred = true;
        childCell.setTextBody && childCell.setTextBody({ paragraphs: [] });
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

                // Handle vertical merges using vMerge semantics (restart/continue)
                // Some PPTX files specify vertical merging via <a:vMerge val="restart|continue"> without explicit rowSpan
                if (cell.vMerge === 'start') {
                    // Count how many subsequent rows continue the merge in the same column
                    let spanCount = 1;
                    for (let r = rowIndex + 1; r < rows.length; r++) {
                        const belowCell = table.getCell(r, colIndex);
                        if (!belowCell) {
                            break;
                        }
                        // Treat missing val as continuation per PresentationML behavior (<a:vMerge/> means continue)
                        if (belowCell.vMerge === 'continue') {
                            belowCell.isMergedContinue = true;
                            belowCell.mergeParent = { row: rowIndex, col: colIndex };
                            // Pull text up into the parent cell
                            this.appendMergedCellText(cell, belowCell);
                            spanCount++;
                        } else {
                            break;
                        }
                    }
                    // Ensure the starting cell's rowSpan reflects the number of continued cells
                    if (!cell.rowSpan || cell.rowSpan < spanCount) {
                        cell.rowSpan = spanCount;
                    }
                } else if (cell.vMerge === 'continue') {
                    // If we encounter a continuation that wasn't marked yet (no parent processed above),
                    // attempt to find its merge start upwards and mark accordingly
                    if (!cell.isMergedContinue) {
                        for (let r = rowIndex - 1; r >= 0; r--) {
                            const aboveCell = table.getCell(r, colIndex);
                            if (!aboveCell) { break; }
                            if (aboveCell.vMerge === 'start') {
                                cell.isMergedContinue = true;
                                cell.mergeParent = { row: r, col: colIndex };
                                // Also ensure the parent rowSpan accounts for this continuation
                                const expectedSpan = rowIndex - r + 1;
                                if (!aboveCell.rowSpan || aboveCell.rowSpan < expectedSpan) {
                                    aboveCell.rowSpan = expectedSpan;
                                }
                                break;
                            } else if (aboveCell.vMerge !== 'continue') {
                                // Hit a non-merged cell; stop searching
                                break;
                            }
                        }
                    }
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
                // Move textual content from continuation into the parent cell
                const parentCell = table.getCell(rowIndex, startCol);
                if (parentCell) {
                    this.appendMergedCellText(parentCell, targetCell);
                }
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
                // Move textual content from continuation into the parent cell
                const parentCell = table.getCell(startRow, colIndex);
                if (parentCell) {
                    this.appendMergedCellText(parentCell, targetCell);
                }
            }
        }
    }

    /**
     * Validate table structure - Updated for merged cells
     */
    validateTableStructure(table) {
        const rows = table.getRows();
        if (rows.length === 0) {return false;}
        
        const expectedCols = table.getTableGrid().length || table.getColumnCount();
        if (expectedCols === 0) {return false;}
        
        // For merged cell tables, validate using row-available columns (subtract vMerge coverage from above)
        for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
            const row = rows[rowIndex];
            const cells = row.getCells();
            let totalSpan = 0;
            
            for (const cell of cells) {
                // Skip merged continuation cells from span calculation (includes hMerge placeholders and vMerge continues)
                if (!cell.isMergedContinue) {
                    const gridSpan = Math.max(1, cell.gridSpan || 1);
                    totalSpan += gridSpan;
                }
            }
            
            // Determine how many columns in this row are already covered by rowSpans from previous rows
            let columnsCoveredFromAbove = 0;
            if (this.tableRowData) {
                const covered = new Array(expectedCols).fill(false);
                for (let prevRow = 0; prevRow < rowIndex; prevRow++) {
                    const prevCells = this.tableRowData[prevRow] || [];
                    for (const prevCell of prevCells) {
                        if (prevCell && prevCell.rowSpan > 1) {
                            const endRow = prevRow + prevCell.rowSpan - 1;
                            if (endRow >= rowIndex) {
                                const startCol = Math.max(0, prevCell.logicalColIndex || 0);
                                const endCol = Math.min(expectedCols, startCol + Math.max(1, prevCell.gridSpan || 1));
                                for (let c = startCol; c < endCol; c++) {
                                    covered[c] = true;
                                }
                            }
                        }
                    }
                }
                columnsCoveredFromAbove = covered.filter(Boolean).length;
            }
            const availableColsThisRow = Math.max(0, expectedCols - columnsCoveredFromAbove);
            
            // If all columns are covered from above, this row is valid even with totalSpan 0/low
            if (availableColsThisRow === 0) {
                continue;
            }
            
            // Compare against available columns for this row
            if (totalSpan < availableColsThisRow * 0.5 || totalSpan > availableColsThisRow * 2.0) {
                console.warn(`Row ${rowIndex}: total span ${totalSpan} mismatches available columns ${availableColsThisRow} (grid cols ${expectedCols}, covered from above ${columnsCoveredFromAbove})`);
                try {
                    const details = cells.map((c, i) => `c${i}{gs:${c.gridSpan||1},rs:${c.rowSpan||1},cont:${!!c.isMergedContinue},v:${c.vMerge||''},lc:${c.logicalColIndex||0}}`).join(' | ');
                    console.warn(`[TableProcessor][DEBUG] Row ${rowIndex} cells(${cells.length}): ${details}`);
                } catch(_e) {}
                return false;
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
     * Parse cell borders from tcPr (table format)
     */
    parseCellBordersFromTcPr(tcPr) {
        const borders = {};
        const borderMappings = {
            'lnL': 'left',
            'lnR': 'right', 
            'lnT': 'top',
            'lnB': 'bottom'
        };
        
        for (const [xmlName, borderName] of Object.entries(borderMappings)) {
            let borderElement = tcPr.querySelector(xmlName);
            if (!borderElement) {borderElement = tcPr.querySelector(`a\\:${xmlName}`);}
            
            if (borderElement) {
                const wAttr = parseInt(borderElement.getAttribute('w')) || 0; // EMU
                const color = this.parseColor(borderElement);
                // FIXED: Use proper EMU to pixel conversion
                // 914400 EMU = 1 inch = 96 pixels at 96 DPI
                const widthPx = wAttr > 0 ? Math.max(1, Math.round(this.convertEMUToPixels(wAttr))) : 1;
                
                // Only create border object if color is not null AND width > 0
                if (color !== null && widthPx > 0) {
                    borders[borderName] = {
                        style: 'single',
                        width: widthPx,
                        color: color
                    };
                }
            }
        }
        
        return Object.keys(borders).length > 0 ? borders : null;
    }

    /**
     * Parse cell borders (legacy format)
     */
    parseCellBorders(bordersElement) {
        return this.parseTableBorders(bordersElement);
    }

    /**
     * Parse border properties
     */
    parseBorderProperties(borderElement) {
        // Extract color first
        const color = this.parseColor(borderElement);
        // Prefer EMU width if provided via attribute 'w'
        const wAttr = parseInt(borderElement.getAttribute('w')) || 0; // EMU
        let widthPx = 0;
        if (wAttr > 0) {
            // FIXED: Use proper EMU to pixel conversion
            // 914400 EMU = 1 inch = 96 pixels at 96 DPI
            widthPx = Math.max(1, Math.round(this.convertEMUToPixels(wAttr)));
        } else {
            // Fallback to 'sz' which is in 1/8 pt units for borders; convert to px
            const sz = parseInt(borderElement.getAttribute('sz')) || 0; // eights of a point
            if (sz > 0) {
                const points = sz / 8; // convert to points
                widthPx = Math.max(1, Math.round(points * (96 / 72)));
            }
        }

        return {
            style: borderElement.getAttribute('val') || 'single',
            width: widthPx, // pixels for drawing
            size: parseInt(borderElement.getAttribute('sz')) || 0,
            color: color,
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
        
        let fillColor = null;
        
        // Check if the element itself is a solidFill
        if (shadingElement.tagName === 'a:solidFill' || shadingElement.tagName === 'solidFill') {
            fillColor = this.parseColor(shadingElement);
        } else {
            // Look for solid fill within the shading element
            const solidFill = shadingElement.querySelector('solidFill') || shadingElement.querySelector('a\\:solidFill');
            if (solidFill) {
                fillColor = this.parseColor(solidFill);
            }
        }
        
        return {
            fill: fill,
            color: color ? this.parseColorFromHex(color) : null,
            fillColor: fillColor
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
            console.warn('Error extracting table structure:', error);
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
     * Extract cell text body
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
     * Extract cell content
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
        // This follows a generic content extraction pattern

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
     * Get table data by shape ID
     */
    getTableData(shapeId) {
        // This method should be implemented to work with the slide processor
        // It will need access to the slides data
        return null;
    }

    /**
     * Get all tables in the presentation
     */
    getAllTableData(slides) {
        const tables = [];
        
        for (let slideIndex = 0; slideIndex < slides.length; slideIndex++) {
            const slide = slides[slideIndex];
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
     * Parse text body from element (helper method) - Enhanced with vertical alignment
     */
    parseTextBodyFromElement(element) {
        let txBody = element.querySelector('txBody');
        if (!txBody) {txBody = element.querySelector('a\\:txBody');}
        
        if (!txBody) {return null;}
        
        const textBody = { paragraphs: [] };

        // Parse body properties (a:bodyPr) for wrap, margins, and vertical alignment
        try {
            let bodyPr = txBody.querySelector('bodyPr');
            if (!bodyPr) {bodyPr = txBody.querySelector('a\\:bodyPr');}
            if (bodyPr) {
                const props = { wrap: true };
                const wrapAttr = bodyPr.getAttribute('wrap');
                if (wrapAttr && wrapAttr.toLowerCase() === 'none') {
                    props.wrap = false;
                }
                
                // ENHANCED: Vertical alignment (anchor)
                // Values: 't' (top), 'ctr' (center), 'b' (bottom)
                const anchor = bodyPr.getAttribute('anchor');
                if (anchor) {
                    props.anchor = anchor;
                }
                
                // ENHANCED: Text rotation
                const rot = bodyPr.getAttribute('rot');
                if (rot) {
                    props.rotation = parseInt(rot) || 0; // In 60000ths of a degree
                }
                
                // Margins (EMU)
                const lIns = bodyPr.getAttribute('lIns');
                const rIns = bodyPr.getAttribute('rIns');
                const tIns = bodyPr.getAttribute('tIns');
                const bIns = bodyPr.getAttribute('bIns');
                if (lIns) {props.leftMargin = parseInt(lIns) || 0;}
                if (rIns) {props.rightMargin = parseInt(rIns) || 0;}
                if (tIns) {props.topMargin = parseInt(tIns) || 0;}
                if (bIns) {props.bottomMargin = parseInt(bIns) || 0;}
                textBody.bodyProperties = props;
            }
        } catch(_e) {}
        
        let paragraphs = txBody.querySelectorAll('p');
        if (paragraphs.length === 0) {
            paragraphs = txBody.querySelectorAll('a\\:p');
        }
        
        paragraphs.forEach(pElement => {
            const paragraph = { runs: [], properties: {} };
            
            // Parse paragraph properties (pPr) for alignment
            let pPr = pElement.querySelector('pPr');
            if (!pPr) {pPr = pElement.querySelector('a\\:pPr');}
            if (pPr) {
                const algnAttr = pPr.getAttribute('algn');
                if (algnAttr) {
                    paragraph.properties.alignment = algnAttr; // 'l', 'ctr', 'r', etc.
                }
            }
            
            let runs = pElement.querySelectorAll('r');
            if (runs.length === 0) {
                runs = pElement.querySelectorAll('a\\:r');
            }
            
            runs.forEach(rElement => {
                const run = { text: '', properties: {} };
                
                // Get text content
                let tElement = rElement.querySelector('t');
                if (!tElement) {tElement = rElement.querySelector('a\\:t');}
                if (tElement) {
                    run.text = tElement.textContent || '';
                }
                
                // Parse run properties - Enhanced with more text formatting options
                let rPr = rElement.querySelector('rPr');
                if (!rPr) {rPr = rElement.querySelector('a\\:rPr');}
                if (rPr) {
                    // Text color
                    const solidFill = rPr.querySelector('solidFill') || rPr.querySelector('a\\:solidFill');
                    if (solidFill) {
                        run.properties.color = this.parseColor(solidFill);
                    }
                    
                    // Font properties - sz is in half-points (e.g., sz="1400" = 14pt)
                    const sizeAttr = rPr.getAttribute('sz');
                    if (sizeAttr) {
                        run.properties.fontSize = parseInt(sizeAttr) / 100; // Convert half-points to points
                    } else {
                        run.properties.fontSize = 12; // Default
                    }
                    
                    // Bold
                    run.properties.bold = rPr.getAttribute('b') === '1';
                    
                    // Italic
                    run.properties.italic = rPr.getAttribute('i') === '1';
                    
                    // ENHANCED: Underline
                    const uAttr = rPr.getAttribute('u');
                    if (uAttr && uAttr !== 'none') {
                        run.properties.underline = true;
                    }
                    
                    // ENHANCED: Strikethrough
                    const strikeAttr = rPr.getAttribute('strike');
                    if (strikeAttr && strikeAttr !== 'noStrike') {
                        run.properties.strikethrough = true;
                    }
                    
                    // Font family - check multiple possible locations
                    let fontFamily = rPr.getAttribute('typeface') || rPr.getAttribute('fontFamily');
                    
                    // Also check for <a:latin> element which is the standard way in PPTX
                    if (!fontFamily) {
                        const latinElement = rPr.querySelector('latin') || rPr.querySelector('a\\:latin');
                        if (latinElement) {
                            fontFamily = latinElement.getAttribute('typeface');
                        }
                    }
                    
                    if (fontFamily) {
                        run.properties.fontFamily = fontFamily;
                    } else {
                        run.properties.fontFamily = 'Arial'; // Default
                    }
                    
                }
                
                paragraph.runs.push(run);
            });
            
            textBody.paragraphs.push(paragraph);
        });
        
        return textBody.paragraphs.length > 0 ? textBody : null;
    }

    /**
     * Parse color from element (helper method)
     */
    parseColor(element) {
        if (!element) {return null;}
        
        // FIXED: Check for noFill first - if present, return null (no color/no border)
        const noFill = element.querySelector('noFill') || element.querySelector('a\\:noFill');
        if (noFill) {
            return null;
        }
        
        // Look for solid fill with sRGB or theme scheme color
        const solidFill = element.querySelector('solidFill') || element.querySelector('a\\:solidFill');
        if (solidFill) {
            const srgbClr = solidFill.querySelector('srgbClr') || solidFill.querySelector('a\\:srgbClr');
            if (srgbClr) {
                const val = srgbClr.getAttribute('val');
                if (val) {
                    const color = this.parseColorFromHex(val);
                    if (color) {return color;}
                }
            }
            const schemeClr = solidFill.querySelector('schemeClr') || solidFill.querySelector('a\\:schemeClr');
            if (schemeClr) {
                const schemeVal = schemeClr.getAttribute('val');
                const color = this.resolveSchemeColor(schemeVal);
                if (color) {return color;}
            }
        }
        
        // Look for direct sRGB color
        const srgbClr = element.querySelector('srgbClr') || element.querySelector('a\\:srgbClr');
        if (srgbClr) {
            const val = srgbClr.getAttribute('val');
            if (val) {
                const color = this.parseColorFromHex(val);
                if (color) {return color;}
            }
        }
        
        // Look for direct schemeClr
        const schemeClr = element.querySelector('schemeClr') || element.querySelector('a\\:schemeClr');
        if (schemeClr) {
            const schemeVal = schemeClr.getAttribute('val');
            const color = this.resolveSchemeColor(schemeVal);
            if (color) {return color;}
        }
        
        // Look for color attribute directly
        const colorAttr = element.getAttribute('color');
        if (colorAttr) {
            const color = this.parseColorFromHex(colorAttr);
            if (color) {return color;}
        }
        
        // If unable to resolve color (e.g., unknown scheme), return null so callers can skip drawing
        return null;
    }

    /**
     * Map PresentationML schemeClr names to default RGB colors.
     * This is a reasonable fallback when theme colors aren't loaded.
     */
    resolveSchemeColor(schemeName) {
        if (!schemeName) {return null;}
        const mapping = {
            tx1: '#000000',
            tx2: '#FFFFFF',
            bg1: '#FFFFFF',
            bg2: '#EEECE1',
            accent1: '#5B9BD5',
            accent2: '#ED7D31',
            accent3: '#A5A5A5',
            accent4: '#FFC000',
            accent5: '#4472C4',
            accent6: '#70AD47'
        };
        const hex = mapping[schemeName];
        return hex ? this.parseColorFromHex(hex) : null;
    }

    /**
     * Parse color from hex (helper method)
     */
    parseColorFromHex(hexColor) {
        if (!hexColor || typeof hexColor !== 'string') {
            return null;
        }
        
        // Remove # if present
        hexColor = hexColor.replace('#', '');
        
        // Handle 6-digit hex
        if (hexColor.length === 6) {
            const r = parseInt(hexColor.substring(0, 2), 16);
            const g = parseInt(hexColor.substring(2, 4), 16);
            const b = parseInt(hexColor.substring(4, 6), 16);
            return { r, g, b, a: 255 };
        }
        
        // Handle 3-digit hex
        if (hexColor.length === 3) {
            const r = parseInt(hexColor.charAt(0) + hexColor.charAt(0), 16);
            const g = parseInt(hexColor.charAt(1) + hexColor.charAt(1), 16);
            const b = parseInt(hexColor.charAt(2) + hexColor.charAt(2), 16);
            return { r, g, b, a: 255 };
        }
        
        return null;
    }

    /**
     * Convert EMU (English Metric Units) to pixels
     * 1 EMU = 1/914400 inch, 1 inch = 96 pixels (at 96 DPI)
     */
    convertEMUToPixels(emuValue) {
        if (!emuValue || typeof emuValue !== 'number') {
            return 3.78; // Default ~1mm in pixels
        }
        // 914400 EMU = 1 inch = 96 pixels
        return (emuValue / 914400) * 96;
    }
    
    /**
     * Cleanup resources
     */
    cleanup() {
        this.tableCache.clear();
    }
    
    // ===== VALIDATION AND HELPER METHODS =====
    
    /**
     * Get validator instance with fallback
     */
    _getValidator() {
        return (typeof window !== 'undefined' && window.DataValidator) ? 
            window.DataValidator : class { validate() { return { valid: true, errors: [] }; } };
    }
    
    /**
     * Get error boundary class with fallback
     */
    _getErrorBoundary() {
        return (typeof window !== 'undefined' && window.ErrorBoundary) ? 
            window.ErrorBoundary : class { 
                constructor() {} 
                wrap(fn) { return fn; }
            };
    }
    
    /**
     * Validate configuration options
     */
    _validateConfig(config) {
        const schema = {
            type: 'object',
            properties: {
                maxTableSize: { type: 'number', minimum: 1, maximum: 10000 },
                maxCellTextLength: { type: 'number', minimum: 1, maximum: 100000 },
                enableCaching: { type: 'boolean' },
                enableValidation: { type: 'boolean' },
                renderTimeout: { type: 'number', minimum: 1000, maximum: 60000 }
            }
        };
        
        try {
            const validator = new (this._getValidator())({ throwOnError: false });
            const result = validator.validate(config, schema);
            if (!result.valid && this.logger) {
                this.logger.log("warn", this.constructor.name, 'TableProcessor', 'Invalid configuration:', result.errors);
            }
        } catch (error) {
            // Validation unavailable, use defaults
        }
        
        return config;
    }
    
    /**
     * Validate render input parameters
     */
    _validateRenderInput(graphics, shape, x, y, w, h, options) {
        const errors = [];
        
        // Graphics validation
        if (!graphics || typeof graphics !== 'object') {
            errors.push('Graphics object is required');
        } else {
            const requiredMethods = ['SaveGrState', 'RestoreGrState', 'b_color1', 'TableRect'];
            for (const method of requiredMethods) {
                if (typeof graphics[method] !== 'function') {
                    errors.push(`Graphics object missing required method: ${method}`);
                }
            }
        }
        
        // Shape validation
        if (!shape || typeof shape !== 'object') {
            errors.push('Shape object is required');
        }
        
        // Coordinate validation
        if (typeof x !== 'number' || !isFinite(x)) {
            errors.push('X coordinate must be a finite number');
        }
        if (typeof y !== 'number' || !isFinite(y)) {
            errors.push('Y coordinate must be a finite number');
        }
        if (typeof w !== 'number' || !isFinite(w) || w <= 0) {
            errors.push('Width must be a positive finite number');
        }
        if (typeof h !== 'number' || !isFinite(h) || h <= 0) {
            errors.push('Height must be a positive finite number');
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
     * Validate table XML structure
     */
    _validateTableXML(tableXml) {
        if (typeof tableXml !== 'string' || tableXml.trim().length === 0) {
            throw this._createError('Table XML must be a non-empty string', {
                name: 'ValidationError',
                code: 'INVALID_TABLE_XML'
            });
        }
        
        // Basic XML structure validation
        if (!tableXml.includes('<') || !tableXml.includes('>')) {
            throw this._createError('Table XML does not appear to be valid XML', {
                name: 'ValidationError',
                code: 'MALFORMED_TABLE_XML'
            });
        }
        
        // Check for required table elements
        const requiredElements = ['tbl', 'tr', 'tc'];
        const hasRequiredElements = requiredElements.some(element => 
            tableXml.includes(`<${element}`) || tableXml.includes(`<a:${element}`)
        );
        
        if (!hasRequiredElements) {
            throw this._createError('Table XML missing required table elements', {
                name: 'ValidationError',
                code: 'INCOMPLETE_TABLE_XML'
            });
        }
    }
    
    /**
     * Validate parsed table structure
     */
    _validateParsedTable(table) {
        if (!table) {
            throw this._createError('Parsed table is null or undefined', {
                name: 'ValidationError',
                code: 'NULL_PARSED_TABLE'
            });
        }
        
        // Check if table has required methods
        const requiredMethods = ['getRowCount', 'getColumnCount', 'getCell'];
        for (const method of requiredMethods) {
            if (typeof table[method] !== 'function') {
                throw this._createError(`Parsed table missing required method: ${method}`, {
                    name: 'ValidationError',
                    code: 'INVALID_TABLE_STRUCTURE'
                });
            }
        }
        
        // Validate table size constraints
        const rowCount = table.getRowCount();
        const colCount = table.getColumnCount();
        
        if (rowCount <= 0 || colCount <= 0) {
            throw this._createError(`Invalid table dimensions: ${rowCount}x${colCount}`, {
                name: 'ValidationError',
                code: 'INVALID_TABLE_DIMENSIONS'
            });
        }
        
        if (rowCount * colCount > this.config.maxTableSize) {
            throw this._createError(`Table size ${rowCount * colCount} exceeds maximum ${this.config.maxTableSize}`, {
                name: 'ValidationError',
                code: 'TABLE_TOO_LARGE'
            });
        }
    }
    
    /**
     * Validate table structure for rendering
     */
    _validateTableStructure(table) {
        try {
            const rowCount = table.getRowCount();
            const colCount = table.getColumnCount();
            
            // Check for reasonable table dimensions
            if (rowCount > 100 || colCount > 50) {
                if (this.logger) {
                    this.logger.log("warn", this.constructor.name, 'TableProcessor', `Large table detected: ${rowCount}x${colCount}`);
                }
            }
            
            // Validate cell structure
            let cellCount = 0;
            for (let row = 0; row < Math.min(rowCount, 10); row++) { // Sample first 10 rows
                for (let col = 0; col < colCount; col++) {
                    const cell = table.getCell(row, col);
                    if (cell) {
                        cellCount++;
                        
                        // Validate cell text length
                        if (cell.textBody && typeof cell.textBody === 'string') {
                            if (cell.textBody.length > this.config.maxCellTextLength) {
                                if (this.logger) {
                                    this.logger.log("warn", this.constructor.name, 'TableProcessor', 
                                        `Cell text length ${cell.textBody.length} exceeds maximum ${this.config.maxCellTextLength}`);
                                }
                            }
                        }
                    }
                }
            }
            
            if (cellCount === 0) {
                throw this._createError('Table has no valid cells', {
                    name: 'ValidationError',
                    code: 'EMPTY_TABLE'
                });
            }
            
        } catch (error) {
            throw this._enhanceError(error, 'table_structure_validation');
        }
    }
    
    /**
     * Safe wrapper for drawing table background and outer border
     */
    async _safeDrawTableBackgroundAndOuterBorder(graphics, table, x, y, w, h) {
        try {
            this.drawTableBackgroundAndOuterBorder(graphics, table, x, y, w, h);
        } catch (error) {
            if (this.logger) {
                this.logger.log("warn", this.constructor.name, 'TableProcessor', 'Error drawing table background/border:', error);
            }
            // Continue with rendering - this is not critical
        }
    }
    
    /**
     * Safe wrapper for drawing cell backgrounds
     */
    async _safeDrawCellsBackground(graphics, table, x, y, w, h) {
        try {
            this.drawCellsBackground(graphics, table, x, y, w, h);
        } catch (error) {
            if (this.logger) {
                this.logger.log("warn", this.constructor.name, 'TableProcessor', 'Error drawing cell backgrounds:', error);
            }
            // Continue with rendering - this is not critical
        }
    }
    
    /**
     * Safe wrapper for drawing cell content
     */
    async _safeDrawCellsContent(graphics, table, x, y, w, h) {
        try {
            this.drawCellsContent(graphics, table, x, y, w, h);
        } catch (error) {
            console.error('[TableProcessor] DETAILED ERROR drawing cell content:');
            console.error('[TableProcessor] Error message:', error.message);
            console.error('[TableProcessor] Error stack:', error.stack);
            console.error('[TableProcessor] Error object:', error);
            if (this.logger) {
                this.logger.logError(this.constructor.name, 'TableProcessor', 'Error drawing cell content:', error);
            }
            // Cell content is critical - rethrow error
            throw error;
        }
    }
    
    /**
     * Safe wrapper for drawing cell borders
     */
    async _safeDrawCellsBorders(graphics, table, x, y, w, h) {
        try {
            this.drawCellsBorders(graphics, table, x, y, w, h);
        } catch (error) {
            if (this.logger) {
                this.logger.log("warn", this.constructor.name, 'TableProcessor', 'Error drawing cell borders:', error);
            }
            // Continue with rendering - this is not critical
        }
    }
    
    /**
     * Handle validation errors
     */
    _handleValidationError(operation, errors) {
        const errorMessage = `Validation failed for ${operation}: ${errors.join(', ')}`;
        
        if (this.logger) {
            this.logger.logError(this.constructor.name, 'TableProcessor', errorMessage);
        }
        
        // Create validation error but don't throw - let caller handle
        const error = this._createError(errorMessage, {
            name: 'ValidationError',
            code: 'INPUT_VALIDATION_FAILED',
            details: { operation, errors }
        });
        
        return error;
    }
    
    /**
     * Update performance metrics
     */
    _updatePerformanceMetrics(startTime) {
        this.performanceMetrics.tablesProcessed++;
        
        const renderTime = performance.now() - startTime;
        const totalTime = this.performanceMetrics.averageRenderTime * (this.performanceMetrics.tablesProcessed - 1);
        this.performanceMetrics.averageRenderTime = (totalTime + renderTime) / this.performanceMetrics.tablesProcessed;
        
        if (this.logger && renderTime > 1000) { // Log slow renders
            this.logger.log("warn", this.constructor.name, 'TableProcessor', `Slow table render: ${Math.round(renderTime)}ms`);
        }
    }
    
    /**
     * Create enhanced error with context
     */
    _createError(message, options = {}) {
        const error = new Error(message);
        error.name = options.name || 'TableProcessingError';
        error.code = options.code || null;
        error.context = 'TableProcessor';
        error.timestamp = new Date().toISOString();
        error.recoverable = options.recoverable !== false;
        
        if (options.details) {
            error.details = options.details;
        }
        
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
        enhanced.tableProcessorContext = context;
        enhanced.timestamp = new Date().toISOString();
        enhanced.enhanced = true;
        
        return enhanced;
    }
    
    /**
     * Get performance metrics
     */
    getPerformanceMetrics() {
        return {
            ...this.performanceMetrics,
            cacheHitRatio: this.performanceMetrics.cacheHits / 
                (this.performanceMetrics.cacheHits + this.performanceMetrics.cacheMisses) || 0
        };
    }
    
    /**
     * Reset performance metrics
     */
    resetPerformanceMetrics() {
        this.performanceMetrics = {
            tablesProcessed: 0,
            cacheHits: 0,
            cacheMisses: 0,
            averageRenderTime: 0
        };
    }
}

// Export the class
if (typeof globalThis !== 'undefined') {
    globalThis.TableProcessor = TableProcessor;
}

// ES Module export
// export { TableProcessor }; 
