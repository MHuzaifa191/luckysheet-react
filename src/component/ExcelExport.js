import * as XLSX from 'xlsx';

class ExcelExport {
    
    // Convert Luckysheet data to Excel format
    static exportToExcel(filename = 'luckysheet-export') {
        try {
            // Get all sheet data from Luckysheet
            const allSheets = window.luckysheet.getAllSheets();
            
            if (!allSheets || allSheets.length === 0) {
                alert('No data to export');
                return;
            }

            // Create new workbook
            const workbook = XLSX.utils.book_new();

            // Process each sheet
            allSheets.forEach((sheet, index) => {
                const sheetName = sheet.name || `Sheet${index + 1}`;
                const worksheet = this.convertLuckysheetToWorksheet(sheet);
                XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
            });

            // Save the file
            XLSX.writeFile(workbook, `${filename}.xlsx`);
            
        } catch (error) {
            console.error('Export failed:', error);
            alert('Export failed. Please try again.');
        }
    }

    // Convert Luckysheet sheet data to Excel worksheet
    static convertLuckysheetToWorksheet(sheet) {
        const data = sheet.data || [];
        const config = sheet.config || {};
        
        // Convert cell data to 2D array
        const excelData = [];
        const maxRow = data.length;
        const maxCol = data.reduce((max, row) => Math.max(max, row ? row.length : 0), 0);

        for (let row = 0; row < maxRow; row++) {
            const rowData = [];
            for (let col = 0; col < maxCol; col++) {
                const cell = data[row] && data[row][col];
                if (cell) {
                    // Handle different cell types
                    let value = cell.v;
                    
                    // Handle dates
                    if (cell.ct && cell.ct.t === 'd') {
                        value = new Date(cell.m.replace(/-/g, '/'));
                    }
                    // Handle percentages
                    else if (cell.ct && cell.ct.t === 'n' && cell.m && cell.m.includes('%')) {
                        value = parseFloat(cell.v) / 100;
                    }
                    // Use display value for other types
                    else {
                        value = cell.m || cell.v || '';
                    }
                    
                    rowData.push(value);
                } else {
                    rowData.push('');
                }
            }
            excelData.push(rowData);
        }

        // Create worksheet from array
        const worksheet = XLSX.utils.aoa_to_sheet(excelData);

        // Apply merged cells if available
        if (config.merge && Object.keys(config.merge).length > 0) {
            worksheet['!merges'] = this.convertMergeConfig(config.merge);
        }

        // Set column widths if available
        if (config.columnlen && Object.keys(config.columnlen).length > 0) {
            worksheet['!cols'] = this.convertColumnWidths(config.columnlen, maxCol);
        }

        return worksheet;
    }

    // Convert Luckysheet merge config to Excel format
    static convertMergeConfig(mergeConfig) {
        const merges = [];
        
        for (const key in mergeConfig) {
            const merge = mergeConfig[key];
            const startRow = merge.r;
            const startCol = merge.c;
            const endRow = merge.r + (merge.rs || 1) - 1;
            const endCol = merge.c + (merge.cs || 1) - 1;
            
            merges.push({
                s: { r: startRow, c: startCol },
                e: { r: endRow, c: endCol }
            });
        }
        
        return merges;
    }

    // Convert column widths
    static convertColumnWidths(columnlen, maxCol) {
        const cols = [];
        
        for (let col = 0; col < maxCol; col++) {
            const width = columnlen[col];
            if (width) {
                cols.push({ wch: Math.round(width / 7) }); // Convert px to character width
            } else {
                cols.push({ wch: 10 }); // Default width
            }
        }
        
        return cols;
    }

    // Export current sheet only
    static exportCurrentSheet(filename = 'current-sheet') {
        try {
            const sheetData = window.luckysheet.getSheetData();
            const config = window.luckysheet.getConfig();
            
            if (!sheetData || sheetData.length === 0) {
                alert('No data to export');
                return;
            }

            const sheet = {
                data: sheetData,
                config: config,
                name: 'Sheet1'
            };

            const workbook = XLSX.utils.book_new();
            const worksheet = this.convertLuckysheetToWorksheet(sheet);
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
            
            XLSX.writeFile(workbook, `${filename}.xlsx`);
            
        } catch (error) {
            console.error('Export failed:', error);
            alert('Export failed. Please try again.');
        }
    }

    // Export selected range
    static exportRange(filename = 'selected-range') {
        try {
            const range = window.luckysheet.getRange();
            
            if (!range || range.length === 0) {
                alert('No range selected');
                return;
            }

            const rangeData = window.luckysheet.getRangeValue();
            
            if (!rangeData || rangeData.length === 0) {
                alert('No data in selected range');
                return;
            }

            // Convert range data to simple 2D array
            const excelData = rangeData.map(row => 
                row.map(cell => cell ? (cell.m || cell.v || '') : '')
            );

            const workbook = XLSX.utils.book_new();
            const worksheet = XLSX.utils.aoa_to_sheet(excelData);
            XLSX.utils.book_append_sheet(workbook, worksheet, 'SelectedRange');
            
            XLSX.writeFile(workbook, `${filename}.xlsx`);
            
        } catch (error) {
            console.error('Export failed:', error);
            alert('Export failed. Please try again.');
        }
    }
}

export default ExcelExport; 