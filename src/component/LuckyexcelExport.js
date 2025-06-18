import LuckyExcel from 'luckyexcel';

class LuckyexcelExport {
    
    // Export using Luckyexcel library (designed specifically for Luckysheet)
    static exportWithLuckyexcel(filename = 'luckysheet-export') {
        try {
            // Get the complete Luckysheet data
            const luckysheetData = window.luckysheet.toJson();
            
            if (!luckysheetData || !luckysheetData.sheets || luckysheetData.sheets.length === 0) {
                alert('No data to export');
                return;
            }

            // Note: Luckyexcel is primarily for import (Excel to Luckysheet)
            // For export (Luckysheet to Excel), we need to use the reverse process
            // This is a conceptual implementation - you may need to adapt based on your needs
            
            console.log('Luckysheet data to export:', luckysheetData);
            
            // Since Luckyexcel is mainly for import, we'll use our custom method
            // with the data structure that Luckyexcel expects
            this.exportLuckysheetFormat(luckysheetData, filename);
            
        } catch (error) {
            console.error('Export with Luckyexcel failed:', error);
            alert('Export failed. Please try again.');
        }
    }

    // Custom export method that handles Luckysheet format
    static exportLuckysheetFormat(data, filename) {
        // This method processes the Luckysheet JSON format and converts it
        // to a downloadable format
        
        try {
            // Create a comprehensive export
            const exportData = {
                info: data.info || { name: filename, creator: 'Luckysheet' },
                sheets: data.sheets || []
            };

            // Convert to JSON and download
            const jsonString = JSON.stringify(exportData, null, 2);
            const blob = new Blob([jsonString], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = `${filename}.json`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
            
            console.log('Exported Luckysheet format successfully');
            
        } catch (error) {
            console.error('Export failed:', error);
            throw error;
        }
    }

    // Method to prepare data for potential server-side export
    static prepareDataForServerExport() {
        try {
            const data = window.luckysheet.toJson();
            return {
                success: true,
                data: data,
                message: 'Data prepared for export'
            };
        } catch (error) {
            return {
                success: false,
                data: null,
                message: error.message
            };
        }
    }

    // Method to export raw cell data
    static exportRawData(filename = 'raw-data') {
        try {
            const allSheets = window.luckysheet.getAllSheets();
            
            if (!allSheets || allSheets.length === 0) {
                alert('No data to export');
                return;
            }

            const exportData = allSheets.map(sheet => ({
                name: sheet.name,
                data: sheet.data,
                config: sheet.config,
                index: sheet.index
            }));

            const jsonString = JSON.stringify(exportData, null, 2);
            const blob = new Blob([jsonString], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = `${filename}.json`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
            
        } catch (error) {
            console.error('Export raw data failed:', error);
            alert('Export failed. Please try again.');
        }
    }
}

export default LuckyexcelExport; 