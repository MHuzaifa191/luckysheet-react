import React from 'react';
import ExcelExport from './ExcelExport';
import LuckyexcelExport from './LuckyexcelExport';

class ExportDemo extends React.Component {

    // Method to show all available export options
    showExportOptions = () => {
        const options = [
            'Export All Sheets to Excel (.xlsx)',
            'Export Current Sheet to Excel (.xlsx)', 
            'Export Selected Range to Excel (.xlsx)',
            'Export to CSV (.csv)',
            'Export to JSON (Luckysheet format)',
            'Export Raw Data (JSON)'
        ];
        
        console.log('Available export options:', options);
        alert('Check console for available export options');
    }

    // Advanced Excel export with custom filename
    exportWithCustomName = () => {
        const filename = prompt('Enter filename (without extension):');
        if (filename) {
            ExcelExport.exportToExcel(filename);
        }
    }

    // Export with date timestamp
    exportWithTimestamp = () => {
        const now = new Date();
        const timestamp = now.toISOString().slice(0, 19).replace(/[:-]/g, '');
        const filename = `luckysheet_export_${timestamp}`;
        ExcelExport.exportToExcel(filename);
    }

    // Method to get export data without downloading (for API calls)
    getExportData = () => {
        try {
            const data = window.luckysheet.toJson();
            console.log('Export data:', data);
            
            // You can send this data to your server
            // Example:
            // fetch('/api/save-spreadsheet', {
            //     method: 'POST',
            //     headers: { 'Content-Type': 'application/json' },
            //     body: JSON.stringify(data)
            // });
            
            alert('Export data logged to console. Check developer tools.');
        } catch (error) {
            console.error('Failed to get export data:', error);
            alert('Failed to get export data');
        }
    }

    // Method to export specific sheet by index
    exportSpecificSheet = () => {
        const sheetIndex = prompt('Enter sheet index (0-based):');
        if (sheetIndex !== null) {
            try {
                const index = parseInt(sheetIndex);
                const allSheets = window.luckysheet.getAllSheets();
                
                if (allSheets && allSheets[index]) {
                    const sheet = allSheets[index];
                    console.log('Exporting sheet:', sheet.name);
                    
                    // Create a temporary workbook with just this sheet
                    const XLSX = require('xlsx');
                    const workbook = XLSX.utils.book_new();
                    const worksheet = ExcelExport.convertLuckysheetToWorksheet(sheet);
                    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name);
                    XLSX.writeFile(workbook, `${sheet.name}.xlsx`);
                } else {
                    alert('Invalid sheet index');
                }
            } catch (error) {
                console.error('Export failed:', error);
                alert('Export failed');
            }
        }
    }

    render() {
        const buttonStyle = {
            margin: '5px',
            padding: '10px 15px',
            fontSize: '14px',
            backgroundColor: '#2196F3',
            color: 'white',
            border: 'none',
            borderRadius: '5px',
            cursor: 'pointer',
            minWidth: '200px'
        };

        const containerStyle = {
            padding: '20px',
            backgroundColor: '#f5f5f5',
            borderRadius: '10px',
            margin: '20px',
            boxShadow: '0 4px 8px rgba(0,0,0,0.1)'
        };

        return (
            <div style={containerStyle}>
                <h3>📊 Luckysheet Export Options</h3>
                
                <div style={{ marginBottom: '20px' }}>
                    <h4>📋 Basic Export Methods:</h4>
                    <button onClick={() => ExcelExport.exportToExcel()} style={buttonStyle}>
                        📊 Export All Sheets (Excel)
                    </button>
                    <button onClick={() => ExcelExport.exportCurrentSheet()} style={buttonStyle}>
                        📄 Export Current Sheet
                    </button>
                    <button onClick={() => ExcelExport.exportRange()} style={buttonStyle}>
                        🔲 Export Selected Range
                    </button>
                </div>

                <div style={{ marginBottom: '20px' }}>
                    <h4>⚙️ Advanced Export Methods:</h4>
                    <button onClick={this.exportWithCustomName} style={buttonStyle}>
                        ✏️ Export with Custom Name
                    </button>
                    <button onClick={this.exportWithTimestamp} style={buttonStyle}>
                        🕒 Export with Timestamp
                    </button>
                    <button onClick={this.exportSpecificSheet} style={buttonStyle}>
                        📋 Export Specific Sheet
                    </button>
                </div>

                <div style={{ marginBottom: '20px' }}>
                    <h4>💾 Data Export Methods:</h4>
                    <button onClick={() => LuckyexcelExport.exportWithLuckyexcel()} style={buttonStyle}>
                        📦 Export Luckysheet Format
                    </button>
                    <button onClick={() => LuckyexcelExport.exportRawData()} style={buttonStyle}>
                        🔧 Export Raw Data
                    </button>
                    <button onClick={this.getExportData} style={buttonStyle}>
                        📡 Get Data for API
                    </button>
                </div>

                <div>
                    <h4>ℹ️ Information:</h4>
                    <button onClick={this.showExportOptions} style={buttonStyle}>
                        📝 Show All Options
                    </button>
                </div>

                <div style={{ marginTop: '20px', padding: '10px', backgroundColor: '#e3f2fd', borderRadius: '5px' }}>
                    <h4>💡 Usage Tips:</h4>
                    <ul style={{ fontSize: '12px', color: '#666' }}>
                        <li>Excel export preserves cell formatting, merged cells, and formulas</li>
                        <li>CSV export is simpler but loses formatting</li>
                        <li>JSON export preserves all Luckysheet-specific data</li>
                        <li>Selected range export only works when cells are selected</li>
                        <li>Custom name export allows you to specify the filename</li>
                    </ul>
                </div>
            </div>
        );
    }
}

export default ExportDemo; 