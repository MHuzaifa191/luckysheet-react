import React from 'react';
import axios from 'axios';
import ExcelExport from './ExcelExport';
import LuckyexcelExport from './LuckyexcelExport';

class Luckysheet extends React.Component {

    constructor(props) {
        super(props);
        this.state = {
            urlParams: {}
        };
    }

    componentDidMount() {
        const urlParams = this.getUrlParameters();
        this.setState({ urlParams }, async () => {
            // Load existing spreadsheet after URL params are set
            await this.loadExistingSpreadsheet();
            
        const luckysheet = window.luckysheet;
        luckysheet.create({
            container: "luckysheet",
                plugins:['chart'],
                showinfobar: false,
                title: '',
            });
        });
    }

    getUrlParameters = () => {
        const params = new URLSearchParams(window.location.search);
        const urlParams = {};
        
        for (let [key, value] of params.entries()) {
            urlParams[key] = value;
        }
        
        return urlParams;
    }

    // Method to load existing spreadsheet
    loadExistingSpreadsheet = async () => {
        try {
            const activityId = this.state.urlParams.activityId;
            const userId = this.state.urlParams.user;

            if (!activityId || !userId) {
                console.log('No activityId or userId provided, starting with empty spreadsheet');
                return;
            }

            console.log('Loading existing spreadsheet for:', { activityId, userId });
            
            // Directly fetch the file content from your backend
            const response = await axios.get(`http://127.0.0.1:8000/activity/spreadsheets/save`, {
                params: {
                    activityId: activityId,
                    userId: userId
                },
                responseType: 'arraybuffer' // Important: get binary data
            });

            console.log('File loaded successfully, size:', response.data.byteLength, 'bytes');
            
            // Ensure XLSX library is loaded before processing
            await this.ensureXLSXLoaded();
            
            // Convert Excel to Luckysheet format
            const workbook = window.XLSX.read(response.data, { type: 'array' });
            const luckysheetData = this.convertExcelToLuckysheet(workbook);

            console.log('Converted to Luckysheet format:', luckysheetData);
            
            // Update Luckysheet with the loaded data
            if (window.luckysheet && luckysheetData.length > 0) {
                window.luckysheet.destroy();
                setTimeout(() => {
                    window.luckysheet.create({
                        container: "luckysheet",
                        plugins: ['chart'],
                        showinfobar: false,
                        title: '',
                        data: luckysheetData
                    });
                }, 100);
            } else {
                console.log('No valid data found, starting with empty spreadsheet');
            }
            
        } catch (error) {
            if (error.response) {
                console.log('Backend response:', error.response.status, error.response.data);
                if (error.response.status === 404) {
                    console.log('No existing spreadsheet found, starting with empty sheet');
                } else {
                    console.error('Error loading spreadsheet:', error.response.data);
                }
            } else {
                console.log('Network error or no existing spreadsheet:', error.message);
            }
            // Continue with empty spreadsheet if none exists or error occurs
        }
    }

    // Helper method to ensure XLSX library is loaded
    ensureXLSXLoaded = () => {
        return new Promise((resolve, reject) => {
            if (typeof window.XLSX !== 'undefined') {
                resolve();
                return;
            }

            console.log('Loading XLSX library...');
            const script = document.createElement('script');
            script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
            script.onload = () => {
                console.log('XLSX library loaded successfully');
                resolve();
            };
            script.onerror = () => {
                reject(new Error('Failed to load XLSX library'));
            };
            document.head.appendChild(script);
        });
    }

    // Helper method to convert Excel workbook to Luckysheet format
    convertExcelToLuckysheet = (workbook) => {
        const luckysheetData = [];
        
        workbook.SheetNames.forEach((sheetName, index) => {
            const worksheet = workbook.Sheets[sheetName];
            const range = window.XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
            
            const celldata = [];
            
            for (let row = range.s.r; row <= range.e.r; row++) {
                for (let col = range.s.c; col <= range.e.c; col++) {
                    const cellAddress = window.XLSX.utils.encode_cell({ r: row, c: col });
                    const cell = worksheet[cellAddress];
                    
                    if (cell) {
                        celldata.push({
                            r: row,
                            c: col,
                            v: {
                                v: cell.v,
                                t: cell.t
                            }
                        });
                    }
                }
            }
            
            luckysheetData.push({
                name: sheetName,
                index: index,
                status: index === 0 ? 1 : 0, // First sheet is active
                order: index,
                celldata: celldata
            });
        });
        
        return luckysheetData;
    }

    // Method to export all sheets to Excel
    exportToExcel = () => {
        ExcelExport.exportToExcel('my-spreadsheet');
    }

    // Method to export current sheet only
    exportCurrentSheet = () => {
        ExcelExport.exportCurrentSheet('current-sheet');
    }

    // Method to export selected range
    exportSelectedRange = () => {
        ExcelExport.exportRange('selected-range');
    }

    // Method to export using Luckyexcel format
    exportLuckyexcelFormat = () => {
        LuckyexcelExport.exportWithLuckyexcel('luckysheet-export');
    }

    // Method to export raw data
    exportRawData = () => {
        LuckyexcelExport.exportRawData('raw-data');
    }

    // Method to save as JSON (Luckysheet format)
    saveAsJson = () => {
        const data = window.luckysheet.toJson();
        const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = 'luckysheet-data.json';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }

    // Updated save method to handle better error messages
    saveToBackend = async () => {
        try {
            console.log('URL Params:', this.state.urlParams);
            const activityId = this.state.urlParams.activityId;
            const userId = this.state.urlParams.user;
            const excelBlob = await this.generateExcelBlob();

            if (!activityId || !userId) {
                alert('Activity ID and User ID are required');
                return;
            }
            
            const formData = new FormData();
            formData.append('activityId', activityId);
            formData.append('userId', userId);
            formData.append('excelFile', excelBlob, `spreadsheet.xlsx`);

            // Send POST request to backend using axios
            const response = await axios.post('http://127.0.0.1:8000/activity/spreadsheets/save', formData, {
                headers: {
                    'Content-Type': 'multipart/form-data',
                }
            });

            alert('Spreadsheet saved successfully!');
            console.log('Save response:', response.data);
        } catch (error) {
            console.error('Error saving spreadsheet:', error);
            if (error.response) {
                alert(`Failed to save: ${error.response.data.error || error.response.statusText}`);
            } else {
                alert('Failed to save spreadsheet. Please try again.');
            }
        }
    }

    // Helper method to generate Excel blob
    generateExcelBlob = async () => {
        return new Promise((resolve, reject) => {
            try {
                // Import XLSX library dynamically if not already available
                if (typeof window.XLSX === 'undefined') {
                    const script = document.createElement('script');
                    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
                    script.onload = () => this.createExcelBlob(resolve, reject);
                    script.onerror = () => reject(new Error('Failed to load XLSX library'));
                    document.head.appendChild(script);
                } else {
                    this.createExcelBlob(resolve, reject);
                }
            } catch (error) {
                reject(error);
            }
        });
    }

    // Create Excel blob from Luckysheet data
    createExcelBlob = (resolve, reject) => {
        try {
            const luckysheetData = window.luckysheet.getAllSheets();
            const XLSX = window.XLSX; // Get XLSX from window object
            const workbook = XLSX.utils.book_new();

            luckysheetData.forEach((sheet, index) => {
                const sheetName = sheet.name || `Sheet${index + 1}`;
                const cellData = sheet.celldata || [];
                
                // Convert Luckysheet cell data to worksheet format
                const worksheet = {};
                let range = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };

                cellData.forEach(cell => {
                    const row = cell.r;
                    const col = cell.c;
                    const cellValue = cell.v && cell.v.v !== undefined ? cell.v.v : '';
                    
                    const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                    worksheet[cellAddress] = { v: cellValue, t: typeof cellValue === 'number' ? 'n' : 's' };
                    
                    // Update range
                    if (range.s.r > row) range.s.r = row;
                    if (range.s.c > col) range.s.c = col;
                    if (range.e.r < row) range.e.r = row;
                    if (range.e.c < col) range.e.c = col;
                });

                worksheet['!ref'] = XLSX.utils.encode_range(range);
                XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
            });

            // Generate Excel file as blob
            const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            resolve(blob);
        } catch (error) {
            reject(error);
        }
    }

    render() {
        const luckyCss = {
            margin: '0px',
            padding: '0px',
            position: 'absolute',
            width: '100%',
            height: 'calc(100% - 50px)', // Adjust height to account for navbar
            left: '0px',
            top: '50px' // Start below navbar
        }

        const navbarStyle = {
            position: 'absolute',
            top: '0px',
            left: '0px',
            width: '100%',
            height: '50px',
            backgroundColor: '#2c3e50',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between', // This pushes content to opposite ends
            padding: '0 15px',
            boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
            zIndex: 1001
        };

        const brandStyle = {
            color: 'white',
            fontSize: '18px',
            fontWeight: 'bold'
        };

        const buttonContainerStyle = {
            display: 'flex',
            alignItems: 'center',
            gap: '10px'
        };

        const exportDropdownStyle = {
            position: 'relative',
            display: 'inline-block'
        };

        const dropdownButtonStyle = {
            backgroundColor: '#3498db',
            color: 'white',
            padding: '8px 16px',
            fontSize: '14px',
            border: 'none',
            borderRadius: '4px',
            cursor: 'pointer',
            display: 'flex',
            alignItems: 'center',
            gap: '5px'
        };

        const saveButtonStyle = {
            backgroundColor: '#27ae60',
            color: 'white',
            padding: '8px 16px',
            fontSize: '14px',
            border: 'none',
            borderRadius: '4px',
            cursor: 'pointer',
            display: 'flex',
            alignItems: 'center',
            gap: '5px'
        };

        const dropdownContentStyle = {
            display: 'none',
            position: 'absolute',
            backgroundColor: 'white',
            minWidth: '200px',
            boxShadow: '0px 8px 16px 0px rgba(0,0,0,0.2)',
            zIndex: 1,
            borderRadius: '4px',
            overflow: 'hidden',
            top: '100%',
            left: '0'
        };

        const dropdownItemStyle = {
            color: '#333',
            padding: '12px 16px',
            textDecoration: 'none',
            display: 'block',
            cursor: 'pointer',
            border: 'none',
            backgroundColor: 'transparent',
            width: '100%',
            textAlign: 'left',
            fontSize: '14px'
        };

        return (
            <div>
                {/* Navbar */}
                <nav style={navbarStyle}>

                    
                    {/* Button Container - moved to the right */}
                    <div style={buttonContainerStyle}>
                        {/* Export Dropdown */}
                        <div 
                            style={exportDropdownStyle}
                            onMouseEnter={(e) => {
                                const dropdown = e.currentTarget.querySelector('.dropdown-content');
                                if (dropdown) dropdown.style.display = 'block';
                            }}
                            onMouseLeave={(e) => {
                                const dropdown = e.currentTarget.querySelector('.dropdown-content');
                                if (dropdown) dropdown.style.display = 'none';
                            }}
                        >
                            <button style={dropdownButtonStyle}>
                                📤 Export
                                <span style={{ fontSize: '10px' }}>▼</span>
                            </button>
                            
                            <div className="dropdown-content" style={dropdownContentStyle}>
                                <button 
                                    onClick={this.exportToExcel} 
                                    style={dropdownItemStyle}
                                    onMouseEnter={(e) => e.target.style.backgroundColor = '#f1f1f1'}
                                    onMouseLeave={(e) => e.target.style.backgroundColor = 'transparent'}
                                >
                                    📊 Export All Sheets (Excel)
                                </button>
                                
                                <button 
                                    onClick={this.exportCurrentSheet} 
                                    style={dropdownItemStyle}
                                    onMouseEnter={(e) => e.target.style.backgroundColor = '#f1f1f1'}
                                    onMouseLeave={(e) => e.target.style.backgroundColor = 'transparent'}
                                >
                                    📄 Export Current Sheet
                                </button>
                                
                                <button 
                                    onClick={this.exportSelectedRange} 
                                    style={dropdownItemStyle}
                                    onMouseEnter={(e) => e.target.style.backgroundColor = '#f1f1f1'}
                                    onMouseLeave={(e) => e.target.style.backgroundColor = 'transparent'}
                                >
                                    🔲 Export Selected Range
                                </button>
                                
                                <hr style={{ margin: '5px 0', border: 'none', borderTop: '1px solid #eee' }} />
                                
                                <button 
                                    onClick={this.exportLuckyexcelFormat} 
                                    style={dropdownItemStyle}
                                    onMouseEnter={(e) => e.target.style.backgroundColor = '#f1f1f1'}
                                    onMouseLeave={(e) => e.target.style.backgroundColor = 'transparent'}
                                >
                                    📦 Export Luckyexcel Format
                                </button>
                                
                                <button 
                                    onClick={this.exportRawData} 
                                    style={dropdownItemStyle}
                                    onMouseEnter={(e) => e.target.style.backgroundColor = '#f1f1f1'}
                                    onMouseLeave={(e) => e.target.style.backgroundColor = 'transparent'}
                                >
                                    🔧 Export Raw Data (JSON)
                                </button>
                                
                                <button 
                                    onClick={this.saveAsJson} 
                                    style={dropdownItemStyle}
                                    onMouseEnter={(e) => e.target.style.backgroundColor = '#f1f1f1'}
                                    onMouseLeave={(e) => e.target.style.backgroundColor = 'transparent'}
                                >
                                    💾 Save as JSON
                                </button>
                            </div>
                        </div>
                        
                        {/* Save Button */}
                        <button 
                            onClick={this.saveToBackend}
                            style={saveButtonStyle}
                            onMouseEnter={(e) => e.target.style.backgroundColor = '#219a52'}
                            onMouseLeave={(e) => e.target.style.backgroundColor = '#27ae60'}
                        >
                            💾 Save
                        </button>
                    </div>
                </nav>
                
                {/* Luckysheet Container */}
            <div
            id="luckysheet"
            style={luckyCss}
            ></div>
            </div>
        )
    }
}

export default Luckysheet