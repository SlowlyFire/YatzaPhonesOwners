const XLSX = require('xlsx');
const fs = require('fs');

function createSpecialFormatExcel(inputFile) {
    try {
        // Read the input file
        const fileData = fs.readFileSync(inputFile);
        const workbook = XLSX.read(fileData, {
            cellStyles: true,
            cellFormulas: true,
            cellDates: true,
            cellNF: true,
            sheetStubs: true
        });
        
        // Get the data
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(worksheet);
        
        // Filter for special format entries
        const specialFormatEntries = data
            .filter(row => {
                const phoneString = row['טלפון בעלים'].toString();
                return /\d+\s+\d+/.test(phoneString) || 
                       /\b\d{2,3}\s+\d{6,7}\b/.test(phoneString) ||
                       phoneString.includes(';') ||
                       phoneString.includes('tel:') ||
                       /\d+\s+\d{2,3}\s+\d+/.test(phoneString);
            })
            .sort((a, b) => a['חפ'] - b['חפ']);
        
        // Create new workbook
        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.json_to_sheet(specialFormatEntries);
        
        // Add column widths for better readability
        const colWidths = [
            { wch: 15 },  // Column A width
            { wch: 50 }   // Column B width
        ];
        newWorksheet['!cols'] = colWidths;
        
        // Add the worksheet to the workbook
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Special Format Entries');
        
        // Save the file with timestamp
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const outputFile = `special_format_entries_${timestamp}.xlsx`;
        XLSX.writeFile(newWorkbook, outputFile);
        
        console.log(`Created Excel file: ${outputFile}`);
        console.log(`Total entries: ${specialFormatEntries.length}`);
        
    } catch (error) {
        console.error('Error creating Excel file:', error);
    }
}

// Use the default filename
const inputFile = 'calls.xlsx';
createSpecialFormatExcel(inputFile);