const XLSX = require('xlsx');
const _ = require('lodash');
const fs = require('fs');
const path = require('path');

function mergeExcelColumns(inputFile) {
    try {
        // Create unique filename for output using timestamp
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const outputFile = `merged_calls_${timestamp}.xlsx`;
        
        console.log(`Reading file: ${inputFile}`);
        const fileData = fs.readFileSync(inputFile);
        const workbook = XLSX.read(fileData, {
            cellStyles: true,
            cellFormulas: true,
            cellDates: true,
            cellNF: true,
            sheetStubs: true
        });
        
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const data = XLSX.utils.sheet_to_json(worksheet);
        
        const groupedData = _.groupBy(data, row => row['חפ']);
        
        const mergedData = _.map(groupedData, (group, companyId) => {
            // Improved phone number cleaning function
            const cleanPhoneNumbers = (phoneString) => {
                // First split by commas to separate multiple numbers
                return phoneString.toString()
                    .split(',')
                    .map(number => {
                        // Clean each individual number
                        return number
                            .trim()  // Remove leading/trailing spaces
                            .replace(/\s+/g, '')  // Remove spaces between digits
                            .replace(/^0*(\d{2})-*\s*(\d+)$/, '$1-$2')  // Format area code numbers
                            .replace(/^0*(\d+)$/, '$1');  // Handle numbers without area code
                    })
                    .filter(number => number.length > 0);  // Remove empty strings
            };
            
            // Process all phone numbers for this company
            const phoneNumbers = group
                .map(row => row['טלפון בעלים'])
                .flatMap(phones => cleanPhoneNumbers(phones));
            
            // Create merged entry with unique, cleaned numbers
            return {
                'חפ': companyId,
                'טלפון בעלים': _.uniq(phoneNumbers).join(', ')
            };
        });
        
        // Create and save the new workbook
        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.json_to_sheet(mergedData);
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Merged Data');
        XLSX.writeFile(newWorkbook, outputFile);
        
        // Print processing summary
        console.log('\nProcessing Summary:');
        console.log(`Original file: ${inputFile}`);
        console.log(`New merged file: ${outputFile}`);
        console.log(`Original rows: ${data.length}`);
        console.log(`Merged rows: ${mergedData.length}`);
        
        // Show details of merged entries
        const duplicates = Object.entries(groupedData)
            .filter(([_, group]) => group.length > 1);
            
        if (duplicates.length > 0) {
            console.log('\nMerged Entries:');
            duplicates.forEach(([companyId, entries]) => {
                console.log(`\nCompany ID (חפ): ${companyId}`);
                console.log('Original entries:');
                entries.forEach(entry => {
                    console.log(`  - ${entry['טלפון בעלים']}`);
                });
                
                const mergedEntry = mergedData.find(row => row['חפ'].toString() === companyId);
                console.log('Merged into:');
                console.log(`  → ${mergedEntry['טלפון בעלים']}`);
            });
        }
        
        console.log('\nProcessing completed successfully!');
        
    } catch (error) {
        console.error('Error processing Excel file:', error);
        if (error.code === 'ENOENT') {
            console.error('File not found. Please check if the file exists in the correct location.');
        }
    }
}

// Get input file name from command line or use default
const inputFile = process.argv[2] || 'calls.xlsx';

// Execute the function
mergeExcelColumns(inputFile);