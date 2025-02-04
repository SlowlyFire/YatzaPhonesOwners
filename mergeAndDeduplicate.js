const XLSX = require('xlsx');
const fs = require('fs');

function mergeAndDeduplicate(inputFiles) {
    try {
        console.log('Starting merge process...\n');
        
        // This array will hold all rows from all input files
        let allRows = [];
        
        // Process each input file
        inputFiles.forEach((file, index) => {
            console.log(`Reading file ${index + 1}: ${file}`);
            
            // Read the Excel file
            const fileData = fs.readFileSync(file);
            const workbook = XLSX.read(fileData, {
                cellStyles: true,
                cellFormulas: true,
                cellDates: true,
                cellNF: true,
                sheetStubs: true
            });
            
            // Get the first sheet from the workbook
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(worksheet);
            
            // Add these rows to our collection
            allRows = allRows.concat(rows);
            
            console.log(`Found ${rows.length} rows in file ${index + 1}`);
        });
        
        console.log(`\nTotal rows from all files: ${allRows.length}`);
        
        // Use a Map to store unique entries
        // The key will be a combination of company ID and phone number
        const uniqueRows = new Map();
        
        // Process each row and keep only unique combinations
        allRows.forEach(row => {
            // Create a unique key by combining ID and phone number
            const key = `${row['חפ']}_${row['טלפון בעלים']}`;
            uniqueRows.set(key, row);
        });
        
        // Convert back to array and sort
        const mergedRows = Array.from(uniqueRows.values())
            .sort((a, b) => {
                // Sort by company ID first
                if (a['חפ'] !== b['חפ']) {
                    return a['חפ'] - b['חפ'];
                }
                // Then by phone number if IDs are the same
                return a['טלפון בעלים'].localeCompare(b['טלפון בעלים']);
            });
        
        // Create new workbook for the merged data
        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.json_to_sheet(mergedRows);
        
        // Set column widths for better readability
        const colWidths = [
            { wch: 15 },  // Company ID column width
            { wch: 15 }   // Phone number column width
        ];
        newWorksheet['!cols'] = colWidths;
        
        // Add the worksheet to the workbook
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Merged Data');
        
        // Create output filename with timestamp
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const outputFile = `merged_unique_${timestamp}.xlsx`;
        
        // Save the merged file
        XLSX.writeFile(newWorkbook, outputFile);
        
        // Print summary statistics
        console.log('\nMerging Summary:');
        console.log(`Total input files: ${inputFiles.length}`);
        console.log(`Total rows before deduplication: ${allRows.length}`);
        console.log(`Unique rows after deduplication: ${mergedRows.length}`);
        console.log(`Duplicates removed: ${allRows.length - mergedRows.length}`);
        console.log(`\nOutput saved to: ${outputFile}`);
        
        // Show example of merged data
        console.log('\nFirst few entries from merged data:');
        mergedRows.slice(0, 3).forEach(row => {
            console.log(`ID: ${row['חפ']}, Phone: ${row['טלפון בעלים']}`);
        });
        
    } catch (error) {
        console.error('\nError during merge process:');
        if (error.code === 'ENOENT') {
            console.error('One or more input files not found. Please check filenames and paths.');
        } else {
            console.error(error.message);
        }
        process.exit(1);
    }
}

// Get input filenames from command line arguments
const inputFiles = process.argv.slice(2);

// Check if we have any input files
if (inputFiles.length === 0) {
    console.error('Please provide at least one input file.');
    console.error('Usage: node mergeExcels.js file1.xlsx file2.xlsx [file3.xlsx ...]');
    process.exit(1);
}

// Execute the merge with all provided files
mergeAndDeduplicate(inputFiles);