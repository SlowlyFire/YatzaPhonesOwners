const XLSX = require('xlsx');
const _ = require('lodash');
const fs = require('fs');

function processFinalPhoneNumbers(inputFile) {
    try {
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const outputFile = `final_processed_calls_${timestamp}.xlsx`;
        
        console.log(`Reading file: ${inputFile}`);
        const fileData = fs.readFileSync(inputFile);
        const workbook = XLSX.read(fileData, {
            cellStyles: true,
            cellFormulas: true,
            cellDates: true,
            cellNF: true,
            sheetStubs: true
        });
        
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(worksheet);
        
        function formatPhoneNumber(phone) {
            // Clean the input first
            let number = phone
                .toString()
                .trim()
                .replace(/^tel:\s*/, '')         // Remove 'tel:' prefix
                .replace(/\*+/g, '')             // Remove asterisks
                .replace(/nan/gi, '')            // Remove 'nan'
                .replace(/\/+/g, '')             // Remove forward slashes
                .replace(/[^\d\s+\-]/g, '')      // Keep only digits, spaces, plus signs, and hyphens
                .trim();
            
            if (!number || number.length < 2) return '';
            
            // Handle international format and clean up
            number = number
                .replace(/^\+/, '')              // Remove leading plus
                .replace(/^972/, '')             // Remove country code
                .replace(/^0*/, '')              // Remove leading zeros
                .replace(/\-/g, '')              // Remove existing hyphens
                .replace(/\s+/g, '');            // Remove all spaces
            
            // Add leading zero if missing
            if (number && !number.startsWith('0')) {
                number = '0' + number;
            }
            
            // Format based on the number format
            if (number.length === 10 && number.startsWith('05')) {
                return `${number.slice(0, 3)}-${number.slice(3)}`;  // Mobile format: 05X-XXXXXXX
            } else if (number.length === 9 && number.startsWith('0')) {
                return `${number.slice(0, 2)}-${number.slice(2)}`;  // Landline format: 0X-XXXXXXX
            }
            
            return '';
        }
        
        function extractPhoneNumbers(phoneString) {
            if (!phoneString) return [];
            
            // First normalize different separators to commas
            const normalized = phoneString
                .replace(/\s*;\s*/g, ',')        // Convert semicolons to commas
                .replace(/\s*\/\/\s*/g, ',')     // Convert double slashes to commas
                .replace(/\s*\|\s*/g, ',')       // Convert pipes to commas
                .trim();
            
            // Array to store all found numbers
            let numbers = new Set();
            
            // Split by commas first
            const parts = normalized.split(',');
            
            parts.forEach(part => {
                part = part.trim();
                
                // Handle space-separated numbers
                const spaceParts = part.split(/\s+/);
                
                if (spaceParts.length > 1) {
                    // Check for complete numbers in parts
                    spaceParts.forEach(sp => {
                        if (/\d/.test(sp)) {
                            const formatted = formatPhoneNumber(sp);
                            if (formatted) numbers.add(formatted);
                        }
                    });
                    
                    // Check for area code + number combinations
                    for (let i = 0; i < spaceParts.length - 1; i++) {
                        if (/^\d{2,3}$/.test(spaceParts[i]) && /^\d{6,7}$/.test(spaceParts[i + 1])) {
                            const formatted = formatPhoneNumber(spaceParts[i] + spaceParts[i + 1]);
                            if (formatted) numbers.add(formatted);
                        }
                    }
                    
                    // Handle numbers that start with digits (possible complete numbers)
                    spaceParts.filter(part => /^\d{8,10}$/.test(part))
                        .forEach(num => {
                            const formatted = formatPhoneNumber(num);
                            if (formatted) numbers.add(formatted);
                        });
                } else if (/\d/.test(part)) {
                    const formatted = formatPhoneNumber(part);
                    if (formatted) numbers.add(formatted);
                }
            });
            
            return Array.from(numbers);
        }
        
        // Process all rows
        const processedRows = [];
        
        data.forEach(row => {
            const companyId = row['חפ'];
            const phoneString = row['טלפון בעלים'].toString();
            
            // Get all valid phone numbers
            const phoneNumbers = extractPhoneNumbers(phoneString);
            
            // Create a new row for each valid phone number
            phoneNumbers.forEach(phone => {
                processedRows.push({
                    'חפ': companyId,
                    'טלפון בעלים': phone
                });
            });
        });
        
        // Sort by company ID and phone number
        const sortedData = _.sortBy(processedRows, ['חפ', 'טלפון בעלים']);
        
        // Create new workbook
        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.json_to_sheet(sortedData);
        
        // Add the worksheet to the workbook
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Processed Numbers');
        
        // Write to file
        XLSX.writeFile(newWorkbook, outputFile);
        
        // Print processing summary
        console.log('\nProcessing Summary:');
        console.log(`Original file: ${inputFile}`);
        console.log(`New file created: ${outputFile}`);
        console.log(`Original rows: ${data.length}`);
        console.log(`Processed rows: ${sortedData.length}`);
        
        // Print sample of processed entries
        console.log('\nSample of processed entries:');
        sortedData.slice(0, 5).forEach(row => {
            console.log(`ID: ${row['חפ']}, Phone: ${row['טלפון בעלים']}`);
        });
        
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
processFinalPhoneNumbers(inputFile);