const XLSX = require('xlsx');
const _ = require('lodash');
const fs = require('fs');

function processPhoneNumbers(inputFile) {
    try {
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const outputFile = `processed_calls_${timestamp}.xlsx`;
        
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
        
        function formatPhoneNumber(phone) {
            let number = phone
                .trim()
                .replace(/^tel:\s*/, '')         // Remove 'tel:' prefix
                .replace(/\*+/g, '')             // Remove asterisks
                .replace(/nan/gi, '')            // Remove 'nan' (case insensitive)
                .replace(/\/+/g, '')             // Remove forward slashes
                .trim();
            
            if (!number || number.length < 2) return '';
            
            number = number
                .replace(/^\+/, '')              // Remove leading plus
                .replace(/^972/, '')             // Remove country code
                .replace(/^0*/, '');             // Remove all leading zeros
            
            number = number.replace(/[^\d\s]/g, '');
            
            const parts = number.trim().split(/\s+/);
            if (parts.length === 2) {
                number = parts.join('');
            } else {
                number = number.replace(/\s+/g, '');
            }
            
            if (number && !number.startsWith('0')) {
                number = '0' + number;
            }
            
            if (number.length === 10 && number.startsWith('05')) {
                return `${number.slice(0, 3)}-${number.slice(3)}`;
            } else if (number.length === 9 && number.startsWith('0')) {
                return `${number.slice(0, 2)}-${number.slice(2)}`;
            }
            
            return '';
        }
        
        function extractPhoneNumbers(phoneString) {
            let numbers = [];
            
            // First, handle complex number patterns (space separated with potential area codes)
            const spaceMatches = phoneString.match(/\d+\s+\d{2,3}\s+\d+/g) || [];
            spaceMatches.forEach(match => {
                const parts = match.trim().split(/\s+/);
                if (parts.length >= 2) {
                    // Handle first number if it looks complete
                    if (parts[0].length >= 8) {
                        numbers.push(parts[0]);
                    }
                    // Handle area code + number combinations
                    for (let i = 0; i < parts.length - 1; i++) {
                        if (/^\d{2,3}$/.test(parts[i]) && /^\d{6,7}$/.test(parts[i + 1])) {
                            numbers.push(parts[i] + parts[i + 1]);
                        }
                    }
                }
            });
            
            // Then handle comma/semicolon separated numbers
            const separatorSplit = phoneString
                .replace(/\s*;\s*/g, ',')        // Convert semicolons to commas
                .replace(/\s*\/\/\s*/g, ',')     // Convert double slashes to commas
                .split(',');
                
            separatorSplit.forEach(part => {
                const trimmed = part.trim();
                if (trimmed) {
                    if (trimmed.includes(' ')) {
                        // Handle space-separated numbers not caught earlier
                        const spaceParts = trimmed.split(/\s+/);
                        if (spaceParts.length === 2 && /^\d{2,3}$/.test(spaceParts[0]) && /^\d{6,7}$/.test(spaceParts[1])) {
                            numbers.push(spaceParts.join(''));
                        } else {
                            // Add each part that looks like a number
                            spaceParts.forEach(sp => {
                                if (/\d/.test(sp)) numbers.push(sp);
                            });
                        }
                    } else {
                        // Add the whole part if it contains digits
                        if (/\d/.test(trimmed)) numbers.push(trimmed);
                    }
                }
            });
            
            // Format each number but preserve duplicates
            return numbers
                .map(num => formatPhoneNumber(num))
                .filter(num => num.length > 0);
        }
        
        // Process each row
        const splitRows = [];
        data.forEach(row => {
            const companyId = row['חפ'];
            const phoneString = row['טלפון בעלים'].toString();
            
            // Get all phone numbers, including duplicates
            const phoneNumbers = extractPhoneNumbers(phoneString);
            
            // Create a new row for each phone number
            phoneNumbers.forEach(phone => {
                splitRows.push({
                    'חפ': companyId,
                    'טלפון בעלים': phone
                });
            });
        });
        
        // Sort data by company ID (keeping duplicate entries)
        const sortedData = _.sortBy(splitRows, ['חפ', 'טלפון בעלים']);
        
        // Create and save the new workbook
        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.json_to_sheet(sortedData);
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Processed Data');
        XLSX.writeFile(newWorkbook, outputFile);
        
        // Print processing summary with detailed counts
        console.log('\nProcessing Summary:');
        console.log(`Original file: ${inputFile}`);
        console.log(`New processed file: ${outputFile}`);
        console.log(`Original rows: ${data.length}`);
        console.log(`Processed rows: ${sortedData.length}`);
        
        // Print some example duplicate entries for verification
        const duplicates = _.groupBy(sortedData, row => `${row['חפ']}_${row['טלפון בעלים']}`);
        const duplicateEntries = Object.entries(duplicates)
            .filter(([_, group]) => group.length > 1)
            .slice(0, 5);
            
        if (duplicateEntries.length > 0) {
            console.log('\nExample duplicate entries preserved:');
            duplicateEntries.forEach(([key, group]) => {
                const [id, phone] = key.split('_');
                console.log(`ID ${id}: ${phone} (appears ${group.length} times)`);
            });
        }
        
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
processPhoneNumbers(inputFile);