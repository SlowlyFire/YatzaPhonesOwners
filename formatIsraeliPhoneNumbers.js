const XLSX = require('xlsx');
const fs = require('fs');

function formatIsraeliPhoneNumbers(inputFile) {
    try {
        console.log(`Reading file: ${inputFile}`);
        
        // Read the Excel file with the special format entries
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
        
        // Function to clean and format a single phone number
        function formatPhoneNumber(phone) {
            let number = phone.toString()
                .trim()
                .replace(/^tel:\s*/, '')         // Remove 'tel:' prefix
                .replace(/\*+/g, '')             // Remove asterisks
                .replace(/nan/gi, '')            // Remove 'nan'
                .replace(/\/+/g, '')             // Remove forward slashes
                .replace(/[^\d\s+\-]/g, '')      // Keep only digits, spaces, plus signs, and hyphens
                .trim();
            
            // Handle international format
            number = number
                .replace(/^\+/, '')              // Remove leading plus
                .replace(/^972/, '')             // Remove country code
                .replace(/^0*/, '')              // Remove all leading zeros
                .replace(/\-/g, '')              // Remove existing hyphens
                .replace(/\s+/g, '');            // Remove all spaces
            
            // Add leading zero if missing
            if (number && !number.startsWith('0')) {
                number = '0' + number;
            }
            
            // Return empty string for invalid numbers
            if (!number || number.length < 8) return '';
            
            // Format based on number pattern
            if (number.length === 10 && number.startsWith('05')) {
                return `${number.slice(0, 3)}-${number.slice(3)}`;  // Mobile format
            } else if (number.length === 9 && number.startsWith('0')) {
                return `${number.slice(0, 2)}-${number.slice(2)}`;  // Landline format
            }
            
            return '';
        }
        
        // Function to extract all possible phone numbers from a string
        function extractPhoneNumbers(phoneString) {
            if (!phoneString) return [];
            
            // First normalize separators
            let normalized = phoneString
                .replace(/\s*;\s*/g, ',')        // Convert semicolons to commas
                .replace(/\s*\/\/\s*/g, ',')     // Convert double slashes to commas
                .replace(/\s*\|\s*/g, ',');      // Convert pipes to commas
            
            let numbers = [];
            
            // Split by commas first
            const parts = normalized.split(',');
            
            parts.forEach(part => {
                // Handle space-separated numbers
                const spaceParts = part.trim().split(/\s+/);
                
                if (spaceParts.length > 1) {
                    // Add complete numbers that might be in the parts
                    spaceParts.forEach(sp => {
                        if (/\d/.test(sp)) {
                            const formatted = formatPhoneNumber(sp);
                            if (formatted) numbers.push(formatted);
                        }
                    });
                    
                    // Look for area code + number combinations
                    for (let i = 0; i < spaceParts.length - 1; i++) {
                        if (/^\d{2,3}$/.test(spaceParts[i]) && /^\d{6,7}$/.test(spaceParts[i + 1])) {
                            const formatted = formatPhoneNumber(spaceParts[i] + spaceParts[i + 1]);
                            if (formatted) numbers.push(formatted);
                        }
                    }
                } else if (/\d/.test(part)) {
                    const formatted = formatPhoneNumber(part);
                    if (formatted) numbers.push(formatted);
                }
            });
            
            // Remove duplicates while preserving order
            return Array.from(new Set(numbers));
        }
        
        // Process each row and create new entries
        const formattedRows = [];
        
        data.forEach(row => {
            const companyId = row['חפ'];
            const phoneString = row['טלפון בעלים'].toString();
            
            // Get all valid phone numbers for this entry
            const phoneNumbers = extractPhoneNumbers(phoneString);
            
            // Create a new row for each valid phone number
            phoneNumbers.forEach(phone => {
                formattedRows.push({
                    'חפ': companyId,
                    'טלפון בעלים': phone
                });
            });
        });
        
        // Sort by company ID and phone number
        const sortedRows = formattedRows.sort((a, b) => {
            if (a['חפ'] !== b['חפ']) {
                return a['חפ'] - b['חפ'];
            }
            return a['טלפון בעלים'].localeCompare(b['טלפון בעלים']);
        });
        
        // Create new workbook
        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.json_to_sheet(sortedRows);
        
        // Set column widths
        const colWidths = [
            { wch: 15 },  // Column A width
            { wch: 15 }   // Column B width
        ];
        newWorksheet['!cols'] = colWidths;
        
        // Add the worksheet
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Formatted Numbers');
        
        // Create output filename
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const outputFile = `formatted_numbers_${timestamp}.xlsx`;
        
        // Save the file
        XLSX.writeFile(newWorkbook, outputFile);
        
        // Print summary
        console.log(`\nProcessing Summary:`);
        console.log(`Input file: ${inputFile}`);
        console.log(`Output file: ${outputFile}`);
        console.log(`Original entries: ${data.length}`);
        console.log(`Formatted entries: ${sortedRows.length}`);
        
        // Show some examples
        console.log('\nExample transformations:');
        for (let i = 0; i < Math.min(3, data.length); i++) {
            const originalRow = data[i];
            const formattedRows = sortedRows.filter(row => row['חפ'] === originalRow['חפ']);
            
            console.log(`\nOriginal: ID ${originalRow['חפ']}, Numbers: ${originalRow['טלפון בעלים']}`);
            console.log('Formatted into:');
            formattedRows.forEach(row => {
                console.log(`- ${row['טלפון בעלים']}`);
            });
        }
        
    } catch (error) {
        console.error('Error formatting phone numbers:', error);
        if (error.code === 'ENOENT') {
            console.error('File not found. Please check the filename and path.');
        }
    }
}

// Get input filename from command line argument, or use default
const inputFile = process.argv[2];

if (!inputFile) {
    console.error('Please provide an input file name.');
    console.error('Usage: node formatIsraeliPhoneNumbers.js <input-file.xlsx>');
    process.exit(1);
}

// Execute the function
formatIsraeliPhoneNumbers(inputFile);