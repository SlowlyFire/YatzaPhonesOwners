const XLSX = require('xlsx');
const _ = require('lodash');
const fs = require('fs');

function cleanPhoneNumber(phone) {
    // Remove the 'tel:' prefix and trim spaces
    let cleaned = phone
        .replace(/^tel:\s*/, '')
        .trim();

    // Remove asterisks and 'nan'
    cleaned = cleaned
        .replace(/\*+/g, '')
        .replace(/nan/gi, '')
        .trim();

    // If we're left with nothing after cleaning, return empty string
    if (!cleaned) return '';

    // Handle international format (+972)
    cleaned = cleaned
        .replace(/\+/, '')        // Remove plus sign
        .replace(/^972/, '')      // Remove country code
        .replace(/^0*/, '');      // Remove leading zeros

    // If we have a valid number after cleaning
    if (cleaned.length >= 8) {
        // Add leading zero
        cleaned = '0' + cleaned;
        
        // Format based on length
        const digits = cleaned.replace(/\D/g, '');
        if (digits.length === 10 && digits.startsWith('05')) {
            // Mobile number format (05X-XXXXXXX)
            return `${digits.slice(0, 3)}-${digits.slice(3)}`;
        } else if (digits.length === 9 && digits.startsWith('0')) {
            // Landline format (0X-XXXXXXX)
            return `${digits.slice(0, 2)}-${digits.slice(2)}`;
        }
    }

    return '';
}

function splitPhoneNumbers(inputFile) {
    try {
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const outputFile = `split_phones_${timestamp}.xlsx`;
        
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
        
        const splitRows = [];
        
        data.forEach(row => {
            const companyId = row['חפ'];
            const phoneString = row['טלפון בעלים'].toString();
            
            // Split by both spaces and commas for international format
            const phoneNumbers = phoneString
                .split(/[\s,]+/)
                .map(phone => phone.trim())
                .map(phone => cleanPhoneNumber(phone))
                .filter(phone => phone.length > 0);  // Remove empty results
            
            // Create a new row for each valid phone number
            phoneNumbers.forEach(phone => {
                splitRows.push({
                    'חפ': companyId,
                    'טלפון בעלים': phone
                });
            });
        });
        
        const sortedData = _.sortBy(splitRows, ['חפ']);
        
        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.json_to_sheet(sortedData);
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Split Data');
        XLSX.writeFile(newWorkbook, outputFile);
        
        console.log('\nProcessing Summary:');
        console.log(`Original file: ${inputFile}`);
        console.log(`New split file: ${outputFile}`);
        console.log(`Original rows: ${data.length}`);
        console.log(`Split rows: ${sortedData.length}`);
        
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
splitPhoneNumbers(inputFile);