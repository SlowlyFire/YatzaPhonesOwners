const XLSX = require('xlsx');
const _ = require('lodash');
const fs = require('fs');
const path = require('path');

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
        
        // Enhanced phone number cleaning function with proper leading zero handling
        const cleanPhoneNumber = (number) => {
            // First clean the number of unwanted characters
            let cleanNumber = number
                .trim()
                .replace(/\*+/g, '')         // Remove asterisks
                .replace(/nan/gi, '')        // Remove 'nan' (case insensitive)
                .replace(/[^\d+]/g, '');     // Keep only digits and plus sign

            // Handle international formats
            cleanNumber = cleanNumber
                .replace(/^(\+)?972/, '')    // Remove +972 or 972 prefix
                .replace(/^0*/, '');         // Remove all leading zeros temporarily

            // Now we'll add back exactly one leading zero and format the number
            if (cleanNumber && cleanNumber.length >= 8) {
                // Add leading zero if it's missing
                cleanNumber = '0' + cleanNumber;
                
                // Format with proper hyphens for area codes
                // This regex captures the area code (2-3 digits after the leading zero)
                const areaCodeMatch = cleanNumber.match(/^0(\d{2,3})(\d+)$/);
                if (areaCodeMatch) {
                    return `0${areaCodeMatch[1]}-${areaCodeMatch[2]}`;
                }
            }
            
            return cleanNumber;
        };

        // Process and unnest the data
        const unnested_data = [];
        
        data.forEach(row => {
            // Get all phone numbers for this row
            const phoneNumbers = row['טלפון בעלים'].toString()
                .split(',')
                .map(num => cleanPhoneNumber(num))
                .filter(num => {
                    // Validate phone numbers
                    const digits = num.replace(/-/g, '');
                    return digits.length >= 9 && digits.length <= 10 && digits.startsWith('0');
                });
            
            // Create a new row for each valid phone number
            phoneNumbers.forEach(phone => {
                unnested_data.push({
                    'חפ': row['חפ'],
                    'טלפון בעלים': phone
                });
            });
        });

        // Sort the data by company ID and phone number
        const sortedData = _.sortBy(unnested_data, ['חפ', 'טלפון בעלים']);
        
        // Create and save the new workbook
        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.json_to_sheet(sortedData);
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Processed Data');
        XLSX.writeFile(newWorkbook, outputFile);
        
        // Print processing summary
        console.log('\nProcessing Summary:');
        console.log(`Original file: ${inputFile}`);
        console.log(`New processed file: ${outputFile}`);
        console.log(`Original rows: ${data.length}`);
        console.log(`Processed rows: ${sortedData.length}`);
        
        // Show examples of the processing
        console.log('\nExample of processed phone numbers:');
        const examples = Object.entries(_.groupBy(sortedData.slice(0, 10), row => row['חפ']))
            .slice(0, 3);
        
        examples.forEach(([companyId, entries]) => {
            console.log(`\nCompany ID (חפ): ${companyId}`);
            console.log('Processed entries:');
            entries.forEach(entry => {
                console.log(`  - ${entry['טלפון בעלים']}`);
            });
        });
        
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
processPhoneNumbers(inputFile);