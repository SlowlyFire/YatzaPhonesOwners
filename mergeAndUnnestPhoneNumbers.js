const XLSX = require('xlsx');
const _ = require('lodash');
const fs = require('fs');
const path = require('path');

function mergeAndUnnestPhoneNumbers(inputFile) {
    try {
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const outputFile = `unnested_calls_${timestamp}.xlsx`;
        
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
        
        // First, let's clean and process all phone numbers
        const cleanPhoneNumber = (number) => {
            // Handle the cleaning of a single phone number
            let cleanNumber = number
                .trim()
                .replace(/\*+/g, '')         // Remove asterisks
                .replace(/nan/gi, '')        // Remove 'nan' (case insensitive)
                .replace(/[^\d+]/g, '');     // Keep only digits and plus sign

            // Handle international formats
            cleanNumber = cleanNumber
                .replace(/^(\+)?972/, '')    // Remove +972 or 972 prefix
                .replace(/^0*/, '');         // Remove leading zeros

            // Add leading zero if missing
            if (cleanNumber && !cleanNumber.startsWith('0')) {
                cleanNumber = '0' + cleanNumber;
            }

            // Format with proper hyphens
            if (cleanNumber.length >= 9) {
                const areaCodeMatch = cleanNumber.match(/^0(\d{2,3})(\d+)$/);
                if (areaCodeMatch) {
                    return `${areaCodeMatch[1]}-${areaCodeMatch[2]}`;
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
                    // Keep only valid phone numbers
                    const digits = num.replace(/-/g, '');
                    return digits.length >= 8 && digits.length <= 10;
                });
            
            // Create a new row for each phone number
            phoneNumbers.forEach(phone => {
                unnested_data.push({
                    'חפ': row['חפ'],
                    'טלפון בעלים': phone
                });
            });
        });

        // Sort the data by company ID for better readability
        const sortedData = _.sortBy(unnested_data, ['חפ', 'טלפון בעלים']);
        
        // Create and save the new workbook
        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.json_to_sheet(sortedData);
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Unnested Data');
        XLSX.writeFile(newWorkbook, outputFile);
        
        // Print processing summary
        console.log('\nProcessing Summary:');
        console.log(`Original file: ${inputFile}`);
        console.log(`New unnested file: ${outputFile}`);
        console.log(`Original rows: ${data.length}`);
        console.log(`Unnested rows: ${sortedData.length}`);
        
        // Show some examples of the unnesting
        console.log('\nExample of unnested phone numbers:');
        const examples = Object.entries(_.groupBy(sortedData.slice(0, 10), row => row['חפ']))
            .slice(0, 3);
        
        examples.forEach(([companyId, entries]) => {
            if (entries.length > 1) {
                console.log(`\nCompany ID (חפ): ${companyId}`);
                console.log('Unnested entries:');
                entries.forEach(entry => {
                    console.log(`  - ${entry['טלפון בעלים']}`);
                });
            }
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
mergeAndUnnestPhoneNumbers(inputFile);