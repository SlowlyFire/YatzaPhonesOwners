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
        
        // Enhanced phone number cleaning function with Israeli format handling
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

            // Add leading zero
            if (cleanNumber && cleanNumber.length >= 8) {
                cleanNumber = '0' + cleanNumber;
                
                // Now format based on the number length
                const digits = cleanNumber.replace(/-/g, '');
                
                if (digits.length === 10) {
                    // Mobile phone format (05X-XXXXXXX)
                    if (digits.startsWith('05')) {
                        return `${digits.slice(0, 3)}-${digits.slice(3)}`;
                    }
                } else if (digits.length === 9) {
                    // Landline format (0X-XXXXXXX)
                    return `${digits.slice(0, 2)}-${digits.slice(2)}`;
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
                    // Must be either 9 digits (landline) or 10 digits (mobile)
                    return (digits.length === 9 || digits.length === 10) && digits.startsWith('0');
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
        
        // Print processing summary with examples
        console.log('\nProcessing Summary:');
        console.log(`Original file: ${inputFile}`);
        console.log(`New processed file: ${outputFile}`);
        console.log(`Original rows: ${data.length}`);
        console.log(`Processed rows: ${sortedData.length}`);
        
        // Show examples of both mobile and landline formats
        console.log('\nExample of processed phone numbers:');
        const examples = Object.entries(_.groupBy(sortedData.slice(0, 15), row => row['חפ']))
            .filter(([_, entries]) => entries.some(e => e['טלפון בעלים'].startsWith('05')) ||
                                    entries.some(e => !e['טלפון בעלים'].startsWith('05')))
            .slice(0, 3);
        
        examples.forEach(([companyId, entries]) => {
            console.log(`\nCompany ID (חפ): ${companyId}`);
            console.log('Processed entries:');
            entries.forEach(entry => {
                console.log(`  - ${entry['טלפון בעלים']} (${entry['טלפון בעלים'].startsWith('05') ? 'Mobile' : 'Landline'})`);
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