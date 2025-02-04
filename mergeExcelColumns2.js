const XLSX = require('xlsx');
const _ = require('lodash');
const fs = require('fs');
const path = require('path');

function mergeExcelColumns(inputFile) {
    try {
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
            // Enhanced phone number cleaning function with international number handling
            const cleanPhoneNumbers = (phoneString) => {
                return phoneString.toString()
                    .split(',')
                    .map(number => {
                        // First, handle international format
                        let cleanNumber = number
                            .trim()
                            .replace(/\*+/g, '')         // Remove asterisks
                            .replace(/nan/gi, '')        // Remove 'nan' (case insensitive)
                            .replace(/[^\d+]/g, '');     // Keep only digits and plus sign

                        // Handle international formats (972 or +972)
                        cleanNumber = cleanNumber
                            .replace(/^(\+)?972/, '')    // Remove +972 or 972 prefix
                            .replace(/^0*/, '');         // Remove leading zeros

                        // Add leading zero if it's missing (for local format)
                        if (cleanNumber && !cleanNumber.startsWith('0')) {
                            cleanNumber = '0' + cleanNumber;
                        }

                        // Format with proper hyphens for area codes
                        if (cleanNumber.length >= 9) {  // Valid phone number length
                            // Handle different area code lengths (2 or 3 digits)
                            const areaCodeMatch = cleanNumber.match(/^0(\d{2,3})(\d+)$/);
                            if (areaCodeMatch) {
                                return `${areaCodeMatch[1]}-${areaCodeMatch[2]}`;
                            }
                        }
                        
                        return cleanNumber;
                    })
                    .filter(number => {
                        // Keep only valid phone numbers
                        const digits = number.replace(/-/g, '');
                        return digits.length >= 8 && digits.length <= 10;
                    });
            };
            
            // Process phone numbers for this company
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
        
        // Show examples of international number handling
        console.log('\nExample of cleaned phone numbers:');
        const examples = Object.entries(groupedData)
            .filter(([_, group]) => 
                group.some(entry => 
                    entry['טלפון בעלים'].includes('972') || 
                    entry['טלפון בעלים'].includes('+972') ||
                    entry['טלפון בעלים'].includes('nan') || 
                    entry['טלפון בעלים'].includes('*')
                )
            )
            .slice(0, 5);  // Show first 5 examples
            
        examples.forEach(([companyId, entries]) => {
            console.log(`\nCompany ID (חפ): ${companyId}`);
            console.log('Original entries:');
            entries.forEach(entry => {
                console.log(`  - ${entry['טלפון בעלים']}`);
            });
            
            const mergedEntry = mergedData.find(row => row['חפ'].toString() === companyId);
            console.log('Cleaned and merged into:');
            console.log(`  → ${mergedEntry['טלפון בעלים']}`);
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
mergeExcelColumns(inputFile);