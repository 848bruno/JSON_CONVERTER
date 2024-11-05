import XLSX from 'xlsx';
import { readFileSync } from 'fs';
import { exportToExcel } from './src/utils/excelUtils.js';

try {
  // Read and parse JSON data
  const jsonData = JSON.parse(readFileSync('data.json', 'utf8'));

  // Export to Excel with enhanced formatting
  exportToExcel(jsonData, 'formatted-output.xlsx');

  console.log('Excel file has been created successfully with improved formatting!');
} catch (error) {
  console.error('Error creating Excel file:', error.message);
}