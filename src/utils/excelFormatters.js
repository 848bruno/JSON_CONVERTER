/**
 * Utility functions for Excel formatting
 */

/**
 * Formats cell value for Excel
 * @param {any} value - The cell value to format
 * @returns {string} Formatted cell value
 */
export function formatCellValue(value) {
  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'boolean') {
    return value ? 'Yes' : 'No';
  }
  if (typeof value === 'number') {
    return Number.isInteger(value) ? value.toString() : value.toFixed(2);
  }
  if (value instanceof Date) {
    return value.toISOString().split('T')[0];
  }
  return String(value).trim();
}

/**
 * Calculates optimal column widths based on content
 * @param {Array} data - The worksheet data
 * @param {Array} headers - Column headers
 * @returns {Array} Array of column width specifications
 */
export function calculateColumnWidths(data, headers) {
  return headers.map(header => {
    const maxContentLength = Math.max(
      header.length,
      ...data.slice(1).map(row => {
        const value = row[headers.indexOf(header)];
        return value ? String(value).length : 0;
      })
    );

    return {
      wch: Math.min(Math.max(maxContentLength + 2, 15), 50) // Min 15, Max 50
    };
  });
}

/**
 * Prepares worksheet data with headers and values
 * @param {Array} data - The transformed data
 * @returns {Object} Worksheet data and headers
 */
export function prepareWorksheetData(data) {
  // Extract unique headers from data
  const headers = [...new Set(data.flatMap(obj => Object.keys(obj)))].sort();
  
  // Clean data by filtering out null or undefined rows
  const cleanedData = data.filter(row => 
    Object.values(row).some(value => value !== null && value !== undefined)
  );

  // Prepare worksheet data without any duplication of headers
  const wsData = [
    headers,
    ...cleanedData.map(obj =>
      headers.map(header => formatCellValue(obj[header]))
    )
  ];

  return { wsData, headers };
}
