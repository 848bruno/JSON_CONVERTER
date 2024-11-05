import * as XLSX from 'xlsx';

/**
 * Creates and formats an Excel workbook from structured data
 * @param {Array} data - The data to convert
 * @returns {Object} XLSX workbook
 */
export function createExcelWorkbook(data) {
  const workbook = XLSX.utils.book_new();

  // Handle flat array of objects
  if (Array.isArray(data)) {
    const headers = [...new Set(data.flatMap(Object.keys))];
    
    // Prepare sheet data with headers and rows
    const sheetData = [];
    sheetData.push(headers);

    data.forEach(item => {
      const row = headers.map(header => {
        const value = item[header];
        
        // Format different value types
        if (value === null || value === undefined) return '';
        if (typeof value === 'boolean') return value ? 'Yes' : 'No';
        if (Array.isArray(value)) return value.join(', ');
        if (typeof value === 'number') {
          return Number.isInteger(value) ? value : value.toFixed(2);
        }
        return String(value).trim();
      });

      // Add the formatted row to sheetData
      sheetData.push(row);
    });

    // Create worksheet
    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);

    // Set column widths
    worksheet['!cols'] = headers.map(header => ({
      wch: Math.min(Math.max(
        header.length,
        ...sheetData.slice(1).map(row => 
          String(row[headers.indexOf(header)] || '').length
        )
      ) + 2, 50)
    }));

    // Add header styling
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let col = 0; col <= range.e.c; col++) {
      const headerCell = worksheet[XLSX.utils.encode_cell({ r: 0, c: col })];
      headerCell.s = {
        font: { bold: true },
        fill: { fgColor: { rgb: "E6E6E6" } },
        alignment: { horizontal: "center" }
      };
    }

    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');
  }

  return workbook;
}

/**
 * Exports data to Excel file
 * @param {Array} data - The data to export
 * @param {string} filename - Output filename
 */
export function exportToExcel(data, filename) {
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error('Invalid or empty data array');
  }

  const workbook = createExcelWorkbook(data);
  
  if (!workbook.SheetNames.length) {
    throw new Error('No data to export');
  }

  const blob = XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'array'
  });

  return new Blob([blob], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
}
