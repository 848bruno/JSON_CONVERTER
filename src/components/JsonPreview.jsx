import React from 'react';

export default function JsonPreview({ data }) {
  if (!data || !Array.isArray(data) || data.length === 0) {
    return null;
  }

  const headers = [...new Set(data.flatMap(obj => Object.keys(obj)))];

  return (
    <div className="space-y-4">
      <div className="terminal-text">
        <span className="prompt">$</span> DATA_PREVIEW
      </div>
      <div className="terminal-window overflow-hidden">
        <div className="terminal-content">
          <div className="overflow-x-auto">
            <table className="terminal-table w-full">
              <thead>
                <tr>
                  {headers.map((header, index) => (
                    <th key={`header-${index}`} className="px-4 py-2 text-left">
                      {header}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {data.map((entry, rowIndex) => (
                  <tr key={`row-${rowIndex}`}>
                    {headers.map((header, colIndex) => (
                      <td key={`cell-${rowIndex}-${colIndex}`} className="px-4 py-2">
                        {entry[header] || '---'}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}