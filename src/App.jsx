import React, { useState } from 'react';
import { saveAs } from 'file-saver';
import FileUploader from './components/FileUploader';
import JsonPreview from './components/JsonPreview';
import ConvertButton from './components/ConvertButton';
import { parseDocument } from './utils/documentParser';
import { exportToExcel } from './utils/excelUtils';

export default function App() {
  const [dragActive, setDragActive] = useState(false);
  const [jsonData, setJsonData] = useState(null);
  const [error, setError] = useState('');

  const handleDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(e.type === "dragenter" || e.type === "dragover");
  };

  const handleDrop = async (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    setError('');

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      await processFile(e.dataTransfer.files[0]);
    }
  };

  const handleFileChange = async (e) => {
    if (e.target.files && e.target.files[0]) {
      await processFile(e.target.files[0]);
    }
  };

  const processFile = async (file) => {
    try {
      setError('');
      const isWordDoc = file.name.toLowerCase().match(/\.(doc|docx)$/);
      const data = await parseDocument(
        isWordDoc ? file : await file.text(),
        isWordDoc
      );
      setJsonData(data);
    } catch (err) {
      setError(err.message || 'Error processing file');
      setJsonData(null);
    }
  };

  const convertToExcel = () => {
    if (!jsonData || !Array.isArray(jsonData) || jsonData.length === 0) {
      setError('No valid data to convert');
      return;
    }

    try {
      const blob = exportToExcel(jsonData, 'converted-data.xlsx');
      saveAs(blob, 'converted-data.xlsx');
      setError('');
    } catch (err) {
      setError(err.message || 'Error converting to Excel');
    }
  };

  return (
    <div className="min-h-screen p-4 terminal-container">
      <div className="max-w-6xl mx-auto space-y-8">
        <div className="text-center space-y-4">
          <div className="terminal-window">
            <div className="terminal-header">
              <div className="terminal-buttons">
                <span></span>
                <span></span>
                <span></span>
              </div>
              <div className="terminal-title">json_converter.exe</div>
            </div>
            <div className="terminal-content">
              <div className="ascii-art">
                {`
 _____ _____ _____ _____ 
|     |   __|     |   | |
| | | |__   |  |  | | | |
|_|_|_|_____|_____|_|___|
                `}
              </div>
              <div className="terminal-text">
                <span className="prompt">root@system</span>
                <span className="path">~/json-converter</span>
                <span className="cursor">█</span>
              </div>
            </div>
          </div>
        </div>

        <div className="space-y-8">
          <FileUploader 
            dragActive={dragActive}
            onDrag={handleDrag}
            onDrop={handleDrop}
            onFileChange={handleFileChange}
          />

          {error && (
            <div className="terminal-error">
              <span className="error-prefix">[ERROR]</span> {error}
            </div>
          )}

          {jsonData && (
            <div className="space-y-8">
              <ConvertButton onClick={convertToExcel} />
              <JsonPreview data={jsonData} />
            </div>
          )}
        </div>
      </div>
    </div>
  );
}