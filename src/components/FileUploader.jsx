import React from 'react';

export default function FileUploader({ dragActive, onDrag, onDrop, onFileChange }) {
  return (
    <div 
      className={`file-drop-zone ${dragActive ? 'bg-opacity-20' : ''}`}
      onDragEnter={onDrag}
      onDragLeave={onDrag}
      onDragOver={onDrag}
      onDrop={onDrop}
    >
      <div className="flex items-center justify-center">
        <label className="cursor-pointer">
          <div className="flex flex-col items-center space-y-4">
            <div className="text-2xl">[SYSTEM]</div>
            <div className="text-lg">
              <span className="prompt">$</span> UPLOAD_FILE
            </div>
            <div className="text-sm opacity-70">
              <span className="prompt">&gt;</span> SUPPORTED_FORMATS: .json .doc .docx
            </div>
          </div>
          <input
            type="file"
            className="hidden"
            accept=".json,.doc,.docx"
            onChange={onFileChange}
          />
        </label>
      </div>
    </div>
  );
}