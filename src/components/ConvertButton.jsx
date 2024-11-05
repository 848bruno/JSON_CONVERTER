import React from 'react';

export default function ConvertButton({ onClick }) {
  return (
    <div className="space-y-4">
      <div className="terminal-window">
        <div className="terminal-content">
          <span className="prompt">$</span> FILE_STATUS: READY
          <br />
          <span className="prompt">$</span> CONVERSION_ENGINE: INITIALIZED
        </div>
      </div>
      
      <button
        onClick={onClick}
        className="terminal-button w-full"
      >
        $ EXECUTE_CONVERSION.exe
      </button>
    </div>
  );
}