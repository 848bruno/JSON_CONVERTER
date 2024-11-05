import mammoth from 'mammoth';

function findJsonObjects(text) {
  // Enhanced regex to better detect JSON boundaries
  const jsonRegex = /(\{(?:[^{}]|(?:\{[^{}]*\}))*\}|\[(?:[^\[\]]|(?:\[.*?\]))*\])/g;
  const potentialMatches = text.match(jsonRegex) || [];
  const results = [];
  
  for (const match of potentialMatches) {
    try {
      const cleaned = match
        .replace(/[\u200B-\u200D\uFEFF]/g, '')
        .replace(/[\u2018\u2019]/g, "'")
        .replace(/[\u201C\u201D]/g, '"')
        .replace(/\n\s*(["}\\])/g, '$1')
        .replace(/([{[]),\s*\n\s*/g, '$1')
        .replace(/\\n/g, ' ')
        .replace(/\\r/g, ' ');

      const parsed = JSON.parse(cleaned);
      
      // Handle both array and object formats
      if (Array.isArray(parsed)) {
        results.push(...parsed);
      } else {
        results.push(parsed);
      }
    } catch (e) {
      // Continue to next match if parsing fails
      continue;
    }
  }
  
  return results;
}

function extractKeyValuePairs(text) {
  // Split text into blocks based on common separators
  const blocks = text.split(/(?:\r?\n\s*\r?\n|\{|\}|\[|\])+/).filter(block => block.trim());
  const entries = [];
  
  for (const block of blocks) {
    const lines = block.split(/[\n;]+/).filter(line => line.trim());
    let currentEntry = {};
    let hasValidPair = false;

    for (const line of lines) {
      // Look for key-value patterns
      const match = line.match(/^([^:=]+)[:=]\s*(.+)$/);
      if (match) {
        const [, key, value] = match;
        const cleanKey = key.trim()
          .replace(/['"]/g, '')
          .replace(/[\s_]+/g, '_')
          .replace(/[^\w\s-]/g, '')
          .toLowerCase();

        const cleanValue = value.trim()
          .replace(/^["']|["']$/g, '')
          .replace(/\s+/g, ' ');

        if (cleanKey && cleanValue) {
          currentEntry[cleanKey] = cleanValue;
          hasValidPair = true;
        }
      }
    }

    if (hasValidPair) {
      entries.push(currentEntry);
      currentEntry = {};
    }
  }

  return entries;
}

function flattenObject(obj, prefix = '') {
  const flattened = {};

  for (const [key, value] of Object.entries(obj)) {
    const newKey = prefix ? `${prefix}_${key}` : key;

    if (value === null || value === undefined) {
      flattened[newKey] = '';
    } else if (typeof value === 'object' && !Array.isArray(value)) {
      Object.assign(flattened, flattenObject(value, newKey));
    } else if (Array.isArray(value)) {
      if (value.length > 0 && typeof value[0] === 'object') {
        value.forEach((item, index) => {
          Object.assign(flattened, flattenObject(item, `${newKey}_${index}`));
        });
      } else {
        flattened[newKey] = value.map(item => String(item || '')).join(', ');
      }
    } else {
      flattened[newKey] = String(value);
    }
  }

  return flattened;
}

function normalizeData(data) {
  if (!Array.isArray(data)) {
    data = [data];
  }

  // Handle array of arrays (table-like data)
  if (data.length > 0 && Array.isArray(data[0])) {
    const headers = data[0];
    return data.slice(1).map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[String(header).toLowerCase()] = row[index] || '';
      });
      return obj;
    });
  }

  // Flatten and normalize all entries
  return data.map(entry => {
    if (typeof entry !== 'object' || entry === null) {
      return { value: String(entry) };
    }
    return flattenObject(entry);
  });
}

function removeDuplicates(entries) {
  const seen = new Set();
  return entries.filter(entry => {
    const key = JSON.stringify(entry);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function mergeEntries(entries) {
  // Remove duplicates first
  const uniqueEntries = removeDuplicates(entries);
  
  // Get all unique keys
  const allKeys = new Set();
  uniqueEntries.forEach(entry => {
    Object.keys(entry).forEach(key => allKeys.add(key));
  });

  // Normalize all entries to have the same structure
  return uniqueEntries.map(entry => {
    const normalized = {};
    for (const key of allKeys) {
      normalized[key] = entry[key] || '';
    }
    return normalized;
  });
}

export async function parseDocument(input, isWordDoc = false) {
  try {
    let text;
    let entries = [];
    
    if (isWordDoc) {
      const arrayBuffer = await input.arrayBuffer();
      const result = await mammoth.extractRawText({ arrayBuffer });
      text = result.value;
      
      // Find JSON objects first
      const jsonEntries = findJsonObjects(text);
      if (jsonEntries.length > 0) {
        entries.push(...normalizeData(jsonEntries));
      }
      
      // Look for key-value pairs in non-JSON sections
      const kvEntries = extractKeyValuePairs(text);
      if (kvEntries.length > 0) {
        entries.push(...kvEntries);
      }
    } else {
      // Handle JSON file
      text = input;
      try {
        const jsonData = JSON.parse(text);
        entries = normalizeData(jsonData);
      } catch (e) {
        // If JSON parsing fails, try key-value extraction
        const jsonEntries = findJsonObjects(text);
        if (jsonEntries.length > 0) {
          entries.push(...normalizeData(jsonEntries));
        }
        
        const kvEntries = extractKeyValuePairs(text);
        if (kvEntries.length > 0) {
          entries.push(...kvEntries);
        }
      }
    }

    if (entries.length === 0) {
      throw new Error('No valid data entries found in the document');
    }

    // Merge and normalize all entries, removing duplicates
    return mergeEntries(entries);
  } catch (error) {
    throw new Error(
      'Could not parse the document. Please ensure:\n' +
      '- The document contains valid data (JSON or key-value pairs)\n' +
      '- Each entry is properly formatted\n' +
      '- There are no special characters or formatting issues'
    );
  }
}