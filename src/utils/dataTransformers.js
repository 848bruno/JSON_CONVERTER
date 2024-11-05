/**
 * Utility functions for data transformation
 */

/**
 * Formats array values for Excel
 * @param {Array} arr - Array to format
 * @returns {string} Formatted string
 */
function formatArrayValue(arr) {
  return arr.map(item => String(item).trim()).join(', ');
}

/**
 * Flattens nested objects and arrays within data
 * @param {Object} item - The object to flatten
 * @returns {Object} Flattened object
 */
export function flattenObject(item) {
  const flattened = {};

  for (const [key, value] of Object.entries(item)) {
    if (Array.isArray(value)) {
      flattened[key] = formatArrayValue(value);
    } else if (value && typeof value === 'object') {
      const nested = flattenObject(value);
      for (const [nestedKey, nestedValue] of Object.entries(nested)) {
        flattened[`${key}_${nestedKey}`] = nestedValue;
      }
    } else {
      flattened[key] = value ?? '';
    }
  }

  return flattened;
}

/**
 * Transforms data into a format suitable for Excel
 * @param {Array|Object} data - The data to transform
 * @returns {Object} Transformed data by category
 */
export function transformData(data) {
  if (!data || typeof data !== 'object') {
    throw new Error('Invalid data format');
  }

  // Handle single array of data
  if (Array.isArray(data)) {
    return {
      'Data': data.map(item => flattenObject(item))
    };
  }

  // Handle categorized data
  const transformed = {};
  for (const [category, items] of Object.entries(data)) {
    if (Array.isArray(items)) {
      transformed[category] = items
        .filter(item => item && typeof item === 'object')
        .map(item => flattenObject(item));
    }
  }

  return transformed;
}