/**
 * Fuzzy LOOKUP Google Apps Script
 * Provides fuzzy matching functionality similar to VLOOKUP but optimized for account/contact names
 * Usage: =LOOKUP(search_value, table_array, col_index_num, show_confidence)
 * 
 * @param {string} searchValue - The value to search for
 * @param {Array} tableArray - The range to search in (2D array)
 * @param {number} colIndexNum - Column index to return (1-based)
 * @param {boolean} showConfidence - Whether to show confidence level (true/false)
 * @return {string|Array} The matched value or array with confidence
 */

function LOOKUP(searchValue, tableArray, colIndexNum, showConfidence = false) {
  try {
    // Input validation
    if (!searchValue || !tableArray || !colIndexNum) {
      return "#ERROR: Missing required parameters";
    }
    
    if (!Array.isArray(tableArray) || tableArray.length === 0) {
      return "#ERROR: Invalid table array";
    }
    
    if (colIndexNum < 1 || colIndexNum > tableArray[0].length) {
      return "#ERROR: Column index out of range";
    }
    
    // Convert search value to string and normalize
    const normalizedSearchValue = normalizeString(searchValue.toString());
    
    let bestMatch = null;
    let bestScore = 0;
    let bestIndex = -1;
    
    // Search through the first column of the table array
    for (let i = 0; i < tableArray.length; i++) {
      const currentValue = tableArray[i][0];
      if (currentValue === null || currentValue === undefined) continue;
      
      const normalizedCurrentValue = normalizeString(currentValue.toString());
      
      // Calculate fuzzy match score using specialized algorithm for names
      const score = calculateNameMatchScore(normalizedSearchValue, normalizedCurrentValue);
      
      if (score > bestScore) {
        bestScore = score;
        bestMatch = tableArray[i][colIndexNum - 1]; // Convert to 0-based index
        bestIndex = i;
      }
    }
    
    // If no reasonable match found (threshold of 0.3 for names)
    if (bestScore < 0.3) {
      return showConfidence ? ["#N/A", "0%"] : "#N/A";
    }
    
    // Return result with or without confidence
    if (showConfidence) {
      const confidencePercent = Math.round(bestScore * 100) + "%";
      return [bestMatch || "", confidencePercent];
    } else {
      return bestMatch || "";
    }
    
  } catch (error) {
    return "#ERROR: " + error.toString();
  }
}

/**
 * Normalize string for better matching (handles names specifically)
 */
function normalizeString(str) {
  return str
    .toLowerCase()
    .trim()
    .replace(/[^\w\s]/g, '') // Remove punctuation
    .replace(/\s+/g, ' ') // Normalize whitespace
    .replace(/\b(inc|llc|corp|ltd|company|co)\b/g, '') // Remove common business suffixes
    .replace(/\b(mr|mrs|ms|dr|prof)\b/g, '') // Remove titles
    .trim();
}

/**
 * Calculate match score optimized for account/contact names
 * Combines multiple algorithms for best results with names
 */
function calculateNameMatchScore(str1, str2) {
  if (str1 === str2) return 1.0;
  if (!str1 || !str2) return 0.0;
  
  // Weight different matching algorithms
  const exactWeight = 0.4;
  const levenshteinWeight = 0.3;
  const jaroWinklerWeight = 0.2;
  const tokenWeight = 0.1;
  
  // Exact substring matching (high weight for names)
  const exactScore = calculateExactSubstringScore(str1, str2);
  
  // Levenshtein distance (good for typos)
  const levenshteinScore = calculateLevenshteinScore(str1, str2);
  
  // Jaro-Winkler (excellent for name variations)
  const jaroWinklerScore = calculateJaroWinklerScore(str1, str2);
  
  // Token-based matching (good for reordered names)
  const tokenScore = calculateTokenMatchScore(str1, str2);
  
  return (exactScore * exactWeight) + 
         (levenshteinScore * levenshteinWeight) + 
         (jaroWinklerScore * jaroWinklerWeight) + 
         (tokenScore * tokenWeight);
}

/**
 * Calculate exact substring matching score
 */
function calculateExactSubstringScore(str1, str2) {
  const longer = str1.length > str2.length ? str1 : str2;
  const shorter = str1.length > str2.length ? str2 : str1;
  
  if (longer.includes(shorter)) {
    return shorter.length / longer.length;
  }
  
  // Check for common substrings
  let maxCommonLength = 0;
  for (let i = 0; i < shorter.length; i++) {
    for (let j = i + 1; j <= shorter.length; j++) {
      const substring = shorter.substring(i, j);
      if (longer.includes(substring) && substring.length > maxCommonLength) {
        maxCommonLength = substring.length;
      }
    }
  }
  
  return maxCommonLength / Math.max(str1.length, str2.length);
}

/**
 * Calculate Levenshtein distance-based score
 */
function calculateLevenshteinScore(str1, str2) {
  const distance = levenshteinDistance(str1, str2);
  const maxLength = Math.max(str1.length, str2.length);
  return maxLength === 0 ? 1 : (maxLength - distance) / maxLength;
}

/**
 * Levenshtein distance calculation
 */
function levenshteinDistance(str1, str2) {
  const matrix = [];
  
  for (let i = 0; i <= str2.length; i++) {
    matrix[i] = [i];
  }
  
  for (let j = 0; j <= str1.length; j++) {
    matrix[0][j] = j;
  }
  
  for (let i = 1; i <= str2.length; i++) {
    for (let j = 1; j <= str1.length; j++) {
      if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1
        );
      }
    }
  }
  
  return matrix[str2.length][str1.length];
}

/**
 * Calculate Jaro-Winkler similarity score
 */
function calculateJaroWinklerScore(str1, str2) {
  const jaro = calculateJaroScore(str1, str2);
  
  if (jaro < 0.7) return jaro;
  
  // Calculate common prefix length (up to 4 characters)
  let prefixLength = 0;
  for (let i = 0; i < Math.min(str1.length, str2.length, 4); i++) {
    if (str1[i] === str2[i]) {
      prefixLength++;
    } else {
      break;
    }
  }
  
  return jaro + (0.1 * prefixLength * (1 - jaro));
}

/**
 * Calculate Jaro similarity score
 */
function calculateJaroScore(str1, str2) {
  if (str1 === str2) return 1.0;
  
  const len1 = str1.length;
  const len2 = str2.length;
  
  if (len1 === 0 || len2 === 0) return 0.0;
  
  const matchWindow = Math.floor(Math.max(len1, len2) / 2) - 1;
  if (matchWindow < 0) return 0.0;
  
  const str1Matches = new Array(len1).fill(false);
  const str2Matches = new Array(len2).fill(false);
  
  let matches = 0;
  let transpositions = 0;
  
  // Find matches
  for (let i = 0; i < len1; i++) {
    const start = Math.max(0, i - matchWindow);
    const end = Math.min(i + matchWindow + 1, len2);
    
    for (let j = start; j < end; j++) {
      if (str2Matches[j] || str1[i] !== str2[j]) continue;
      str1Matches[i] = true;
      str2Matches[j] = true;
      matches++;
      break;
    }
  }
  
  if (matches === 0) return 0.0;
  
  // Count transpositions
  let k = 0;
  for (let i = 0; i < len1; i++) {
    if (!str1Matches[i]) continue;
    while (!str2Matches[k]) k++;
    if (str1[i] !== str2[k]) transpositions++;
    k++;
  }
  
  return (matches / len1 + matches / len2 + (matches - transpositions / 2) / matches) / 3.0;
}

/**
 * Calculate token-based matching score (good for reordered names)
 */
function calculateTokenMatchScore(str1, str2) {
  const tokens1 = str1.split(/\s+/).filter(token => token.length > 0);
  const tokens2 = str2.split(/\s+/).filter(token => token.length > 0);
  
  if (tokens1.length === 0 || tokens2.length === 0) return 0;
  
  let matchingTokens = 0;
  const used = new Set();
  
  for (const token1 of tokens1) {
    for (let i = 0; i < tokens2.length; i++) {
      if (used.has(i)) continue;
      
      const token2 = tokens2[i];
      // Allow partial token matching for names
      if (token1.includes(token2) || token2.includes(token1) || 
          calculateLevenshteinScore(token1, token2) > 0.8) {
        matchingTokens++;
        used.add(i);
        break;
      }
    }
  }
  
  return matchingTokens / Math.max(tokens1.length, tokens2.length);
}

/**
 * Helper function to insert confidence values in adjacent cells
 * This function can be called manually to update confidence values
 */
function insertConfidenceValues() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getDataRange();
  const values = range.getValues();
  
  // Find cells with LOOKUP formulas and insert confidence values
  for (let row = 0; row < values.length; row++) {
    for (let col = 0; col < values[row].length; col++) {
      const cell = sheet.getRange(row + 1, col + 1);
      const formula = cell.getFormula();
      
      if (formula && formula.toUpperCase().includes('LOOKUP(')) {
        // Check if this is a confidence-enabled lookup
        if (formula.includes('TRUE') || formula.includes('true')) {
          // The formula should return an array, confidence will be in the next column
          Logger.log(`Found LOOKUP formula with confidence at ${cell.getA1Notation()}`);
        }
      }
    }
  }
}

/**
 * Test function to demonstrate usage
 */
function testFuzzyLookup() {
  // Example usage:
  const testData = [
    ['John Smith Inc', 'CEO', 'john@example.com'],
    ['Jane Doe LLC', 'Manager', 'jane@example.com'],
    ['Bob Johnson Corp', 'Director', 'bob@example.com'],
    ['Alice Brown Company', 'VP', 'alice@example.com']
  ];
  
  // Test searches
  console.log('Testing fuzzy LOOKUP:');
  console.log('Search "Jon Smith" result:', LOOKUP('Jon Smith', testData, 2, false));
  console.log('Search "Jon Smith" with confidence:', LOOKUP('Jon Smith', testData, 2, true));
  console.log('Search "Jane D" result:', LOOKUP('Jane D', testData, 3, false));
  console.log('Search "Jane D" with confidence:', LOOKUP('Jane D', testData, 3, true));
}
