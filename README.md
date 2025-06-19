# Fuzzy LOOKUP for Google Sheets

A powerful Google Apps Script that brings intelligent fuzzy matching to Google Sheets, specifically optimized for account and contact name lookups. Think VLOOKUP, but smarter and more forgiving of typos, variations, and inconsistencies.

![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=flat&logo=google&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Version](https://img.shields.io/badge/Version-1.0.0-blue.svg)

## 🚀 Features

- **🎯 Smart Name Matching**: Optimized algorithms specifically designed for account and contact names
- **📊 Confidence Scoring**: Optional confidence percentages to gauge match reliability
- **🔧 VLOOKUP-like Syntax**: Familiar function structure with `LOOKUP()` instead of `VLOOKUP()`
- **🧠 Multi-Algorithm Approach**: Combines 4 different matching algorithms for superior results
- **⚡ Typo Tolerant**: Handles common misspellings, variations, and formatting differences
- **🏢 Business Name Aware**: Automatically handles business suffixes (Inc, LLC, Corp, etc.)
- **👤 Title Recognition**: Ignores common titles (Mr, Mrs, Dr, Prof, etc.)

## 📋 Table of Contents

- [Installation](#-installation)
- [Usage](#-usage)
- [Examples](#-examples)
- [Algorithm Details](#-algorithm-details)
- [Configuration](#-configuration)
- [Contributing](#-contributing)
- [License](#-license)

## 🛠️ Installation

### Method 1: Direct Script Installation

1. Open your Google Sheets document
2. Navigate to **Extensions** → **Apps Script**
3. Delete any existing code in the editor
4. Copy and paste the entire script from `fuzzy-lookup.gs`
5. Save the project (Ctrl+S or Cmd+S)
6. Return to your Google Sheet

### Method 2: Clone and Deploy

```bash
git clone https://github.com/yourusername/fuzzy-lookup-google-sheets.git
cd fuzzy-lookup-google-sheets
```

Then follow the steps above to add the script to your Google Apps Script project.

## 🎯 Usage

The `LOOKUP()` function follows a familiar syntax pattern:

```javascript
=LOOKUP(search_value, table_array, col_index_num, show_confidence)
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `search_value` | String | ✅ | The value you want to find |
| `table_array` | Range | ✅ | The range to search in (like A1:C100) |
| `col_index_num` | Number | ✅ | Column index to return (1-based) |
| `show_confidence` | Boolean | ❌ | `TRUE` to show confidence %, `FALSE` or omit for just results |

## 📚 Examples

### Basic Usage

```javascript
// Simple lookup without confidence
=LOOKUP("John Smith", A1:C100, 2, FALSE)

// Same as above (FALSE is default)
=LOOKUP("John Smith", A1:C100, 2)
```

### With Confidence Scoring

```javascript
// Returns both result and confidence percentage
=LOOKUP("John Smith", A1:C100, 2, TRUE)
// Result: ["john.smith@company.com", "85%"]
```

### Real-World Scenarios

#### Customer Database Matching
```javascript
// Find email for slightly misspelled customer name
=LOOKUP("Jon Smyth", CustomerData!A:D, 3, TRUE)
```

#### Account Reconciliation
```javascript
// Match company names with variations
=LOOKUP("ABC Corp", Accounts!A:E, 2, FALSE)
// Matches: "ABC Corporation", "A.B.C. Corp.", "ABC Corp Inc"
```

#### Contact Deduplication
```javascript
// Find duplicate contacts with different formatting
=LOOKUP("Dr. Jane Smith", Contacts!A:C, 2, TRUE)
// Matches: "Jane Smith MD", "Dr Jane Smith", "Smith, Jane Dr."
```

## 🧠 Algorithm Details

### Multi-Algorithm Approach

The fuzzy matching combines four specialized algorithms:

| Algorithm | Weight | Purpose | Best For |
|-----------|---------|---------|----------|
| **Exact Substring** | 40% | Direct partial matches | Abbreviated names, partial entries |
| **Levenshtein Distance** | 30% | Character-level differences | Typos, minor misspellings |
| **Jaro-Winkler** | 20% | String similarity with prefix bonus | Name variations, reorderings |
| **Token Matching** | 10% | Word-level matching | Different name orders, middle names |

### Name-Specific Optimizations

- **Business Suffix Removal**: Automatically ignores Inc, LLC, Corp, Ltd, Company, Co
- **Title Normalization**: Removes Mr, Mrs, Ms, Dr, Prof from comparisons
- **Punctuation Handling**: Normalizes dots, commas, and other punctuation
- **Whitespace Normalization**: Handles multiple spaces and formatting differences

### Confidence Thresholds

- **Minimum Match Threshold**: 30% (optimized for name matching)
- **High Confidence**: 80%+ (very likely correct match)
- **Medium Confidence**: 50-79% (probable match, review recommended)
- **Low Confidence**: 30-49% (possible match, manual verification needed)

## ⚙️ Configuration

### Adjusting Match Sensitivity

You can modify the minimum threshold in the `LOOKUP` function:

```javascript
// Change this line in the script (around line 45):
if (bestScore < 0.3) {  // Change 0.3 to your preferred threshold
```

### Algorithm Weight Tuning

Modify the weights in the `calculateNameMatchScore` function:

```javascript
const exactWeight = 0.4;        // Exact substring matching
const levenshteinWeight = 0.3;  // Typo tolerance
const jaroWinklerWeight = 0.2;  // Name variations
const tokenWeight = 0.1;        // Word reordering
```

## 🧪 Testing

The script includes a built-in test function. To run tests:

1. Open the Apps Script editor
2. Select the `testFuzzyLookup` function
3. Click the **Run** button
4. Check the console for test results

### Sample Test Data

```javascript
const testData = [
  ['John Smith Inc', 'CEO', 'john@example.com'],
  ['Jane Doe LLC', 'Manager', 'jane@example.com'],
  ['Bob Johnson Corp', 'Director', 'bob@example.com'],
  ['Alice Brown Company', 'VP', 'alice@example.com']
];
```

## 🔧 Troubleshooting

### Common Issues

#### Formula Returns "#ERROR"
- Check that all required parameters are provided
- Verify the table range exists and contains data
- Ensure column index is within the table range

#### No Matches Found (Returns "#N/A")
- Try lowering the confidence threshold
- Check for hidden characters in your data
- Verify the search value format

#### Performance Issues
- Limit table ranges to necessary data only
- Consider breaking large datasets into smaller chunks
- Use exact matching for high-volume operations

### Error Messages

| Error | Cause | Solution |
|-------|-------|----------|
| `#ERROR: Missing required parameters` | Missing search value, table, or column index | Provide all required parameters |
| `#ERROR: Invalid table array` | Table range is empty or invalid | Check your range reference |
| `#ERROR: Column index out of range` | Column number exceeds table width | Use valid column index (1-based) |

## 🤝 Contributing

We welcome contributions! Here's how you can help:

### Reporting Bugs
1. Use the [Issues](https://github.com/yourusername/fuzzy-lookup-google-sheets/issues) tab
2. Provide a clear description of the problem
3. Include sample data and expected vs actual results
4. Mention your Google Sheets version

### Suggesting Features
1. Check existing [Issues](https://github.com/yourusername/fuzzy-lookup-google-sheets/issues) first
2. Create a new issue with the "enhancement" label
3. Describe the use case and expected behavior

### Contributing Code
1. Fork the repository
2. Create a feature branch: `git checkout -b feature-name`
3. Make your changes and test thoroughly
4. Submit a pull request with a clear description

### Development Setup
```bash
# Clone your fork
git clone https://github.com/yourusername/fuzzy-lookup-google-sheets.git

# Create a new branch
git checkout -b feature-branch-name

# Make changes and test in Google Apps Script

# Commit and push
git commit -m "Add new feature"
git push origin feature-branch-name
```

## 📈 Performance

### Benchmarks

| Dataset Size | Average Lookup Time | Memory Usage |
|--------------|-------------------|--------------|
| 100 rows | <100ms | Minimal |
| 1,000 rows | <500ms | Low |
| 10,000 rows | 2-5 seconds | Moderate |

### Optimization Tips

- **Use smaller ranges**: Limit your table array to necessary columns and rows
- **Cache results**: For repeated lookups, consider copying results to avoid recalculation
- **Batch operations**: Process multiple lookups together when possible

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Inspired by the need for better name matching in business applications
- Built on Google Apps Script platform
- Uses established string similarity algorithms (Levenshtein, Jaro-Winkler)

## 📞 Support

- **Documentation**: This README and inline code comments
- **Issues**: Use GitHub Issues for bug reports and feature requests
- **Discussions**: Use GitHub Discussions for questions and community support

---

**⭐ If this project helps you, please consider giving it a star!**

Built with ❤️ for the Google Sheets community
