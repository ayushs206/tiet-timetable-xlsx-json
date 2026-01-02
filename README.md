# TIET Timetable XLSX → JSON Converter

Convert *Thapar University (TIET)* timetable XLSX files into a JSON format compatible with **[ACMTimeTable](https://github.com/ACM-Thapar/ACMTimeTable)**.

## 🔧 What It Does

This tool takes an XLSX file following the Thapar timetable format and outputs a structured JSON file that can be used directly with the ACMTimeTable project.

Currently supported:

- ✅ Even Semester
- ✅ 1st Year
- ✅ **A Pool** only

⚠️ Note: Other semesters, years, and pools are not yet supported. You can extend the parser logic to add them (see *Extending Support* below).

## 📦 Files Included

| File | Description |
|------|-------------|
| `timetable.xlsx` | Example input XLSX for testing |
| `index.js` | Core conversion script |
| `grid.json` | Fetching data from xlsx (raw) |
| `result.json` | Actual output to get saved |
| `result2.json` | Saves the index of days |
| `package.json` | Project metadata & dependencies |

## 🚀 Installation

```bash
# Clone the repo
git clone https://github.com/ayushs206/tiet-timetable-xlsx-json.git
cd tiet-timetable-xlsx-json

# Install dependencies
npm install
