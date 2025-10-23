# Excel File Merger

A Python solution to merge multiple Excel files from subdirectories into consolidated files, one per subdirectory.

## 🎯 Objective

This project processes a root directory containing multiple subdirectories, where each subdirectory contains several Excel files, and merges all Excel files within each subdirectory into a single consolidated Excel file.

## 📁 Project Structure

```
FileMerger/
├── src/                           # Source code
│   ├── __init__.py
│   ├── excel_merger.py           # Core merger logic
│   └── merge_excel_files.py      # Main Python script
├── notebooks/                     # Jupyter notebooks
│   └── merge_excel_files.ipynb   # Interactive notebook
├── Data/                          # Excel files (ignored by git)
│   ├── Export 1 - Daily Delivery/
│   ├── Export 2 - Hour Delivery/
│   ├── Export 4 - Target profile/
│   ├── Export 6 - Engagement (FB - IG)/
│   └── TikTok/
├── requirements.txt               # Dependencies
├── .gitignore                    # Git ignore rules
├── .python-version               # Python 3.13.2
└── README.md                     # This file
```

## 🚀 Features

- ✅ **Merges all Excel files** within each subdirectory
- ✅ **Handles column mismatches** gracefully by adding missing columns with NaN values
- ✅ **Preserves all data** without information loss
- ✅ **Supports multiple formats** (.xlsx and .xls files)
- ✅ **Progress feedback** with detailed status messages
- ✅ **Error handling** and validation
- ✅ **Two interfaces**: Command-line script and Jupyter notebook

## 📦 Installation

1. **Clone or download** this repository
2. **Create a virtual environment** (recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # On macOS/Linux
   ```
3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

## 🔧 Usage

### Command-Line Interface

```bash
# Basic usage
python src/merge_excel_files.py /path/to/your/data

# With verbose output
python src/merge_excel_files.py /path/to/your/data --verbose

# Examples
python src/merge_excel_files.py ./Data
python src/merge_excel_files.py "C:\\Users\\Data"
```

### Jupyter Notebook Interface

1. **Start Jupyter**:
   ```bash
   jupyter notebook notebooks/merge_excel_files.ipynb
   ```

2. **Follow the interactive steps** in the notebook to:
   - Set your data directory path
   - Initialize the Excel merger
   - Process all subdirectories
   - View results and preview merged files

## 📊 Expected Output

For each subdirectory, the tool creates a merged file with the naming convention:
- `{subdirectory_name}_merged.xlsx`

**Example**: If you have a subdirectory named "Sales_Q1", the output file will be "Sales_Q1_merged.xlsx"

## 🔍 How It Works

1. **Directory Validation**: Checks that the provided directory exists and contains subdirectories
2. **File Discovery**: Finds all Excel files (.xlsx, .xls) in each subdirectory
3. **Data Reading**: Reads each Excel file into a pandas DataFrame
4. **Column Handling**: Automatically handles column mismatches by adding missing columns with NaN values
5. **Merging**: Uses `pandas.concat()` with `sort=False` to preserve column order
6. **Output**: Saves the merged DataFrame as a new Excel file in the same subdirectory

## ⚠️ Error Handling

The tool handles various error scenarios:

- **Invalid directory paths**: Reports if the directory doesn't exist
- **Empty subdirectories**: Gracefully handles subdirectories with no Excel files
- **Unreadable files**: Reports files that cannot be read due to formatting issues
- **Non-Excel files**: Automatically skips files that aren't Excel format
- **Save errors**: Reports any issues when saving merged files

## 📋 Requirements

- **Python 3.7+**
- **pandas** - Data manipulation and merging
- **openpyxl** - Excel file reading/writing (.xlsx files)
- **xlwings** - Advanced Excel operations and additional formats

## 🧪 Testing

Test the tool with your data:

```bash
# Test with the included Data directory
python src/merge_excel_files.py ./Data
```

## 📝 Notes

- **Original files are preserved**: The tool never modifies or deletes original files
- **Independent processing**: Each subdirectory is processed independently
- **Column preservation**: All columns from all files are preserved in the merged output
- **Data integrity**: No data is lost during the merging process

## 🤝 Contributing

Feel free to submit issues, feature requests, or pull requests to improve this tool.

## 📄 License

This project is open source and available under the MIT License.
