# File Merger (Excel & CSV)

A Python solution to merge multiple Excel and CSV files from subdirectories into consolidated files, one per subdirectory. Supports both direct Python execution and Docker containerization for easy distribution.

## 🎯 Objective

This project processes a root directory containing multiple subdirectories, where each subdirectory contains several Excel and CSV files, and merges all files within each subdirectory into a single consolidated file. Supports both direct Python execution and Docker containerization.

## 📁 Project Structure

```
FileMerger/
├── src/                           # Source code
│   ├── __init__.py
│   ├── file_merger.py            # Core merger logic
│   └── merge_excel_files.py      # Main Python script
├── notebooks/                     # Jupyter notebooks
│   └── merge_excel_files.ipynb   # Interactive notebook
├── Data/                          # Excel/CSV files (ignored by git)
│   ├── Export 1 - Daily Delivery/
│   ├── Export 2 - Hour Delivery/
│   ├── Export 4 - Target profile/
│   ├── Export 6 - Engagement (FB - IG)/
│   └── TikTok/
├── requirements.txt               # Python dependencies
├── Dockerfile                     # Docker container configuration
├── docker-compose.yml            # Docker Compose configuration
├── build.sh                      # Docker build script
├── .dockerignore                 # Docker ignore rules
├── .gitignore                    # Git ignore rules
├── .python-version               # Python 3.13.2
└── README.md                     # This file
```

## 🚀 Features

- ✅ **Merges Excel and CSV files** within each subdirectory
- ✅ **Handles column mismatches** gracefully by adding missing columns with NaN values
- ✅ **Preserves all data** without information loss
- ✅ **Supports multiple formats** (.xlsx, .xls, .csv files)
- ✅ **Column cleaning** with configurable options (whitespace, case, special characters)
- ✅ **Single-file optimization** - skips merging if only one file exists
- ✅ **Progress feedback** with detailed status messages
- ✅ **Error handling** and validation
- ✅ **Multiple interfaces**: Command-line script, Jupyter notebook, and Docker container
- ✅ **Cross-platform support** with Docker containerization

## 📦 Installation

### Option 1: Docker (Recommended for Easy Setup)

**Prerequisites:**
- Docker Desktop installed ([Download here](https://www.docker.com/products/docker-desktop/))

**Steps:**
1. **Clone the repository**:
   ```bash
   git clone https://github.com/yourusername/FileMerger.git
   cd FileMerger
   ```

2. **Build the Docker image**:
   ```bash
   ./build.sh
   ```
   
   Or manually:
   ```bash
   docker build -t file-merger:latest .
   ```

3. **Run with your data**:
   ```bash
   docker run -v $(pwd)/Data:/app/data file-merger:latest /app/data
   ```

### Option 2: Direct Python Installation

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

### Docker Usage (Recommended)

```bash
# Basic usage
docker run -v $(pwd)/Data:/app/data file-merger:latest /app/data

# With verbose output
docker run -v $(pwd)/Data:/app/data file-merger:latest /app/data --verbose

# With column cleaning options
docker run -v $(pwd)/Data:/app/data file-merger:latest /app/data --lowercase-columns
docker run -v $(pwd)/Data:/app/data file-merger:latest /app/data --remove-special-chars

# Windows users
docker run -v C:\MyData:/app/data file-merger:latest /app/data

# Using docker-compose
docker-compose up
```

### Direct Python Usage

```bash
# Basic usage
python src/merge_excel_files.py /path/to/your/data

# With verbose output
python src/merge_excel_files.py /path/to/your/data --verbose

# With column cleaning options
python src/merge_excel_files.py /path/to/your/data --lowercase-columns
python src/merge_excel_files.py /path/to/your/data --remove-special-chars

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

### For Direct Python Installation:
- **Python 3.7+**
- **pandas** - Data manipulation and merging
- **openpyxl** - Excel file reading/writing (.xlsx files)
- **xlrd** - Excel file reading (.xls files)
- **xlwings** - Advanced Excel operations (Windows/Mac only)
- **numpy** - Numerical operations

### For Docker:
- **Docker Desktop** - Container runtime
- **No Python installation required** - Everything is included in the container

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
- **Single-file optimization**: Subdirectories with only one file are skipped (no merge needed)

## ⚠️ Important Docker Limitation

**When using Docker, some Excel files may fail to process** if they require xlwings for reading. This happens because:

- **xlwings requires Excel application** to be installed and running
- **Docker containers run Linux** (even on Windows/Mac), which doesn't have Excel
- **Some problematic Excel files** can only be read by xlwings, not by pandas/openpyxl

**If you encounter files that fail in Docker but work with direct Python execution, use the Python version instead:**

```bash
# Install Python dependencies
pip install -r requirements.txt

# Run directly (with full xlwings support)
python src/merge_excel_files.py Data --verbose
```

**Recommendation**: Use Docker for convenience, but fall back to direct Python execution if you encounter files that fail to process.

## 🤝 Contributing

Feel free to submit issues, feature requests, or pull requests to improve this tool.

## 📄 License

This project is open source and available under the MIT License.
