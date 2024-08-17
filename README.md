# Excel Header Logo Replacement Script

This Python script automates the process of replacing logos in the headers of multiple Excel files. It leverages the `xlwings` library to access and modify the headers of Excel sheets, removing any existing images and inserting a new logo.

## Features

- **Bulk Processing**: Replaces logos across all Excel files in a specified directory.
- **Header Manipulation**: Targets and modifies the headers of all sheets within each workbook.
- **Error Handling**: Skips temporary files and handles exceptions gracefully.

## Requirements

- Python 3.6 or higher
- The following Python libraries:
  - `xlwings`
  - `os`

You can install the required libraries using pip:

```bash
pip install xlwings
```

## Usage

1. **Clone the Repository**:  
   Clone the repository to your local machine:

   ```bash
   git clone https://github.com/yourusername/repository-name.git
   ```

2. **Prepare the Script**:  
   Ensure that the paths to your Excel files and the new logo image are correctly set in the script:

   ```python
   folder_path = r"C:\path\to\your\excel\files"
   new_logo_path = r"C:\path\to\your\new\logo.jpg"
   ```

3. **Run the Script**:  
   Execute the script using Python:

   ```bash
   python app.py
   ```

   The script will process all `.xlsx` and `.xlsm` files in the specified directory, replacing any existing header logos with the new one.

4. **Check the Results**:  
   Once the script completes, the logos in the headers of your Excel files will be updated. You can open the Excel files to verify the changes.

## Error Handling

- The script automatically skips temporary Excel files (those starting with `~$`).
- Any errors encountered during the process will be printed to the console, and the script will continue processing the remaining files.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your changes.

