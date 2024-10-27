
# PRINTCHECK

**PRINTCHECK** is a tool designed to generate structured checklists for STL files used in 3D printing. It automates the process of creating a comprehensive checklist with previews for each STL file, organized by folder and subfolder structure. The goal is to help users ensure all necessary files are ready and reviewed before 3D printing, streamlining the workflow and minimizing errors.

## Features

- **Automatic Checklist Generation**: Creates a checklist from STL files found in a specified directory.
- **3D Preview Generation**: Generates 3D preview images for each STL file, color-coded based on naming conventions.
- **Folder Structure Retention**: Preserves the original folder and subfolder organization within the checklist.
- **Missing File Alerts**: Identifies STL files that fail to generate previews and lists them at the top of the checklist for easy tracking.
- **Excel Output**: Exports the checklist to an Excel file with columns for file name, preview, and status placeholders.

## Color-Coding

The preview images are color-coded based on the file names:
- **Red**: Files containing `[a]` in their name.
- **White**: Files containing `[c]` in their name.
- **Black**: All other files.

## Installation

### Step 1: Clone the Repository
   ```bash
   git clone https://github.com/yourusername/PRINTCHECK.git
   cd PRINTCHECK
   ```

### Step 2: Create a Virtual Environment
   It is recommended to use a virtual environment to manage dependencies and keep the project isolated from other Python installations.
   ```bash
   # On Windows
   python -m venv venv

   # On macOS/Linux
   python3 -m venv venv
   ```

### Step 3: Activate the Virtual Environment
   ```bash
   # On Windows
   .\venv\Scripts\activate

   # On macOS/Linux
   source ./venv/bin/activate
   ```

### Step 4: Install the Required Dependencies
   Once the virtual environment is activated, install the necessary packages using `requirements.txt`:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Running the Script
1. **Provide the STL Directory Path**: The script can accept the STL folder path as a command-line argument, or it will prompt the user to enter the path manually if no argument is provided.
   ```bash
   python create_checklist.py <path_to_stl_folder>
   ```
   Example:
   ```bash
   python create_checklist.py ./Stls
   ```

2. **Output**: The script processes all STL files in the specified directory, generating previews and creating an Excel checklist named `STL_Checklist_Structured.xlsx`. Any STL files that failed to generate previews are listed in the Excel file and logged.

### Configuration

#### Folder Structure
The script expects the STL files to be located in a user-specified directory. The output will be saved to a temporary folder during processing and will be deleted afterward.

#### File Naming Conventions
- Files containing `[a]` in the name will be rendered in red.
- Files containing `[c]` in the name will be rendered in white.
- All other files will be rendered in black.

## Example Output

The generated Excel file includes:
- **File Name**: The name of the STL file.
- **Preview**: A 200x200 pixel image of the STL file's 3D preview.
- **Checked and Available**: Placeholder for marking if the file has been reviewed.
- **Not Needed**: Placeholder for marking if the file is not required.

Missing previews will be highlighted at the top of the Excel sheet, showing both the file name and the relative folder path.

## Error Handling

The script provides verbose output to help diagnose any issues during execution. If a file fails to generate a preview (e.g., due to an invalid STL format), it will retry up to 10 times. Failed previews will be reported in the log file and noted in the checklist.

## Version History

- **v1.0.0 (2024-10-26)**: Initial release
    - Added automatic checklist generation with 3D previews.
    - Implemented color-coding based on file name patterns.
    - Provided Excel output with structured folder organization.

- **v2.0.0 (2024-10-27)**: Updates and improvements
    - Added support for entering the STL folder path via command line or manual input.
    - Improved logging: log messages are collected and written to a timestamped log file.
    - Created a "logs" directory for storing log files.
    - Changed to use a temporary directory for storing STL previews, which are deleted after execution.
    - Added progress bar using tqdm for better user feedback during STL file processing.
    - Added retries to create_3d_preview to make it more stable, as the backend pyglet can fail occasionally.

## Contributing

Contributions are welcome! If you have any suggestions or find bugs, please open an issue or submit a pull request.

## License

This project is licensed under the MIT License.
