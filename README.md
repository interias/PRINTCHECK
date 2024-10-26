
# PRINTCHECK

**PRINTCHECK** is a tool designed to generate structured checklists for STL files used in 3D printing. It automates the process of creating a comprehensive checklist with previews for each STL file, organized by folder and subfolder structure. The goal is to help users ensure all necessary files are ready and reviewed before and after 3D printing, streamlining the workflow and minimizing errors.

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

1. **Clone the Repository:**
   ```bash
   git clone https://github.com/yourusername/PRINTCHECK.git
   cd PRINTCHECK
   ```

2. **Install Required Dependencies:**
   Make sure you have Python installed, then install the required packages:
   ```bash
   pip install trimesh openpyxl
   ```

## Usage

1. **Prepare Your Directory Structure:**
   - Place all STL files in a main directory (`Stls`) with any subfolders as needed.
   - The script will process all files in the specified main directory and its subfolders.

2. **Run the Script:**
   ```bash
   python printcheck.py
   ```
   - The script will search for STL files in the `Stls` folder and create previews for each file.
   - It will then generate an Excel checklist (`STL_Checklist_Structured.xlsx`) containing the previews, file names, and status columns.

3. **Review the Output:**
   - Open the generated Excel file to see the checklist, organized by folders.
   - Any STL files that failed to generate previews will be listed at the top, along with their folder paths.

## Configuration

### Folder Structure
The script expects the STL files to be located in a folder named `Stls` in the same directory as the script. The output will be saved to an `stl_previews` folder.

### File Naming Conventions
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

The script provides verbose output to help diagnose any issues during execution. If a file fails to generate a preview (e.g., due to an invalid STL format), it will be reported in the console, and the file will still appear in the checklist with "Preview Missing" noted.

## Contributing

Contributions are welcome! If you have any suggestions or find bugs, please open an issue or submit a pull request.

## License

This project is licensed under the BSD 3-Clause License.
