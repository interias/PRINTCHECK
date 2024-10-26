import os
from pathlib import Path
import trimesh
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font

# Set the directory where the script is located
base_dir = Path(os.path.dirname(__file__))
stls_dir = base_dir / 'Stls'
output_dir = base_dir / 'stl_previews'
output_dir.mkdir(exist_ok=True)  # Create directory for previews if it doesn't exist

# Gather all STL files in the 'Stls' directory and its subdirectories
stl_files = []
for root, _, files in os.walk(stls_dir):
    for file in files:
        if file.endswith('.stl'):
            full_path = Path(root) / file
            stl_files.append(full_path)

print(f"Found {len(stl_files)} STL files in '{stls_dir}'.")

# Function to create a 3D preview with object-centered view and auto-scaling
def create_3d_preview(stl_file_path, save_path, image_size=(200, 200)):
    try:
        mesh = trimesh.load_mesh(stl_file_path)
        scene = mesh.scene()

        # Apply color based on file name condition
        if '[a]' in stl_file_path.name.lower():
            mesh.visual.vertex_colors = [180, 0, 0, 255]  # Mild red
            print(f"Rendering {stl_file_path.name} in red.")
        elif '[c]' in stl_file_path.name.lower():
            mesh.visual.vertex_colors = [255, 255, 255, 255]  # White
            print(f"Rendering {stl_file_path.name} in white.")
        else:
            mesh.visual.vertex_colors = [50, 50, 50, 255]  # Mild black
            print(f"Rendering {stl_file_path.name} in black.")

        # Calculate optimal camera distance to fit the entire model
        bounding_box = mesh.bounding_box_oriented.bounds
        scale = max(bounding_box[1] - bounding_box[0])  # Size of the model's longest axis
        camera_distance = scale * 2.5  # Adjust zoom level based on model size

        # Position the camera directly above the center of the model
        scene.set_camera(
            angles=[0.7, -0.3, 0.3],  # Front-top angle
            distance=camera_distance,
            center=mesh.centroid  # Center the view on the model
        )

        # Set ambient lighting for better contrast and depth
        scene.ambient_light = [0.3, 0.3, 0.3, 1.0]
        scene.background = [220, 220, 220, 255]  # Light gray background

        # Save the image with enhanced settings
        image = scene.save_image(resolution=image_size, visible=True)
        with open(save_path, 'wb') as f:
            f.write(image)
        print(f"Preview successfully created for {stl_file_path.name}")
        return True
    except Exception as e:
        print(f"Error creating preview for {stl_file_path}: {e}")
        return False

# Initialize Excel workbook
wb = Workbook()
ws = wb.active
ws.title = "STL Checklist"

# List to track missing previews
missing_previews = []

# Set headers
ws.append(["Filename", "Preview", "Checked and Available", "Not Needed"])

# Adjust column widths
ws.column_dimensions['A'].width = 30
ws.column_dimensions['B'].width = 30
ws.column_dimensions['C'].width = 20
ws.column_dimensions['D'].width = 15

# Make headers bold
for cell in ws[1]:
    cell.font = Font(bold=True)

# Track the current folder to add section headers
current_folder = None

# Process each STL file
for stl_path in sorted(stl_files):
    # Determine the folder structure relative to 'Stls' directory
    folder = stl_path.parent.relative_to(stls_dir)
    
    # Add a new header for each folder if it changes
    if folder != current_folder:
        # Insert a blank row for spacing before new folder
        ws.append([""])
        
        # Add folder header row
        folder_row = ws.max_row + 1
        ws.append([f"Folder: {folder}"])
        ws[f"A{folder_row}"].font = Font(bold=True)
        
        current_folder = folder

    # Generate preview image path
    preview_path = output_dir / f"{stl_path.name}.png"

    # Generate and save preview image
    if create_3d_preview(stl_path, preview_path):
        # Add filename and placeholders for checkboxes
        row = [
            stl_path.name,
            "",  # Placeholder for the image
            "",  # Placeholder for "Checked and Available"
            "",  # Placeholder for "Not Needed"
        ]
        ws.append(row)
        
        # Insert preview image into the Excel sheet
        img = Image(str(preview_path))
        img.width, img.height = 200, 200  # Resize image to 200x200 pixels
        img_cell = f"B{ws.max_row}"  # Column B, current row
        ws.add_image(img, img_cell)
        
        # Set the row height to 150 for rows with images
        ws.row_dimensions[ws.max_row].height = 150
    else:
        # Log missing preview and add entry to missing list with full path
        missing_info = f"{folder}/{stl_path.name}"
        print(f"Preview generation failed for {missing_info}")
        missing_previews.append(missing_info)

        # Add a row in the main table even if preview is missing
        ws.append([stl_path.name, "Preview Missing", "", ""])

# Insert missing previews information at the top of the sheet if any previews are missing
if missing_previews:
    ws.insert_rows(1)
    ws.insert_rows(1)
    ws["A1"] = "Warning: The following STL files failed to generate a preview:"
    ws["A1"].font = Font(bold=True, color="FF0000")
    for idx, missing_file in enumerate(missing_previews, start=2):
        ws[f"A{idx}"] = missing_file
    print("Warning: Some STL previews are missing. Details added to the Excel file.")

# Save the Excel file
excel_output_path = base_dir / "STL_Checklist_Structured.xlsx"
wb.save(excel_output_path)
print(f"Checklist successfully saved at: {excel_output_path}")
