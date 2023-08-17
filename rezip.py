import zipfile
import os
import shutil

# Specify paths and directories
source_file1 = ".\\MathcadUnzip\\mathcad\\xaml\\FlowDocument1.XamlPackage"
destination_directory = "Xaml"

# Ensure the destination directory exists
if not os.path.exists(destination_directory):
    os.makedirs(destination_directory)

# Extract all contents of FlowDocument1.XamlPackage
with zipfile.ZipFile(source_file1, "r") as zip_ref:
    zip_ref.extractall(destination_directory)

# Print the members to inspect
print(f"Members in {source_file1}:", os.listdir(destination_directory))

# Replace the Document.xaml file with your modified one
# Assuming you have a modified 'Document1.xaml' in the base directory
shutil.copy('Document1.xaml', os.path.join(destination_directory, 'Xaml/Document.xaml'))

# Re-zip the entire content into FlowDocument4.XamlPackage
destination_zip_path4 = ".\\MathcadUnzip\\mathcad\\xaml\\FlowDocument4.XamlPackage"
with zipfile.ZipFile(destination_zip_path4, "w") as new_zip:
    for foldername, subfolders, filenames in os.walk(destination_directory):
        for filename in filenames:
            file_path = os.path.join(foldername, filename)
            new_zip.write(file_path, arcname=file_path[len(destination_directory)+1:])

# Create the final .mcdx file
source_dir = "MathcadUnzip"
output_filename = "mcdx/TestOutputUnzip"

# Create a zip archive from the MathcadUnzip directory
shutil.make_archive(output_filename, 'zip', source_dir)

# Rename the file to have the .mcdx extension
os.rename(output_filename + ".zip", output_filename + ".mcdx")

print(f"Created the final .mcdx file at {output_filename + '.mcdx'}")
