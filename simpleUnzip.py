import zipfile
import shutil
import os

path_to_zip_file = "mcdx/TestOutput.mcdx"

# path_to_zip_file = "TestStackedOperations.mcdx"
# path_to_zip_file = "TestInsert.mcdx"
# path_to_zip_file = "TestWorksheetForUnzipping2.mcdx"
if path_to_zip_file == "mcdx/TestOutput.mcdx":
    directory_path = "MathcadUnzipEdited"
else:
    directory_path = "MathcadUnzip"


# Iterate over the contents of the directory
for filename in os.listdir(directory_path):
    file_path = os.path.join(directory_path, filename)
    if os.path.isfile(file_path):
        os.remove(file_path)
    elif os.path.isdir(file_path):
        shutil.rmtree(file_path)
with zipfile.ZipFile(path_to_zip_file, "r") as zip_ref:
    zip_ref.extractall(directory_path)
import zipfile
import shutil
import os

path_to_zip_file = "mcdx/complexAssignment.mcdx"

# path_to_zip_file = "TestStackedOperations.mcdx"
# path_to_zip_file = "TestInsert.mcdx"
# path_to_zip_file = "TestWorksheetForUnzipping2.mcdx"
if path_to_zip_file == "mcdx/TestOutput.mcdx":
    directory_path = "MathcadUnzipEdited"
else:
    directory_path = "MathcadUnzip"


# Iterate over the contents of the directory
for filename in os.listdir(directory_path):
    file_path = os.path.join(directory_path, filename)
    if os.path.isfile(file_path):
        os.remove(file_path)
    elif os.path.isdir(file_path):
        shutil.rmtree(file_path)
with zipfile.ZipFile(path_to_zip_file, "r") as zip_ref:
    zip_ref.extractall(directory_path)


output_file_path = "mcdx/TestOutput.mcdx"
