import os
import win32com.client


def dwg_to_pdf(file_path, output_folder):
    acad = win32com.client.Dispatch(
        "AutoCAD.Application"
    )  # Create an AutoCAD application instance.

    if not acad:  # Check if the AutoCAD instance is live.
        print("Cannot connect to AutoCAD.")
        return

    doc = acad.Documents.Open(file_path)  # Open the .dwg file.

    if not doc:
        print(f"Cannot open document: {file_path}")
        return

    output_path = os.path.join(
        output_folder, f"{os.path.splitext(os.path.basename(file_path))[0]}.pdf"
    )

    # Use Publish method to save as PDF
    dwg_files = win32com.client.Dispatch("AutoCAD.DwgFileCollection")
    dwg_files.Add(file_path)

    # Create a new Publisher object
    publisher = acad.Publisher
    publisher.PublishDwgFiles(
        dwg_files, output_path, "DWG To PDF.pc3"
    )  # You might need to adjust the plot configuration according to your AutoCAD setup

    print(f"Exported: {output_path}")
    doc.Close()  # Close the document.


def convert_folder(folder_path, output_folder):
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".dwg"):
                dwg_to_pdf(os.path.join(root, file), output_folder)


convert_folder(
    r"C:\Users\rcarlson\Desktop\DWG Test", r"C:\Users\rcarlson\Desktop\DWG Test"
)
