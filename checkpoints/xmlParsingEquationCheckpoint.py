from lxml import etree as ET
import zipfile
import io

input_file_path = "TestWorksheetForUnzipping2.mcdx"
output_file_path = "TestOutput.mcdx"
replacement_file_path = "mathcad_modified.xml"

state = {"region_id": 0, "top": 64}  # the initial value of 'top'


def register_namespaces():
    ns_dict = {
        "default": "http://schemas.mathsoft.com/worksheet50",
        "ml": "http://schemas.mathsoft.com/math50",
        "ve": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "ws": "http://schemas.mathsoft.com/worksheet50",
        "u": "http://schemas.mathsoft.com/units10",
        "p": "http://schemas.mathsoft.com/provenance10",
    }

    for key, value in ns_dict.items():
        ET.register_namespace(key, value)


def parse_xml(file):
    register_namespaces()
    parser = ET.XMLParser(remove_blank_text=True)
    tree = ET.parse(file, parser)
    return tree.getroot()


def print_defined_variables(root):
    namespaces = {"ml": "http://schemas.mathsoft.com/math50"}
    for region in root.findall(".//ml:define", namespaces):
        try:
            variable_name = region.find("ml:id", namespaces).text
            variable_value = region.find("ml:apply/ml:real", namespaces).text
            print(f"Variable Name: {variable_name}, Value: {variable_value}")
        except AttributeError:
            continue


def create_variable(name, value, unit):
    ws_namespace = "http://schemas.mathsoft.com/worksheet50"
    ml_namespace = "http://schemas.mathsoft.com/math50"
    region_id = state["region_id"]
    top = state["top"]

    region = ET.Element(
        "region",
        {
            "region-id": str(region_id),
            "actualWidth": "134.49333333333334",
            "actualHeight": "25.6",
            "top": str(top),
            "left": "134.40000000000003",
        },
    )
    math = ET.SubElement(region, "math")
    define = ET.SubElement(math, "{" + ml_namespace + "}define")
    variable_id = ET.SubElement(define, "{" + ml_namespace + "}id")
    variable_id.text = name
    variable_id.set("labels", "VARIABLE")
    variable_id.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    apply = ET.SubElement(define, "{" + ml_namespace + "}apply")
    scale = ET.SubElement(apply, "{" + ml_namespace + "}scale")
    real = ET.SubElement(apply, "{" + ml_namespace + "}real")
    real.text = str(value)
    unit_id = ET.SubElement(apply, "{" + ml_namespace + "}id")
    unit_id.text = unit
    unit_id.set("labels", "UNIT")
    unit_id.set("label-is-contextual", "true")
    unit_id.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    state["region_id"] += 1
    state["top"] += 25.6

    return region


def add_variables(root, var_names, var_values, var_units):
    regions_tag = root.find("{http://schemas.mathsoft.com/worksheet50}regions")
    for name, value, unit in zip(var_names, var_values, var_units):
        regions_tag.append(create_variable(name, value, unit))


def write_xml(root, filename):
    tree = ET.ElementTree(root)
    tree.write(filename, encoding="utf-8", xml_declaration=True, pretty_print=True)


def modify_mcdx_file():
    with zipfile.ZipFile(input_file_path, "r") as myzip:
        data = {}
        for filename in myzip.namelist():
            if filename == "mathcad/result.xml":
                continue
            elif filename == "mathcad/worksheet.xml":
                with open(replacement_file_path, "rb") as f:
                    data[filename] = f.read()
            else:
                with myzip.open(filename) as myfile:
                    data[filename] = myfile.read()
    with zipfile.ZipFile(output_file_path, "w") as zipf:
        for name, content in data.items():
            zipf.writestr(name, content)


# Update the main function to test the new create_operation function:
def create_operation(var_name1, var_name2, operation):
    ws_namespace = "http://schemas.mathsoft.com/worksheet50"
    ml_namespace = "http://schemas.mathsoft.com/math50"
    region_id = state["region_id"]
    top = state["top"]

    region = ET.Element(
        "region",
        {
            "region-id": str(region_id),
            "actualWidth": "216.41666666666669",
            "actualHeight": "25.6",
            "top": str(top),
            "left": "172.8",
        },
    )
    math = ET.SubElement(region, "math")
    eval = ET.SubElement(math, "{" + ml_namespace + "}eval")
    apply = ET.SubElement(eval, "{" + ml_namespace + "}apply")

    if operation == "add":
        operator = ET.SubElement(apply, "{" + ml_namespace + "}plus")
    elif operation == "multiply":
        operator = ET.SubElement(apply, "{" + ml_namespace + "}mult")
    elif operation == "divide":
        operator = ET.SubElement(apply, "{" + ml_namespace + "}div")
    else:
        raise ValueError(f"Unknown operation {operation}")

    id1 = ET.SubElement(apply, "{" + ml_namespace + "}id")
    id1.text = var_name1
    id1.set("labels", "VARIABLE")
    id1.set("label-is-contextual", "true")
    id1.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    id2 = ET.SubElement(apply, "{" + ml_namespace + "}id")
    id2.text = var_name2
    id2.set("labels", "VARIABLE")
    id2.set("label-is-contextual", "true")
    id2.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    unitOverride = ET.SubElement(eval, "{" + ml_namespace + "}unitOverride")
    placeholder = ET.SubElement(unitOverride, "{" + ml_namespace + "}placeholder")

    state["region_id"] += 1
    state["top"] += 25.6

    return region


def main():
    with zipfile.ZipFile(input_file_path, "r") as myzip:
        with myzip.open("mathcad/worksheet.xml") as myfile:
            root = parse_xml(io.BytesIO(myfile.read()))

    write_xml(root, "worksheet.xml")  # write the original worksheet.xml for reference

    print_defined_variables(root)

    variable_names = ["NewVar1", "NewVar2", "NewVar3", "NewVar4", "NewVar5", "NewVar6"]
    variable_values = [10, 20, 30, 40, 50, 60]
    variable_units = ["in", "ft", "cm", "in", "ft", "cm"]

    add_variables(root, variable_names, variable_values, variable_units)

    # Test the new create_operation function:
    regions_tag = root.find("{http://schemas.mathsoft.com/worksheet50}regions")
    regions_tag.append(create_operation("NewVar1", "NewVar2", "add"))
    regions_tag.append(create_operation("NewVar1", "NewVar2", "multiply"))
    regions_tag.append(create_operation("NewVar1", "NewVar2", "divide"))

    write_xml(root, replacement_file_path)  # write the modified xml file

    modify_mcdx_file()


if __name__ == "__main__":
    main()
