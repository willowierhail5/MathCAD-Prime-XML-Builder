# We are not going to mess with flow documents
# implement "top" like on excel for all methods
# define requires units to function currently
# I can add operation var12 and it functions as add eval. no longer need add eval.
# add write excel functionality
# does not support complex assignment. like "var1:=excelVar1*force". Should not be too hard to do
# "  File "c:\Users\rcarlson\Desktop\Programs\ChimneyProgramFULLReferences\RCPython\currentxlmParsingAndZip2.py", line 400, in parse_excel_input
#     top = float(row[2])
#           ^^^^^^^^^^^^^
# ValueError: could not convert string to float: '=C2+1000'"
# sheet!page number can we do that in readexcel


# Arguments WRITEEXCEL("file", M, [rows, [cols]], [“range”])—Writes matrix M to the defined range within the Excel file you specified.
# rows or cols (optional) are either scalars specifying the first row or column of matrix M to write, or two-element vectors specifying the range of rows or columns (inclusive) of matrix M to write. If you omit this argument, WRITEEXCEL writes out every row and column of the matrix to the specified file.
# •“file” is a string containing the filename or the full pathname and filename. You must include the XLS or XLSX file extension, for example, heat.xlsx. Non-absolute pathnames are relative to the current working directory.
# •“range” (optional) is a string containing the cell range. If this argument is omitted, then READEXCEL reads all the data in "Sheet1" of the specified file and WRITEEXCEL writes all the data in the specified matrix to "Sheet1" of the specified file.
# You can specify range in one of the following forms:
# ◦"Sheet1!A1:B3" specifying the worksheet name, the top left cell, and the bottom right cell. "Sheet1!A1" means cell A1 of Sheet1, and "Sheet1!" means the entire worksheet.
# ◦"[1]A1:B3" specifying the worksheet number, the top left cell, and the bottom right cell. "[1]A1" means cell A1 of Sheet1, and "[1]" means the entire worksheet.
# •emptyfill (optional) is a string, scalar, or NaN (default), which is substituted for missing entries in the data file.
# •“blankrows” (optional) is a string that specifies what to do when encountering a blank line:
# ◦skip—Skips the current line.
# ◦read—(default) Reads the blank line.
# ◦stop—Stops the reading process.
# •M is a matrix of scalars. If M contains units, functions, or embedded matrices, PTC Mathcad cannot write the file.
# •rows or cols (optional) are either scalars specifying the first row or column of matrix M to write, or two-element vectors specifying the range of rows or columns (inclusive) of matrix M to write. If you omit this argument, WRITEEXCEL writes out every row and column of the matrix to the specified file.

# Import necessary libraries for XML processing, algebraic parsing, etc.
import zipfile
import io
from lxml import etree as ET
from sympy.parsing.sympy_parser import parse_expr

from sympy import Add, Mul, Pow, Integer, Symbol

import sys
import openpyxl
import re

# Define global constants and initial state.
input_file_path = "mcdx/blank.mcdx"  # a blank mathcad file to edit. Needs to exist before script is ran
output_file_path = "mcdx/TestOutput2.mcdx"  # output file path. Does not need to exist before script is ran

# TODO: organize functions into classes (like a class for reading in data so I can change the source later, a class for parsing xml, a class for writing xml, etc.)
# TODO: will help to really understand order of operations, try to reorganize the functions so that they are in order of execution (roughly)

# TODO: currently, top is read in from spreadsheet and left is not used. I probably want to just dynamically change top based on some testing. Set left to some fixed number so that they are all aligned. Also need to look into figuring out how to assign a unit to a calculated equation. Currently, a placeholder is used - can I just replace that with a unit item, or is it some custom thing?

state = {"region_id": 0, "top": 128}  # the initial value of 'top'

# Define namespaces for various XML schemas used.
d_ns = "http://schemas.mathsoft.com/worksheet50"
ml_ns = "http://schemas.mathsoft.com/math50"  # uses mathml namespace with a twist. It is similar to RPN in the background, (* 4 (*3 2)) to do 4 * (3 * 2)
ve_ns = "http://schemas.openxmlformats.org/markup-compatibility/2006"
r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ws_ns = d_ns
u_ns = "http://schemas.mathsoft.com/units10"
p_ns = "http://schemas.mathsoft.com/provenance10"

id_s = "{http://www.w3.org/XML/1998/namespace}space"
r_s = "{http://schemas.mathsoft.com/worksheet50}regions"
ap_s = "{" + ml_ns + "}apply"
df_s = "{" + ml_ns + "}define"


def create_id(parent, name, labels, preserve="preserve"):
    """Function to create an 'id' element in XML.

    The 'id' element will contain the provided name and labels. The attribute "preserve" is optional and defaults to "preserve".

    Args:
        parent (_type_): _description_
        name (_type_): _description_
        labels (_type_): _description_
        preserve (str, optional): _description_. Defaults to "preserve".

    Returns:
        _type_: _description_
    """
    id_attrs = {"labels": labels, id_s: preserve}
    id_element = create_element(parent, "{" + ml_ns + "}id", id_attrs, name)
    return id_element


def create_id_with_contextual_label(
    parent, text, labels, preserve="preserve", label_is_contextual="true"
):
    """Function to create an 'id' element in XML with contextual label attribute.

    This element is used to label math elements contextually. For instance, a variable with a specific unit.

    Args:
        parent (_type_): _description_
        text (_type_): _description_
        labels (_type_): _description_
        preserve (str, optional): _description_. Defaults to "preserve".
        label_is_contextual (str, optional): _description_. Defaults to "true".

    Returns:
        _type_: _description_
    """
    id_attrs = {
        "labels": labels,
        id_s: preserve,
        "label-is-contextual": label_is_contextual,
    }
    return create_element(parent, "{" + ml_ns + "}id", id_attrs, text)


def create_id_no_label(parent, text):
    """Function to create an 'id' element in XML without any label.

    This function generates a simple ID element with only the "preserve" attribute.

    Args:
        parent (_type_): _description_
        text (_type_): _description_

    Returns:
        _type_: _description_
    """
    id_attrs = {id_s: "preserve"}
    return create_element(parent, "{" + ml_ns + "}id", id_attrs, text)


def create_region(region_id, width, height, top, left):
    """Function to create a region element in XML. This element describes a specific region in the worksheet layout (given by its width, height, top, and left coordinates).

    Args:
        region_id (_type_): _description_
        width (_type_): _description_
        height (_type_): _description_
        top (_type_): _description_
        left (_type_): _description_

    Returns:
        _type_: _description_
    """
    return ET.Element(
        "region",
        {
            "region-id": str(region_id),
            "actualWidth": str(width),
            "actualHeight": str(height),
            "top": str(top),
            "left": str(left),
        },
    )


def create_scale(parent):
    """Function to create a 'scale' element in XML under the provided parent element.

    Args:
        parent (_type_): _description_

    Returns:
        _type_: _description_
    """
    return create_element(parent, "{" + ml_ns + "}scale")


def create_math(parent):
    """Function to create a 'math' element in XML under the provided parent element.

    Args:
        parent (_type_): _description_

    Returns:
        _type_: _description_
    """
    return ET.SubElement(parent, "math")


def create_eval(parent):
    """Function to create an 'eval' element in XML under the provided parent element.

    Args:
        parent (_type_): _description_

    Returns:
        _type_: _description_
    """
    return ET.SubElement(parent, "{" + ml_ns + "}eval")


def create_element(parent, name, attrs={}, text=None):
    """Generic function to create a new XML element with provided name, attributes, and text.

    Args:
        parent (_type_): _description_
        name (_type_): _description_
        attrs (dict, optional): _description_. Defaults to {}.
        text (_type_, optional): _description_. Defaults to None.

    Returns:
        _type_: _description_
    """
    element = ET.SubElement(parent, name, attrs)
    if text:
        element.text = text
    return element


def create_def(parent):
    """Function to create a 'define' element in XML under the provided parent element.

    Args:
        parent (_type_): _description_

    Returns:
        _type_: _description_
    """
    return create_element(parent, df_s)


def create_real(parent, value):
    """Function to create a 'real' element in XML to represent a real number.

    Args:
        parent (_type_): _description_
        value (_type_): _description_

    Returns:
        _type_: _description_
    """
    return create_element(parent, "{" + ml_ns + "}real", {}, str(value))


def create_placeholder(parent):
    """Function to create a 'placeholder' element in XML under the provided parent element.

    Args:
        parent (_type_): _description_

    Returns:
        _type_: _description_
    """
    return create_element(parent, "{" + ml_ns + "}placeholder")


def create_string(parent, text, preserve="preserve"):
    """Function to create a 'string' element in XML with provided text and the "preserve" attribute.

    Args:
        parent (_type_): _description_
        text (_type_): _description_
        preserve (str, optional): _description_. Defaults to "preserve".

    Returns:
        _type_: _description_
    """
    str_attrs = {id_s: preserve}
    return create_element(parent, "{" + ml_ns + "}str", str_attrs, text)


def parse_xml(file):
    """Function to parse an XML file and return its root element.

    Args:
        file (_type_): _description_

    Returns:
        _type_: _description_
    """
    register_namespaces()
    parser = ET.XMLParser(remove_blank_text=True)
    tree = ET.parse(file, parser)
    return tree.getroot()


def register_namespaces():
    """Function to register the necessary namespaces for XML parsing and generation."""
    ns_dict = {
        "default": d_ns,
        "ml": ml_ns,
        "ve": ve_ns,
        "r": r_ns,
        "ws": d_ns,
        "u": u_ns,
        "p": p_ns,
    }

    for key, value in ns_dict.items():
        ET.register_namespace(key, value)


def create_matrix_from_name_matrix(var_name, matrix):
    """Function to create a matrix from a given matrix name and a matrix of values. This function will generate an XML structure representing the matrix and will update the state's "region_id" and "top" values.

    Args:
        var_name (_type_): _description_
        matrix (_type_): _description_

    Returns:
        _type_: _description_
    """
    region_id = state["region_id"]
    top = state["top"]

    rows = len(matrix)
    cols = len(matrix[0]) if rows > 0 else 0

    region = create_region(region_id, "67.3", "380", top, "96.14")
    math = create_math(region)

    define = create_def(math)
    create_id(define, var_name, "VARIABLE")

    matrix_element = create_matrix(define, rows, cols)

    for row in matrix:
        for element in row:
            create_real(matrix_element, element)

    state["region_id"] += 1
    state["top"] += 25.6 * 7

    return region


def create_matrix(parent, rows, cols):
    """Function to create a matrix element in XML with the provided number of rows and columns.

    Args:
        parent (_type_): _description_
        rows (_type_): _description_
        cols (_type_): _description_

    Returns:
        _type_: _description_
    """
    matrix_attrs = {"rows": str(rows), "cols": str(cols)}
    return create_element(parent, "{" + ml_ns + "}matrix", matrix_attrs)


def create_variable(name, value, unit, top):
    """Function to create a variable in XML representation with its name, value, and unit. This function also updates the state's "region_id" and "top" values.

    Args:
        name (_type_): _description_
        value (_type_): _description_
        unit (_type_): _description_
        top (_type_): _description_

    Returns:
        _type_: _description_
    """
    region_id = state["region_id"]
    top = top

    region = create_region(region_id, "134.5", "25.6", top, "134.4")
    math = create_math(region)

    define = create_def(math)
    create_id(define, name, "VARIABLE")

    if unit:  # if unit is not an empty string
        apply = create_element(define, "{" + ml_ns + "}apply")
        create_scale(apply)
        create_real(apply, value)
        create_id_with_contextual_label(apply, unit, "UNIT")
    else:  # if unit is an empty string
        create_real(define, value)

    state["region_id"] += 1
    state["top"] += 25.6

    return region


def create_operation(root, item):
    """Function to create an operation in XML format. This function generates the XML structure for mathematical operations on the worksheet.

    Args:
        root (_type_): _description_
        item (_type_): _description_

    Returns:
        _type_: _description_
    """
    region_id = state["region_id"]
    top = item["top"]
    region = create_region(region_id, "216.4", "25.6", top, "172.8")
    math = create_math(region)
    # eval = create_eval(math) #below allows us to define variable with an expression if var_name exists in the definition
    if item.get("var_name"):
        define_node = create_def(math)
        create_id(define_node, item["var_name"], "VARIABLE")
        parent_node = create_eval(define_node)
    else:
        parent_node = create_eval(math)

    def create_var_node(var_name):
        var_node = ET.Element("{" + ml_ns + "}id")
        var_node.set("labels", "VARIABLE")
        var_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        var_node.set("label-is-contextual", "true")
        var_node.text = str(var_name)
        return var_node

    def create_real_node(value):
        real_node = ET.Element("{" + ml_ns + "}real")
        real_node.text = str(value)
        return real_node

    def process_expression(expr):
        if isinstance(expr, Add):
            return handle_addition(expr)
        elif isinstance(expr, Mul):
            return handle_multiplication(expr)
        elif isinstance(expr, Pow):
            return handle_power(expr)
        elif isinstance(expr, Symbol):
            return create_var_node(expr.name)
        elif isinstance(expr, Integer):
            return create_real_node(expr)

    def handle_sequence(terms):
        # Base case: if only one term, return its processed expression
        if len(terms) == 1:
            return process_expression(terms[0])

        # If first term is division, create division node and handle its operands
        if isinstance(terms[0], Pow) and terms[0].args[1] == -1:
            apply_node = ET.Element("{" + ml_ns + "}apply")
            ET.SubElement(apply_node, "{" + ml_ns + "}div")
            apply_node.append(process_expression(terms[0].args[0]))
            apply_node.append(handle_sequence(terms[1:]))
            return apply_node
        else:
            # If first term is multiplication or any other term,
            # create multiplication node and nest remaining terms
            apply_node = ET.Element("{" + ml_ns + "}apply")
            ET.SubElement(apply_node, "{" + ml_ns + "}mult")
            apply_node.append(process_expression(terms[0]))
            apply_node.append(handle_sequence(terms[1:]))
            return apply_node

    def handle_addition(expr):
        apply_node = ET.Element("{" + ml_ns + "}apply")
        ET.SubElement(apply_node, "{" + ml_ns + "}plus")
        for arg in expr.args:
            apply_node.append(process_expression(arg))
        return apply_node

    def handle_multiplication(expr):
        # Split the terms into multiplication and division terms
        multiplication_terms = [
            arg for arg in expr.args if not (isinstance(arg, Pow) and arg.args[1] == -1)
        ]
        division_terms = [
            arg.args[0]
            for arg in expr.args
            if isinstance(arg, Pow) and arg.args[1] == -1
        ]

        if division_terms:
            # Create primary division node if there are division terms
            apply_node = ET.Element("{" + ml_ns + "}apply")
            ET.SubElement(apply_node, "{" + ml_ns + "}div")
            apply_node.append(handle_sequence(multiplication_terms))
            apply_node.append(handle_sequence(division_terms))
            return apply_node
        else:
            # Otherwise, return only the multiplication terms
            return handle_sequence(multiplication_terms)

    def handle_power(expr):
        apply_node = ET.Element("{" + ml_ns + "}apply")
        ET.SubElement(apply_node, "{" + ml_ns + "}pow")
        apply_node.append(process_expression(expr.args[0]))
        apply_node.append(process_expression(expr.args[1]))
        return apply_node

    # eval.append(process_expression(item["expr"]))
    parent_node.append(process_expression(item["expr"]))

    unitOverride = ET.Element("{" + ml_ns + "}unitOverride")
    # eval.append(unitOverride)
    placeholder = ET.Element("{" + ml_ns + "}placeholder")
    unitOverride.append(placeholder)
    parent_node.append(unitOverride)

    state["region_id"] += 1
    state["top"] += 25.6
    root.append(region)
    return region


def create_write_excel(var_name, file_name, row_num, col_num, range, top, matrix):
    """Function to create an XML structure for writing data to an Excel file. It constructs the "write" operation specifying the file name, variable name, and range.

    Args:
        var_name (_type_): _description_
        file_name (_type_): _description_
        row_num (_type_): _description_
        col_num (_type_): _description_
        range (_type_): _description_
        top (_type_): _description_
        matrix (_type_): _description_

    Returns:
        _type_: _description_
    """
    region_id = state["region_id"]
    top = top

    region = create_region(region_id, "489.247", "25.6", top, "230.403")
    math = create_math(region)

    define = create_def(math)
    create_id(define, var_name, "VARIABLE")

    apply = create_element(define, "{" + ml_ns + "}apply")
    create_id(apply, "WRITEEXCEL", "FUNCTION")

    sequence = create_element(apply, "{" + ml_ns + "}sequence")

    create_string(sequence, file_name)

    # Create the contextual variable representation
    create_contextual_variable(sequence, matrix)

    create_real(sequence, row_num)
    create_real(sequence, col_num)

    create_string(sequence, range)

    state["region_id"] += 1
    state["top"] += 25.6

    return region


def create_contextual_variable(parent, var_name):
    """Function to create the 'contextual variable' XML element.

    Args:
        parent (_type_): _description_
        var_name (_type_): _description_

    Returns:
        _type_: _description_
    """
    return create_element(
        parent,
        "{" + ml_ns + "}id",
        {
            "labels": "VARIABLE",
            "label-is-contextual": "true",
            "{http://www.w3.org/XML/1998/namespace}space": "preserve",
        },
        var_name,
    )


def create_read_excel(var_name, file_name, range, top):
    """Function to create an XML structure for reading data from an Excel file. It constructs the "read" operation specifying the file name, variable name, and range.

    Args:
        var_name (_type_): _description_
        file_name (_type_): _description_
        range (_type_): _description_
        top (_type_): _description_

    Returns:
        _type_: _description_
    """
    region_id = state["region_id"]
    top = top

    region = create_region(region_id, "489.247", "25.6", top, "230.403")
    math = create_math(region)

    define = create_def(math)
    create_id(define, var_name, "VARIABLE")

    apply = create_element(define, "{" + ml_ns + "}apply")
    create_id(apply, "READEXCEL", "FUNCTION")

    sequence = create_element(apply, "{" + ml_ns + "}sequence")

    create_string(sequence, file_name)
    create_string(sequence, range)

    state["region_id"] += 1
    state["top"] += 25.6

    return region


def append_read_excels(root, read_excel_data):
    """Function to append multiple "read" operations to the XML root element based on the given read_excel_data.

    Args:
        root (_type_): _description_
        read_excel_data (_type_): _description_
    """
    regions_tag = root.find(r_s)
    for item in read_excel_data:
        if not item["var_name"] or not item["file_name"] or not item["range"]:
            continue
        regions_tag.append(
            create_read_excel(
                item["var_name"], item["file_name"], item["range"], item["top"]
            )
        )


# Function to append multiple "write" operations to the XML root element based on the given read_excel_data.


def append_write_excels(root, write_excel_data):
    """Function to append multiple "write" operations to the XML root element based on the given read_excel_data.

    Args:
        root (_type_): _description_
        write_excel_data (_type_): _description_
    """
    regions_tag = root.find(r_s)
    for item in write_excel_data:
        if not item["var_name"] or not item["file_name"] or not item["range"]:
            continue
        regions_tag.append(
            create_write_excel(
                item["var_name"],
                item["file_name"],
                item["row_num"],
                item["col_num"],
                item["range"],
                item["top"],
                item["matrix"],
            )
        )


def append_matrices(root, var_names, matrices):
    """Function to append matrix definitions to the XML root element based on the given variable names and matrices.

    Args:
        root (_type_): _description_
        var_names (_type_): _description_
        matrices (_type_): _description_
    """
    regions_tag = root.find("ws:regions", namespaces=root.nsmap)
    for name, matrix in zip(var_names, matrices):
        if not name or not matrix:
            continue  # skip this iteration if any value is empty or None

        regions_tag.append(create_matrix_from_name_matrix(name, matrix))


def append_variables(root, define_variables_data):
    """Function to append variable definitions to the XML root element based on the given data for defining variables.

    Args:
        root (_type_): _description_
        define_variables_data (_type_): _description_
    """
    regions_tag = root.find("ws:regions", namespaces=root.nsmap)
    for item in define_variables_data:
        if not item["var_name"] or not item["value"] or not item["unit"]:
            continue  # skip this iteration if any value is empty or None

        regions_tag.append(
            create_variable(item["var_name"], item["value"], item["unit"], item["top"])
        )


def append_operations(root, operations_data):
    """Function to append operations to the XML root element based on the given operations_data.

    Args:
        root (_type_): _description_
        operations_data (_type_): _description_
    """
    regions = root.find(r_s)
    for item in operations_data:
        if item["expr"]:
            create_operation(regions, item)


def get_max_region_id_from_root(root):
    """Function to find the maximum region ID present in the XML root element. This can be useful to avoid ID conflicts when adding new regions.

    Args:
        root (_type_): _description_

    Returns:
        _type_: _description_
    """
    max_region_id = max(
        int(region.get("region-id"))
        for region in root.findall(".//{http://schemas.mathsoft.com/worksheet50}region")
    )
    return max_region_id


# Function to write the given data into a zip archive at the specified output path.


def write_data_to_zip(output_file_path, data):
    """Function to write the given data into a zip archive at the specified output path.

    Args:
        output_file_path (_type_): _description_
        data (_type_): _description_
    """
    with zipfile.ZipFile(output_file_path, "w") as zip_out:
        for name, content in data.items():
            zip_out.writestr(name, content)


def parse_assignment(data, top):
    """Parsing Utilities

    This function parses assignment strings and extracts variable details. Given a string like "variable := 12.34 unit", it will extract the variable name, its assigned value, and its unit.

    Args:
        data (_type_): _description_
        top (_type_): _description_

    Returns:
        _type_: _description_
    """
    # Split data into variable and value
    var, value = data.split(":=")
    var = var.strip()

    # Check if there's a unit associated with the value
    match = re.match(r"([0-9.]+)\s*([a-zA-Z]*)", value)
    if match:
        numeric_value, unit = match.groups()
        return {
            "var_name": var,
            "value": float(numeric_value),
            "unit": unit,
            "top": top,
        }
    else:
        return {"var_name": var, "value": float(value), "unit": "", "top": top}


def parse_excel_input(file_name):
    """This function reads an Excel workbook to extract assignment and operation details. It recognizes patterns like READ and WRITE Excel operations, as well as simple assignments. Returns separate lists of data for different operations.

    Args:
        file_name (_type_): path to the Excel file

    Returns:
        _type_: _description_
    """
    workbook = openpyxl.load_workbook(file_name)
    sheet = workbook.active

    operations_data = []
    read_excel_data = []
    write_excel_data = []  # Initialize this list to store WRITEEXCEL matches

    READEXCEL_PATTERN = r"([a-zA-Z0-9_]+):=READEXCEL\(\"(.*\.xlsx|.*\.xlsm|.*\.csv)\",\"([A-Z0-9:]+)\"\)"
    WRITEEXCEL_PATTERN = r"([a-zA-Z0-9_]+):=WRITEEXCEL\(\"(.*\.xlsx|.*\.xlsm|.*\.csv)\",([a-zA-Z0-9_]+),(\d+),(\d+),\"([^\"]+)\"\)"

    define_variables_data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        # skipping any blank data rows
        if not row[1]:
            continue
        data = row[1]
        # replacing any blank space (TODO: .strip maybe instead?)
        data = data.replace(" ", "")
        # storing vertical location of element
        top = float(row[2])
        if data is None:
            break
        print(data)
        if ":=" in data:
            print("Matching data:", data)

            match_readexcel = re.match(READEXCEL_PATTERN, data)
            print("Match read? : " + str(match_readexcel))
            match_writeexcel = re.match(WRITEEXCEL_PATTERN, data)
            print("MATCH WRITE?: " + str(match_writeexcel))
            if match_writeexcel:
                print("DOING WRITE EXCEL LOGIC")
                (
                    var_name,
                    file_name,
                    matrix,
                    row_num,
                    col_num,
                    data_range,
                ) = match_writeexcel.groups()
                write_excel_data.append(
                    {
                        "var_name": var_name,
                        "file_name": file_name,
                        "matrix": matrix,  # This is simplistic. The matrix variable will need further parsing.
                        "row_num": row_num,
                        "col_num": col_num,
                        "range": data_range,  # This might also need further splitting
                        "top": top if top is not None else None,
                    }
                )
                print("END WRITE EXCEL LOGIC: " + str(write_excel_data))

            elif match_readexcel:
                var_name, file_name, data_range = match_readexcel.groups()
                read_excel_data.append(
                    {
                        "var_name": var_name,
                        "file_name": file_name,
                        "range": data_range,
                        "top": top if top is not None else None,
                    }
                )
            else:  # checking whether a basic variable assignment or defining variable with an expression
                # Split left/right side of assignment
                var, rhs = data.split(":=")
                var = var.strip()
                rhs = rhs.strip()

                # Check if RHS is a simple numeric assignment or an expression
                if re.match(r"^[0-9.]+([a-zA-Z]*)$", rhs):
                    # Simple numeric value (possibly with unit)
                    variable_data = parse_assignment(data, top)
                    define_variables_data.append(variable_data)
                else:
                    # Expression-based variable definition
                    try:
                        expr = parse_expr(
                            rhs, evaluate=False
                        )  #! NOTE: this code uses sympy parser as the background format. May want to stick with this moving forward, but need to look into making sure that it can handle unicode characters when written into xml somehow... Can we export sympy expressions to latex somehow to get the special characters? and then unicode from latex?
                        variable_data = {"var_name": var, "expr": expr, "top": top}
                        operations_data.append(variable_data)
                    except Exception as e:
                        print(f"Error parsing expression for {var}: {rhs} -> {e}")
                        ###############
        else:  # standalone expression  # TODO: figure out how to add units to the operation
            try:
                op_data = {
                    "expr": parse_expr(
                        data.strip().replace("=", ""),
                        evaluate=False,
                    ),
                    "top": top,
                }
                operations_data.append(op_data)
            except Exception as e:
                print(f"Error parsing expression: {data} -> {e}")

    print("RETURNING DEFINE: " + str(define_variables_data))
    print("END RETURNING DEFINE")

    print("WRITING DATA: " + str(write_excel_data))
    print("END WRITING DATA")

    print("READING DATA: " + str(read_excel_data))
    print("END READING DATA")

    print("EXPRESSIONS: " + str(operations_data))
    print("END EXPRESSIONS")
    return (
        define_variables_data,
        operations_data,
        read_excel_data,
        write_excel_data,  # return the write_excel_data as well
    )


def read_and_modify_zip(
    input_file_path,
    define_variables_data,
    matrix_names,
    matrices,
    read_excel_data,
    write_excel_data,
    output_file_path,
    operations_data,
):
    """ZIP File Handling

    Given the paths of an input and output ZIP file, this function extracts XML content from the input, modifies the XML based on provided operations and variables, and writes the modified XML to the output ZIP. It appends data like variables, matrices, operations, and Excel-related functions to the XML.

    Args:
        input_file_path (_type_): _description_
        define_variables_data (_type_): _description_
        matrix_names (_type_): _description_
        matrices (_type_): _description_
        read_excel_data (_type_): _description_
        write_excel_data (_type_): _description_
        output_file_path (_type_): _description_
        operations_data (_type_): _description_
    """
    with zipfile.ZipFile(input_file_path, "r") as myzip:
        with zipfile.ZipFile(
            output_file_path, "w"
        ) as myzip_out:  # open output zip file
            for filename in myzip.namelist():  # looping through xml files in zip
                if filename == "mathcad/result.xml":  # skipping result.xml file
                    continue
                elif filename == "mathcad/worksheet.xml":  # the file to write to
                    # TODO: organize this printed output better/set up proper logging that can be toggled on/off whether debugging or not (if debugging, log to console, if not, log to file with timestamps, etc.)
                    print(
                        input_file_path,
                        define_variables_data,
                        matrix_names,
                        matrices,
                        read_excel_data,
                        write_excel_data,
                        output_file_path,
                        operations_data,
                    )
                    xml_data = myzip.read(filename)
                    root = parse_xml(io.BytesIO(xml_data))
                    state["region_id"] = get_max_region_id_from_root(root) + 1
                    print(define_variables_data)
                    append_variables(root, define_variables_data)
                    append_matrices(root, matrix_names, matrices)

                    append_operations(root, operations_data)
                    append_read_excels(
                        root, read_excel_data
                    )  # passing the read_excel_data directly
                    append_write_excels(root, write_excel_data)
                    # using a list comprehension to extract var_names from read_excel_data
                    # write modified XML data to output zip file
                    xmlstr = ET.tostring(root).decode()
                    myzip_out.writestr(filename, xmlstr)
                else:
                    myzip_out.writestr(filename, myzip.read(filename))


def main():
    """Main Execution

    This is the main driver function for the script. It begins by reading and parsing an Excel workbook for operations and variable assignments. Following that, it reads and modifies the content of a ZIP file based on the parsed data.
    """
    excel_path = "mcdx/excelInput copy.xlsm"
    (
        define_variables_data,
        operations_data,
        read_excel_data,
        write_excel_data,
    ) = parse_excel_input(excel_path)
    # matrix_names = ["mat1", "mat2"]
    # matrices = [[[1, 2], [3, 4]], [[5, 6], [7, 8]]]
    matrix_names = []
    matrices = []

    read_and_modify_zip(
        input_file_path,
        define_variables_data,
        matrix_names,
        matrices,
        read_excel_data,
        write_excel_data,
        output_file_path,
        operations_data,
    )


# Entry Point

# If the script is executed as the main module, it calls the main function to start processing.


if __name__ == "__main__":
    # include some logic for debugging when this is the main file. When in production, this will be imported as a module and main will be called directly.
    main()

# Note. Exponents are **. Not ^.
