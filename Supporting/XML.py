import zipfile
import io
from lxml import etree as ET

from sympy import Add, Mul, Pow, Integer, Symbol, Float

import re

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

ns_dict = {
    "default": d_ns,
    "ml": ml_ns,
    "ve": ve_ns,
    "r": r_ns,
    "ws": d_ns,
    "u": u_ns,
    "p": p_ns,
}
state = {
    "region_id": 0,
    "top": 128,
}  # the initial value of 'top' #! will want to play around with having this be variable somehow


def read_and_modify_zip(
    define_variables_data,
    read_excel_data,
    write_excel_data,
    operations_data,
    matrix_names=[],
    matrices=[],
    input_file_path="mcdx/blank.mcdx",
    output_file_path="mcdx/TestOutput3.mcdx",
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


def parse_xml(file: str) -> ET.Element:
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


def register_namespaces() -> None:
    """Function to register the necessary namespaces for XML parsing and generation."""
    for key, value in ns_dict.items():
        ET.register_namespace(key, value)


def get_max_region_id_from_root(root: ET.Element) -> int:
    """Function to find the maximum region ID present in the XML root element. This can be useful to avoid ID conflicts when adding new regions.

    Args:
        root (_type_): _description_

    Returns:
        _type_: _description_
    """
    try:
        max_region_id = max(
            int(region.get("region-id"))
            for region in root.findall(
                ".//{http://schemas.mathsoft.com/worksheet50}region"
            )
        )
        return max_region_id
    except ValueError:
        # value error is thrown if there are no existing regions. In that case, return -1 so that the first region id when it increments is 0 to kick things off
        return -1


def append_variables(root: ET.Element, define_variables_data) -> None:
    """Function to append variable definitions to the XML root element based on the given data for defining variables.

    Args:
        root (_type_): _description_
        define_variables_data (_type_): _description_
    """
    regions_tag = root.find("ws:regions", namespaces=root.nsmap)
    for item in define_variables_data:
        if not item["var_name"] or not item["value"]:
            continue  # skip this iteration if any value is empty or None

        regions_tag.append(
            create_variable(item["var_name"], item["value"], item["unit"], item["top"])
        )


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


def create_math(parent):
    """Function to create a 'math' element in XML under the provided parent element.

    Args:
        parent (_type_): _description_

    Returns:
        _type_: _description_
    """
    return ET.SubElement(parent, "math")


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
    id_element = ET.SubElement(parent, "{" + ml_ns + "}id", id_attrs)

    return add_subscript_to_id(id_element, name)


def add_subscript_to_id(id_element, name):
    """Helper function for creating id elements. If there is a subscript (checked via regex), adds the necessary XML elements in the Mathcad specific format.

    Args:
        id_element (_type_): _description_
        name (_type_): _description_
    """

    """
    <ml:define>
        <ml:id labels="VARIABLE" xml:space="preserve">
            <Span xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                xmlns:pw="clr-namespace:Ptc.Wpf;assembly=Ptc.Core">b<pw:Subscript>test2</pw:Subscript>
            </Span>
        </ml:id>
        <ml:real>6</ml:real>
    </ml:define>
    """

    # checking for a prime after the subscript and reorganizing string so it goes before
    if "_" in name and "'" == name[-1]:
        name_split = name.split("_")
        name = name_split[0] + "'_" + name_split[1][:-1]

    sub_match = re.match(r"(.+?)_(?:\{)?(.+?)(?:\})?$", name)
    if sub_match:
        base, sub = sub_match.groups()

        # Define namespace mappings for lxml
        nsmap = {
            None: "http://schemas.microsoft.com/winfx/2006/xaml/presentation",
            "pw": "clr-namespace:Ptc.Wpf;assembly=Ptc.Core",
        }

        # Create <Span> element with namespace map (not xmlns attributes or else it throws an error)
        span = ET.Element(
            "{http://schemas.microsoft.com/winfx/2006/xaml/presentation}Span",
            nsmap=nsmap,
        )
        span.text = base

        # Add the <pw:Subscript> element under this namespace
        pw_ns = "clr-namespace:Ptc.Wpf;assembly=Ptc.Core"
        pw_sub = ET.Element(f"{{{pw_ns}}}Subscript")
        pw_sub.text = sub
        span.append(pw_sub)
        id_element.append(span)
    else:
        id_element.text = name
    return id_element


def create_def(parent):
    """Function to create a 'define' element in XML under the provided parent element.

    Args:
        parent (_type_): _description_

    Returns:
        _type_: _description_
    """
    return create_element(parent, df_s)


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


def create_scale(parent):
    """Function to create a 'scale' element in XML under the provided parent element.

    Args:
        parent (_type_): _description_

    Returns:
        _type_: _description_
    """
    return create_element(parent, "{" + ml_ns + "}scale")


def create_real(parent, value):
    """Function to create a 'real' element in XML to represent a real number.

    Args:
        parent (_type_): _description_
        value (_type_): _description_

    Returns:
        _type_: _description_
    """
    return create_element(parent, "{" + ml_ns + "}real", {}, str(value))


def create_id_with_contextual_label(
    parent, text, labels, preserve="preserve", label_is_contextual="true"
):
    """Function to create an 'id' element in XML with contextual label attribute.

    This element is used to label math elements contextually. For instance, a variable with a specific unit.

    ONLY USED FOR UNITS

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

        return add_subscript_to_id(var_node, var_name)

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
        elif isinstance(expr, Integer) or isinstance(expr, Float):
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
    # placeholder = ET.Element("{" + ml_ns + "}placeholder")
    # unitOverride.append(placeholder)
    # parent_node.append(unitOverride)
    #! allows us to set final calculation result units if desired, otherwise placeholder (i.e., no units).
    if item.get("unit"):
        # If a unit is provided, create an <ml:id> instead of a placeholder
        create_id_with_contextual_label(unitOverride, item["unit"], "UNIT")
    else:
        placeholder = ET.Element("{" + ml_ns + "}placeholder")
        unitOverride.append(placeholder)
    parent_node.append(unitOverride)

    state["region_id"] += 1
    state["top"] += 25.6
    root.append(region)
    return region


def create_eval(parent):
    """Function to create an 'eval' element in XML under the provided parent element.

    Args:
        parent (_type_): _description_

    Returns:
        _type_: _description_
    """
    return ET.SubElement(parent, "{" + ml_ns + "}eval")


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


def main():
    print("Hello world!")


if __name__ == "__main__":
    main()
