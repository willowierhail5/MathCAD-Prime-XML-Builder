import zipfile
import io
from lxml import etree as ET
from sympy.parsing.sympy_parser import parse_expr
from sympy import simplify
from sympy import expand
from sympy import pretty
from sympy import symbols

from sympy import Add, Mul, Pow

input_file_path = "blank.mcdx"
output_file_path = "TestOutput.mcdx"

state = {"region_id": 0, "top": 128}  # the initial value of 'top'

d_ns = "http://schemas.mathsoft.com/worksheet50"
ml_ns = "http://schemas.mathsoft.com/math50"
ve_ns = "http://schemas.openxmlformats.org/markup-compatibility/2006"
r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ws_ns = d_ns
u_ns = "http://schemas.mathsoft.com/units10"
p_ns = "http://schemas.mathsoft.com/provenance10"

id_s = "{http://www.w3.org/XML/1998/namespace}space"
r_s = "{http://schemas.mathsoft.com/worksheet50}regions"
ap_s = "{" + ml_ns + "}apply"
df_s = "{" + ml_ns + "}define"


def create_scale(parent):
    return create_element(parent, "{" + ml_ns + "}scale")


def create_real(parent, value, preserve="preserve"):
    real_attrs = {id_s: preserve}
    return create_element(parent, "{" + ml_ns + "}real", real_attrs, str(value))


def create_id_with_contextual_label(
    parent, text, labels, preserve="preserve", label_is_contextual="true"
):
    id_attrs = {
        "labels": labels,
        id_s: preserve,
        "label-is-contextual": label_is_contextual,
    }
    return create_element(parent, "{" + ml_ns + "}id", id_attrs, text)


def create_region(region_id, width, height, top, left):
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
    return ET.SubElement(parent, "math")


def create_eval(parent):
    return ET.SubElement(parent, "{" + ml_ns + "}eval")


def create_element(parent, name, attrs={}, text=None):
    element = ET.SubElement(parent, name, attrs)
    if text:
        element.text = text
    return element


def create_def(parent):
    return create_element(parent, df_s)


def create_id(parent, name, labels, preserve="preserve"):
    id_attrs = {"labels": labels, id_s: preserve}
    id_element = create_element(parent, "{" + ml_ns + "}id", id_attrs, name)
    return id_element


def create_matrix(parent, rows, cols):
    matrix_attrs = {"rows": str(rows), "cols": str(cols)}
    return create_element(parent, "{" + ml_ns + "}matrix", matrix_attrs)


def create_real(parent, value):
    return create_element(parent, "{" + ml_ns + "}real", {}, str(value))


def create_eval(parent):
    return create_element(parent, "{" + ml_ns + "}eval")


def create_placeholder(parent):
    return create_element(parent, "{" + ml_ns + "}placeholder")


def create_string(parent, text, preserve="preserve"):
    str_attrs = {id_s: preserve}
    return create_element(parent, "{" + ml_ns + "}str", str_attrs, text)


def parse_xml(file):
    register_namespaces()
    parser = ET.XMLParser(remove_blank_text=True)
    tree = ET.parse(file, parser)
    return tree.getroot()


def write_data_to_zip(output_file_path, data):
    with zipfile.ZipFile(output_file_path, "w") as zip_out:
        for name, content in data.items():
            zip_out.writestr(name, content)


def register_namespaces():
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


def define_matrix(var_name, matrix):
    region_id = state["region_id"]
    top = state["top"]

    rows = len(matrix)
    cols = len(matrix[0]) if rows > 0 else 0

    region = create_region(region_id, "67.3", "380", top, "96.14")
    math = create_math(region)

    define = create_def(math)
    variable_id = create_id(define, var_name, "VARIABLE")

    matrix_element = create_matrix(define, rows, cols)

    for row in matrix:
        for element in row:
            create_real(matrix_element, element)

    state["region_id"] += 1
    state["top"] += 25.6 * 7

    return region


def evaluate_var(var_name):
    region_id = state["region_id"]
    top = state["top"]

    region = create_region(region_id, "74.7", "192.7", top, "192.3")
    math = create_math(region)

    eval = create_eval(math)

    id = create_id_with_contextual_label(eval, var_name, "VARIABLE")

    unitOverride = create_element(eval, "{" + ml_ns + "}unitOverride")
    placeholder = create_placeholder(unitOverride)

    state["region_id"] += 1
    state["top"] += 25.6

    return region


def create_variable(name, value, unit):
    region_id = state["region_id"]
    top = state["top"]

    region = create_region(region_id, "134.5", "25.6", top, "134.4")
    math = create_math(region)

    define = create_def(math)
    variable_id = create_id(define, name, "VARIABLE")

    if unit:  # if unit is not an empty string
        apply = create_element(define, "{" + ml_ns + "}apply")
        scale = create_scale(apply)
        real = create_real(apply, value)
        unit_id = create_id_with_contextual_label(apply, unit, "UNIT")
    else:  # if unit is an empty string
        real = create_real(define, value)

    state["region_id"] += 1
    state["top"] += 25.6

    return region


def add_variables(root, var_names, var_values, var_units):
    regions_tag = root.find("ws:regions", namespaces=root.nsmap)
    for name, value, unit in zip(var_names, var_values, var_units):
        regions_tag.append(create_variable(name, value, unit))


def add_matrices(root, var_names, matrices):
    regions_tag = root.find("ws:regions", namespaces=root.nsmap)
    for name, matrix in zip(var_names, matrices):
        regions_tag.append(define_matrix(name, matrix))


def add_operations(root, var_names_list, operations_list):
    regions_tag = root.find(r_s)
    for var_names, operations in zip(var_names_list, operations_list):
        regions_tag.append(create_operation(var_names, operations))


import sympy


def create_operation(root, parsed_expr):
    region_id = state["region_id"]
    top = state["top"]
    region = create_region(region_id, "216.4", "25.6", top, "172.8")
    math = create_math(region)
    eval = create_eval(math)

    op_map = {
        "Add": "{" + ml_ns + "}plus",
        "Mul": "{" + ml_ns + "}mult",
        "Pow": "{" + ml_ns + "}pow",
    }

    def create_var_node(parent, var_name):
        return create_id_with_contextual_label(parent, var_name, "VARIABLE")

    def create_operation_node(parent, operation, *operands):
        apply = ET.SubElement(parent, "{" + ml_ns + "}apply")
        ET.SubElement(apply, op_map[operation])
        for operand in operands:
            if (
                isinstance(operand, ET._Element)
                and operand.tag == "{" + ml_ns + "}apply"
            ):
                for child in list(
                    operand
                ):  # create a copy of the children list to avoid modification issues
                    apply.append(child)
            else:
                if operand is not None:
                    apply.append(operand)
        return apply

    def process_expression(parent, expr, is_outermost=True):
        if (
            isinstance(expr, Mul)
            and len(expr.args) == 2
            and isinstance(expr.args[1], Pow)
            and expr.args[1].args[1] == -1
        ):
            operand1 = process_expression(parent, expr.args[0], is_outermost=False)
            operand2 = process_expression(
                parent, expr.args[1].args[0], is_outermost=False
            )
            apply = ET.SubElement(parent, "{" + ml_ns + "}apply")
            ET.SubElement(apply, "{" + ml_ns + "}div")
            apply.append(operand1)
            apply.append(operand2)
        elif isinstance(expr, (Mul, Add, Pow)):
            operands = [
                process_expression(parent, arg, is_outermost=False) for arg in expr.args
            ]
            operation_node = create_operation_node(
                parent, expr.func.__name__, *operands
            )
            if is_outermost:
                parent.append(operation_node)
        elif isinstance(expr, sympy.core.symbol.Symbol):
            var_node = ET.SubElement(parent, "{" + ml_ns + "}id")
            var_node.set("labels", "VARIABLE")
            var_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            var_node.set("label-is-contextual", "true")
            var_node.text = str(expr)
            return var_node
        elif isinstance(expr, sympy.core.numbers.Integer):
            num_node = ET.SubElement(parent, "{" + ml_ns + "}real")
            num_node.text = str(expr)
            return num_node
        elif isinstance(expr, sympy.core.numbers.NegativeOne):
            neg_node = ET.SubElement(parent, "{" + ml_ns + "}neg")
            num_node = ET.SubElement(neg_node, "{" + ml_ns + "}real")
            num_node.text = str(abs(expr))  # abs() used to remove the negative sign
            return neg_node
        else:
            print("Unmatched expression type: ", type(expr))
            print("Expression: ", expr)

    process_expression(eval, parsed_expr)

    unitOverride = create_element(eval, "{" + ml_ns + "}unitOverride")
    create_placeholder(unitOverride)

    state["region_id"] += 1
    state["top"] += 25.6
    if region is not None:
        root.append(region)
    return region


def add_evals(root, var_names):
    regions_tag = root.find("ws:regions", namespaces=root.nsmap)
    for name in var_names:
        regions_tag.append(evaluate_var(name))


def add_excels(root, var_names, file_names, ranges):
    regions_tag = root.find(r_s)
    for var_name, file_name, range in zip(var_names, file_names, ranges):
        regions_tag.append(read_excel(var_name, file_name, range))


def read_excel(var_name, file_name, range):
    region_id = state["region_id"]
    top = state["top"]

    region = create_region(region_id, "489.247", "25.6", top, "230.403")
    math = create_math(region)

    define = create_def(math)
    variable_id = create_id(define, var_name, "VARIABLE")

    apply = create_element(define, "{" + ml_ns + "}apply")
    function_id = create_id(apply, "READEXCEL", "FUNCTION")

    sequence = create_element(apply, "{" + ml_ns + "}sequence")

    create_string(sequence, file_name)
    create_string(sequence, range)

    state["region_id"] += 1
    state["top"] += 25.6

    return region


def get_max_region_id_from_root(root):
    max_region_id = max(
        int(region.get("region-id"))
        for region in root.findall(".//{http://schemas.mathsoft.com/worksheet50}region")
    )
    return max_region_id


def add_operations(root, parsed_expressions):
    regions = root.find(r_s)
    for parsed_expr in parsed_expressions:
        create_operation(regions, parsed_expr)


def read_and_modify_zip(
    input_file_path,
    var_names,
    var_values,
    var_units,
    matrix_names,
    matrices,
    excel_var_names,  # new parameter
    excel_file_names,  # new parameter
    excel_ranges,  # new parameter
    output_file_path,  # new parameter
):
    with zipfile.ZipFile(input_file_path, "r") as myzip:
        with zipfile.ZipFile(
            output_file_path, "w"
        ) as myzip_out:  # open output zip file
            for filename in myzip.namelist():
                if filename == "mathcad/result.xml":
                    continue
                elif filename == "mathcad/worksheet.xml":
                    xml_data = myzip.read(filename)
                    root = parse_xml(io.BytesIO(xml_data))
                    state["region_id"] = get_max_region_id_from_root(root) + 1

                    add_variables(root, var_names, var_values, var_units)
                    # add_matrices(root, matrix_names, matrices)
                    # add_evals(root, var_names)

                    # get the <regions> element
                    regions = root.find(r_s)

                    parsed_expr = parse_expr("var12*var23/var34", evaluate=False)

                    add_operations(root, [parsed_expr])
                    add_excels(root, excel_var_names, excel_file_names, excel_ranges)
                    add_evals(root, excel_var_names)
                    # parsed_expr = parse_expr("excelVar1*var12", evaluate=False)
                    # create_operation(regions, parsed_expr)
                    # write modified XML data to output zip file
                    xmlstr = ET.tostring(root).decode()
                    myzip_out.writestr(filename, xmlstr)
                else:
                    # copy other files to output zip file
                    myzip_out.writestr(filename, myzip.read(filename))


def main():
    var_names = ["var12", "var23", "var34"]
    var_values = [10, 20, 30]
    var_units = ["m", "m", "s"]
    matrix_names = ["mat1", "mat2"]
    matrices = [[[1, 2], [3, 4]], [[5, 6], [7, 8]]]
    excel_var_names = ["excelVar1"]
    excel_file_names = ["./mathcadExcel.xlsm"]
    excel_ranges = ["B1:B100"]

    read_and_modify_zip(
        input_file_path,
        var_names,
        var_values,
        var_units,
        matrix_names,
        matrices,
        excel_var_names,
        excel_file_names,
        excel_ranges,
        output_file_path,  # pass output_file_path to read_and_modify_zip
    )


if __name__ == "__main__":
    main()

# Note. Exponents are **. Not ^.
