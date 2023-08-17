# next step is in GPT.
from lxml import etree as ET
import zipfile
import io
from sympy.parsing.sympy_parser import parse_expr
from sympy import simplify
from sympy import expand
from sympy import pretty
from sympy import symbols

input_file_path = "testOperationStackParentheses.mcdx"
output_file_path = "TestOutput.mcdx"

state = {"region_id": 0, "top": 128}  # the initial value of 'top'


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


def define_matrix(var_name, matrix):
    ml_namespace = "http://schemas.mathsoft.com/math50"

    region_id = state["region_id"]
    top = state["top"]

    rows = len(matrix)
    cols = len(matrix[0]) if rows > 0 else 0

    region = ET.Element(
        "region",
        {
            "region-id": str(region_id),
            "actualWidth": "67.3",
            "actualHeight": "380",
            "top": str(top),
            "left": "96.14",
        },
    )
    math = ET.SubElement(region, "math")
    define = ET.SubElement(math, "{" + ml_namespace + "}define")

    id = ET.SubElement(define, "{" + ml_namespace + "}id")
    id.text = var_name
    id.set("labels", "VARIABLE")
    id.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    matrix_element = ET.SubElement(
        define, "{" + ml_namespace + "}matrix", {"rows": str(rows), "cols": str(cols)}
    )

    for row in matrix:
        for element in row:
            real = ET.SubElement(matrix_element, "{" + ml_namespace + "}real")
            real.text = str(element)

    state["region_id"] += 1
    state["top"] += 25.6 * 7

    return region


def evaluate_var(var_name):
    ml_namespace = "http://schemas.mathsoft.com/math50"

    region_id = state["region_id"]
    top = state["top"]

    region = ET.Element(
        "region",
        {
            "region-id": str(region_id),
            "actualWidth": "74.7",
            "actualHeight": "192.7",
            "top": str(top),
            "left": "192.3",
        },
    )
    math = ET.SubElement(region, "math")
    eval = ET.SubElement(math, "{" + ml_namespace + "}eval")

    id = ET.SubElement(eval, "{" + ml_namespace + "}id")
    id.text = var_name
    id.set("labels", "VARIABLE")
    id.set("label-is-contextual", "true")
    id.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    unitOverride = ET.SubElement(eval, "{" + ml_namespace + "}unitOverride")
    placeholder = ET.SubElement(unitOverride, "{" + ml_namespace + "}placeholder")

    state["region_id"] += 1
    state["top"] += 25.6

    return region


def create_variable(name, value, unit):
    ml_namespace = "http://schemas.mathsoft.com/math50"

    region_id = state["region_id"]
    top = state["top"]

    region = ET.Element(
        "region",
        {
            "region-id": str(region_id),
            "actualWidth": "134.5",
            "actualHeight": "25.6",
            "top": str(top),
            "left": "134.4",
        },
    )
    math = ET.SubElement(region, "math")
    define = ET.SubElement(math, "{" + ml_namespace + "}define")
    variable_id = ET.SubElement(define, "{" + ml_namespace + "}id")
    variable_id.text = name
    variable_id.set("labels", "VARIABLE")
    variable_id.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    if unit:  # if unit is not an empty string
        apply = ET.SubElement(define, "{" + ml_namespace + "}apply")
        scale = ET.SubElement(apply, "{" + ml_namespace + "}scale")
        real = ET.SubElement(apply, "{" + ml_namespace + "}real")
        real.text = str(value)
        unit_id = ET.SubElement(apply, "{" + ml_namespace + "}id")
        unit_id.text = unit
        unit_id.set("labels", "UNIT")
        unit_id.set("label-is-contextual", "true")
        unit_id.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    else:  # if unit is an empty string
        real = ET.SubElement(define, "{" + ml_namespace + "}real")
        real.text = str(value)

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
    regions_tag = root.find("{http://schemas.mathsoft.com/worksheet50}regions")
    for var_names, operations in zip(var_names_list, operations_list):
        regions_tag.append(create_operation(var_names, operations))


from sympy import Add, Mul, Pow


def create_operation(root, parsed_expr):
    ml_namespace = "http://schemas.mathsoft.com/math50"
    region_id = state["region_id"]
    top = state["top"]

    region = ET.SubElement(
        root,
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

    op_map = {
        "Add": "{" + ml_namespace + "}plus",
        "Mul": "{" + ml_namespace + "}mult",
        "Div": "{" + ml_namespace + "}div",
        "Pow": "{" + ml_namespace + "}pow",
    }

    def create_var_node(parent, var_name):
        id = ET.SubElement(parent, "{" + ml_namespace + "}id")
        id.text = var_name
        id.set("labels", "VARIABLE")
        id.set("label-is-contextual", "true")
        id.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        return id

    def create_operation_node(parent, operation, operand1, operand2):
        apply = ET.SubElement(parent, "{" + ml_namespace + "}apply")
        ET.SubElement(apply, op_map[operation])
        apply.append(operand1)
        apply.append(operand2)
        return apply

    def process_expression(parent, expr):
        if (
            expr.func == Mul
            and len(expr.args) == 2
            and isinstance(expr.args[1], Pow)
            and expr.args[1].args[1] == -1
        ):
            operand1 = process_expression(parent, expr.args[0])
            operand2 = process_expression(parent, expr.args[1].args[0])
            return create_operation_node(parent, "Div", operand1, operand2)
        elif expr.func in [Mul, Add, Pow]:
            operand1 = process_expression(parent, expr.args[0])
            operand2 = process_expression(parent, expr.args[1])
            return create_operation_node(parent, expr.func.__name__, operand1, operand2)
        else:
            return create_var_node(parent, str(expr))

    process_expression(eval, parsed_expr)

    unitOverride = ET.SubElement(eval, "{" + ml_namespace + "}unitOverride")
    placeholder = ET.SubElement(unitOverride, "{" + ml_namespace + "}placeholder")

    state["region_id"] += 1
    state["top"] += 25.6

    return region


def add_evals(root, var_names):
    regions_tag = root.find("ws:regions", namespaces=root.nsmap)
    for name in var_names:
        regions_tag.append(evaluate_var(name))


def add_excels(root, var_names, file_names, ranges):
    regions_tag = root.find("{http://schemas.mathsoft.com/worksheet50}regions")
    for var_name, file_name, range in zip(var_names, file_names, ranges):
        regions_tag.append(read_excel(var_name, file_name, range))


def read_excel(var_name, file_name, range):
    ws_namespace = "http://schemas.mathsoft.com/worksheet50"
    ml_namespace = "http://schemas.mathsoft.com/math50"
    region_id = state["region_id"]
    top = state["top"]

    region = ET.Element(
        "region",
        {
            "region-id": str(region_id),
            "actualWidth": "489.247",
            "actualHeight": "25.6",
            "top": str(top),
            "left": "230.403",
        },
    )
    math = ET.SubElement(region, "math")
    define = ET.SubElement(math, "{" + ml_namespace + "}define")

    id = ET.SubElement(define, "{" + ml_namespace + "}id")
    id.text = var_name
    id.set("labels", "VARIABLE")
    id.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    apply = ET.SubElement(define, "{" + ml_namespace + "}apply")
    id = ET.SubElement(apply, "{" + ml_namespace + "}id")
    id.text = "READEXCEL"
    id.set("labels", "FUNCTION")
    id.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    sequence = ET.SubElement(apply, "{" + ml_namespace + "}sequence")

    str1 = ET.SubElement(sequence, "{" + ml_namespace + "}str")
    str1.text = file_name
    str1.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    str2 = ET.SubElement(sequence, "{" + ml_namespace + "}str")
    str2.text = range
    str2.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    state["region_id"] += 1
    state["top"] += 25.6

    return region


def get_max_region_id_from_root(root):
    max_region_id = max(
        int(region.get("region-id"))
        for region in root.findall(".//{http://schemas.mathsoft.com/worksheet50}region")
    )
    return max_region_id


def write_data_to_zip(output_file_path, data):
    with zipfile.ZipFile(output_file_path, "w") as zip_out:
        for name, content in data.items():
            zip_out.writestr(name, content)


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

                    # get the <regions> element
                    regions = root.find(
                        "{http://schemas.mathsoft.com/worksheet50}regions"
                    )

                    parsed_expr = parse_expr(
                        "var12**var23 +var12/var23", evaluate=False
                    )

                    print(type(parsed_expr))
                    print(parsed_expr.func)
                    print(parsed_expr.args)

                    create_operation(regions, parsed_expr)  # pass the <regions> element
                    # add_excels(root, excel_var_names, excel_file_names, excel_ranges)
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
    var_units = ["", "", ""]
    matrix_names = ["mat1", "mat2"]
    matrices = [[[1, 2], [3, 4]], [[5, 6], [7, 8]]]
    excel_var_names = ["excelVar1"]
    excel_file_names = ["./copyRunningMathcad.xlsm"]
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
