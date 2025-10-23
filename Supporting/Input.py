from sympy.parsing.sympy_parser import parse_expr
from sympy.parsing.latex import parse_latex
from sympy import Symbol, symbols
import openpyxl
import re
import Supporting.XML


def LatexInput(
    latex_str: str,
    input_file_path: str = "mcdx/blank.mcdx",
    output_file_path: str = "mcdx/TestOutput3.mcdx",
    open_file: bool = False,
) -> str:
    """_summary_

    Args:
        latex_str (str): _description_
        input_file_path (str, optional): _description_. Defaults to "mcdx/blank.mcdx".
        output_file_path (str, optional): _description_. Defaults to "mcdx/TestOutput3.mcdx".

    Returns:
        str: _description_
    """
    (
        define_variables_data,
        operations_data,
        read_excel_data,
        write_excel_data,
    ) = parse_latex_input(latex_str)

    Supporting.XML.read_and_modify_zip(
        define_variables_data=define_variables_data,
        read_excel_data=read_excel_data,
        write_excel_data=write_excel_data,
        operations_data=operations_data,
        input_file_path=input_file_path,
        output_file_path=output_file_path,
    )

    if open_file:
        import subprocess
        import os

        subprocess.run(
            [
                # r"C:\Program Files\PTC\Mathcad Prime 11.0.1.0\MathcadPrime.exe",
                r"C:\Program Files\PTC\Mathcad Prime 9.0.0.0\MathcadPrime.exe",
                os.path.abspath(output_file_path),
            ]
        )

    return output_file_path


def parse_latex_input(latex_string):
    """Parsing a multiline latex string into the necessary data structures to match the excel input function. Needs additional tweaking and may change some things like removing the expressions. Currently, it assumes these:

        - Uses '=' as both define and equality.
        - If LHS is a single symbol (optionally with subscript/Greek), it's a definition.
        - Otherwise, it's an equation (evaluation).
    Args:
        latex_string (_type_): _description_

    Returns:
        _type_: _description_
    """
    define_variables_data = []
    operations_data = []
    read_excel_data = []
    write_excel_data = []

    equations = re.split(r"(?:\\\\|\n)+", latex_string.strip())

    top = 128.0
    line_spacing = 51.2

    for line in equations:
        line = line.strip()
        if not line:
            continue

        # Must contain '='
        if "=" not in line:
            continue

        lhs, rhs = map(str.strip, line.split("=", 1))

        # Try parsing both sides with Sympy
        try:
            lhs_expr = parse_latex(lhs)
            rhs_expr = parse_latex(rhs)
        except Exception as e:
            print(f"[WARN] Could not parse LaTeX line '{line}': {e}")
            continue

        # Detect if LHS is a simple symbol → definition
        if isinstance(lhs_expr, Symbol):
            var_name = str(lhs_expr)
            expr = rhs_expr
        else:
            var_name = None
            expr = lhs_expr - rhs_expr

        replacements = {}
        for s in expr.free_symbols:
            if "_" in s.name:
                # ? funky parsing with multi-character subscripts. It will recognize that it is one arg overall, but will insert a * between the subscript characters so just stripping that for now. Not sure if anything else is needed to clean up the parsing
                # new_name = s.name.replace("{", "").replace("*", "").replace("}", "")
                new_name = s.name.replace("*", "")
                replacements[s] = symbols(new_name)

        expr = expr.subs(replacements)

        operations_data.append(
            {"var_name": var_name, "expr": expr, "unit": None, "top": top}
        )

        top += line_spacing

    return define_variables_data, operations_data, read_excel_data, write_excel_data


def ExcelInput(
    filePath: str,
    input_file_path: str = "mcdx/blank.mcdx",
    output_file_path: str = "mcdx/TestOutput3.mcdx",
    open_file: bool = False,
) -> str:
    """_summary_

    Args:
        filePath (str): _description_
        input_file_path (str, optional): _description_. Defaults to "mcdx/blank.mcdx".
        output_file_path (str, optional): _description_. Defaults to "mcdx/TestOutput3.mcdx".

    Returns:
        str: _description_
    """
    (
        define_variables_data,
        operations_data,
        read_excel_data,
        write_excel_data,
    ) = parse_excel_input(filePath)

    Supporting.XML.read_and_modify_zip(
        define_variables_data=define_variables_data,
        read_excel_data=read_excel_data,
        write_excel_data=write_excel_data,
        operations_data=operations_data,
        input_file_path=input_file_path,
        output_file_path=output_file_path,
    )

    if open_file:
        import subprocess
        import os

        subprocess.run(
            [
                # r"C:\Program Files\PTC\Mathcad Prime 11.0.1.0\MathcadPrime.exe",
                r"C:\Program Files\PTC\Mathcad Prime 9.0.0.0\MathcadPrime.exe",
                os.path.abspath(output_file_path),
            ]
        )
    return output_file_path


def parse_excel_input(file_name: str):
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
        # replacing any blank space
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
                        unit_match = re.match(r"(.+?)(?:\[(.+)\])?$", rhs)
                        expr_part = unit_match.group(1).strip()
                        unit_part = (
                            unit_match.group(2).strip() if unit_match.group(2) else None
                        )

                        expr = parse_expr(
                            expr_part.strip().replace("=", ""), evaluate=False
                        )  #! NOTE: this code uses sympy parser as the background format. May want to stick with this moving forward, but need to look into making sure that it can handle unicode characters when written into xml somehow... Can we export sympy expressions to latex somehow to get the special characters? and then unicode from latex?
                        variable_data = {
                            "var_name": var,
                            "expr": expr,
                            "top": top,
                            "unit": unit_part,
                        }
                        operations_data.append(variable_data)
                    except Exception as e:
                        #! known exception - right now this is set up to parse python formatting, so no curly brackets for subscripts, ** instead of ^ for exponents, etc. Will eventually split things up to try to allow for different input formats like latex
                        print(f"Error parsing expression for {var}: {rhs} -> {e}")
                        ###############
        else:  # standalone expression
            try:
                unit_match = re.match(r"(.+?)(?:\[(.+)\])?$", data)
                expr_part = unit_match.group(1).strip()
                unit_part = unit_match.group(2).strip() if unit_match.group(2) else None

                op_data = {
                    "expr": parse_expr(
                        expr_part.strip().replace("=", ""),
                        evaluate=False,
                    ),
                    "top": top,
                    "unit": unit_part,
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


# TODO: may want to have this file ONLY convert inputs into a common format like RPN or MathML (with potentially even an intermediate format like sympy expressions taking advantage of the sympy parser, though need to keep track of custom characters somehow...). Can import from an excel file, latex file directly, potentially python somehow to convert a qmd or jupyter file to mathcad, or can call directly and pass in a string from another source.
# Then have another file that takes that common format and builds the XML structure specific to mathcad, Converter.py
# https://docs.sympy.org/latest/modules/parsing.html
# https://pypi.org/project/sympy-latex-parser/ or https://github.com/purdue-tlt/latex2sympy

# https://github.com/sbrisard/pymathml a converter for python to mathml
# https://github.com/dhmath/RPN/tree/master a converter for python formatted string (i.e., ^ is **) to RPN (this is from the below reddit post and not a formal library)
# https://www.reddit.com/r/CritiqueMyCode/comments/30spg2/python_convert_a_mathematical_expression_into/
# https://csis.pace.edu/~wolf/CS122/infix-postfix.htm
# https://stackoverflow.com/questions/26191707/transform-an-algebraic-expression-with-brackets-into-rpn-reverse-polish-notatio


def main():
    print("Hello world!")

    from sympy import pprint, symbols

    from sympy.parsing.latex import parse_latex
    import unicodeit
    import re

    # to avoid converting math expressions, only want to convert letters (expand to other things)
    [
        unicodeit.data.REPLACEMENTS.remove(x)
        for x in unicodeit.data.REPLACEMENTS
        if x[0] == r"\sqrt"
    ]

    # note that if I move forward with using sympy to keep track of expressions, I need to keep two versions of the equation, one that has not had unicode characters converted since sympy can't parse those, and one that has had them converted for display purposes
    # expr = parse_latex(unicodeit.replace(r"\frac {1 + \sqrt {\alpha}} {b}"))
    latex = r"I_{e} = \left(\frac{M_{cr}}{M_{a}}\right)^{3} I_{g} + \left[1 - \left(\frac{M_{cr}}{M_{a}}\right)^{3}\right] I_{cr}"
    latex_clean = re.sub(r"([A-Za-z])_\{([A-Za-z0-9]+)\}", r"\1_\2", latex)
    expr = parse_latex(latex)
    pprint(expr)
    print(rf"{expr}")

    replacements = {}
    for s in expr.free_symbols:
        if "_" in s.name:
            # ? funky parsing with multi-character subscripts. It will recognize that it is one arg overall, but will insert a * between the subscript characters so just stripping that for now. Not sure if anything else is needed to clean up the parsing
            # new_name = s.name.replace("{", "").replace("*", "").replace("}", "")
            new_name = s.name.replace("*", "")
            replacements[s] = symbols(new_name)

    expr = expr.subs(replacements)
    pprint(expr)
    print(expr.args)

    # maybe try except to see if can parse with traditional parser or if latex parser is needed
    # from sympy.parsing.sympy_parser import parse_expr
    # expr = parse_expr(r"\frac {1 + \sqrt {\alpha}} {b}")
    # pprint(expr)
    # print(rf"{expr}")
    # https://github.com/svenkreiss/unicodeit python library to convert latex to unicdode


if __name__ == "__main__":
    main()
