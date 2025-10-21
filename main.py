def main():
    print("Hello from mathcad-prime-xml-builder!")
    import currentxlmParsingAndZip2

    currentxlmParsingAndZip2.main()
    # import latex2mathml.converter

    # mathml_output = latex2mathml.converter.convert("\\frac{a}{b} + c")
    # print(mathml_output)


if __name__ == "__main__":
    main()
    import subprocess

    subprocess.run(
        [
            r"C:\Program Files\PTC\Mathcad Prime 11.0.1.0\MathcadPrime.exe",
            r"C:\Users\Danny\source\repos\MathCAD-Prime-XML-Builder\mcdx\TestOutput2.mcdx",
        ]
    )

# TODO: first, just look into programmatically adding a placeholder equation that is an input so we can modify it easily through the api.
# TODO: then,

# hierarchytree of functions used in main.py
# ┌─ main()
# │
# │  ├─ parse_excel_input(file_name)
# │  │   ├─ parse_assignment()
# │  │   └─ sympy.parse_expr()
# │  │
# │  ├─ read_and_modify_zip(
# │  │       input_file_path,
# │  │       define_variables_data,
# │  │       matrix_names,
# │  │       matrices,
# │  │       read_excel_data,
# │  │       write_excel_data,
# │  │       output_file_path,
# │  │       operations_data)
# │  │
# │  │   ├─ zipfile.ZipFile(...)
# │  │   ├─ parse_xml()
# │  │   │   └─ register_namespaces()
# │  │   ├─ get_max_region_id_from_root()
# │  │   ├─ append_variables()
# │  │   │   └─ create_variable()
# │  │   │       ├─ create_region()
# │  │   │       ├─ create_math()
# │  │   │       ├─ create_def()
# │  │   │       ├─ create_id()
# │  │   │       │   └─ create_element()
# │  │   │       ├─ create_scale()
# │  │   │       ├─ create_real()
# │  │   │       └─ create_id_with_contextual_label()
# │  │   │           └─ create_element()
# │  │   │
# │  │   ├─ append_matrices()
# │  │   │   └─ create_matrix_from_name_matrix()
# │  │   │       ├─ create_region()
# │  │   │       ├─ create_math()
# │  │   │       ├─ create_def()
# │  │   │       ├─ create_id()
# │  │   │       ├─ create_matrix()
# │  │   │       └─ create_real()
# │  │   │
# │  │   ├─ append_operations()
# │  │   │   └─ create_operation()
# │  │   │       ├─ internal helpers: create_var_node(), create_real_node()
# │  │   │       ├─ handle_addition(), handle_multiplication(), handle_power()
# │  │   │       └─ recursive process_expression()
# │  │   │
# │  │   ├─ append_read_excels()
# │  │   │   └─ create_read_excel()
# │  │   │       ├─ create_region()
# │  │   │       ├─ create_math()
# │  │   │       ├─ create_def()
# │  │   │       ├─ create_id()
# │  │   │       ├─ create_element()
# │  │   │       └─ create_string()
# │  │   │
# │  │   ├─ append_write_excels()
# │  │   │   └─ create_write_excel()
# │  │   │       ├─ create_region()
# │  │   │       ├─ create_math()
# │  │   │       ├─ create_def()
# │  │   │       ├─ create_id()
# │  │   │       ├─ create_element()
# │  │   │       ├─ create_string()
# │  │   │       ├─ create_contextual_variable()
# │  │   │       └─ create_real()
# │  │   │
# │  │   └─ write modified XML to new zip
# │  │
# │  └─ (end)
# │
# └─ if __name__ == "__main__": main()
