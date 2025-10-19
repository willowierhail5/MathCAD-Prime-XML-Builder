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
