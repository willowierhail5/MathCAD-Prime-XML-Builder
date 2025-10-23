def main():
    # import currentxlmParsingAndZip2

    # currentxlmParsingAndZip2.main()
    import Supporting.Input

    excel_path = "mcdx/excelInput copy.xlsm"
    # out = Supporting.Input.ExcelInput(excel_path)

    latex_input = r"I_{e} = \left(\frac{M_{cr}}{M_{a}}\right)^{3} I_{g} + \left[1 - \left(\frac{M_{cr}}{M_{a}}\right)^{3}\right] I_{cr}"
    out = Supporting.Input.LatexInput(latex_input)

    import subprocess
    import os

    subprocess.run(
        [
            r"C:\Program Files\PTC\Mathcad Prime 11.0.1.0\MathcadPrime.exe",
            os.path.abspath(out),
        ]
    )


if __name__ == "__main__":
    main()

# TODO: first, just look into programmatically adding a placeholder equation that is an input so we can modify it easily through the api.
# TODO: then,
