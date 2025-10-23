def main():
    # import currentxlmParsingAndZip2

    # currentxlmParsingAndZip2.main()
    import Supporting.Input

    excel_path = "mcdx/excelInput copy.xlsm"
    # out = Supporting.Input.ExcelInput(excel_path)

    latex_input = r"I_{e} = \left(\frac{M_{cr}}{M_{a}}\right)^{3} I_{g} + \left[1 - \left(\frac{M_{cr}}{M_{a}}\right)^{3}\right] I_{cr}"
    latex_input = r"T_n = \frac{2A_o A_i f_{yi}}{s} \cot \theta"
    # v_s = \frac{A_v f_{yt}}{b_o s}
    # v_c = 3.5 \lambda \sqrt{f_c'}
    # v_{c} = \left(1.5 + \frac{\alpha_{s}d}{b_{o}}\right) \lambda \sqrt{f_{c}'} + 0.3 f_{pc} + \frac{V_{p}}{b_{o}d}
    out = Supporting.Input.LatexInput(latex_input)

    import subprocess
    import os

    try:
        subprocess.run(
            [
                # r"C:\Program Files\PTC\Mathcad Prime 11.0.1.0\MathcadPrime.exe",
                r"C:\Program Files\PTC\Mathcad Prime 9.0.0.0\MathcadPrime.exe",
                os.path.abspath(out),
            ]
        )
    except FileNotFoundError:
        subprocess.run(
            [
                r"C:\Program Files\PTC\Mathcad Prime 11.0.1.0\MathcadPrime.exe",
                # r"C:\Program Files\PTC\Mathcad Prime 9.0.0.0\MathcadPrime.exe",
                os.path.abspath(out),
            ]
        )


if __name__ == "__main__":
    main()

# TODO: first, just look into programmatically adding a placeholder equation that is an input so we can modify it easily through the api.
# TODO: then,
