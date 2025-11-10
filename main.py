import Supporting.Input

from sympy.parsing.latex import parse_latex

# from sympy.parsing.latex.lark import parse_latex_lark
import latex2mathml.converter


def main():
    # import currentxlmParsingAndZip2

    # currentxlmParsingAndZip2.main()

    excel_path = "mcdx/excelInput copy.xlsm"
    # out = Supporting.Input.ExcelInput(excel_path)

    # latex_input = r"I_{e} = \left(\frac{M_{cr}}{M_{a}}\right)^{3} I_{g} + \left[1 - \left(\frac{M_{cr}}{M_{a}}\right)^{3}\right] I_{cr}"
    # latex_input = r"T_n = \frac{2A_o^3 A_i f_{yi}}{s} \cot \theta"
    # latex_input = r"v_s = \frac{A_v f_{yt}}{b_o s}"
    latex_input = r"v_c = 3.5 \lambda \sqrt{f_c'}"
    # latex_input = r"v_{c} = \left(1.5 + \frac{\alpha_{s}d}{b_{o}}\right) \lambda \sqrt{f_{c}'} + 0.3 f_{pc} + \frac{V_{p}}{b_{o}d}"

    antlr_output = parse_latex(latex_input)
    # lark_output = parse_latex_lark(latex_input) can't really utilize as it does not try to fix 'bad' latex, which may naturally occur with the ocr
    mathml_output = latex2mathml.converter.convert(latex_input)
    # need to test out purdue's fork of latex to sympy... but probably the same as the default essentially
    print(antlr_output)
    # print(lark_output)
    print(
        mathml_output
    )  # to test out mathml output - https://www-archive.mozilla.org/projects/mathml/demo/tester, https://www.tutorialspoint.com/compilers/online-mathml-editor.htm

    # I think I have decided that I would rather use mathml as the backend for this... And then modify the code to parse mathml into the necessary format for mathcad. This way I don't have to rely on sympy parsing the latex and potentially messing up the order of operations (e.g., converting a/b*c to a*c/b or something funky. Also things like subscripts). Because I don't need to be able to convert it to something that python could use, though I guess that could be a future extension, latex to python. But that is not the current goal of this.

    # import subprocess
    # import os

    # out = Supporting.Input.LatexInput(latex_input)
    # try:
    #     subprocess.run(
    #         [
    #             # r"C:\Program Files\PTC\Mathcad Prime 11.0.1.0\MathcadPrime.exe",
    #             r"C:\Program Files\PTC\Mathcad Prime 9.0.0.0\MathcadPrime.exe",
    #             os.path.abspath(out),
    #         ]
    #     )
    # except FileNotFoundError:
    #     subprocess.run(
    #         [
    #             r"C:\Program Files\PTC\Mathcad Prime 11.0.1.0\MathcadPrime.exe",
    #             # r"C:\Program Files\PTC\Mathcad Prime 9.0.0.0\MathcadPrime.exe",
    #             os.path.abspath(out),
    #         ]
    #     )


if __name__ == "__main__":
    main()

# TODO: first, just look into programmatically adding a placeholder equation that is an input so we can modify it easily through the api.
# TODO: then,
