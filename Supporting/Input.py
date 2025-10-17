def main():
    print("Hello world!")

    from sympy import pprint

    from sympy.parsing.latex import parse_latex
    import unicodeit

    # to avoid converting math expressions, only want to convert letters (expand to other things)
    [
        unicodeit.data.REPLACEMENTS.remove(x)
        for x in unicodeit.data.REPLACEMENTS
        if x[0] == r"\sqrt"
    ]

    # note that if I move forward with using sympy to keep track of expressions, I need to keep two versions of the equation, one that has not had unicode characters converted since sympy can't parse those, and one that has had them converted for display purposes
    # expr = parse_latex(unicodeit.replace(r"\frac {1 + \sqrt {\alpha}} {b}"))
    expr = parse_latex(r"\frac {1 + \sqrt {\alpha}} {b}")
    pprint(expr)
    print(rf"{expr}")

    # maybe try except to see if can parse with traditional parser or if latex parser is needed
    # from sympy.parsing.sympy_parser import parse_expr
    # expr = parse_expr(r"\frac {1 + \sqrt {\alpha}} {b}")
    # pprint(expr)
    # print(rf"{expr}")
    # https://github.com/svenkreiss/unicodeit python library to convert latex to unicdode


if __name__ == "__main__":
    main()


class Converter:
    """Class for converting to the necessary format for Mathcad

    This could be converting a standard algebraic equation or a latex formula
    """

    def __init__(self):
        self.stack = []

    def push(self, value):
        self.stack.append(value)

    def pop(self):
        if not self.stack:
            raise IndexError("pop from empty stack")
        return self.stack.pop()

    def peek(self):
        if not self.stack:
            raise IndexError("peek from empty stack")
        return self.stack[-1]

    def v2(self, stack):
        # this is just an rpn calculator in python
        # https://rosettacode.org/wiki/Parsing/RPN_calculator_algorithm#Python
        a = []
        b = {
            "+": lambda x, y: y + x,
            "-": lambda x, y: y - x,
            "*": lambda x, y: y * x,
            "/": lambda x, y: y / x,
            "^": lambda x, y: y**x,
        }
        stack = "3 4 2 * 1 5 - 2 3 ^ ^ / +"
        for c in stack.split():
            if c in b:
                a.append(b[c](a.pop(), a.pop()))
            else:
                a.append(float(c))
            print(c, a)
        return


# TODO: may want to have this file ONLY convert inputs into a common format like RPN or MathML (with potentially even an intermediate format like sympy expressions taking advantage of the sympy parser, though need to keep track of custom characters somehow...). Can import from an excel file, latex file directly, potentially python somehow to convert a qmd or jupyter file to mathcad, or can call directly and pass in a string from another source.
# Then have another file that takes that common format and builds the XML structure specific to mathcad, Converter.py
# https://docs.sympy.org/latest/modules/parsing.html
# https://pypi.org/project/sympy-latex-parser/ or https://github.com/purdue-tlt/latex2sympy

# https://github.com/sbrisard/pymathml a converter for python to mathml
# https://github.com/dhmath/RPN/tree/master a converter for python formatted string (i.e., ^ is **) to RPN (this is from the below reddit post and not a formal library)
# https://www.reddit.com/r/CritiqueMyCode/comments/30spg2/python_convert_a_mathematical_expression_into/
# https://csis.pace.edu/~wolf/CS122/infix-postfix.htm
# https://stackoverflow.com/questions/26191707/transform-an-algebraic-expression-with-brackets-into-rpn-reverse-polish-notatio
