from sympy.parsing.sympy_parser import parse_expr
from sympy import symbols


def traverse_tree(expr, operations=None):
    print("Traversing tree: " + str(expr))
    if operations is None:
        operations = []
    if len(expr.args) >= 1:  # Change this line
        operations.append((expr.func, expr.args))
        for arg in expr.args:
            traverse_tree(arg, operations)
    return operations


var1, var2 = symbols("var1 var2")
tree = parse_expr("sin(cos(var1 + var2))", evaluate=False)

operations = traverse_tree(tree)
for operation in operations:
    print("Operation: ", operation[0])
    print("Operands: ", operation[1])


# This will give you a syntax tree for the expression
# tree = parse_expr("var1*var2+var1*(((var1+var2))*var1+var2)", evaluate=False)
# print(tree)
# print(simplify(tree))
# print(expand(tree))
# print(pretty(tree))
# print("End Test Prints")
# print(tree.args[1].args)


# You can also traverse the tree by using the args property, which returns a tuple of all arguments of an expression.
# For example, for the expression x + y, args would return (x, y). If an argument is also an expression, it will have its own args, allowing you to traverse the entire tree.

# For your purposes, you would need to create a recursive function that traverses the tree and generates the corresponding XML
# . You can differentiate between different types of nodes by using the func property, which gives the type of the operation, e.g.,
#  Add for addition, Mul for multiplication, etc. Note that in sympy, operands are always treated as arguments, so an addition x + y is represented as Add(x, y).
