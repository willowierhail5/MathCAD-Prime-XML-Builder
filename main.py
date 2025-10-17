def main():
    print("Hello from mathcad-prime-xml-builder!")
    import currentxlmParsingAndZip2

    # currentxlmParsingAndZip2.main()
    import latex2mathml.converter

    mathml_output = latex2mathml.converter.convert("\\frac{a}{b} + c")
    print(mathml_output)


if __name__ == "__main__":
    main()
