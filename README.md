# MathCAD-Prime-XML-Builder
Not currently finished. But it kind of works. currentxlmParsingAndZip2.py is my current code

Expanding on original code for my purposes. Read ^ for my comments on the items I want to try to take a crack at eventually. Readme will eventually be updated with this list and a list of what I have already modified in the code

Next steps/ideas (in no particular order):
- Create package from code to be imported and used in other places
- Parse and work with symbols (symbols for variables, think delta, or even the ' in f'c, and symbols for math, like sqrt doesn't seem to work)
    - convert inequalities somehow
    - For math symbols, create test mathcad sheet and see what the xml looks like. May need functions to create them like the functions for multiplcation, division, etc.
- Add logging
- Look into why placement of equations is off
- Look into adding units to variables within equations, not just overall result units (this needs a lot more thought into it with how to parse the input equation)
- Create various blank templates to use (with different margins, look into some headers, footers, etc.)
- Try to 'calculate' top based on previous regions or make it a fixed amount of cells (if we can somehow read the cell height from within the mcdx template file used)

- (Existing File Modifications) Look into basing top off of lowest region in the template file used if it has any regions in it already defined.
- Look into modifying existing variables/equations in a file instead of just appending new ones at the end (though this may be more of an API thing)
- Look into copying sheets and formatting programmatically. Say I create a calculation manually. But later, I want to make it a bit more formal to submit, so I want to add a typical header/footer and change the margins all at once.
    - Additionally, maybe I want to repeat this calculation x amount of times with different inputs, can I at least copy the sheet progammatically to avoid the hassle of setting everything up? And then can use the api to modify the inputs for each sheet copy if possible
    - Following a similar train of thought, maybe I have a file already set up, can I create a table of contents programmatically with links to text formatted as a header? Maybe an option for text styles to include in this. Or create a table of contents when the program copies sheets as an option for easier navigation

- (API) Look into setting equations as input/output here so that we can modify them later through the Mathcad API if needed. May be easier to use this tool for initial creation and then api stuff for extraction still if set up properly here

- (Unsure) Decide whether to continue with sympy parsing or try to move to something else in the backend like mathml... Biggest concern is handling different formats of input equations, there is a parse latex function in sympy, but unsure about how good it is and some additional formatting is still needed after parsing... Will likely just stick with sympy for now, but may revisit later
  - Sticking with it for now, but long-term would like to move away from it or at least explore my options a bit. It messes with the formatting too much for no apparent reason, adding * in subscripts if there are multiple characters (even though it still recognizes it as one large variable, not actual multiplication at least), rearranges the math, etc.  (though may want to try the latex2sympy libraries out there, currently only testing with the ootb latex parser that is noted as in experimental in the [documentation](https://docs.sympy.org/latest/modules/parsing.html#experimental-mathrm-latex-parsing) - on that note, there is also another backend that for the ootb one that it seems will be used in the future so may want to explore that)
  - This would require many changes to the code if we moved away from it so hopefully I can just find a better solution that still makes use of sympy