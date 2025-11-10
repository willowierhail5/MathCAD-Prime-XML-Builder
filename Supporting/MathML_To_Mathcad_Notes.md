# MathML to Mathcad XML Conversions

## Resources

Some links used for figuring out what MathML elements were or what certain unicode corresponded to:

- MathML elements - https://developer.mozilla.org/en-US/docs/Web/MathML/Reference/Element
- unicode - https://gist.github.com/ngs/2782436
- html character encoding - https://www.tutorialbrain.com/html_tutorial/html_character_encoding/
- html entities - https://tools.w3cub.com/html-entities

## Conversions

The typical format for the code blocks below will be `MathML` followed by the `Mathcad XML` equivalent.

### Equal Sign

```html
<mo>&#x0003D;</mo>

<ml:id labels="VARIABLE" xml:space="preserve">
    <Span xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:pw="clr-namespace:Ptc.Wpf;assembly=Ptc.Core">v<pw:Subscript>c</pw:Subscript>
    </Span>
</ml:id>
<ml:apply>
    (expression here)
</ml:apply> (the variable name followed by its definition)
```

### Subscripts

```html
<msub>
    <mi>v</mi>
    <mrow>
        <mi>c</mi>
    </mrow>
</msub>
<ml:id labels="VARIABLE" xml:space="preserve">
    <Span xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:pw="clr-namespace:Ptc.Wpf;assembly=Ptc.Core">v<pw:Subscript>c</pw:Subscript>
    </Span>
</ml:id>
```

### Superscripts

```html
Same as subscript, just msup instead for the tag
<ml:pow />
```

### Combined Subscripts and Superscripts

```html
<msubsup>
    <mi>A</mi>
    <mi>o</mi>
    <mn>3</mn>
</msubsup> Combines subscript with superscript (subscript goes before superscript like name implies)
In mathcad, there is no special operator for this. Just combine subscript and superscript as shown above
```

### Parentheses and Brackets

```html
left - <mo>&#x0002B;</mo> (or [)
right - <mo stretchy="true" fence="true" form="postfix">&#x00029;</mo> (or ])
<ml:parens>
</ml:parens>
```

### Generic Characters/Variables/Functions

```html
<mi>&#x003BB;</mi> - the html entity, need to check if we can feed this into mathcad directly or need to convert somehow to the character like below?
<ml:id xml:space="preserve">λ</ml:id>

Seemingly, this is how both handle special functions like trig functions. I guess Mathcad just has something that parses it out properly??
```

### Numbers

```html
<mn>1.5</mn>
<ml:real>1.5</ml:real>
```

## Math Expressions

### Addition

```html
<mo>&#x0002B;</mo> plus sign
<ml:plus />
```

### Subtraction

```html
<mo>&#x02212;</mo> minus sign
<ml:plus />
```

### Multiplication

```html
I do not think there is any explicit multiplication symbol, I think it just uses the mi for characters and places the • or x there so this may be a bit tricky. For now, ignore and assume any variables next to each other without anything are multiplied?
<ml:mult />

Altenative - <ml:scale /> if multiplying just one number and one variable (i.e., scaling the variable)
```

### Division

```html
<mfrac></mfrac>
<ml:div />
```

### Square/nth Root

```html
<msqrt></msqrt>

<ml:nthRoot />
<ml:placeholder /> - placeholder just means sqrt, can put a number here for nth root
<variable or number here>
```
