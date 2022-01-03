# docx2css

Given a Word (docx) document, _docx2css_ will produce a "clean" CSS stylesheet.


## Basic Usage

```python
import docx2css
stylesheet = docx2css.open_docx('path/to/word_file.docx') 
print(stylesheet.cssText)
```

## Supported Run Properties

The following Run properties (see sect. 17.3.2 of ooxml specs) are supported:

* b (Bold)
* bdr (Border) _[with some quirks]_
* caps (All Caps)
* color (Font Color) _[with some quirks]_
* dstrike (Double Strike) _[with some quirks]_
* emboss (Embossing)
* highlight (Text Highlighting) _[with some quirks]_
* i (Italics)
* imprint (Imprinting)
* kern (Kerning) _[with some quirks]_
* outline (Display Character Outline)
* position (Vertically Raised or Lowered Text) _[with some quirks]_
* rFonts (Run Fonts)  
* shadow (Shadow)  
* shd (Run Shading) _[with some quirks]_
* smallCaps (Small Caps)
* strike (Single Strikethrough) _[with some quirks]_
* sz (Non-Complex Script Font Size)
* u (Underline) _[with some quirks]_
* vanish (Hidden Text)
* vertAlign (Subscript/Superscript Text) _[with some quirks]_


## Unupported Run Properties

The following Run properties (see sect. 17.3.2 of ooxml specs) are **NOT** supported:

* bCs (Complex Script Bold)
* bdo (Bidirectional Override)
* cs (Complex Script Formatting)
* dir (Bidirectional Embedded)
* eastAsianLayout (East Asian Typography)
* effects (Animated Text Effects)
* em (Emphasis Mark)
* fitText (Manual Run Width) (It's not clear how Word handles this property anyway...)
* iCs (Complex Script Italics)
* lang (Languages for Run Content)
* noProof (Do Not Check Spelling or Grammar)  
* oMath (Office Open XML Math)
* rtl (Right To Left Text)
* snapToGrid (Use Document Grid Settings For Inter-Character Spacing)
* specVanish (Paragraph Mark Is Always Hidden)
* szCs (Complex Script Font Size)
* w (Expanded/Compressed Text)  
* webHidden (Web Hidden Text)

## Run Quirks and Limitations

* Strike, Double Strike and Underline docx properties are all translated to the same CSS properties:
  _text-decoration-line_ and _text-decoration-style_. 
A docx allows different line styles for strike and underline, CSS does not. In this case, the transposed CSS
  will use the strike or dstrike line style (that is, single or double) to the detriment of the underline.
  
* Not all underline styles are available in CSS.
  
### Character Borders
* Not all docx borders have an equivalent in CSS. Some borders styles are mapped to something that is relatively close.
* 3D box (the frame attribute) is ignored. In a docx, this attribute only applies to styles with two or more
lines of a different weight and CSS has no such styles.
  
### Fonts
* rFonts _hint_ attribute is not supported.
* When using the shade or tint attribute of a theme color on a font, the results computed are not exactly the
  same as those calculated by Word 365. However, the computed result should be close enough that the difference is
  barely noticeable.
* Font kerning does not differentiate on the basis of the font size,
despite ooxml specs providing it should.
* superscript/subscript and position use the same CSS property: _vertical-align_
If they are mixed in the same style, the value of _position_ will be used
  for vertical-align, but the font-size will be set to smaller just as
  for superscript/subscript.
  
### Highlighting and Shading

* Highlighting is supported even though the Word UI does not allow to a
highlight value in a character style.
* Support for shading is partial. Only fill color is recognized. Pattern
fill pattern and color are ignored.
* If both highlight and shading are in the same style, the highlight color
takes precedence.
  
## Supported Paragraph Properties

* bottom (Paragraph Border Below Identical Paragraphs)
* ind (Paragraph Indentation) _[with some quirks]_
* jc (Paragraph Alignment)
* keepLines (Keep All Lines On One Page)
* keepNext (Keep Paragraph With Next Paragraph)
* left (Left Paragraph Border)
* pageBreakBefore (Start Paragraph on Next Page)
* right (Right Paragraph Border)
* shd (Paragraph Shading)
* spacing (Spacing Between Lines and Above/Below Paragraph) _[with some quirks]_
* top (Paragraph Border Above Identical Paragraphs)
* widowControl (Allow First/Last Line to Display on a Separate Page)

## Unsupported Paragraph Properties

* adjustRightInd (Automatically Adjust Right Indent When Using Document Grid)
* autoSpaceDE (Automatically Adjust Spacing of Latin and East Asian Text)
* autoSpaceDN (Automatically Adjust Spacing of East Asian Text and Numbers)
* bar (Paragraph Border Between Facing Pages)
* between (Paragraph Border Between Identical Paragraphs)
*  bidi (Right to Left Paragraph Layout)
* cnfStyle (Paragraph Conditional Formatting)
* contextualSpacing (Ignore Spacing Above and Below When Using Identical Styles)
* divId (Associated HTML div ID)
* framePr (Text Frame Properties)
* kinsoku (Use East Asian Typography Rules for First and Last Character per Line)
* overflowPunct (Allow Punctuation to Extend Past Text Extents)
* snapToGrid (Use Document Grid Settings for Inter-Line Paragraph Spacing)
* suppressLineNumbers (Suppress Line Numbers for Paragraph)
* suppressOverlap (Prevent Text Frames From Overlapping)
* tab (Custom Tab Stop)
* textAlignment (Vertical Character Alignment on Line)
* textboxTightWrap (Allow Surrounding Paragraphs to Tight Wrap to Text Box
Contents)
* textDirection (Paragraph Text Flow Direction)
* topLinePunct (Compress Punctuation at Start of a Line)
* wordWrap (Allow Line Breaking At Character Level)

## Paragraph Quirks and Limitations

* Indentation attributes _endChars_, _firstChars_, and _startChars_ are ignored.
* Paragraph spacing attributes _afterLines_ and _beforeLines_ are not supported.