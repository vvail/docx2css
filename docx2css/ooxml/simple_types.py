class TwoWayDict:

    @classmethod
    def css_value(cls, docx_value):
        return cls.docx2css.get(docx_value, None)

    @classmethod
    def docx_value(cls, css_value):
        for key, value in reversed(cls.docx2css.items()):
            if value == css_value:
                return key


class ST_Border(TwoWayDict):
    """Border types. Currently, only line borders are supported. Art borders
    might be supported in the future.
    """

    docx2css = {
        'dashDotStroked': 'dashed',
        'dashSmallGap': 'dashed',
        'dotDash': 'dotted',
        'dotDotDash': 'dotted',
        'doubleWave': 'solid',
        'nil': 'none',
        'thick': 'solid',
        'wave': 'solid',

        # Double borders
        'thickThinLargeGap': 'double',
        'thickThinMediumGap': 'double',
        'thickThinSmallGap': 'double',
        'thinThickLargeGap': 'double',
        'thinThickMediumGap': 'double',
        'thinThickSmallGap': 'double',

        # Triple borders
        'thinThickThinLargeGap': 'double',
        'thinThickThinMediumGap': 'double',
        'thinThickThinSmallGap': 'double',
        'triple': 'double',

        'dashed': 'dashed',
        'dotted': 'dotted',
        'double': 'double',
        'inset': 'inset',
        'none': 'none',
        'outset': 'outset',
        'single': 'solid',
        'threeDEmboss': 'ridge',
        'threeDEngrave': 'groove',
    }


class ST_FontFamily(TwoWayDict):
    docx2css = {
        'decorative': 'fantasy',
        'modern': 'monospace',
        'roman': 'serif',
        'script': 'cursive',
        'swiss': 'sans-serif',
    }


class ST_Jc(TwoWayDict):
    docx2css = {
        'left': 'start',
        'right': 'end',
        'both': 'justify',
        'center': 'center',
        'distribute': 'justify',
        'end': 'end',
        'start': 'start',
    }


class ST_NumberFormat(TwoWayDict):
    docx2css = {
        'ordinal': 'decimal',
        'cardinalText': 'decimal',
        'ordinalText': 'decimal',

        'none': 'none',
        'decimal': 'decimal',
        'decimalZero': 'decimal-leading-zero',
        'upperRoman': 'upper-roman',
        'lowerRoman': 'lower-roman',
        'upperLetter': 'upper-alpha',
        'lowerLetter': 'lower-alpha',
        'bullet': '',
    }


ST_Theme = {
    'majorAscii',
    'majorBidi',
    'majorEastAsia',
    'majorHAnsi',
    'minorAscii',
    'minorBidi',
    'minorEastAsia',
    'minorHAnsi',
}


class ST_Underline(TwoWayDict):

    docx2css = {
        'dashDotDotHeavy': 'dashed',
        'dashDotHeavy': 'dashed',
        'dashedHeavy': 'dashed',
        'dashLong': 'dashed',
        'dashLongHeavy': 'dashed',
        'dotDash': 'dotted',
        'dotDotDash': 'dotted',
        'dottedHeavy': 'dotted',
        'none': 'none',
        'thick': 'solid',
        'wavyDouble': 'wavy',
        'wavyHeavy': 'wavy',
        'words': 'solid',

        'dash': 'dashed',
        'dotted': 'dotted',
        'double': 'double',
        'single': 'solid',
        'wave': 'wavy',
    }
