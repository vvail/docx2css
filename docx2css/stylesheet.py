from docx2css.api import BodyStyle, PageStyle, BaseStyle


class Stylesheet:
    page_style: PageStyle = PageStyle()
    body_style: BodyStyle = BodyStyle()

    def __init__(self):
        self.span_styles = dict()
        self.paragraph_styles = dict()
        self.table_styles = dict()

    def __get_type_dict(self, element):
        return {
            'span': self.span_styles,
            'p': self.paragraph_styles,
            'table': self.table_styles,
        }[element]

    def add_style(self, style: BaseStyle):
        key = style.qualified_id
        element, _, class_name = key.partition('.')
        type_dict = self.__get_type_dict(element)
        type_dict[class_name] = style
        if style.parent_id is not None:
            parent_key = style.parent_id
            parent = type_dict.get(parent_key, None)
            if parent:
                style.parent = parent
                parent.children.append(style)
