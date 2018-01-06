
from PIL import Image # type: ignore
from docx import Document # type: ignore
from docx.shared import Cm, Pt # type: ignore
from docx.enum.dml import MSO_THEME_COLOR_INDEX as MSO_THEME_COLOR # type: ignore
from docx.enum.section import WD_SECTION_START # type: ignore
from docx.enum.section import WD_ORIENTATION # type: ignore
from docx.oxml import OxmlElement # type: ignore
from docx.oxml.ns import qn # type: ignore
from typing import Callable, List, cast
from docx.text.paragraph import Paragraph # type: ignore
from docx.table import Table # type: ignore
import abc

ref_t = Callable[[str], str]

class IDocxParRenderer(object, metaclass=abc.ABCMeta):
    @abc.abstractmethod
    def render(self, p: Paragraph, doc: Document) -> None:
        raise NotImplementedError('users must define render to use this base class')

class DocxEntityParagraphNormal(IDocxParRenderer):
    def __init__(self, text: str, style: str=None) -> None:
        self.text, self.style = text, style
    def render(self, p: Paragraph, _: Document) -> None:
        p.add_run(self.text, self.style)

class DocxEntityParagraphItalic(IDocxParRenderer):
    def __init__(self, text: str, style: str=None) -> None:
        self.text, self.style = text, style
    def render(self, p: Paragraph, _: Document) -> None:
        p.add_run(self.text, self.style).italic = True

class DocxEntityParagraphBold(IDocxParRenderer):
    def __init__(self, text: str, style: str=None) -> None:
        self.text, self.style = text, style
    def render(self, p: Paragraph, _: Document) -> None:
        p.add_run(self.text, self.style).bold = True

class DocxEntityParagraphItem(IDocxParRenderer):
    def __init__(self, text: str, style: str="List Paragraph") -> None:
        self.text, self.style = text, style
    def render(self, _: Paragraph, doc: Document) -> None:
        doc.add_paragraph(self.text, self.style)

class DocxEntityParagraphRef(IDocxParRenderer):
    def __init__(self, ref: ref_t, key: str) -> None:
        self.key, self.ref = key, ref
    def render(self, p: Paragraph, _: Document) -> None:
        if self.ref:
            p.add_run(self.ref(self.key))

class IDocxEntityRenderer(object, metaclass=abc.ABCMeta):
    @abc.abstractmethod
    def render(self, doc: Document):
        raise NotImplementedError('users must define render to use this base class')

class DocxEntityParagraph(IDocxEntityRenderer):
    def __init__(self, text: str, style: str=None, _ref: ref_t=None) -> None:
        self.text, self.style = text, style
        self._ref = _ref
        self.subs = [] # type: List[IDocxParRenderer]
    def b(self, text: str, style: str=None) -> 'DocxEntityParagraph': # forward declaration
        self.subs.append(DocxEntityParagraphBold(text, style))
        return self
    def it(self, text: str, style: str=None) -> 'DocxEntityParagraph':
        self.subs.append(DocxEntityParagraphItalic(text, style))
        return self
    def n(self, text: str, style: str=None) -> 'DocxEntityParagraph':
        self.subs.append(DocxEntityParagraphNormal(text, style))
        return self
    def item(self, text: str="", style: str="Bullet0") -> 'DocxEntityParagraph':
        self.subs.append(DocxEntityParagraphItem(text, style))
        return self
    def ref(self, key: str) -> 'DocxEntityParagraph':
        self.subs.append(DocxEntityParagraphRef(self._ref, key))
        return self
    def render(self, doc: Document) -> None:
        p = doc.add_paragraph(self.text, self.style)
        for obj in self.subs:
            obj.render(p, doc)

class DocxEntityPicture(IDocxEntityRenderer):

    DEFAULT_WIDTH = Cm(17.8) # type: Cm
    DEFAULT_WIDTH_LANDSCAPE = Cm(25.5) # type: Cm
    MAX_HEIGHT = Cm(13) # type: Cm
    MAX_HEIGHT_PORTRAIT = Cm(19.5) # type: Cm 
    
    def __init__(self, filename: str, width: Cm = DEFAULT_WIDTH, height : Cm= None, caption: str=None) -> None:
        self.width, self.height, self.filename, self.caption = width, height, filename, caption
        self.key = "?"
    
    def render(self, doc: Document) -> None:
        imagePath = 'images/' + self.filename
        
        # autosize en fonction de l'orientation
        last_section = doc.sections[-1]
        if last_section.orientation == WD_ORIENTATION.LANDSCAPE:  # @UndefinedVariable pylint: disable=no-member
            width = self.DEFAULT_WIDTH_LANDSCAPE
        else:
            width = self.width
            
        # ajuster de sorte que l'image ne sorte pas de la page verticalement
        im = Image.open(imagePath)
        ratio = im.size[0] / float(im.size[1])
        computedHeight = width / ratio
        if last_section.orientation == WD_ORIENTATION.LANDSCAPE:  # @UndefinedVariable pylint: disable=no-member
            if computedHeight > self.MAX_HEIGHT or (self.height and self.height > self.MAX_HEIGHT):
                width = None
                height = self.MAX_HEIGHT
            else:
                height = self.height
        else:
            if computedHeight > int(self.MAX_HEIGHT_PORTRAIT) or (self.height and self.height > self.MAX_HEIGHT_PORTRAIT):
                width = None
                height = self.MAX_HEIGHT_PORTRAIT
            else:
                height = self.height
        
        doc.add_picture(imagePath, width, height)
        if self.caption:
            doc.add_paragraph(Docx.DEFAULT_FIGURE+self.key+" : "+ self.caption, Docx.DEFAULT_STYLE_LEGENDE_FIGURE)

class DocsEntityPageSection(IDocxEntityRenderer):
    def __init__(self, start_type = WD_SECTION_START.NEW_PAGE, orientation = WD_ORIENTATION.PORTRAIT):  # @UndefinedVariable pylint: disable=no-member
        self.start_type, self.orientation = start_type, orientation
    def render(self, doc: Document) -> None:
        last_section = doc.sections[-1]
        last_orientation = last_section.orientation
        s = doc.add_section(self.start_type)
        if not last_orientation == self.orientation: 
            new_width, new_height = s.page_height, s.page_width
            s.orientation, s.page_height, s.page_width = self.orientation, new_height, new_width

class DocsEntityPageBreak(IDocxEntityRenderer):
    def render(self, doc: Document) -> None:
        doc.add_page_break()

class DocxEntityDocumentTitle(IDocxEntityRenderer):
    def __init__(self, title, level = 0, style=None):
        self.title, self.level, self.style = title, level, style
    def render(self, doc):
        doc.add_heading(self.title, level=self.level, style=self.style)
        
class DocxEntitySection(DocxEntityDocumentTitle):
    def __init__(self, title, style=None):
        DocxEntityDocumentTitle.__init__(self, title, 1, style)
        
class DocxEntitySubSection(DocxEntityDocumentTitle):
    def __init__(self, title, style=None):
        DocxEntityDocumentTitle.__init__(self, title, 2, style)
        
class DocxEntitySubSubSection(DocxEntityDocumentTitle):
    def __init__(self, title, style=None):
        DocxEntityDocumentTitle.__init__(self, title, 3, style)

class TableConstructor(object, metaclass=abc.ABCMeta):
    @abc.abstractmethod
    def render(self, t: Table) -> None:
        raise NotImplementedError('users must define render to use this base class')

class DocxEntityTable((IDocxEntityRenderer)):
    def __init__(self, callback: TableConstructor, rows: int, cols: int, caption: str=None, style: str=None) -> None:
        self.callback, self.caption, self.rows, self.cols, self.style = callback, caption, rows, cols, style
        self.key = "?"
    def render(self, doc: Document) -> None:
        table = doc.add_table(self.rows, self.cols, self.style)
        self.callback.render(table)
        if self.caption:
            doc.add_paragraph(Docx.DEFAULT_TABLEAU+self.key+" : "+ self.caption, Docx.DEFAULT_STYLE_LEGENDE_TABLEAU)

class DocxEntityTOC(IDocxEntityRenderer):
    def __init__(self, titre:str, command:str) -> None:
        self.command, self.titre = command, titre
    def render(self, doc: Document) -> None:
        if self.titre:
            doc.add_paragraph(self.titre, "Illustration Index Heading")
        paragraph = doc.add_paragraph()
        run = paragraph.add_run()
        fldChar = OxmlElement('w:fldChar')  # creates a new element
        fldChar.set(qn('w:fldCharType'), 'begin')  # sets attribute on element
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')  # sets attribute on element
        instrText.text = self.command   # change 1-3 depending on heading levels you need
    
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'separate')
        fldChar3 = OxmlElement('w:t')
        fldChar3.text = "Right-click to update field."
        fldChar2.append(fldChar3)
    
        fldChar4 = OxmlElement('w:fldChar')
        fldChar4.set(qn('w:fldCharType'), 'end')
    
        r_element = run._r
        r_element.append(fldChar)
        r_element.append(instrText)
        r_element.append(fldChar2)
        r_element.append(fldChar4)
        p_element = paragraph._p

class DocxEntity(IDocxEntityRenderer):
    def __init__(self, filename: str) -> None:
        self.filename = filename
        self.entities = [] # type: List[IDocxEntityRenderer]
        
    def getFilename(self) -> str:
        return self.filename
    
    def append(self, obj: IDocxEntityRenderer) -> IDocxEntityRenderer:
        self.entities.append(obj)
        return obj
        
    def initialize(self, doc: Document) -> None:
        # caption defaults
        style = doc.styles['Caption']
        #font = style.font
        #font.name = 'Calibri'
        #font.size = Pt(10)
        #font.italic = True
        #font.bold = False
        #print(font.color.theme_color)
        #font.color.theme_color = MSO_THEME_COLOR.ACCENT_2  # @UndefinedVariable pylint: disable=no-member
        paragraph_format = style.paragraph_format
        paragraph_format.space_before = Pt(1)
        paragraph_format.space_after = Pt(12)
        # text defaults
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Calibri'
        font.size = Pt(11)
        font.italic = False
        font.bold = False
        # section defaults
        #s = doc.sections[-1]
        #s.left_margin = Cm(1)
        #s.right_margin = Cm(1)
        #s.top_margin = Cm(1)
        #s.bottom_margin = Cm(1)
        # heading 1 default
        style = doc.styles['Heading 1']
        paragraph_format = style.paragraph_format
        paragraph_format.space_before = Pt(6)
        paragraph_format.space_after = Pt(12)
        # heading 2 default
        style = doc.styles['Heading 2']
        paragraph_format = style.paragraph_format
        paragraph_format.space_before = Pt(6)
        paragraph_format.space_after = Pt(12)
        
    def render(self, doc: Document) -> None:
        self.initialize(doc)
        for obj in self.entities:
            obj.render(doc)

class Docx(object):
    
    # https://support.office.com/en-us/article/Field-codes-TOC-Table-of-Contents-field-1f538bc4-60e6-4854-9f64-67754d78d05c?ui=en-US&rs=en-US&ad=US
    # \\o... Builds a table of contents from paragraphs formatted with built-in heading styles. 
    # For example, { TOC \o "1-3" } lists only headings formatted with the styles Heading 1 through Heading 3. 
    # If no heading range is specified, all heading levels used in the document are listed. Enclose the 
    # range numbers in quotation marks.
    # \\h Inserts TOC entries as hyperlinks.
    # \\z Hides tab leader and page numbers in Web layout view.
    # \\u Builds a table of contents by using the applied paragraph TE000128012.
    DEFAULT_COMMAND = 'TOC \\o "1-3" \\h \\z \\u'
    DEFAULT_TOC_TITLE = "Table des matières"
    DEFAULT_STYLE_LEGENDE_TABLEAU = "Légende Tableau"
    DEFAULT_STYLE_LEGENDE_FIGURE = "Légende Figure"
    DEFAULT_FIGURE = "Figure "
    DEFAULT_TABLEAU = "Tableau "
    CHAPITRE_STYLE = "Chapitre"
    
    class ORIENT:
        LANDSCAPE = WD_ORIENTATION.LANDSCAPE  # @UndefinedVariable pylint: disable=no-member
        PORTRAIT = WD_ORIENTATION.PORTRAIT  # @UndefinedVariable pylint: disable=no-member
    
    class START_TYPE:
        NEW_PAGE = WD_SECTION_START.NEW_PAGE  # @UndefinedVariable pylint: disable=no-member
        CONTINUOUS = WD_SECTION_START.CONTINUOUS  # @UndefinedVariable pylint: disable=no-member
        ODD_PAGE = WD_SECTION_START.ODD_PAGE  # @UndefinedVariable pylint: disable=no-member
    
    def __init__(self, filename: str, _ref: Callable[[str], str] = None) -> None:
        self.filename = filename
        self._ref = _ref
        self.entity = DocxEntity(filename)

    def toc(self, titre: str=DEFAULT_TOC_TITLE, command: str= DEFAULT_COMMAND) -> None:
        self.entity.append(DocxEntityTOC(titre, command))
        
    def title(self, title: str, style: str=CHAPITRE_STYLE) -> DocxEntityDocumentTitle:
        return cast(DocxEntityDocumentTitle, self.entity.append(DocxEntityDocumentTitle(title, style=style)))
        
    def sec(self, title: str, style: str=None) -> DocxEntitySection:
        return cast(DocxEntitySection, self.entity.append(DocxEntitySection(title, style)))
        
    def subsec(self, title: str, style: str=None) -> DocxEntitySubSection:
        return cast(DocxEntitySubSection, self.entity.append(DocxEntitySubSection(title, style)))
        
    def subsubsec(self, title:str, style:str=None) -> DocxEntitySubSubSection:
        return cast(DocxEntitySubSubSection, self.entity.append(DocxEntitySubSubSection(title, style)))
        
    def pic(self, filename: str, width:Cm=DocxEntityPicture.DEFAULT_WIDTH, height:Cm=None, caption:str=None) -> DocxEntityPicture:
        """
        Utiliser plutôt add_pic de chapters.metier.pictures_tables
        """
        return cast(DocxEntityPicture, self.entity.append(DocxEntityPicture(filename, width, height, caption)))
        
    def par(self, text:str = "", style:str=None) -> DocxEntityParagraph:
        return cast(DocxEntityParagraph, self.entity.append(DocxEntityParagraph(text, style, self._ref)))
        
    def b(self, text:str = "", style:str=None, pretext:str="") -> DocxEntityParagraph:
        p = DocxEntityParagraph(pretext, style)
        return cast(DocxEntityParagraph, self.entity.append(p.b(text)))
        
    def pageBreak(self) -> DocsEntityPageBreak:
        return cast(DocsEntityPageBreak, self.entity.append(DocsEntityPageBreak()))
        
    def pageSection(self, start_type: 'START_TYPE' = START_TYPE.NEW_PAGE, orientation: 'ORIENT' = ORIENT.PORTRAIT) -> DocsEntityPageSection:
        return cast(DocsEntityPageSection, self.entity.append(DocsEntityPageSection(start_type, orientation)))
        
    def table(self, callback: TableConstructor, rows: int, cols: int, caption=None, style=None) -> DocxEntityTable:
        return cast(DocxEntityTable, self.entity.append(DocxEntityTable(callback, rows, cols, caption, style)))

    def save(self, target:str=None, pre: Callable[[Document],None]=None) -> None:
        if self.filename:
            d = Document(self.filename)
        else:
            d = Document()
        #for i in d.styles:
        #    print(i)
        if pre:
            pre(d)
        self.entity.render(d)
        if target:
            filename = target
        else:
            filename = self.filename
        d.save(filename)

