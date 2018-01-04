from docx import Document
from docx.shared import Cm, Pt
from docx.enum.dml import MSO_THEME_COLOR_INDEX as MSO_THEME_COLOR
from docx.enum.section import WD_SECTION_START
from docx.enum.section import WD_ORIENTATION
from PIL import Image

class DocxEntityParagraphNormal:
    def __init__(self, text, style=None):
        self.text, self.style = text, style
    def render(self, p, doc):
        p.add_run(self.text, self.style)

class DocxEntityParagraphItalic:
    def __init__(self, text, style=None):
        self.text, self.style = text, style
    def render(self, p, doc):
        p.add_run(self.text, self.style).italic = True

class DocxEntityParagraphBold:
    def __init__(self, text, style=None):
        self.text, self.style = text, style
    def render(self, p, doc):
        p.add_run(self.text, self.style).bold = True

class DocxEntityParagraphItem:
    def __init__(self, text, style="List Paragraph"):
        self.text, self.style = text, style
    def render(self, p, doc):
        doc.add_paragraph(self.text, self.style)

class DocxEntityParagraphRef:
    def __init__(self, ref, key):
        self.key, self.ref = key, ref
    def render(self, p, doc):
        if self.ref:
            p.add_run(self.ref(self.key))

class DocxEntityParagraph:
    def __init__(self, text, style=None, _ref=None):
        self.text, self.style = text, style
        self._ref = _ref
        self.subs = []
    def b(self, text, style=None):
        self.subs.append(DocxEntityParagraphBold(text, style))
        return self
    def it(self, text, style=None):
        self.subs.append(DocxEntityParagraphItalic(text, style))
        return self
    def n(self, text, style=None):
        self.subs.append(DocxEntityParagraphNormal(text, style))
        return self
    def item(self, text="", style="List Paragraph"):
        self.subs.append(DocxEntityParagraphItem(text, style))
        return self
    def ref(self, key):
        self.subs.append(DocxEntityParagraphRef(self._ref, key))
        return self
    def render(self, doc):
        p = doc.add_paragraph(self.text, self.style)
        for obj in self.subs:
            obj.render(p, doc)

class DocxEntityPicture:
    
    DEFAULT_WIDTH = Cm(17.8)
    DEFAULT_WIDTH_LANDSCAPE = Cm(25.5)
    MAX_HEIGHT = Cm(13)
    MAX_HEIGHT_PORTRAIT = Cm(26)
    
    def __init__(self, filename, width = DEFAULT_WIDTH, height = None, caption=None):
        self.width, self.height, self.filename, self.caption = width, height, filename, caption
        self.imageRoot = 'images/'
        self.key = "?";
    
    def render(self, doc):
        imagePath = self.imageRoot + self.filename
        
        # autosize en fonction de l'orientation
        last_section = doc.sections[-1]
        if last_section.orientation == WD_ORIENTATION.LANDSCAPE:  # @UndefinedVariable
            print("Autoresize image in landscape : "+self.filename)
            width = self.DEFAULT_WIDTH_LANDSCAPE
        else:
            width = self.width
            
        # ajuster de sorte que l'image ne sorte pas de la page verticalement
        im = Image.open(imagePath)
        ratio = im.size[0] / float(im.size[1])
        computedHeight = width / ratio
        #print(ratio, computedHeight, self.MAX_HEIGHT)
        if last_section.orientation == WD_ORIENTATION.LANDSCAPE:  # @UndefinedVariable
            if computedHeight > self.MAX_HEIGHT or (self.height and self.height > self.MAX_HEIGHT):
                print("debordement en hauteur : "+ self.filename)
                width = None
                height = self.MAX_HEIGHT
            else:
                height = self.height
        else:
            if computedHeight > int(self.MAX_HEIGHT_PORTRAIT) or (self.height and self.height > self.MAX_HEIGHT_PORTRAIT):
                print("debordement en hauteur")
                width = None
                height = self.MAX_HEIGHT
            else:
                height = self.height
        
        doc.add_picture(imagePath, width, height)
        if self.caption:
            doc.add_paragraph('Figure '+self.key+" : "+ self.caption, 'Caption')

class DocsEntityPageSection:
    def __init__(self, start_type = WD_SECTION_START.NEW_PAGE, orientation = WD_ORIENTATION.PORTRAIT):  # @UndefinedVariable
        self.start_type, self.orientation = start_type, orientation
    def render(self, doc):
        last_section = doc.sections[-1]
        last_orientation = last_section.orientation
        s = doc.add_section(self.start_type)
        if not last_orientation == self.orientation: 
            new_width, new_height = s.page_height, s.page_width
            s.orientation, s.page_height, s.page_width = self.orientation, new_height, new_width
            print("changed page orientation")

        

class DocsEntityPageBreak:
    def render(self, doc):
        doc.add_page_break()

class DocxEntityDocumentTitle:
    def __init__(self, title, level = 0):
        self.title, self.level = title, level
    def render(self, doc):
        doc.add_heading(self.title, level=self.level)
        
class DocxEntitySection(DocxEntityDocumentTitle):
    def __init__(self, title):
        DocxEntityDocumentTitle.__init__(self, title, 1)
        
class DocxEntitySubSection(DocxEntityDocumentTitle):
    def __init__(self, title):
        DocxEntityDocumentTitle.__init__(self, title, 2)
        
class DocxEntitySubSubSection(DocxEntityDocumentTitle):
    def __init__(self, title):
        DocxEntityDocumentTitle.__init__(self, title, 3)

class DocxEntityTable:
    def __init__(self, callback, rows, cols, caption=None, style=None):
        self.callback, self.caption, self.rows, self.cols, self.style = callback, caption, rows, cols, style
        self.key = "?"
    def render(self, doc):
        table = doc.add_table(self.rows, self.cols, self.style)
        self.callback(table)
        if self.caption:
            doc.add_paragraph('Tableau '+self.key+" : "+ self.caption, 'Caption')

class DocxEntityListOfFigures:
    def render(self, doc):
        doc.add_paragraph('Liste des figures', 'Illustration Index Heading')
        
        
class DocxEntity:
    def __init__(self, filename):
        self.filename = filename
        self.entities = []
        
    def getFilename(self):
        return self.filename
    
    def append(self, obj):
        self.entities.append(obj)
        return obj
        
    def initialize(self, doc):
        # caption defaults
        style = doc.styles['Caption']
        font = style.font
        font.name = 'Calibri'
        font.size = Pt(10)
        font.italic = True
        font.bold = False
        print(font.color.theme_color)
        font.color.theme_color = MSO_THEME_COLOR.ACCENT_2  # @UndefinedVariable
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
        
    def render(self, doc):
        self.initialize(doc)
        for obj in self.entities:
            obj.render(doc)

class Docx:
    
    class ORIENT:
        LANDSCAPE = WD_ORIENTATION.LANDSCAPE  # @UndefinedVariable
        PORTRAIT = WD_ORIENTATION.PORTRAIT  # @UndefinedVariable
    
    class START_TYPE:
        NEW_PAGE = WD_SECTION_START.NEW_PAGE  # @UndefinedVariable
        CONTINUOUS = WD_SECTION_START.CONTINUOUS  # @UndefinedVariable
        ODD_PAGE = WD_SECTION_START.ODD_PAGE  # @UndefinedVariable
    
    def __init__(self, filename, _ref = None):
        self.filename = filename
        self._ref = _ref
        self.entity = DocxEntity(filename)
        
    def title(self, title):
        self.entity.append(DocxEntityDocumentTitle(title))
        
    def sec(self, title):
        self.entity.append(DocxEntitySection(title))
        
    def subsec(self, title):
        self.entity.append(DocxEntitySubSection(title))
        
    def subsubsec(self, title):
        self.entity.append(DocxEntitySubSubSection(title))
        
    def pic(self, filename, width=DocxEntityPicture.DEFAULT_WIDTH, height=None, caption=None):
        return self.entity.append(DocxEntityPicture(filename, width, height, caption))
        
    def par(self, text = "", style=None):
        return self.entity.append(DocxEntityParagraph(text, style, self._ref))
        
    def b(self, text = "", style=None, pretext=""):
        p = DocxEntityParagraph(pretext, style)
        return self.entity.append(p.b(text))
        
    def pageBreak(self):
        self.entity.append(DocsEntityPageBreak())
        
    def pageSection(self, start_type = START_TYPE.NEW_PAGE, orientation = ORIENT.PORTRAIT):
        self.entity.append(DocsEntityPageSection(start_type, orientation))
        
    def table(self, callback, rows, cols, caption=None, style=None):
        return self.entity.append(DocxEntityTable(callback, rows, cols, caption, style))

    def list_of_figures(self):
        self.entity.append(DocxEntityListOfFigures())
    
    def save(self, target=None):
        if self.filename:
            d = Document(self.filename)
        else:
            d = Document()
        self.entity.render(d)
        if target:
            filename = target
        else:
            filename = self.filename
        d.save(filename)

