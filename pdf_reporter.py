from reportlab.pdfgen.canvas import Canvas
from reportlab.lib import colors 
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, PageBreak, Table, TableStyle, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import KeepTogether
from reportlab.lib.units import inch, mm

from PyPDF2 import PdfMerger
import pandas as pd

def truncate_strings(df, max_length):
    # Truncate all column names to max_length and add "..." if they are truncated
    df.columns = [col[:max_length-3] + '...' if len(col) > max_length else col for col in df.columns]
    # Truncate all string cell values to max_length and add "..." if they are truncated
    for col in df.columns:
        df = df.applymap(lambda x: (x[:max_length-3] + '...' if len(x) > max_length else x) if isinstance(x, str) else x)
    return df
 
def custom_round(x, precision=3): 
    if abs(x) > 1e5 or (0 < abs(x) < 1e-5):
        return "{:.{}e}".format(x, precision)
    else:
        return "{:.{}f}".format(x, precision)
    
class NumberedCanvas(Canvas):
    def showPage(self):
        self.setFont("Helvetica", 9)
        self.setFillColor(colors.grey)
        x = self._pagesize[0] / 2.0 
        self.drawCentredString(x, 0.3 * inch, str(self._pageNumber))
        Canvas.showPage(self)

#report class
class Report:
    def __init__(self, name, title = '', elements = None, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20):
        self.title = title
        if not name.endswith('.pdf'):
            name = name + '.pdf'
        self.elements = [] if elements is None else elements
        self.doc = SimpleDocTemplate(name, rightMargin=rightMargin, leftMargin=leftMargin, topMargin=topMargin, bottomMargin=bottomMargin)
        self.page_width, self.page_height = letter
        self.page_width = self.page_width - rightMargin - leftMargin
        self.page_height = self.page_height - topMargin - bottomMargin
        self.name = name
        self.title_style = getSampleStyleSheet()['Title']
        self.normal_style = getSampleStyleSheet()['Normal']
        self.italic_style = getSampleStyleSheet()['Italic']

    def initialize_report(self, type):
        if self.name.endswith('.pdf'):
            if self.title == '':
               title_text = self.name[:-4]
            else:
                title_text = self.title 

        title_style = self.italic_style.clone('title_style_close', alignment=TA_CENTER, textColor=colors.blue, fontSize=28)
        title = Paragraph(title_text, title_style)
        self.elements.append(title)
        self.elements.append(Spacer(1, 6))
        title_style = self.italic_style.clone('title_style_close', alignment=TA_CENTER, textColor=colors.blue, fontSize=14)
        if type == 'statistics':
            title = Paragraph('<br/>STATISTICS REPORT', title_style)
        elif type == 'processing':
            title = Paragraph('<br/>PROCESSING REPORT', title_style)
        else:
            title = Paragraph('<br/>REPORT', title_style)
        self.elements.append(title)

        normal_style = self.normal_style.clone('normal_style_clone', alignment=TA_CENTER, textColor=colors.blue)
        para = Paragraph('<br/>________________________________________________________________________________________________', normal_style)
        self.elements.append(para)

        if type == 'statistics':
            text0 = 'This report describes the statistics of analysed data: All of the statistics performed will be automatically reported here (references will be added later).'
        elif type == 'processing':
            text0 = 'This report describes the processing of the data: All of the processing performed will be automatically reported here (references will be added later).'
        else:
            text0 = 'This report describes the data and the processes it went through. (references will be added later).'
        text1 = 'All the <b>images</b> are saved in both <b>png</b> and <b>pdf</b> formats under the same name as stated in their respective sections.'
        
        italic_style = self.italic_style.clone('italic_style_clone', alignment=TA_CENTER)
        para = Paragraph(text0, italic_style)
        self.elements.append(para)
        para = Paragraph(text1, italic_style)
        self.elements.append(para)
        normal_style = self.normal_style.clone('normal_style_clone', alignment=TA_CENTER)
        para = Paragraph('<br/>________________________________________________________________________________________________', normal_style)
        self.elements.append(para)

    def finalize_report(self):
        self.doc.build(self.elements, canvasmaker=NumberedCanvas)

    #---------------------------------------------
    def add_pagebreak(self, return_element_only = False):
        if return_element_only:
            return PageBreak()
        else:
            self.elements.append(PageBreak())
    
    def add_line(self, return_element_only = False):
        #align the line to the center
        normal_style = self.normal_style.clone('normal_style_clone', alignment=TA_CENTER)
        para = Paragraph('________________________________________________________________________________________________', normal_style)
        if return_element_only:
            return para
        else:
            self.elements.append(para)
    
    def add_text(self, text, style='normal', alignment='left', font_size = 10, return_element_only=False):
        switcher = {
            'title': self.title_style,
            'normal': self.normal_style,
            'italic': self.italic_style
        }
        alignment_dict = {'left': TA_LEFT, 'center': TA_CENTER, 'right': TA_RIGHT}
        text = '<br/>' + text
        if style == 'bold':
            style = self.normal_style.clone('clone')
            text = '<b>' + text + '</b>'
        else:
            style = switcher[style].clone('clone')
        style.alignment = alignment_dict[alignment]
        style.fontSize = font_size
        para = Paragraph(text, style)
        if return_element_only:
            return [para, Spacer(1, 3)]
        else:
            self.elements.append(para)
            self.elements.append(Spacer(1, 3))
    
    def add_image(self, image, return_element_only = False):
        # if image is just a path to the file
        if isinstance(image, str):
            if not image.endswith('.png'):
                image = image + '.png'
            image = Image(image)

        # Calculate the scale factors for the width and height
        width_scale = self.page_width / image.drawWidth
        height_scale = self.page_height / image.drawHeight

        # Use the smaller scale factor to ensure the image fits within the page
        scale = min(width_scale, height_scale)

        # Scale the width and height of the image
        image.drawWidth *= scale
        image.drawHeight *= scale

        # Calculate the new position of the image to center it on the page
        image.hAlign = 'CENTER'

        if return_element_only:
            return [image, Spacer(1, 6)]
        else:
            self.elements.append(image)
            self.elements.append(Spacer(1, 6))
    
    def add_table(self, data, return_element_only = False):
        data_view = data.copy()

        truncate_nm = 10
        row_height = 20  # Set the height of the cells to 20

        # If the data is too long append a row with "..." in each cell
        if len(data_view) > 10:
            rows_to_show = data_view.index.tolist()[:10]
            data_view = data_view.loc[rows_to_show]
            new_row = pd.Series(['...' for _ in range(len(data_view.columns))], index=data_view.columns, name='...')
            data_view.loc['...'] = new_row
        # If the data is too wide insert a column with "..." in each cell
        if len(data_view.columns) > 10:
            columns_to_show = data_view.columns.tolist()[:5] + data_view.columns.tolist()[-5:]
            data_view = data_view[columns_to_show]
            data_view.insert(5, '...', ['...' for _ in range(len(data_view))])
            col_width = 50
            col_widths = [col_width if i != 5 else 20 for i in range(len(data_view.columns))]
        else:
            truncate_nm = 13 # can be bigger cause we have less columns
            col_width = 60 if len(data_view.columns) <= 8 else 50
            col_widths = [col_width] * len(data_view.columns)

        data_view = data_view.applymap(lambda x: custom_round(x) if isinstance(x, (int, float)) else x)
        data_view = data_view.applymap(str)
        data_view = truncate_strings(data_view, truncate_nm)

        # Convert DataFrame to a list of lists, with each inner list representing a row
        data_list = data_view.values.tolist()
        # Add column names as the first row
        data_list.insert(0, data_view.columns.tolist())

        
        row_heights = [row_height] *(len(data_view) +1) # +1 for the column names

        table = Table(data_list, colWidths=col_widths, rowHeights=row_heights)        
        table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1, colors.grey),
            ('BACKGROUND', (0, 0), (-1, 0), colors.palegreen),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black), 
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ]))
        if return_element_only:
            return [table, Spacer(1, 6)]
        else:
            self.elements.append(table)
            self.elements.append(Spacer(1, 6))


    def add_together(self, elements):   # can behave unexpectedly if adding a pagebreak this way
        keep_together_elements = []
        for element in elements:
            if isinstance(element, tuple):
                type = element[0]
                if type == 'text':
                    element = self.add_text(*element[1:], return_element_only=True)
                elif type == 'image':
                    element = self.add_image(*element[1:], return_element_only=True)
                elif type == 'table':
                    element = self.add_table(*element[1:], return_element_only=True)
                else:
                    raise ValueError('Invalid type: ' + type)
                if isinstance(element, list):
                    keep_together_elements.extend(element)
                else:
                    keep_together_elements.append(element)
            elif isinstance(element, str):
                if element == 'line':
                    keep_together_elements.append(self.add_line(return_element_only=True))
                elif element == 'pagebreak':
                    keep_together_elements.append(self.add_pagebreak(return_element_only=True))
                else:
                    raise ValueError('Invalid type: ' + element)
            else:
                raise ValueError('Invalid element: ' + str(element))
        self.elements.append(KeepTogether(keep_together_elements))

    def merge_pdfs(self, pdfs, output_name):
        merger = PdfMerger()
        for pdf in pdfs:
            merger.append(pdf)
        merger.write(output_name)
        merger.close()
        return output_name
