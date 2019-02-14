from pptx import Presentation
from pptx.util import Inches, Cm, Pt

from pptx.util import Inches, Cm, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
import numbers
import pandas
DIST_METRIC = Inches

def set_cell_font_attrs(cell, **kwargs):
    for p in cell.text_frame.paragraphs:
        for r in p.runs:
            for k, v in kwargs.items():
                if k == 'color_rgb':
                    r.font.color.rgb = v
                else:
                    setattr(r.font, k, v)
                
def format_cell_text(val, float_format = '{:,.0f}', int_format = '{:,d}'):
    if isinstance(val, numbers.Integral):
        return int_format.format(val)
    elif isinstance(val, numbers.Real):
        return float_format.format(val)
    else:
        return unicode(val)
    
def set_cell_appearance(cell):
    cell.fill.background()
    _set_cell_border(cell)
    
def write_pptx_dataframe(dataframe, pptx_table, col_width = 1.0, format_opts = {}, font_attrs = {},
                         header_font_attrs = {'bold': True, 'color_rgb': RGBColor(34, 64, 97)}):
    rows, cols = dataframe.shape
    if isinstance(dataframe.columns, pandas.MultiIndex) or isinstance(dataframe.index, pandas.MultiIndex):
        raise RuntimeError('Cannot yet cope with MultiIndex columns or row index')
    indexes = 1
    headers = 1

    header_font_attrs = dict(header_font_attrs, **font_attrs)
    for c in range(cols):
        #set column widths
        if isinstance(col_width, numbers.Number):
            w = DIST_METRIC(col_width)
        else:
            w = DIST_METRIC(col_width[c + 1])
        pptx_table.columns[c + indexes].width = w

        #headers
        cell = pptx_table.cell(0, c + indexes)
        cell.text = format_cell_text(dataframe.columns[c], **format_opts)
        set_cell_font_attrs(cell, **header_font_attrs)
        set_cell_appearance(cell)

        #body cells
        for r in range(rows):
            cell = pptx_table.cell(r + headers, c + indexes)
            cell.text = format_cell_text(dataframe.iloc[r, c], **format_opts)
            set_cell_font_attrs(cell, **font_attrs)
            set_cell_appearance(cell)

    #index
    for r in range(rows):
        cell = pptx_table.cell(r + headers, 0)
        cell.text = format_cell_text(dataframe.index[r], **format_opts)
        set_cell_font_attrs(cell, **header_font_attrs)
        set_cell_appearance(cell)
    
    pptx_table.columns[0].width = DIST_METRIC(col_width if isinstance(col_width, numbers.Number) else col_width[0])
    
    #index name
    cell = pptx_table.cell(0, 0)
    if dataframe.index.name is not None:
        cell.text = format_cell_text(dataframe.index.name, **format_opts)
    set_cell_font_attrs(cell, **header_font_attrs)
    set_cell_appearance(cell)
        
def create_pptx_table(pptx_slide, dataframe, left, top, col_width, row_height, **write_kwargs):
    rows, cols = dataframe.shape
    if isinstance(dataframe.columns, pandas.MultiIndex) or isinstance(dataframe.index, pandas.MultiIndex):
        raise RuntimeError('Cannot yet cope with MultiIndex columns or row index')
    indexes = 1
    headers = 1

    width = DIST_METRIC(col_width * cols if isinstance(col_width, numbers.Number) else sum(col_width))
    height = DIST_METRIC(row_height * rows)

    table_shape = pptx_slide.shapes.add_table(rows + headers, cols + indexes, 
                             DIST_METRIC(left), DIST_METRIC(top), 
                             width, height)
    
    table = table_shape.table
    write_pptx_dataframe(dataframe, table, col_width = col_width, **write_kwargs)
    return table
    
def SubElement(parent, tagname, **kwargs):
        element = OxmlElement(tagname)
        element.attrib.update(kwargs)
        parent.append(element)
        return element
    
def _set_cell_border(cell, border_color="4f81bd", border_width='12700'):
    """ Hack function to enable the setting of border width and border color
        - left border
        - right border
        - top border
        - bottom border
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    for border in 'LRTB':
        lnL = SubElement(tcPr, 'a:ln' + border, w='3175', cap='flat', cmpd='sng', algn='ctr')
        lnL_solidFill = SubElement(lnL, 'a:solidFill')
        lnL_srgbClr = SubElement(lnL_solidFill, 'a:srgbClr', val=border_color)
        lnL_prstDash = SubElement(lnL, 'a:prstDash', val='solid')
        lnL_round_ = SubElement(lnL, 'a:round')
        lnL_headEnd = SubElement(lnL, 'a:headEnd', type='none', w='med', len='med')
        lnL_tailEnd = SubElement(lnL, 'a:tailEnd', type='none', w='med', len='med')
