from pptx import Presentation
from pptx.util import Inches, Cm, Pt
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
import pptx
import numbers
import pandas
DIST_METRIC = Inches
from six import text_type

from shutil import copyfile

class NoExistingSlideFoundError(RuntimeError): pass

BASE_POSITIONS = {
    'content': [5.2, 4.1, 5, 2.7],
    'index_width': 0.7, 
    'col_width': 0.6, 
    'single_table_margin_top': 4.3,
    'single_table_margin_left': None, # set to value other than None for different table vs chart left position
    'single_item_margin_top': 1.7, 
    'single_item_margin_left': 2.2,
    'multi_item_margin_top': 1.4, 
    'multi_item_margin_left': 1.5,
    'charts_horizontal_gap': 0.7, 
    'charts_vertical_gap': 0.5,
    'table_row_height': 0.2
}
    
class PresentationWriter():
    def __init__(self, pptx_file, pptx_title, pptx_template = None, 
                 default_overwrite_if_present = False,
                 default_overwrite_only = False,
                 default_position_overrides = {}):
        self.pptx_file = pptx_file
        self.pptx_title = pptx_title
        
        if pptx_template is not None:
            copyfile(pptx_template, pptx_file)
        
        self.presentation = Presentation(self.pptx_file)
        core_props = self.presentation.core_properties
        core_props.title = pptx_title
        core_props.keywords = ''
        
        self.default_overwrite_if_present = default_overwrite_if_present
        self.default_overwrite_only = default_overwrite_only
        self.default_positions = dict(**BASE_POSITIONS)
        self.default_positions.update(default_position_overrides)
        
    def chart_to_file(self, chart_obj, img_file):
        if hasattr(chart_obj, 'write_image'):
            #ply_pd plots
            chart_obj.write_image(img_file, scale=2, width = chart_obj.width, height = chart_obj.height)
        else:
            #matplotlib plots...
            chart_obj.figure.savefig(img_file, dpi = 300, bbox_inches = 'tight')
            
    def save_presentation(self):
        self.presentation.save(self.pptx_file)
            
    def write_slide(self, title, charts = [], tables = [], 
                    overwrite_if_present = None,
                    overwrite_only = None,
                    charts_per_row = 2, tables_per_row = 1,
                    position_overrides = {}):
        
        positions = dict(**self.default_positions)
        positions.update(position_overrides)
        
        if not isinstance(charts, (list, tuple)):
            charts = [charts]
        if not isinstance(tables, (list, tuple)):
            tables = [tables]
            
        overwrite_if_present = overwrite_if_present if overwrite_if_present is not None else self.default_overwrite_if_present
        overwrite_only = overwrite_only if overwrite_only is not None else self.default_overwrite_only
        if overwrite_if_present or overwrite_only:
            try:
                self.overwrite_pptx(title, charts, tables)
            except NoExistingSlideFoundError:
                if overwrite_only:
                    raise
            else:
                return
            
        prs = self.presentation
        title_only_slide_layout = prs.slide_layouts[3]
        slide = prs.slides.add_slide(title_only_slide_layout)
        shapes = slide.shapes

        content, subtitle, title_shape, footnotes = [s for s in slide.shapes if s.has_text_frame]
        title_shape.text = self.pptx_title
        subtitle.text = title

        content.left, content.top, content.width, content.height = (Inches(x) for x in positions['content'])

        font_attrs = dict(size = Pt(7), name = 'Calibri')
                
        if charts:
            total_chart_width = positions['multi_item_margin_left'] if len(charts) > 1 else positions['single_item_margin_left']
            initial_width = total_chart_width
            total_chart_height = positions['multi_item_margin_top'] if len(charts) > 2 else positions['single_item_margin_top']
            for i, chrt in enumerate(charts):
                if hasattr(chrt, 'width'):
                    chart_width = chrt.width / 800.0 * 5.0
                    chart_height = chrt.height / 800.0 * 5.0
                else:
                    chart_width = chrt.figure.get_figwidth() #already in inches
                    chart_height = chrt.figure.get_figheight()

                img_file = 'test%d.png' % i
                self.chart_to_file(chrt, img_file)

                pic = slide.shapes.add_picture(img_file, 
                                               left = Inches(total_chart_width), 
                                               top = Inches(total_chart_height), 
                                               width = Inches(chart_width), 
                                               height = Inches(chart_height))
                total_chart_width += chart_width + positions['charts_horizontal_gap']
                if (i % charts_per_row) == (charts_per_row-1):
                    total_chart_height += chart_height + positions['charts_vertical_gap']
                    total_chart_width = initial_width

        if tables:
            total_width = (positions['multi_item_margin_left'] 
                           if len(tables) > 1 
                           else positions['single_table_margin_left'] or positions['single_item_margin_left'])
            initial_width = total_width
            total_height = positions['multi_item_margin_top'] if len(charts) == 0 else positions['single_table_margin_top']
            
            for i, table in enumerate(tables):
                col_widths = [positions['index_width']] + [positions['col_width']]*len(table.columns)
                table_df = table.get_formatted_df()
                table_df.columns = [s.replace('<br>', '\n')
                                    if isinstance(s, str) else s
                                    for s in table_df.columns]
                table = create_pptx_table(slide, table_df, left = total_width, top = total_height, 
                                          col_width = col_widths, row_height = positions['table_row_height'],
                                          font_attrs = font_attrs)
                
                total_width += sum(col_widths) + positions['charts_horizontal_gap']
                if (i % tables_per_row) == (tables_per_row-1):
                    total_height += (positions['table_row_height']*(len(table_df)+1)) + positions['charts_vertical_gap']
                    total_width = initial_width
                    
        self.save_presentation()
    
    def overwrite_pptx(self, slide_title, strats_charts = None, strats_tables = None):
        prs = self.presentation

        for i, slide in enumerate(prs.slides):
            text_boxes = [s for s in slide.shapes if s.has_text_frame]

            matched_title = False
            for tb in text_boxes:
                if tb.text == slide_title:
                    matched_title = True
                    break

            if matched_title:
                break

        if matched_title:
            charts = []
            tables = []
            for s in slide.shapes:
                if isinstance(s, pptx.shapes.picture.Picture):
                    charts.append(s)
                elif isinstance(s, pptx.shapes.graphfrm.GraphicFrame) and s.has_table:
                    tables.append(s)

            if strats_charts:
                if len(strats_charts) != len(charts):
                    raise RuntimeError('Need %d Picture shapes but %d available on slide "%s"'
                                       % (len(strats_chart), len(charts), slide_title))

                for i, strats_chart in enumerate(strats_charts):
                    img_file = 'test%d.png' % i
                    self.chart_to_file(strats_chart, img_file)

                    # Replace image:
                    picture = charts[i]
                    with open(img_file, 'rb') as f:
                        imgBlob = f.read()
                    imgRID = picture._pic.xpath('./p:blipFill/a:blip/@r:embed')[0]
                    imgPart = slide.part.related_parts[imgRID]
                    imgPart._blob = imgBlob

            if strats_tables:
                if len(strats_tables) != len(tables):
                    raise RuntimeError('Need %d Table shapes but %d available on slide "%s"'
                                       % (strats_tables, len(tables), slide_title))
                    
                for i, strats_table in enumerate(strats_tables):
                    table_df = strats_table.get_formatted_df()
                    table_df.columns = [s.replace('<br>', '\n')
                                        for s in table_df.columns
                                        if isinstance(s, str)]
                    write_pptx_dataframe(table_df, tables[i].table, overwrite_formatting = False)

            self.save_presentation()
        else:
            raise NoExistingSlideFoundError('No existing slide titled "%s"' % slide_title)
    
def set_cell_text(cell, text, overwrite_formatting = True):
    if text == '':
        text = "\u00A0" # unicode nbsp - needed to fill empty cells as otherwise formatting is not applied by PPT
    if overwrite_formatting:
        p = cell.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = text
    else:
        p = cell.text_frame.paragraphs[0]
        if p.runs:
            p.runs[0].text = text
        else:
            r = p.add_run()
            r.text = text
        
def set_cell_font_attrs(cell, **kwargs):
    for p in cell.text_frame.paragraphs:
        for r in p.runs:
            for k, v in kwargs.items():
                if k == 'color_rgb':
                    r.font.color.rgb = v
                else:
                    setattr(r.font, k, v)
                
def format_cell_text(val, float_format = '{:.0f}', int_format = '{:d}'):
    if isinstance(val, numbers.Integral):
        return int_format.format(val)
    elif isinstance(val, numbers.Real):
        return float_format.format(val)
    else:
        return text_type(val)
    
def set_cell_appearance(cell):
    cell.fill.background()
    _set_cell_border(cell)
    
def write_pptx_dataframe(dataframe, pptx_table, col_width = 1.0, format_opts = {}, font_attrs = {'size': Pt(8), 'name': 'Calibri'},
                         header_font_attrs = {'bold': True, 'color_rgb': RGBColor(34, 64, 97)},
                         overwrite_formatting = True):
    rows, cols = dataframe.shape
    if isinstance(dataframe.index, pandas.MultiIndex):
        raise RuntimeError('Cannot yet cope with MultiIndex in rows')
    indexes = 1
    
    if isinstance(dataframe.columns, pandas.MultiIndex):
        headers = len(dataframe.columns.levels)
    else:
        headers = 1
        
    if (rows + headers) != len(pptx_table.rows):
        raise RuntimeError('Need %d rows but PPTX table has %d'
                           % (rows + headers, len(pptx_table.rows)))
    if (cols + indexes) != len(pptx_table.columns):
        raise RuntimeError('Need %d columns but PPTX table has %d'
                           % (cols + indexes, len(pptx_table.columns)))
        
    header_font_attrs = dict(header_font_attrs, **font_attrs)
    
    for i in range(headers):
        #headers
        prev_header = no_prev_header = '##special missing value'
        first_merged_cell = 0
        mergeable_cell_count = 0
        for c, header_name in enumerate(dataframe.columns.values):
            col_name = header_name[i] if headers > 1 else header_name
            if prev_header == no_prev_header or prev_header != col_name:
                if mergeable_cell_count > 0:
                    pptx_table.cell(i, first_merged_cell + indexes).merge(pptx_table.cell(i, first_merged_cell + mergeable_cell_count + indexes))
                prev_header = col_name
                first_merged_cell = c
                mergeable_cell_count = 0
                
                cell = pptx_table.cell(i, c + indexes)
                set_cell_text(cell, 
                              format_cell_text(col_name, **format_opts), 
                              overwrite_formatting = overwrite_formatting)
                if overwrite_formatting:
                    set_cell_font_attrs(cell, **header_font_attrs)
                    set_cell_appearance(cell)
            else:
                mergeable_cell_count += 1
        if mergeable_cell_count > 0:
            pptx_table.cell(i, first_merged_cell + indexes).merge(pptx_table.cell(i, first_merged_cell + mergeable_cell_count + indexes))
                
    for c in range(cols):
        #set column widths
        if isinstance(col_width, numbers.Number):
            w = DIST_METRIC(col_width)
        else:
            w = DIST_METRIC(col_width[c + 1])
        if overwrite_formatting:
            pptx_table.columns[c + indexes].width = w

        #body cells
        for r in range(rows):
            cell = pptx_table.cell(r + headers, c + indexes)
            set_cell_text(cell, 
                          format_cell_text(dataframe.iloc[r, c], **format_opts), 
                          overwrite_formatting = overwrite_formatting)
            if overwrite_formatting:
                set_cell_font_attrs(cell, **font_attrs)
                set_cell_appearance(cell)

    #index
    for r in range(rows):
        cell = pptx_table.cell(r + headers, 0)
        set_cell_text(cell, 
                      format_cell_text(dataframe.index[r], **format_opts), 
                      overwrite_formatting = overwrite_formatting)

        if overwrite_formatting:
            set_cell_font_attrs(cell, **header_font_attrs)
            set_cell_appearance(cell)
    
    if overwrite_formatting:
        pptx_table.columns[0].width = DIST_METRIC(col_width if isinstance(col_width, numbers.Number) else col_width[0])
    
    #index name
    for i in range(headers):
        cell = pptx_table.cell(i, 0)
        if dataframe.index.name is not None:
            set_cell_text(cell, 
                          format_cell_text(dataframe.index.name, **format_opts), 
                          overwrite_formatting = overwrite_formatting)
        if overwrite_formatting:
            set_cell_font_attrs(cell, **header_font_attrs)
            set_cell_appearance(cell)
        
def create_pptx_table(pptx_slide, dataframe, left, top, col_width, row_height, **write_kwargs):
    rows, cols = dataframe.shape
    if isinstance(dataframe.index, pandas.MultiIndex):
        raise RuntimeError('Cannot yet cope with MultiIndex rows')
    indexes = 1

    if isinstance(dataframe.columns, pandas.MultiIndex):
        headers = len(dataframe.columns.levels)
    else:
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
    
def _set_cell_border(cell, border_color="4f81bd", border_width='12700', border_scheme_color = 'accent1'):
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
        if border_scheme_color is not None:
            lnL_srgbClr = SubElement(lnL_solidFill, 'a:schemeClr', val=border_scheme_color)
        else:
            lnL_srgbClr = SubElement(lnL_solidFill, 'a:srgbClr', val=border_color)
        lnL_prstDash = SubElement(lnL, 'a:prstDash', val='solid')
        lnL_round_ = SubElement(lnL, 'a:round')
        lnL_headEnd = SubElement(lnL, 'a:headEnd', type='none', w='med', len='med')
        lnL_tailEnd = SubElement(lnL, 'a:tailEnd', type='none', w='med', len='med')
