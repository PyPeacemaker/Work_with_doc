# -*- coding: utf-8 -*-
"""
Created on Sat May 17 15:59:09 2025

@author: KhasyanovDR
"""

import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
from lxml import etree
import copy
import re


def replace_information_in_doc(doc_path: str, output_path: str, replacements={}):
    with TemporaryDirectory() as temp_dir:
        temp_dir = Path(temp_dir)
        
    with zipfile.ZipFile(doc_path, 'r') as zip_ref:
        zip_ref.extractall('Inside_our_file')

def extract_doc(filename):
    with zipfile.ZipFile(f'Блок 1\\{filename}', 'r') as zip_ref:
        zip_ref.extractall('Inside_our_file')
    return None

def create_new_doc(filename):
    fl = Path(filename)
    with zipfile.ZipFile('new_doc.docx', 'w') as zip_out:
        for file in fl.glob('**/*'):
            arcname = str(file.relative_to(fl))
            zip_out.write(file, arcname)
    return None

def get_tables(body):
    ns = body.nsmap
    tables = body.findall('./w:tbl', ns)
    yield from tables
    
def get_table_rows(table):
    ns = table.nsmap
    rows = table.findall('./w:tr', ns)
    yield from rows

def cells(row):
    ns = row.nsmap
    cells = row.findall('./w:tc', ns)
    yield from cells
    
def get_text_cell(cell):
    ns = cell.nsmap
    paragraphes = cell.findall('./w:p', ns)
    txt = ''
    for p in paragraphes:
        runs = cell.findall('.//w:r', ns)
        for run in runs:
            r_txt = run.find('.//w:t', ns).text
            txt += r_txt
    return txt

def width_of_cell(cell):
    ns = cell.nsmap
    width = cell.find('./w:tcPr/w:tcW', ns)            
    return width.attrib[width.attrib.keys()[0]]

def width_of_table(table):
    ns = table.nsmap
    width = table.find('.//w:tblW', ns)            
    return width.attrib[width.attrib.keys()[0]]

def proportions_of_sheet(body):
    ns = body.nsmap
    options_page = body.find('.//w:sectPr', ns)
    proportions_page = options_page.find('.//w:pgSz', ns).attrib
    fields_page = options_page.find('.//w:pgMar', ns).attrib
    width_pg = proportions_page[proportions_page.keys()[0]]
    height_pg = proportions_page[proportions_page.keys()[1]]
    field_top = fields_page[fields_page.keys()[0]]
    field_right = fields_page[fields_page.keys()[1]]
    field_bottom = fields_page[fields_page.keys()[2]]
    field_left = fields_page[fields_page.keys()[3]]
    field_header = fields_page[fields_page.keys()[4]]
    field_footer = fields_page[fields_page.keys()[5]]
    field_gooter = fields_page[fields_page.keys()[6]]
    return {'page': {'width_pg': width_pg, 'height_pg': height_pg}, 
            'fields': {'field_top': field_top,
                       'field_right': field_right,
                       'field_bottom': field_bottom,
                       'field_left': field_left,
                       'field_header': field_header,
                       'field_footer': field_footer,
                       'field_gooter': field_gooter}}

def create_new_run_from_pr(run, txt='', filling=False):
    xml_new_r = copy.deepcopy(run)
    ns = xml_new_r.nsmap
    form = xml_new_r.find('./w:rPr', ns)    
    if txt:
        xml_new_r.find('.//w:t', ns).text = txt
    if filling:
        element1 = etree.Element('{' + f"{ns['w']}" + '}' + "highlight")
        element1.attrib['{' + f"{ns['w']}" + '}' + "val"] = "yellow"
        element2 = etree.Element('{' + f"{ns['w']}" + '}' + "u")
        element2.attrib['{' + f"{ns['w']}" + '}' + "val"] = "single"        
        form.append(element2)
        form.append(element1)    
    return xml_new_r

def divide_into_runs(txt, old_part, new_part, mode: int):
    '''This functions returns list of runs'''
    regime = ['micro', 'macro'][mode]
    if regime == 'micro':
        new = ''
        flag = False
        new_part = iter(new_part)
        for old_chr in old_part:
            new_chr = next(new_part)
            if old_chr == new_chr and flag is False:
                new += new_chr
            elif old_chr != new_chr and flag is False:
                flag = True
                new += '😃' + new_chr
            elif old_chr != new_chr and flag is True:
                new += new_chr
            elif old_chr == new_chr and flag is True:        
                new += new_chr + '😃'
                flag = False
        c = new.count('😃')
        if c % 2 != 0:
            new += '😃'
        new_part = new
    elif regime == 'macro':
        new_part = '😃' + f'{new_part}' + '😃'
    reg = fr'{old_part}' 
    common = re.sub(reg, new_part, txt)

    runs = []
    flag = False
    run = ''
    for char in common:    
        if char == '😃' and flag is False:
            runs.append(run)
            flag = True
            run = '😃'
            continue
        elif char == '😃' and flag is True:
            run += char
            flag = False
            runs.append(run)
            run = ''
            continue          
        run += char
    if run and run not in runs:
        runs.append(run)
    return runs

def update_paragraph(paragraph, old, new, mode, filling=False):
    ns = paragraph.nsmap    
    runs = paragraph.findall('./w:r', ns)
    for run in runs:
        if old in run.find('./w:t', ns).text:
            common = run.find('./w:t', ns).text
            break
        
    divided = divide_into_runs(common, old, new, mode)
    ind = paragraph.index(run)
                         
    for part in divided:
        ind = ind + 1
        if '😃' in part:            
            new_r = create_new_run_from_pr(run, part[1:len(part)-1], filling=filling)
        else:
            new_r = create_new_run_from_pr(run, part)
        paragraph.insert(ind, new_r)
    
    paragraph.remove(run)
    return

def get_new_paragraph_from_another_one(paragraph, txt='', filling=False):
    xml_new_p = copy.deepcopy(paragraph)
    ns = xml_new_p.nsmap    
    runs = xml_new_p.findall('./w:r', ns)
    for run in runs[1:]:
        xml_new_p.remove(run)
    form = xml_new_p.find('.//w:rPr', ns)    
    if txt:
        xml_new_p.find('.//w:t', ns).text = txt
    if filling:
        element1 = etree.Element('{' + f"{ns['w']}" + '}' + "highlight")
        element1.attrib['{' + f"{ns['w']}" + '}' + "val"] = "yellow"
        element2 = etree.Element('{' + f"{ns['w']}" + '}' + "u")
        element2.attrib['{' + f"{ns['w']}" + '}' + "val"] = "single"        
        form.append(element2)
        form.append(element1)    
    return xml_new_p

def paste_new_paragraph(node, paragraph_target, new_paragraph, below=False, above=False):
    ind = node.index(paragraph_target)
    if below:        
        node.insert(ind + 1, new_paragraph)
    elif above:
        node.insert(ind, new_paragraph)
    return

tree = etree.parse('Inside_our_file\\word\\document.xml')
root = tree.getroot()
ns = root.nsmap
body = root.find('./', ns)
'''tables = body.find('./w:tbl', ns)
cells = tables.find('.//w:tc', ns)
width = cells.find('./w:tcPr/w:tcW', ns)
print(type(width.attrib), width.attrib[width.attrib.keys()[0]])'''

'''table = next(get_tables(body))
for row in get_table_rows(table):
    for i, cell in enumerate(cells(row), start=1):
        print(f'cell {i}', get_text_cell(cell))
        q = input()'''
        
'''
table = next(get_tables(body))
row = next(get_table_rows(table))
cell = next(cells(row))
paragraph = cell.find('./w:p', ns)
run = paragraph.find('./w:r', ns)
print(create_new_run_from_pr(run, txt='Number', filling=True))
'''

'''print(proportions_of_sheet(body))'''

#create_new_doc('Inside_our_file — копия')