# -*- coding: utf-8 -*-
"""
Created on Sat Jun  7 11:39:07 2025

@author: User
"""

from pathlib import Path
from tempfile import TemporaryDirectory
from lxml import etree
import copy
import re


'''
data = """<!-- This is comment -->
<contacts xmlns="http://www.un.com/ns/1">
    <contact>
        <name>
            <first>Evgeniy</first>
            <last>Lestopadov</last>
        </name>
        <phone type="mobile">+8789878797</phone>
        <email>elestopadov@mail.ru</email>
    </contact>
    <contact>
        <name>
            <first>Dmitro</first>
            <last>Lee</last>
        </name>
        <phone type="home">+879789893</phone>
        <email>dim@gmail.com</email>
    </contact>
</contacts>"""

doc = etree.fromstring(data)
ns = doc.nsmap
element = doc.find('.//*[@type="mobile"]', ns)
#print(doc.tag, doc.attrib)
print(element.tag)
'''



def tree_Xmldoc(filename):
    '''For instance, "document.xml"'''
    tree = etree.parse(f'{filename}')
    return tree

def update_run(run, txt='', filling=False):
    '''It's the easiest way to update the run'''
    ns = run.nsmap
    form = run.find('./w:rPr', ns)    
    if txt:
        run.find('.//w:t', ns).text = txt
    if filling:
        element1 = etree.Element('{' + f"{ns['w']}" + '}' + "highlight")
        element1.attrib['{' + f"{ns['w']}" + '}' + "val"] = "yellow"
        element2 = etree.Element('{' + f"{ns['w']}" + '}' + "u")
        element2.attrib['{' + f"{ns['w']}" + '}' + "val"] = "single"        
        form.append(element2)
        form.append(element1)
    
    return None

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
    
def get_paragraphs(node):
    ns = node.nsmap
    paragraphes = node.findall('./w:p', ns)
    yield from paragraphes
    
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

def get_text_from_paragraph(paragraph):
    ns = paragraph.nsmap    
    txt = ''    
    runs = paragraph.findall('.//w:r', ns)
    for run in runs:
        r_txt = run.find('.//w:t', ns).text
        txt += r_txt
    return txt

filename = input('Enter filename: ')
tree = tree_Xmldoc(filename)
root = tree.getroot()
ns = root.nsmap
body = root.find('./', ns)
table = next(get_tables(body))


print('Enter part you need to replace')
old = input()
print('Enter new part you need to paste')
new = input()
print('Enter mode for replacing\nwhere\n[0] - micro regime\n[1] - macro regime')
mode = int(input())


for row in get_table_rows(table):
    for cell in cells(row):
        for paragraph in get_paragraphs(cell):
            if old in get_text_from_paragraph(paragraph):
                update_paragraph(paragraph, old, new, mode, filling=True)
                
        
        
with open(filename, 'wb') as f:
    tree.write(f, pretty_print=True, encoding="utf-8")
    
