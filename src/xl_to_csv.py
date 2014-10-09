# -*- coding: utf-8 -*-
import csv
import zipfile
import xml.etree.ElementTree as ElementTree

def unzip(Filename):
    zipfile.ZipFile(Filename).extractall('/tmp')
    return '/tmp/xl'

def take_sharedStrings(Filename):
    xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    result = {}
    running_path = []
    meet_si = False
    t_bag = None
    for event, node in ElementTree.iterparse(Filename, events=('start','end')):
        if event == 'start' and not meet_si:
            running_path.extend([node.tag])
            if node.tag == '{%s}si' % (xmlns):
                meet_si = True
                t_bag = ''
        if meet_si and event == 'end' and node.tag == '{%s}t' % (xmlns):
            t_bag = t_bag + node.text
        if event == 'end' and meet_si:
            running_path = running_path[:-1]
            meet_si = False
            result.update({len(result): t_bag})
        if event == 'end':
            node.clear()
    return result

def take_rows(Filename, sharedStrings):
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns_r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    running_path = []
    running_row = None
    for event, node in ElementTree.iterparse(Filename, events=('start','end')):
        if event == 'start':
            running_path.extend([node.tag])
            if '/'.join(running_path) == '{%s}worksheet/{%s}sheetData/{%s}row' % (xmlns, xmlns, xmlns):
                running_row = []
            if '/'.join(running_path) == '{%s}worksheet/{%s}sheetData/{%s}row/{%s}c' % (xmlns, xmlns, xmlns, xmlns):
               running_type = node.attrib['t'] if 't' in list(node.attrib) else None
        if event == 'end':
            if '/'.join(running_path) == '{%s}worksheet/{%s}sheetData/{%s}row/{%s}c/{%s}v' % (xmlns, xmlns, xmlns, xmlns, xmlns):
                if running_type == 's':
                    running_row.extend([sharedStrings[int(node.text)]])
                else:
                    running_row.extend([node.text])
            if '/'.join(running_path) == '{%s}worksheet/{%s}sheetData/{%s}row/{%s}c' % (xmlns, xmlns, xmlns, xmlns) and len(node._children) == 0:
                running_row.extend([''])
            if '/'.join(running_path) == '{%s}worksheet/{%s}sheetData/{%s}row' % (xmlns, xmlns, xmlns):
                yield(running_row)
            running_path = running_path[:-1]
            node.clear()


Path = unzip('/home/wolfram/work/data/Viva/Huawei_Data/SiteConfigurationDB/VIVA_LTE_Parameters(20140425)-GT.xlsx')
#Path = unzip('/home/wolfram/work/data/Viva/Huawei_Data/SiteConfigurationDB/v.xlsx')

sharedStrings = take_sharedStrings('/tmp/xl/sharedStrings.xml')

f = open('/tmp/VIVA_LTE_Parameters.csv', 'w')
c = csv.writer(f, lineterminator='\n')
for row in take_rows('/tmp/xl/worksheets/sheet1.xml', sharedStrings):
    c.writerows([[s.encode('utf-8') for s in row]])
f.close()

