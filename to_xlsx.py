__author__= 'karec'
# This script is made for convert xlsm files to xlsx
# Used for pass xlsm file to phpExcel

import os
import zipfile
import shutil
import tempfile
import sys
from xml.dom import minidom

def to_xlsx(fname, *filenames):
    """
    This function create a xlsx file from
    xlsm file

    Args:
        fname (string): file to convert

    Filenames:
        name (string): string representing files tu update
    """
    tempdir = tempfile.mkdtemp()
    try:
        tempname = os.path.join(tempdir, 'new.zip')
        with zipfile.ZipFile(fname, 'r') as zipread:
            with zipfile.ZipFile(tempname, 'w') as zipwrite:
                for item in zipread.infolist():
                    if item.filename not in filenames:
                        data = zipread.read(item.filename)
                        zipwrite.writestr(item, data)
                    else:
                        zipwrite.writestr(item, update_files(
                                item.filename, zipread.read(item.filename)
                                ))
        os.remove(fname)
        fname = fname.split('.')[0]
        fname = fname + '.xlsx'
        shutil.move(tempname, fname)
    finally:
        shutil.rmtree(tempdir)


def update_files(filename, data):
    """
    This function dispatch the data for return good xml values
    """
    if '[Content_Types].xml' in filename:
        return update_content_types(data)
    elif 'workbook' in filename:
        return update_workbook(data)


def update_content_types(data):
    """
    This function remove macro and set
    correct file type in headers
    """
    xml = minidom.parseString(data)
    types = xml.getElementsByTagName('Types')[0]
    for item in xml.getElementsByTagName('Types'):
        if item.hasAttribute('PartName') and item.getAttribute('PartName') == '/xl/vbaProject.bin':
            item.parentNode.removeChild(item)
    
    for item in types.getElementsByTagName('Override'):
        if item.hasAttribute('PartName') and item.getAttribute('PartName') == '/xl/workbook.xlk':
            item.setAttribute('ContentType', 'vnd.openxmlformats-officedocument.extended-properties+xml')
    return xml.toxml()

if __name__ == '__main__':
    f = raw_input('File name and location : ')
    if f == '': f = 'final.xlsm'
    to_xlsx(f, '[Content_Types].xml')
    
