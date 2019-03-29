#!/usr/bin/python3
# -*- coding: utf-8 -*-
import os
from os import listdir
from os.path import isfile, join
from PIL import ImageGrab
import win32com.client as win32
import string

def ListFiles(mypath):
    f = []
    t = listdir(mypath)
    print(":: Processing files: ")
    for filenames in t:
        f.append(filenames)
    return f

def GetWorkbookExcelFile(filename):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(filename)
    return wb

def DiscartImage(shape):
    if (shape.Height >= 40 and shape.Height <= 70 and shape.Width >= 130 and shape.Width <= 160):
        return True
    if (shape.Height >= 45 and shape.Height <= 50 and shape.Width >= 45 and shape.Width <= 50):
        return True
    return False

def SaveFileImage(dest, workbook):
    filename = workbook.Name
    id_ordem = filename.replace('.xlsm', '')
    for sheet in workbook.Worksheets:
        listShape = enumerate(sheet.Shapes)
        for i, shape in listShape:
            success = False
            if DiscartImage(shape) == False and shape.Name.startswith('Picture'):
                destination = '{0}{1:04}\\'.format(dest, int(id_ordem))

                if not os.path.exists(destination):
                    os.makedirs(destination)

                n_item = '{0:04}'.format(i+1)
                try:
                    shape.Copy()
                    image = ImageGrab.grabclipboard()
                    image.save('{0}{1}.jpg'.format(destination, n_item), 'jpeg')
                    success = True
                except:
                    continue
                
                print('File was {2}saved in the directory: {0:04} | File: {1}'.format(int(id_ordem), n_item, 'not ' if not success else ''))

#Img dimensão
#shape.Height	60.8333854675293	float
#shape.Width	156.9049530029297	float

#Img blank
#shape.Height	45.073463439941406	float
#shape.Width	143.1664581298828	float

#shape.Height	45.073463439941406	float
#shape.Width	131.1664581298828	float
