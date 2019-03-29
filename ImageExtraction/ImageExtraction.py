#!/usr/bin/python3
# -*- coding: utf-8 -*-
import sys
import datetime
import time
import os
from Constant import *
from FileManager import *

if __name__ == "__main__":
    print("::         Inicio do processamento         ::")
    print("---------------------------------------------")

    dir = os.getcwd()
    path_source = "{0}\\{1}\\".format(dir, "Source")
    path_destination = "{0}\\{1}\\".format(dir, "Destination")

    filelist = ListFiles(path_source)
    qt = filelist.__len__

    for item in filelist:
        wb = GetWorkbookExcelFile("{0}{1}".format(path_source, item))
        SaveFileImage(path_destination, wb)

    print("---------------------------------------------")
    print("::          Fim do processamento           ::")
