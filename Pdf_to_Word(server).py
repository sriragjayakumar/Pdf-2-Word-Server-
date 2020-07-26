# -*- coding: utf-8 -*-
"""
Created on Sun Jul 26 10:15:21 2020

@author: srira
"""

import autoit
import time

program='Winword.EXE' # Find the winword.exe path in program files for MS Word
pdf='Pdf_name.pdf'

autoit.run(program+' '+pdf)
while True:
    try:
    
        autoit.control_send("[Class:OpusApp]","_WwG1","{F12}",0)
        time.sleep(2)
        autoit.control_click("Save As",'Button8')
        autoit.win_close("[CLASS:OpusApp]")
        break
    except: continue