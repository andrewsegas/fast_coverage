#!/usr/bin/env python
# -*- coding: utf-8 -*-
from os import listdir
import os
import xlwt
from datetime import datetime

nlinhas = 0
nacumulado_fonte = 0
num = 0
lcoment_aberto = False

wb = xlwt.Workbook()
ws = wb.add_sheet('Fontes')
style0 = xlwt.easyxf('font: name Arial, color-index red, bold on')
ws.write(0, 0, "Linhas seguidas",style0)
ws.write(0, 1, "Ate a linha",style0)
ws.write(0, 2, "Fonte",style0)
ws.write(0, 3, "Acumulado por fonte",style0)
ws.write(0, 4, "Caminho",style0)
nwrite = 1

for subdir, dirs, files in os.walk("C:\\workspace\\Fontes\\Materiais"):
    for file in files:
        #print os.path.join(subdir, file)
        filepath = subdir + os.sep + file
        nacumulado_fonte = 0  # salva o acumulado por fonte
        if filepath.upper().endswith(".PRW") or filepath.upper().endswith(".PRX"):
            with open(filepath, errors='ignore') as currentFile:
                nlinhas = 0
                text = currentFile.read()
                atext = text.split("\n")
                print(file)
                for at in range(0,len(atext)):
                    if '/*' in atext[at] or lcoment_aberto:  # se tem comentario de bloco
                        lcoment_aberto = True
                        if '*/' in atext[at]:
                            lcoment_aberto = False
                        continue
                    if atext[at].strip() != '' and \
                            not (atext[at].replace('\t','').strip().startswith('//')):

                        ctext_cleaned = atext[at].replace('\t','').strip().upper()
                        if not (ctext_cleaned.startswith('FUNCTION') or
                                ctext_cleaned.startswith('WHILE ') or
                                ctext_cleaned.startswith('IF ') or
                                ctext_cleaned.startswith('ELSE') or
                                ctext_cleaned.startswith('ENDIF') or
                                ctext_cleaned.startswith('ENDDO') or
                                ctext_cleaned.startswith('FOR ') or
                                ctext_cleaned.startswith('DO CASE') or
                                ctext_cleaned.startswith('CASE ') or
                                ctext_cleaned.startswith('RETURN') or
                                ctext_cleaned.startswith('FUNCTION ') or
                                ctext_cleaned.startswith('STATIC FUNCTION ') or
                                ctext_cleaned.startswith('NEXT') or
                                ctext_cleaned.startswith('CLASS ') or
                                ctext_cleaned.startswith('METHOD ') or
                                ctext_cleaned.startswith('WSMETHOD ') or
                                ctext_cleaned.startswith('WSSTRUCT ') or
                                ctext_cleaned.startswith('USER FUNCTION ')):
                            nlinhas = nlinhas + 1
                        else:
                            if nlinhas > 20:
                                ws.write(nwrite, 0, nlinhas)
                                ws.write(nwrite, 1, at)
                                ws.write(nwrite, 2, file)
                                ws.write(nwrite, 3, nacumulado_fonte)
                                ws.write(nwrite, 4, filepath)

                                nwrite = nwrite + 1
                                nacumulado_fonte = nacumulado_fonte + nlinhas
                            num = atext[at].count("\t")
                            nlinhas = 0

wb.save('C:/workspace/searching/searching/result.xls')