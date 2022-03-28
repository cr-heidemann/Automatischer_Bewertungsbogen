# -*- coding: utf-8 -*-
import os
import glob
import pandas as pd
from docxtpl import DocxTemplate
import tkinter as tk
from tkinter import filedialog

def open_file(way, output, entry, template):
    entries=open(entry, encoding='utf8').read().splitlines()
    studierende=[]
    df=pd.read_excel(way)
    df=df.set_index(df.columns[0]).transpose()
    #print(df)
    for column in df:
        studierende.append(column)

    for stud in studierende:     
        data=df[stud].tolist()
        doc = DocxTemplate(template)
        context={}


        for i in range(len(entries)):
            #print(data[i], i, len(entries)-1)
            if isinstance(data[i], float):
                context[str(entries[i])]= '{0:g}'.format(float(data[i]))              
            else:
                context[str(entries[i])]=data[i]
        doc.render(context)
        doc.save(output + '\\' + str(stud) + ".docx")
    return


