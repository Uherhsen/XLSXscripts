# -*- coding: utf-8 -*-
"""
Created on Fri Jan 15 13:58:49 2021

Скопировать из буфера обмена

"""
from tkinter import Tk
import pyperclip

text=Tk().clipboard_get()

text = text.replace('\t\n',', ')
text = text.replace('\n',',')
text = text.replace('\t',' ')
text = text.replace(':','')

replacements = [['us','мкс'],['Ohms','Ом'],['W','Вт'],['Mb/s','Мб/с'],[' Receiver',''],[' Driver',''],
                ['kb/s','кб/с'],['MHz','МГц'],['pF','пФ'],['GHz','ГГц'],['bit','бит'],['dB','дБ'],['to','...'],
                ['kS/s','кГц'],[' V', ' В'],['mm','мм'],['kV','кВ'],['Минимальная рабочая температура','рабочая температура от'],
                [', Максимальная рабочая температура',' до'],['Упаковка / блок','тип корпуса'],['ns,','нс,'],[' mA',' мА'],['дБm','дБм'],
['uA','мкА'],['nA','нA'],[' Input',''],[' Output',''],['Pd - ',''],['ms','мс'],[' Channel',''],[' I/O',''],['kB','кБайт'],['Втire','wire']]

def allreplacements(text, replacements):
    for n in replacements:
        text = text.replace(n[0],n[1])
        
    return text 

text = allreplacements(text, replacements)

print('{!r}'.format(text))

with open('textfile.txt', 'w', encoding='utf-8') as g:
    g.write(text)
    g.close()

pyperclip.copy(text)