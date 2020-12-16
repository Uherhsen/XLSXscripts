# -*- coding: utf-8 -*-
"""
Created on Mon Dec 14 10:42:42 2020

Определение кода транзисторов по мощности.

"""

import openpyxl, os, shutil, re
from openpyxl import Workbook

def work_sheet(path):
    # Открываем Эксель, помещает данные в переменную sheet
    wb = openpyxl.load_workbook(path)#путь к файлу
    sheet = wb.active
    wb.save(path)
    return sheet

# Функция считает все ячейки первого столбца в которых что то написано,до тех пор пока не встретит пустую ячейку "None"
def number_of_articles(sheet):
    """Функция определяет номер первой пустой ячеки в первом столбце файла xlsx и возвращает его"""
    i = 1
    while sheet['A'+str(i)].value != None:
        i+=1
    return i

def floatconverter(value):
    """преобразование строковых значений в float получает числовые символы, может содержащие один знак / (выбирает число после млеша) и преобразует запятую в точку"""
    if value.find('/') != -1:
        splitvalue = value.split('/')
        value = splitvalue[1]
    value = value.replace(',', '.')
    return float(value)
    
        

def DissipPower(text):
    """считаывет небуквенные символы перед буквами в тексте кВт Вт мВт, переводит во float Вт и возвращает """
    result = re.findall(r'(\d*\S*\d+)Вт', text) #(r'Рассеиваемая мощность (\d*,\d*)Вт', text)
    if result == []:
        result = re.findall(r'(\d*\S*\d+)мВт', text)
        if result == []:
            result = re.findall(r'(\d*\S*\d+)кВт', text)
            if result == []:
                print('Мощность неопределена')
                return None
            else:
                finalresult = floatconverter(result[0]) * 1000 # Переводим кВт в Вт
                return finalresult
        else:
            finalresult = floatconverter(result[0]) * 0.001 # Переводим мВт в Вт
            return finalresult 
    else:
        finalresult = floatconverter(result[0])
        return finalresult # Выводим значение в Вт
    
        
def CodeReturn(value):
    
    if value == None:
        return None
        
    elif value >= 1:
        return "8541290000"
    else:
        return "8541210000"
    

def place_сodes(path):
    sheet0 = work_sheet(path)
    rng = number_of_articles(sheet0)
    # Открываем Эксель, помещает данные в переменную sheet
    wb = openpyxl.load_workbook(path)#путь к файлу
    sheet = wb.active
    for i in range(1,rng):
        text = sheet['A'+str(i)].value
        value = DissipPower(text)
        code = CodeReturn(value)
        #print('код ',code)
        if code == None:
            print(i,value ,'Вт, код: NO CODE (0000000000)')
            sheet['B'+str(i)] = "0000000000"
        else:
            print(i, value ,'Вт, код: ',code)
            sheet['B'+str(i)] = code
    wb.save(path)
    print('Коды транзисторов расставлены.')
        
path = 'tranzistorsCode.xlsx'        
place_сodes(path)
#text = 'Кремниевые МОП-транзисторы с N-канальной структурой. Тип транзистора N-MOSFET, Полярность полевой, Напряжение сток-исток 650В, Ток стока 10А, Рассеиваемая мощность 275Вт, Корпус TO220F, Напряжение затвор-исток +\- 30В, Сопротивление в открытом состоянии 0,63Ом, Монтаж в отверстия печатной платы, Заряд затвора 45нC, рабочие температуры от -40 до 85°С, предназначены для использования в радиоэлектронном оборудовании промышленного назначения.'
#print(DissipPower(text))
#print(floatconverter('535,6'))





