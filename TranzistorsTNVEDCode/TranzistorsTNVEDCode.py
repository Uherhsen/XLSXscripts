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

def DissipPower(text):
    
    result = re.findall(r'(\d*\S*\d+)Вт', text) #(r'Рассеиваемая мощность (\d*,\d*)Вт', text)
    if result != []: 
        try:
            result = result[0].replace(',', '.') #требуется разделить строки по символу /
            #print('после реплейса ',result)
            return float(result)
        except ValueError:
            try:
                splires = result.split('/')
                splires = splires[1]
                splires = splires.replace(',', '.')
                return float(splires)
            except Exception as e:
                print('Ошибка, ', e)
                return None 
            
        except IndexError:
            print('Мощность неопределена')
            return None
    else:
        try:
            result = re.findall(r'(\d*\S*\d+)мВт', text)
            result = result[0].replace(',', '.')
            floatresult = float(result)
            Wtresult  = floatresult * 0.001
            return Wtresult
        except ValueError as e:
            print('Ошибка, ', e)
            return None    
        except IndexError:
            print('Мощность неопределена')
            return None
        
        
        
        
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
            print(i,'код ','NO CODE')
            sheet['B'+str(i)] = "0000000000"
        else:
            print(i,'код ',code)
            sheet['B'+str(i)] = code
    wb.save(path)
    print('Коды транзисторов расставлены.')
        
path = 'tranzistorsCode.xlsx'        
place_сodes(path)
#text = 'Кремниевые МОП-транзисторы с N-канальной структурой. Тип транзистора N-MOSFET, Полярность полевой, Напряжение сток-исток 650В, Ток стока 10А, Рассеиваемая мощность 275Вт, Корпус TO220F, Напряжение затвор-исток +\- 30В, Сопротивление в открытом состоянии 0,63Ом, Монтаж в отверстия печатной платы, Заряд затвора 45нC, рабочие температуры от -40 до 85°С, предназначены для использования в радиоэлектронном оборудовании промышленного назначения.'
#print(DissipPower(text))




