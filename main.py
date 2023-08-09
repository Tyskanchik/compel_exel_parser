import os, sys
from pathlib import Path
from GUI import App

"""
________________________________________________
Структура таблиц
________________________________________________

Эксель СДС Компэл
|Имя строки|Наименование|Наименование товара|Производитель товара|Позиционное обозначение|Корпус|Примечание|Кол-во в изделии|Подобрано предложение|Приоритетный товар|Общее кол-во на партию|
|----------|------------|-------------------|--------------------|-----------------------|------|----------|----------------|---------------------|------------------|----------------------|

Эксель Альтиум
|Name|Description|Footprint|Designator|Quantity|Название/закуплено|Количество|Примечание|
|----|-----------|---------|----------|--------|------------------|----------|----------|
"""

if __name__=="__main__":

    try:

        root = App()
        root.mainloop() 

    except Exception as e:

        print("Somthing wrong!!!\n",e)