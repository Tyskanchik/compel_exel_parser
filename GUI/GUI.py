from typing import Optional, Tuple, Union
import customtkinter
import tkinter
import os
import exel_parser

class App(customtkinter.CTk):

    customtkinter.set_appearance_mode("dark")
    customtkinter.set_default_color_theme("green")
    
    def __init__(self, fg_color: str | Tuple[str, str] | None = None, **kwargs):
        super().__init__(fg_color, **kwargs)

        self.folder_exel = "change files" # Путь файлов для изменения

        self.geometry("1024x768")
        self.title("Exel_parser")
        self.grid_columnconfigure(0,weight=1)

        self.xlsx_files_list = exel_parser.get_xlsx_list(dir=self.folder_exel)

        # Создание областей
        self.choose_frame = customtkinter.CTkFrame(self,width=600,height=600)
        self.choose_frame.grid(row=0,column=0,padx=0,pady=10,sticky="ns")

        self.button_frame = customtkinter.CTkFrame(self,width=600,height=300)
        self.button_frame.grid(row=1,column=0,padx=0,pady=10,sticky="ns")
        
        # Создание подписей к выпадающим меню
        self.alt_label = customtkinter.CTkLabel(self.choose_frame,
                                                text="Выберите файл с Altium",
                                                font=customtkinter.CTkFont(size=16,weight="bold"))
        self.alt_label.grid(row=0,column=0,padx=10,pady=10)

        self.sds_label = customtkinter.CTkLabel(self.choose_frame,
                                                text="Выберите файл с Compel",
                                                font=customtkinter.CTkFont(size=16,weight="bold"))
        self.sds_label.grid(row=0,column=1,padx=10,pady=10)

        # Создание выпадающих меню для выбора файла
        self.alt_bom_name = customtkinter.StringVar(value=self.xlsx_files_list[0])
        self.choose_alt_bom = customtkinter.CTkComboBox(self.choose_frame,
                                                     width=300,
                                                     height=40,
                                                     values=self.xlsx_files_list,
                                                     variable=self.alt_bom_name)
        self.choose_alt_bom.grid(row=1,column=0,padx=30,pady=0)

        self.sds_bom_name = customtkinter.StringVar(value=self.xlsx_files_list[0])
        self.choose_sds_bom = customtkinter.CTkComboBox(self.choose_frame,
                                                     width=300,
                                                     height=40,
                                                     values=self.xlsx_files_list,
                                                     variable=self.sds_bom_name)
        self.choose_sds_bom.grid(row=1,column=1,padx=30,pady=0)

        # Создание поля вводя количества плат
        self.qty_pcbs = customtkinter.CTkEntry(self.choose_frame,
                                               width=300,
                                               height=30,
                                               placeholder_text="Количество плат")
        self.qty_pcbs.grid(row=2,column=0,padx=10,pady=40)
        
        # Создание кнопки активации парсинга
        self.parse_button = customtkinter.CTkButton(self.button_frame,text="Преобразовать",command=self.button_change)
        self.parse_button.grid(row=0,column=0,padx=20,pady=10)

        # Создание кнопки обновления папки на наличие файлов
        self.update_button = customtkinter.CTkButton(self.button_frame,text="Обновить",command=self.button_update)
        self.update_button.grid(row=0,column=1,padx=20,pady=10)

    # Кнопка отвечает за выполнение основной функции преобразования файлов
    def button_change(self):
        try:
            if int(self.qty_pcbs.get())>1:
                os.chdir(self.folder_exel)
                exel_parser.exel_transform(sds_exel_name=self.sds_bom_name.get(),
                                           altium_exel_name=self.alt_bom_name.get(),
                                           qty_pcbs=int(self.qty_pcbs.get()))
            else:
                print("Введите количество плат больше 0!")
            print("Файлы успешно преобразованы!")
        except ValueError as e:
            print("Некорректное значение количества плат!")

    # Обновление выпадающих меню при нажатии кнопки
    def button_update(self):
        self.xlsx_files_list = exel_parser.get_xlsx_list(dir=self.folder_exel)
        self.choose_alt_bom.configure(values=self.xlsx_files_list)
        self.choose_sds_bom.configure(values=self.xlsx_files_list)
        print("Folder update!")

