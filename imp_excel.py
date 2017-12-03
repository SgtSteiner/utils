# -*- coding: utf8 -*-

import os
import openpyxl
import logging
import tkinter as tk
from tkinter import scrolledtext
from tkinter import ttk
from tkinter import messagebox as mBox

MASTER_EXCEL = "master.xlsx"


class Application(ttk.Frame):

    def __init__(self, main_win):
        super().__init__(main_win)

        main_win.geometry("600x390")
        self.place(relwidth=1, relheight=1)

        main_win.title("Importación de CV ISBAN")

        self.labelframe_Top = ttk.LabelFrame(self)
        self.labelframe_Top.place(x=5, y=5, relwidth=0.98, height=65)
        self.labelTop = tk.Label(self.labelframe_Top, text="Importación de archivos CV ISBAN")
        self.labelTop.pack()
        self.labelPath = tk.Label(self.labelframe_Top, text="Dir: " + os.getcwd())
        self.labelPath.pack()

        self.labelframe_arch = ttk.LabelFrame(self, text="Archivos")
        self.labelframe_arch.place(x=5, y=80, relwidth=0.98, height=80)
        self.labelImport = ttk.Label(self.labelframe_arch, text="Procesando: ")
        self.labelImport.place(x=5, y=5)
        self.labelCV = ttk.Label(self.labelframe_arch)
        self.labelCV.place(x=80, y=5)
        self.progressbar = ttk.Progressbar(self.labelframe_arch, length=580, maximum=101)
        self.progressbar.place(x=5, y=30)

        self.labelframe_Detalle = ttk.LabelFrame(self, text="Detalle")
        self.labelframe_Detalle.place(x=5, y=165, relwidth=0.98, height=180)
        self.scr_Detalle = scrolledtext.ScrolledText(self.labelframe_Detalle, width=78, height=10,
                                                     font=('courier', 8, 'normal'))
        self.scr_Detalle.pack()
        # self.labelImportCV = ttk.Label(self.labelframe_Detalle, text="Procesando: ")
        # self.labelImportCV.place(x=5, y=5)
        # self.labelCV_cualif = ttk.Label(self.labelframe_Detalle)
        # self.labelCV_cualif.place(x=80, y=5)
        # self.progressbar_Cualif = ttk.Progressbar(self.labelframe_Detalle, length=380, maximum=101)
        # self.progressbar_Cualif.place(x=5, y=30)

        self.inicio_button = ttk.Button(self, text="Inicio", command=self.inicio_button_clicked)
        self.inicio_button.place(x=275, y=355)

    def inicio_button_clicked(self):

        if self.inicio_button["text"] == "Salir":
            self.quit()
            return

        # Define la configuración del archivo de log
        logging.basicConfig(filename="import.log",
                            level=logging.INFO,
                            format="%(asctime)s:%(levelname)s:%(message)s",
                            filemode="w",
                            )

        logging.info("-- INICIO DEL PROCESO --")
        logging.info("Abriendo archivo excel master %s", MASTER_EXCEL)
        try:
            doc_master = openpyxl.load_workbook(MASTER_EXCEL)
        except FileNotFoundError:
            logging.error("No existe el archivo " + MASTER_EXCEL, exc_info=True)
            raise
        logging.info("Obteniendo listado de archivos del directorio")
        archivos_dir = os.listdir(".")

        num_archivos = 0
        for archivo in archivos_dir:
            num_archivos += 1
            self.act_progress(cv_name=archivo, estado_cv=num_archivos * 100 / len(archivos_dir))
            if archivo[-5:].upper() != ".XLSM":
                continue
            logging.info("Abriendo %s", archivo)

            file = openpyxl.load_workbook(archivo)

            # Lectura de datos del candidato
            self.act_progress(cv_name=archivo, cualif_name="Leyendo Datos Generales")
            datos_generales = self.read_seleccion(file, "Datos Generales", 'A4', 'L4')
            self.act_progress(cv_name=archivo, cualif_name="Leyendo Experiencia Laboral")
            datos_experiencia = self.read_seleccion(file, "Experiencia Laboral", 'B7', 'G49')
            self.act_progress(cv_name=archivo, cualif_name="Leyendo Cualificaciones")
            con_funcional = self.read_seleccion(file, "Catálogo Cualificaciones", 'B10', 'E298')
            con_tecnico = self.read_seleccion(file, "Catálogo Cualificaciones", 'H10', 'J108')
            con_prod_santander = self.read_seleccion(file, "Catálogo Cualificaciones", 'M12', 'N52')
            con_prod_mercado = self.read_seleccion(file, "Catálogo Cualificaciones", 'Q12', 'R59')
            con_perfil = self.read_seleccion(file, "Catálogo Cualificaciones", 'U10', 'W126')
            con_idiomas = self.read_seleccion(file, "Catálogo Cualificaciones", 'Z10', 'AA18')

            # Añade los Datos Generales al archivo master
            self.act_progress(cv_name=archivo, cualif_name="Escribiendo Datos Generales")
            self.write_datos_generales(doc_master, datos_generales)
            nombre_candidato = datos_generales[0][0] + " " + datos_generales[0][1]
            # Añade los Datos de Experiencia al archivo master
            self.act_progress(cv_name=archivo, cualif_name="Escribiendo Experiencia Laboral")
            self.write_experiencia(doc_master, datos_experiencia, nombre_candidato)
            # Añade los Datos de Cualificación al archivo master
            self.act_progress(cv_name=archivo, cualif_name="Escribiendo Cualificaciones")
            self.write_cualificacion(doc_master, con_funcional, "Funcional", nombre_candidato)
            self.write_cualificacion(doc_master, con_tecnico, "Técnico", nombre_candidato)
            self.write_cualificacion(doc_master, con_prod_santander, "Prod_Santander", nombre_candidato)
            self.write_cualificacion(doc_master, con_prod_mercado, "Prod_Mercado", nombre_candidato)
            self.write_cualificacion(doc_master, con_perfil, "Perfil", nombre_candidato)
            self.write_cualificacion(doc_master, con_idiomas, "Idiomas", nombre_candidato)

        logging.info("Cerrando archivo excel master %s", MASTER_EXCEL)
        try:
            doc_master.save(MASTER_EXCEL)
        except PermissionError:
            logging.error('Fallo al grabar el archivo excel. Archivo ocupado o Permiso denegado', exc_info=True)
            raise
        self.inicio_button["text"] = "Salir"
        logging.info("Archivo cerrado")
        logging.info("-- FIN DEL PROCESO --")
        mBox.showinfo("Proceso finalizado",
                      "Importación realizada con éxito\nImportados " + str(num_archivos) + " archivos")

    def quit(self):
        main_win.quit()
        main_win.destroy()
        exit()

    def act_progress(self, cv_name, estado_cv=None, cualif_name=None):
        self.labelCV["text"] = cv_name
        if estado_cv is not None:
            self.progressbar.step(estado_cv)
        if cualif_name is not None:
            self.scr_Detalle.insert(tk.INSERT, cv_name + " " + cualif_name + "\n")
        main_win.update()

    def read_seleccion(self, doc, sheet_name, ini, fin):
        """ Devuelve una lista con los valores de las celdas de una fila, según la selección
            Entrada: objeto workbook, nombre de la hoja, selección de celdas (celda inicio y celda fin)
            Salida: lista con los valores de las celdas
            """
        logging.info("Obteniendo datos - %s", sheet_name)
        cualif = []
        sheet = doc.get_sheet_by_name(sheet_name)
        # Importamos el conocimiento funcional
        selection = sheet[ini:fin]
        for rows in selection:
            datos = []
            for column in rows:
                datos.append(column.value)
            cualif.append(datos)

        return cualif

    def write_datos_generales(self, doc, datos_gen):
        """ Añade los datos generales del candidato al archivo excel
            Entrada: objeto workbook, lista de datos generales
            """
        logging.info("Escribiendo datos - Datos Generales")
        sheet = doc.get_sheet_by_name("Datos Generales")
        for dato_gen in datos_gen:
            logging.debug("Escribiendo datos - Datos Generales: %s", dato_gen)
            col = 1
            row = sheet.max_row + 1
            for dato in dato_gen:
                sheet.cell(row=row, column=col).value = dato
                col += 1

    def write_experiencia(self, doc, datos_exp, candidato):
        """ Añade los datos de experiencia del candidato al archivo excel
            Entrada: objeto workbook, lista de datos de experiencia, nombre del candidato
            """
        logging.info("Escribiendo datos - Experiencia")
        sheet = doc.get_sheet_by_name("Experiencia")
        for proyecto in datos_exp:
            if proyecto[0] is not None:
                logging.debug("Escribiendo datos - Experiencia: %s", proyecto)
                col = 1
                row = sheet.max_row + 1
                sheet.cell(row=row, column=col).value = candidato
                for dato in proyecto:
                    col += 1
                    sheet.cell(row=row, column=col).value = dato

    def write_cualificacion(self, doc, catalogos, tipo_catalogo, candidato):
        """ Añade los datos de cualificación del candidato al archivo excel
            Entrada: objeto workbook, lista de cualificaciones, tipo de catalogo, nombre del candidato
            """
        logging.info("Escribiendo datos - %s", tipo_catalogo)
        sheet = doc.get_sheet_by_name("Cualificación")
        conocimiento = ""
        area = ""
        for catalogo in catalogos:
            logging.debug("Escribiendo datos - " + tipo_catalogo + ": " + str(catalogo))
            row = sheet.max_row + 1
            if tipo_catalogo == "Funcional":
                if catalogo[0] is not None:
                    conocimiento = catalogo[0]
                if catalogo[1] is not None:
                    area = conocimiento + " / " + catalogo[1]
            elif tipo_catalogo == "Técnico" or tipo_catalogo == "Perfil":
                if catalogo[0] is not None:
                    conocimiento = catalogo[0]
            if catalogo[-1] is not None:
                sheet.cell(row=row, column=1).value = candidato
                sheet.cell(row=row, column=2).value = tipo_catalogo
                if area:
                    sheet.cell(row=row, column=3).value = area
                else:
                    sheet.cell(row=row, column=3).value = conocimiento
                sheet.cell(row=row, column=4).value = catalogo[-2]
                sheet.cell(row=row, column=5).value = catalogo[-1]


main_win = tk.Tk()
app = Application(main_win)
app.mainloop()