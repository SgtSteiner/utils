# -*- coding: utf8 -*-

import os
import openpyxl
import logging

MASTER_EXCEL = "master.xlsx"


def read_seleccion(doc, sheet_name, ini, fin):
    """ Devuelve una lista con los valores de las celdas de una fila, según la selección
        Entrada: objeto workbook, nombre de la hoja, selección de celdas (celda inicio y celda fin)
        Salida: lista con los valores de las celdas
        """
    logging.debug("Obteniendo datos - " + sheet_name)
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


def write_datos_generales(doc, datos_gen):
    """ Añade los datos generales del candidato al archivo excel
        Entrada: objeto workbook, lista de datos generales
        """
    logging.debug("Escribiendo datos - Datos Generales")
    sheet = doc_master.get_sheet_by_name("Datos Generales")
    for dato_gen in datos_gen:
        col = 1
        row = sheet.max_row+1
        for dato in dato_gen:
            sheet.cell(row=row, column=col).value = dato
            col += 1


def write_experiencia(doc, datos_exp, candidato):
    """ Añade los datos de experiencia del candidato al archivo excel
        Entrada: objeto workbook, lista de datos de experiencia, nombre del candidato
        """
    logging.debug("Escribiendo datos - Experiencia")
    sheet = doc_master.get_sheet_by_name("Experiencia")
    for proyecto in datos_exp:
        if proyecto[0] != None:
            col = 1
            row = sheet.max_row+1
            sheet.cell(row=row, column=col).value = candidato
            for dato in proyecto:
                col += 1
                sheet.cell(row=row, column=col).value = dato


# Define la configuración del archivo de log
logging.basicConfig(filename="import.log",
                    level=logging.DEBUG,
                    format="%(asctime)s:%(levelname)s:%(message)s",
                    filemode="w",
                    )

logging.info("-- INICIO DEL PROCESO --")
logging.info("Abriendo archivo excel master")
doc_master = openpyxl.load_workbook(MASTER_EXCEL)

logging.info("Obteniendo listado de archivos del directorio")
archivos_dir = os.listdir(".")

for archivo in archivos_dir:
    if archivo[-5:] != ".xlsm":
        continue
    logging.info("Abriendo " + archivo)
    file = openpyxl.load_workbook(archivo)

    datos_generales = read_seleccion(file, "Datos Generales", 'A4', 'L4')
    datos_experiencia = read_seleccion(file, "Experiencia Laboral", 'B7', 'G49')
    con_funcional = read_seleccion(file, "Catálogo Cualificaciones", 'B10', 'E298')
    con_tecnico = read_seleccion(file, "Catálogo Cualificaciones", 'H10', 'J108')
    con_prod_santander = read_seleccion(file, "Catálogo Cualificaciones", 'M12', 'N52')
    con_prod_mercado = read_seleccion(file, "Catálogo Cualificaciones", 'Q12', 'R59')
    con_perfil = read_seleccion(file, "Catálogo Cualificaciones", 'U10', 'W126')
    con_idiomas = read_seleccion(file, "Catálogo Cualificaciones", 'Z10', 'AA18')

    # Añade los Datos Generales al archivo master
    write_datos_generales(doc_master, datos_generales)
    nombre_candidato = datos_generales[0][0] + " " + datos_generales[0][1]
    # Añade los Datos de Experiencia al archivo master
    write_experiencia(doc_master, datos_experiencia, nombre_candidato)

logging.info("Cerrando archivo excel master")
doc_master.save(MASTER_EXCEL)
logging.info("-- FIN DEL PROCESO --")
