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
    sheet = doc.get_sheet_by_name("Datos Generales")
    for dato_gen in datos_gen:
        col = 1
        row = sheet.max_row + 1
        for dato in dato_gen:
            sheet.cell(row=row, column=col).value = dato
            col += 1


def write_experiencia(doc, datos_exp, candidato):
    """ Añade los datos de experiencia del candidato al archivo excel
        Entrada: objeto workbook, lista de datos de experiencia, nombre del candidato
        """
    logging.debug("Escribiendo datos - Experiencia")
    sheet = doc.get_sheet_by_name("Experiencia")
    for proyecto in datos_exp:
        if proyecto[0] is not None:
            col = 1
            row = sheet.max_row + 1
            sheet.cell(row=row, column=col).value = candidato
            for dato in proyecto:
                col += 1
                sheet.cell(row=row, column=col).value = dato


def write_cualificacion(doc, catalogos, tipo_catalogo, candidato):
    """ Añade los datos de cualificación del candidato al archivo excel
        Entrada: objeto workbook, lista de cualificaciones, tipo de catalogo, nombre del candidato
        """
    logging.debug("Escribiendo datos - " + tipo_catalogo)
    sheet = doc.get_sheet_by_name("Cualificación")
    conocimiento = ""
    area = ""
    for catalogo in catalogos:
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
    if archivo[-5:].upper() != ".XLSM":
        continue
    logging.info("Abriendo " + archivo)

    file = openpyxl.load_workbook(archivo)

    # Lectura de datos del candidato
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
    # Añade los Datos de Cualificación al archivo master
    write_cualificacion(doc_master, con_funcional, "Funcional", nombre_candidato)
    write_cualificacion(doc_master, con_tecnico, "Técnico", nombre_candidato)
    write_cualificacion(doc_master, con_prod_santander, "Prod_Santander", nombre_candidato)
    write_cualificacion(doc_master, con_prod_mercado, "Prod_Mercado", nombre_candidato)
    write_cualificacion(doc_master, con_perfil, "Perfil", nombre_candidato)
    write_cualificacion(doc_master, con_idiomas, "Idiomas", nombre_candidato)

logging.info("Cerrando archivo excel master")
doc_master.save(MASTER_EXCEL)
logging.info("-- FIN DEL PROCESO --")
