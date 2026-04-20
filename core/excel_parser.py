import openpyxl

def extraer_estudiantes_de_sede(workbook, sheet_name, n_estudiantes, inicio_fila_relativa):
    """
    Lee los datos de los estudiantes de una hoja específica de Excel.
    
    :param workbook: Objeto de openpyxl.
    :param sheet_name: Nombre de la hoja (sede).
    :param n_estudiantes: Cantidad de alumnos a procesar.
    :param inicio_fila_relativa: El índice relativo de inicio (empieza en 1 según la GUI).
    :return: Lista de diccionarios con datos de estudiantes.
    """
    estudiantes = []
    sheet = workbook[sheet_name]
    
    # La lógica original dice que los encabezados están en la fila 11.
    # Los datos reales empiezan en la fila 12.
    
    # Ajuste de índices: si inicio es 1, queremos la fila 12.
    start_row = 12 + (inicio_fila_relativa - 1)
    end_row = start_row + n_estudiantes
    
    # Leer columnas necesarias en el rango de filas
    # Columna 3 en adelante (según ajuste del usuario)
    
    columnas_matricula = list(sheet.iter_cols(min_col=2, max_col=13, min_row=12, max_row=end_row))
    columnas_nombres = list(sheet.iter_cols(min_col=14, max_col=16, min_row=12, max_row=end_row))
    columnas_materias = list(sheet.iter_cols(min_col=17, max_col=20, min_row=12, max_row=end_row))

    # El offset es porque iter_cols con min_row=12 devuelve celdas donde el índice 0 corresponde a la fila 12
    for i in range(inicio_fila_relativa - 1, inicio_fila_relativa - 1 + n_estudiantes):
        try:
            # Extraer Matrícula
            matricula_parts = [str(columnas_matricula[c][i].value) for c in range(len(columnas_matricula))]
            matricula_str = "".join(matricula_parts).replace("None", "").strip()
            
            if not matricula_str:
                continue

            # Extraer Nombre
            nombres_parts = [str(columnas_nombres[c][i].value) for c in range(len(columnas_nombres))]
            nombre_completo = " ".join(nombres_parts).replace("None", "").strip()

            # Extraer Materias
            materias = [str(columnas_materias[c][i].value) for c in range(len(columnas_materias)) 
                       if columnas_materias[c][i].value and str(columnas_materias[c][i].value) != 'None']

            estudiantes.append({
                "matricula": matricula_str,
                "nombre": nombre_completo,
                "materias": materias,
                "fila_excel": 12 + i,
                "sede": sheet_name
            })
        except IndexError:
            # Fin de los datos alcanzado antes de lo esperado
            break
            
    return estudiantes
