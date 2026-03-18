import pyautogui
import time

def ejecutar_captura_siosad(workbook, config, modo_ejecucion, logger, confirmador_manual, on_finish):
    """
    Núcleo del automatizador SIOSAD.
    
    :param workbook: Objeto de openpyxl cargado.
    :param config: Lista de diccionarios con la configuración de Sedes.
    :param modo_ejecucion: "manual" o "auto"
    :param logger: Función de callback para enviar prints a la Consola GUI.
    :param confirmador_manual: Función (titulo, mensaje) que retorna True/False (askokcancel).
    :param on_finish: Función a ejecutar cuando el proceso termina con éxito o error.
    """
    try:
        for item in config:
            sheet_name = item["sheet"]
            n_estudiantes = item["cant"]
            inicio_estudiante = item["inicio"] + 1
            limite_filas = n_estudiantes + 11
            
            sheet = workbook[sheet_name]
            logger(f"\n[SEDE: {sheet_name}] - ({n_estudiantes} estudiante/s)")

            # Leer rangos
            columnas_a_leer = list(sheet.iter_cols(min_col=2, max_col=13, min_row=11, max_row=limite_filas))
            datos_estudiante = list(sheet.iter_cols(min_col=14, max_col=16, min_row=11, max_row=limite_filas))
            columnas_materias = list(sheet.iter_cols(min_col=17, max_col=20, min_row=11, max_row=limite_filas))

            # Procesamiento
            for row_idx in range(inicio_estudiante, limite_filas - 1):
                try:
                    # Extraer Datos
                    matricula_lista = [str(columnas_a_leer[c][row_idx - 1].value) for c in range(len(columnas_a_leer))]
                    matricula_str = "".join(matricula_lista).replace("None", "")
                    
                    nombres_lista = [str(datos_estudiante[c][row_idx - 1].value) for c in range(len(datos_estudiante))]
                    nombre_completo = " ".join(nombres_lista).replace("None", "").strip()

                    materias = [str(columnas_materias[c][row_idx - 1].value) for c in range(len(columnas_materias)) 
                               if columnas_materias[c][row_idx - 1].value and str(columnas_materias[c][row_idx - 1].value) != 'None']

                    logger(f"-> Registro {row_idx - 1}: {nombre_completo} ({matricula_str}) | Materias: {len(materias)}")

                    # Automatización
                    logger("Preparando captura en 2 segundos...")
                    pyautogui.press('f9') # Refresh screen as requested
                    time.sleep(1)
                    
                    pyautogui.click(x=150, y=150)
                    pyautogui.write(matricula_str)
                    pyautogui.press('enter')
                    pyautogui.press('enter')
                    pyautogui.write('404') 
                    pyautogui.press('enter')
                    pyautogui.write('S')
                    pyautogui.write(str(sheet_name))
                    pyautogui.press('enter')
                    pyautogui.press('enter')

                    for materia in materias:
                        pyautogui.write(materia)
                        pyautogui.press('enter')

                    # Confirmación Manual vía GUI Callback
                    if modo_ejecucion == "manual":
                        res = confirmador_manual("Confirmación de Guardado", 
                            "Revise los datos en la ventana de SIOSAD.\n\n¿Presionar Aceptar para GUARDAR su registro (F2)?")
                        if not res:
                            logger("✖ Captura omitida/abortada.")
                            on_finish()
                            return
                    
                    pyautogui.press('f2')
                    pyautogui.press('enter')
                    pyautogui.press('enter')
                    
                    logger("✔ CAPTURA EXITOSA")
                    
                    if modo_ejecucion == "manual":
                        res2 = confirmador_manual("Siguiente", "¿Continuar con el siguiente alumno?")
                        if not res2:
                            logger("✖ Proceso terminado por el usuario.")
                            on_finish()
                            return
                    
                    time.sleep(1)                    
                    pyautogui.press('f9')
                    
                    if modo_ejecucion == "auto":
                        time.sleep(1) # Breve pausa para limpiar pantalla
                        
                except Exception as e:
                    # La manera que el codigo original maneja las listas vacías
                    logger(f"Aviso en registro {row_idx}: {e}")
                    break

        logger("\n======================================")
        logger("✅ PROCESO TOTAL FINALIZADO")
        # El success messagebox se llama desde la GUI original (lo pasamos por callback o el GUI lo detecta, o ponemos un extra success_callback)
        # Para mantener simple, on_finish se ejecutará (que habilita botones) y asume un éxito global si llega hasta el final sin "errors criticos"
        on_finish(success=True)
        
    except Exception as e:
        logger(f"✖ ERROR CRÍTICO: {e}")
        on_finish(success=False, error_msg=str(e))
