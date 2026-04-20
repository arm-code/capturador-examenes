import pyautogui
import time

def ejecutar_automatizacion_siosad(estudiantes, modo_ejecucion, logger, confirmador_manual):
    """
    Ejecuta las acciones de PyAutoGUI para una lista de estudiantes ya validados.
    
    :param estudiantes: Lista de diccionarios con datos de estudiantes.
    :param modo_ejecucion: "manual" o "auto"
    :param logger: Función de callback para enviar logs.
    :param confirmador_manual: Función para pedir confirmación al usuario.
    """
    for est in estudiantes:
        matricula_str = est["matricula"]
        nombre_completo = est["nombre"]
        materias = est["materias"]
        sheet_name = est["sede"]
        
        logger(f"-> Procesando: {nombre_completo} ({matricula_str}) | Materias: {len(materias)}")

        # Automatización
        logger("Preparando captura en 2 segundos...")
        pyautogui.press('f9') # Refresh screen
        time.sleep(1)
        
        # Coordenadas y flujo original
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
                f"Revise los datos de {nombre_completo}.\n\n¿Presionar Aceptar para GUARDAR su registro (F2)?")
            if not res:
                logger("✖ Captura omitida o abortada por usuario.")
                return False # Indica que se abortó
        
        pyautogui.press('f2')
        pyautogui.press('enter')
        pyautogui.press('enter')
        
        logger(f"✔ CAPTURA EXITOSA: {nombre_completo}")
        
        if modo_ejecucion == "manual":
            res2 = confirmador_manual("Siguiente", f"¿Continuar con el siguiente alumno de {sheet_name}?")
            if not res2:
                logger("✖ Proceso terminado por el usuario.")
                return False
        
        time.sleep(1)                    
        pyautogui.press('f9')
        
        if modo_ejecucion == "auto":
            time.sleep(1) # Breve pausa para limpiar pantalla
            
    return True # Finalizó la lista con éxito
