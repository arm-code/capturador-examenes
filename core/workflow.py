from core.excel_parser import extraer_estudiantes_de_sede
from core.automator import ejecutar_automatizacion_siosad
from api.api_client import verificar_estudiante, verificar_materia_aprobada
import time

def ejecutar_workflow_completo(workbook, config, modo_ejecucion, logger, confirmador_manual, on_finish, on_validation_result=None):
    """
    Orquestador principal que valida y luego captura.
    """
    total_validados = []
    total_rechazados = []  # Estudiantes que no existen en la API
    total_omitidos_aprobados = [] # Estudiantes con TODO aprobado
    
    try:
        # FASE 1: EXTRACCIÓN Y VALIDACIÓN
        logger("\n--- FASE 1: VALIDACIÓN DE DATOS (API) ---")
        
        for item in config:
            sheet_name = item["sheet"]
            n_estudiantes = item["cant"]
            inicio_estudiante = item["inicio"]
            
            logger(f"\n[Sede: {sheet_name}] Leyendo {n_estudiantes} registros...")
            
            # Extraer del Excel
            estudiantes_brutos = extraer_estudiantes_de_sede(workbook, sheet_name, n_estudiantes, inicio_estudiante)
            logger(f"  → Extraídos {len(estudiantes_brutos)} registros del Excel.")
            
            if not estudiantes_brutos:
                logger("  ⚠ No se encontraron registros en el rango especificado.")
                continue

            # Validar con API
            for est in estudiantes_brutos:
                matricula = est['matricula']
                logger(f"Validando {matricula} ({est['nombre']})...")
                
                # 1. Verificar existencia del estudiante
                exito, info = verificar_estudiante(matricula)
                if not exito:
                    logger(f"  ✖ Estudiante no válido: {info}")
                    est["error"] = info
                    total_rechazados.append(est)
                    if on_validation_result:
                        on_validation_result({"matricula": matricula, "nombre": est['nombre'], "status": "No encontrado"})
                    continue

                logger(f"  ✔ Estudiante válido: {info.get('Nombre', 'N/A')}")

                # 2. Verificar cada materia
                materias_originales = est['materias']
                materias_a_capturar = []
                materias_ya_aprobadas = []

                for m in materias_originales:
                    logger(f"    - Revisando materia {m}...")
                    if verificar_materia_aprobada(matricula, m):
                        logger(f"      ⚠ Ya aprobada. Se omitirá.")
                        materias_ya_aprobadas.append(m)
                    else:
                        materias_a_capturar.append(m)

                # Actualizar datos del estudiante
                est['materias'] = materias_a_capturar
                est['materias_ya_aprobadas'] = materias_ya_aprobadas

                if not materias_a_capturar:
                    logger(f"  ➡ OMITIDO: {matricula} tiene TODO aprobado ({len(materias_ya_aprobadas)} mat).")
                    total_omitidos_aprobados.append(est)
                    if on_validation_result:
                        on_validation_result({"matricula": matricula, "nombre": est['nombre'], "status": "Aprobado (Omitido)"})
                else:
                    if materias_ya_aprobadas:
                        logger(f"  ✔ Pendiente: Se capturarán {len(materias_a_capturar)} materias (omitiendo {len(materias_ya_aprobadas)} aprobadas).")
                        status = f"Pendiente ({len(materias_a_capturar)} mat)"
                    else:
                        logger(f"  ✔ Pendiente: {len(materias_a_capturar)} materias por capturar.")
                        status = "Pendiente"
                    total_validados.append(est)
                
                # Reportar a la tabla
                if on_validation_result:
                    on_validation_result({
                        "matricula": matricula,
                        "nombre": est['nombre'],
                        "status": status
                    })
        
        # FASE 2: AUTOMATIZACIÓN
        if not total_validados:
            logger("\n======================================")
            logger("⚠ No hay estudiantes válidos para capturar.")
            on_finish(success=True, extra_info={"rechazados": total_rechazados, "capturados": 0})
            return

        logger(f"\n--- FASE 2: CAPTURA AUTOMÁTICA ({len(total_validados)} alumnos) ---")
        if modo_ejecucion == "manual":
            res = confirmador_manual("Iniciar Captura", 
                f"Se validaron {len(total_validados)} estudiantes.\n\n¿Desea iniciar la captura automática ahora?")
            if not res:
                logger("✖ Proceso cancelado antes de iniciar captura.")
                on_finish(success=False, error_msg="Cancelado por usuario")
                return

        # Agrupar por sede si es necesario o procesar la lista plana
        exito_final = ejecutar_automatizacion_siosad(total_validados, modo_ejecucion, logger, confirmador_manual)
        
        # FINALIZACIÓN
        logger("\n======================================")
        if exito_final:
            logger("✅ WORKFLOW FINALIZADO")
        else:
            logger("⚠ WORKFLOW INTERRUMPIDO POR USUARIO")
            
        on_finish(success=exito_final, extra_info={
            "rechazados": total_rechazados,
            "omitidos_aprobados": total_omitidos_aprobados,
            "validados": total_validados, # Contiene info de mat aprobadas internas
            "capturados": len(total_validados) if exito_final else "Parcial"
        })

    except Exception as e:
        logger(f"✖ ERROR CRÍTICO EN WORKFLOW: {e}")
        on_finish(success=False, error_msg=str(e))
