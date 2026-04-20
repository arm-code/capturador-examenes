from core.excel_parser import extraer_estudiantes_de_sede
from core.automator import ejecutar_automatizacion_siosad
from api.api_client import verificar_estudiante
import time

def ejecutar_workflow_completo(workbook, config, modo_ejecucion, logger, confirmador_manual, on_finish):
    """
    Orquestador principal que valida y luego captura.
    """
    total_validados = []
    total_rechazados = [] # Estudiantes que no pasaron la validación API
    
    try:
        # FASE 1: EXTRACCIÓN Y VALIDACIÓN
        logger("\n--- FASE 1: VALIDACIÓN DE MATRÍCULAS ---")
        
        for item in config:
            sheet_name = item["sheet"]
            n_estudiantes = item["cant"]
            inicio_estudiante = item["inicio"]
            
            logger(f"\n[Sede: {sheet_name}] Leyendo {n_estudiantes} registros...")
            
            # Extraer del Excel
            estudiantes_brutos = extraer_estudiantes_de_sede(workbook, sheet_name, n_estudiantes, inicio_estudiante)
            
            # Validar con API
            for est in estudiantes_brutos:
                logger(f"Validando {est['matricula']} ({est['nombre']})...")
                exito, info = verificar_estudiante(est['matricula'])
                
                if exito:
                    logger(f"  ✔ Válido: {info.get('Nombre', 'Estudiante')} ({info.get('estado', 'N/A')})")
                    total_validados.append(est)
                else:
                    logger(f"  ✖ Rechazado: {info}")
                    est["error"] = info
                    total_rechazados.append(est)
        
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
            "capturados": len(total_validados) if exito_final else "Parcial"
        })

    except Exception as e:
        logger(f"✖ ERROR CRÍTICO EN WORKFLOW: {e}")
        on_finish(success=False, error_msg=str(e))
