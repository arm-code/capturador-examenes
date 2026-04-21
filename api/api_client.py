import requests
from config import API_BASE_URL, API_CALIFICACIONES_URL

def verificar_estudiante(matricula):
    """
    Consulta la API para verificar si un estudiante existe y está activo.
    Retorna True si existe (200 OK), False si no existe (404) o hay error.
    """
    try:
        url = f"{API_BASE_URL}/{matricula}?tipoDeBusqueda=activos"
        response = requests.get(url, timeout=5)
        
        if response.status_code == 200:
            # La API devuelve una lista si encuentra datos
            data = response.json()
            if isinstance(data, list) and len(data) > 0:
                return True, data[0] # Retorna éxito y los datos del primero
            return False, "No se encontraron datos en la lista."
        
        elif response.status_code == 404:
            return False, "Matrícula no encontrada (404)."
        
        else:
            return False, f"Error de servidor: {response.status_code}"
            
    except requests.exceptions.RequestException as e:
        return False, f"Error de conexión: {str(e)}"

def verificar_materia_aprobada(matricula, materia):
    """
    Verifica si una materia ya está aprobada para un estudiante.
    Retorna True si está aprobada, False en caso contrario.
    """
    try:
        url = f"{API_CALIFICACIONES_URL}?matricula={matricula}&materia={materia}&plan=modular"
        response = requests.get(url, timeout=5)
        
        if response.status_code == 200:
            data = response.json()
            if isinstance(data, list) and len(data) > 0:
                # Revisar si alguna entrada marca estatus_materia como APROBADA
                for entry in data:
                    if entry.get("estatus_materia") == "APROBADA":
                        return True
            return False
        
        return False # No encontrada o error, asumimos que no está aprobada
            
    except Exception:
        return False # Error de conexión, etc. No bloqueamos la captura por esto
