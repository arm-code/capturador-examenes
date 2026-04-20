import customtkinter as ctk
from tkinter import filedialog, messagebox
import openpyxl
import pyautogui
import time
import threading
import sys
from core.workflow import ejecutar_workflow_completo

# Configuración de apariencia
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("SIOSAD - Capturador Automático (GUI)")
        self.geometry("650x700")

        self.workbook = None
        self.excel_file = ""
        self.sheet_widgets = []
        self.is_running = False

        self.create_widgets()

    def create_widgets(self):
        # Frame superior
        self.top_frame = ctk.CTkFrame(self)
        self.top_frame.pack(pady=10, padx=10, fill="x")

        self.btn_load = ctk.CTkButton(self.top_frame, text="1. Seleccionar Archivo Excel", command=self.load_excel, fg_color="#3B82F6", hover_color="#2563EB")
        self.btn_load.pack(side="left", padx=10, pady=10)

        self.lbl_file = ctk.CTkLabel(self.top_frame, text="Ningún archivo seleccionado", text_color="gray")
        self.lbl_file.pack(side="left", padx=10, pady=10)

        # Modo de Ejecución
        self.modo_frame = ctk.CTkFrame(self.top_frame, fg_color="transparent")
        self.modo_frame.pack(side="right", padx=10)

        self.lbl_modo = ctk.CTkLabel(self.modo_frame, text="Modo:", font=ctk.CTkFont(weight="bold"))
        self.lbl_modo.pack(side="left", padx=5)

        self.modo_var = ctk.StringVar(value="manual")
        self.radio_manual = ctk.CTkRadioButton(self.modo_frame, text="Manual", variable=self.modo_var, value="manual")
        self.radio_manual.pack(side="left", padx=5)
        
        self.radio_auto = ctk.CTkRadioButton(self.modo_frame, text="Auto", variable=self.modo_var, value="auto")
        self.radio_auto.pack(side="left", padx=5)

        # Configuración de Sedes (Scrollable Frame)
        self.lbl_sedes = ctk.CTkLabel(self, text="2. Configurar Sedes a Procesar", font=ctk.CTkFont(size=14, weight="bold"))
        self.lbl_sedes.pack(pady=(10, 0))

        self.scroll_frame = ctk.CTkScrollableFrame(self, height=200)
        self.scroll_frame.pack(pady=5, padx=10, fill="x")
        
        # Label informativo inicial en el frame
        self.lbl_sedes_info = ctk.CTkLabel(self.scroll_frame, text="Cargue un archivo Excel primero para ver las sedes disponibles.", text_color="gray")
        self.lbl_sedes_info.pack(pady=20)

        # Botón Iniciar
        self.btn_start = ctk.CTkButton(self, text="3. INICIAR CAPTURA", fg_color="#10B981", hover_color="#059669", font=ctk.CTkFont(size=16, weight="bold"), command=self.start_capture)
        self.btn_start.pack(pady=15)

        # Consola de Logs
        self.lbl_logs = ctk.CTkLabel(self, text="Consola de Estado", font=ctk.CTkFont(size=14, weight="bold"))
        self.lbl_logs.pack()

        self.textbox = ctk.CTkTextbox(self, height=200, state="disabled", fg_color="#1E1E1E", text_color="#10B981")
        self.textbox.pack(pady=(5, 15), padx=10, fill="both", expand=True)

    def log(self, text):
        # Aseguramos que se ejecute en el hilo principal
        self.after(0, self._log, text)
        
    def _log(self, text):
        self.textbox.configure(state="normal")
        self.textbox.insert("end", text + "\n")
        self.textbox.see("end")
        self.textbox.configure(state="disabled")

    def load_excel(self):
        filepath = filedialog.askopenfilename(title="Seleccione el archivo Excel (.xlsx)", filetypes=[("Archivos Excel", "*.xlsx")])
        if not filepath:
            return
            
        self.excel_file = filepath
        clean_name = filepath.split("/")[-1]
        self.lbl_file.configure(text=clean_name, text_color=("black", "white"))
        
        self.log(f"Cargando archivo: {clean_name}...")
        self.update_idletasks() # Forzar actualización visual al usuario
        
        try:
            self.workbook = openpyxl.load_workbook(filepath, data_only=True)
            self.log("✔ Archivo cargado correctamente.")
            self.build_sheet_ui()
        except Exception as e:
            self.log(f"✖ Error al cargar el archivo: {e}")
            messagebox.showerror("Error", f"No se pudo cargar el archivo Excel:\n{e}")

    def build_sheet_ui(self):
        # Limpiar frame anterior
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
            
        self.sheet_widgets.clear()
        
        # Headers del Grid
        header_frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 5))
        ctk.CTkLabel(header_frame, text="Sede", width=200, anchor="w", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)
        ctk.CTkLabel(header_frame, text="Alumnos", width=80, font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)
        ctk.CTkLabel(header_frame, text="Desde #", width=80, font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)

        for sheet_name in self.workbook.sheetnames:
            row_frame = ctk.CTkFrame(self.scroll_frame)
            row_frame.pack(fill="x", pady=3)
            
            lbl = ctk.CTkLabel(row_frame, text=sheet_name, width=200, anchor="w")
            lbl.pack(side="left", padx=5, pady=5)
            
            # Caja para "Cantidad"
            entry_cant = ctk.CTkEntry(row_frame, width=80)
            entry_cant.insert(0, "0")
            entry_cant.pack(side="left", padx=5)
            
            # Caja para "Iniciar desde #"
            entry_inicio = ctk.CTkEntry(row_frame, width=80)
            entry_inicio.insert(0, "1")
            entry_inicio.pack(side="left", padx=5)
            
            self.sheet_widgets.append({
                "sheet": sheet_name,
                "cant": entry_cant,
                "inicio": entry_inicio
            })

    def start_capture(self):
        if not self.workbook:
            messagebox.showwarning("Atención", "Debes seleccionar un archivo Excel primero.")
            return
            
        if self.is_running:
            return
            
        # Validar y recopilar inputs antes de bloquear interfaz
        config = []
        for w in self.sheet_widgets:
            try:
                cant = int(w["cant"].get())
                inicio = int(w["inicio"].get())
                if cant > 0:
                    config.append({"sheet": w["sheet"], "cant": cant, "inicio": inicio})
            except ValueError:
                self.log(f"✖ Aviso: Ingreso inválido en '{w['sheet']}'. Será omitida.")
                
        if not config:
            messagebox.showinfo("Aviso", "No hay sedes configuradas con alumnos a procesar (> 0).")
            return

        self.is_running = True
        self.btn_start.configure(state="disabled", fg_color="gray", text="PROCESANDO...")
        self.btn_load.configure(state="disabled")
        
        modo = self.modo_var.get()
        self.log(f"\n======================================")
        self.log(f"🚀 INICIANDO AUTOMATIZACIÓN (Modo: {modo.upper()})")
        self.log(f"======================================")
        
        # Ejecutar PyAutoGUI en un Hilo para no congelar la ventana gráfica
        threading.Thread(target=self.run_automation_thread, args=(config, modo), daemon=True).start()

    def run_automation_thread(self, config, modo):
        # Función callback cuando el automator termina
        def on_finish(success=True, error_msg="", extra_info=None):
            if success and not error_msg:
                msg = "El proceso ha terminado."
                if extra_info:
                    rechazados = extra_info.get("rechazados", [])
                    capturados = extra_info.get("capturados", 0)
                    
                    reporte = f"\nResumen final:\n- Capturados: {capturados}\n- No encontrados (API): {len(rechazados)}"
                    self.log(reporte)
                    
                    if rechazados:
                        detalle = "\n\nEstudiantes NO encontrados en API:\n"
                        for r in rechazados:
                            detalle += f"- {r['matricula']}: {r['nombre']} (Error: {r['error']})\n"
                        self.log(detalle)
                        
                        self.after(0, lambda: messagebox.showwarning("Proceso Terminado", 
                            f"Se capturaron {capturados} estudiantes.\n\n{len(rechazados)} alumnos no fueron encontrados en la API y se omitieron. Revisa la consola para el detalle."))
                    else:
                        self.after(0, lambda: messagebox.showinfo("Éxito", f"Se capturaron {capturados} estudiantes correctamente."))
                else:
                    self.after(0, lambda: messagebox.showinfo("Éxito", "El proceso de captura ha terminado."))
            
            elif error_msg:
                self.after(0, lambda e=error_msg: messagebox.showerror("Error", f"Fallo durante la ejecución:\n{e}"))
            
            self.finish_automation()

        # Llamamos al orquestador (Workflow)
        ejecutar_workflow_completo(
            workbook=self.workbook,
            config=config,
            modo_ejecucion=modo,
            logger=self.log,
            confirmador_manual=self.ask_gui_confirmation,
            on_finish=on_finish
        )

    def ask_gui_confirmation(self, title, message):
        """Muestra un messagebox en el Main Thread y pausa el Background Thread hasta tener respuesta."""
        event = threading.Event()
        result = [False]
        
        def show():
            # askokcancel devuelve True (OK) o False (Cancel)
            result[0] = messagebox.askokcancel(title, message, icon="info")
            event.set()
            
        self.after(0, show)
        event.wait()
        return result[0]

    def finish_automation(self):
        self.is_running = False
        def enable_ui():
            self.btn_start.configure(state="normal", fg_color="#10B981", text="3. INICIAR CAPTURA")
            self.btn_load.configure(state="normal")
        self.after(0, enable_ui)

if __name__ == "__main__":
    app = App()
    app.mainloop()
