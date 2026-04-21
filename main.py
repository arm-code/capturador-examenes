import os
import threading
import queue
from datetime import datetime
import openpyxl
from nicegui import ui, app
from core.workflow import ejecutar_workflow_completo

# Configuración de Tailwind v4
# (NiceGUI incluye Tailwind CSS de fábrica, por lo que no es necesario el CDN externo)

class CapturadorApp:
    def __init__(self):
        self.workbook = None
        self.excel_path = None
        self.sheets = []
        self.selected_sheet = None
        
        # Queues for thread-safe communication
        self.log_queue = queue.Queue()
        self.confirmation_queue = queue.Queue()
        self.results_queue = queue.Queue()
        self.notification_queue = queue.Queue()
        
        self.is_running = False
        self.modo_captura = 'manual'
        
        # UI References
        self.sidebar = None
        self.logs_area = None
        self.sheet_config_card = None
        self.validation_table = None
        
    def log(self, message):
        """Enviado desde cualquier hilo."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_queue.put(f"[{timestamp}] {message}")

    def poll_queues(self):
        """Timer que corre en el hilo principal de NiceGUI."""
        # 1. Logs
        while not self.log_queue.empty():
            msg = self.log_queue.get()
            with self.logs_area:
                ui.label(msg).classes('text-xs font-mono text-gray-400 py-0.5 border-b border-gray-800/50')
            # Auto-scroll hack for NiceGUI scroll_area
            ui.run_javascript(f'const el = document.querySelector(".q-scrollarea__container"); if(el) el.scrollTop = el.scrollHeight;')

        # 2. Confirmaciones (Sync workflow thread -> Blocking UI dialog)
        if not self.confirmation_queue.empty():
            req = self.confirmation_queue.get()
            self.show_confirmation_dialog(req['titulo'], req['mensaje'], req['event'], req['result'])

        # 3. Resultados de validación (para la tabla)
        while not self.results_queue.empty():
            res = self.results_queue.get()
            if self.validation_table:
                if isinstance(res, list) and not res:
                    self.validation_table.rows = []
                else:
                    self.validation_table.add_rows([res])

        # 4. Notificaciones
        while not self.notification_queue.empty():
            note = self.notification_queue.get()
            ui.notify(note['msg'], type=note['type'])

    def show_confirmation_dialog(self, titulo, mensaje, event, result_wrapper):
        """Muestra el diálogo y desbloquea el hilo de background al terminar."""
        with ui.dialog() as dialog, ui.card().classes('p-6 bg-slate-900 border border-slate-700 shadow-2xl'):
            ui.label(titulo).classes('text-xl font-bold text-white mb-2')
            ui.label(mensaje).classes('text-gray-400 mb-6')
            with ui.row().classes('w-full justify-end gap-3'):
                ui.button('CANCELAR', on_click=lambda: self.handle_dialog_res(dialog, False, event, result_wrapper), color='red').props('flat')
                ui.button('ACEPTAR', on_click=lambda: self.handle_dialog_res(dialog, True, event, result_wrapper), color='blue')
        dialog.open()

    def handle_dialog_res(self, dialog, val, event, result_wrapper):
        result_wrapper[0] = val
        dialog.close()
        event.set() # Desbloquea el hilo de background

    def confirmador_bridge(self, titulo, mensaje):
        """Este método es llamado por el hilo del Workflow."""
        event = threading.Event()
        result_wrapper = [None]
        self.confirmation_queue.put({
            'titulo': titulo,
            'mensaje': mensaje,
            'event': event,
            'result': result_wrapper
        })
        event.wait() # Bloqueo seguro del hilo de worker
        return result_wrapper[0]

    async def handle_upload(self, e):
        if not e.file: return
        file_path = os.path.join(os.getcwd(), e.file.name)
        
        try:
            # En NiceGUI moderno usamos e.file.save() que es asíncrono
            await e.file.save(file_path)
            
            self.workbook = openpyxl.load_workbook(file_path)
            self.excel_path = file_path
            self.sheets = self.workbook.sheetnames
            self.log(f"📦 Excel cargado: {e.file.name}")
            self.render_sidebar()
            ui.notify('Archivo cargado correctamente', type='positive')
        except Exception as err:
            ui.notify(f'Error: {err}', type='negative')
            self.log(f"❌ Error al cargar Excel: {err}")

    def start_workflow(self, cant, inicio, modo):
        if self.is_running: return
        
        self.is_running = True
        self.results_queue.put([]) # Clear table? No, add_rows appends. 
        # For simplicity, we'll just log. 
        
        config = [{"sheet": self.selected_sheet, "cant": int(cant), "inicio": int(inicio)}]
        
        def on_finish(success, extra_info=None, error_msg=""):
            self.is_running = False
            if success:
                self.log("🏁 Proceso completado exitosamente.")
                self.notification_queue.put({'msg': 'Terminado', 'type': 'positive'})
            else:
                self.log(f"⚠ Proceso detenido/error: {error_msg}")
                self.notification_queue.put({'msg': f'Error: {error_msg}', 'type': 'negative'})
        
        thread = threading.Thread(
            target=ejecutar_workflow_completo,
            args=(self.workbook, config, modo, self.log, self.confirmador_bridge, on_finish, self.results_queue.put),
            daemon=True
        )
        thread.start()

    def render_sidebar(self):
        self.sidebar.clear()
        with self.sidebar:
            ui.label('SEDES').classes('text-blue-500 text-[10px] font-black tracking-[0.3em] mb-6 ml-2')
            for name in self.sheets:
                is_sel = self.selected_sheet == name
                with ui.row().classes(f'p-3 mb-1 rounded-xl cursor-pointer transition-all items-center \
                                        {"bg-blue-600/20 border border-blue-500/50 shadow-lg shadow-blue-900/20" if is_sel else "hover:bg-slate-800/50 text-slate-400"}') \
                        .on('click', lambda n=name: [setattr(self, "selected_sheet", n), self.render_sidebar(), self.render_content()]):
                    ui.icon('business' if is_sel else 'apartment', size='sm').classes('mr-3')
                    ui.label(name).classes('font-bold text-sm')

    def render_content(self):
        self.sheet_config_card.clear()
        if not self.selected_sheet:
            with self.sheet_config_card:
                with ui.column().classes('w-full items-center py-20 opacity-20'):
                    ui.icon('description', size='120px')
                    ui.label('Selecciona una sede para configurar la captura').classes('text-xl font-medium')
            return

        with self.sheet_config_card:
            with ui.row().classes('w-full items-start justify-between p-8'):
                with ui.column().classes('gap-1'):
                    ui.label('Configuración de Sede').classes('text-blue-400 text-xs font-bold tracking-tighter')
                    ui.label(self.selected_sheet).classes('text-4xl font-black text-white italic')
                
                with ui.card().classes('bg-slate-800/50 border border-slate-700 p-6 rounded-2xl'):
                    with ui.row().classes('gap-6 items-end'):
                        n = ui.number('Cantidad', value=5, min=1).classes('w-24').props('dark filled')
                        i = ui.number('Inicio', value=1, min=1).classes('w-24').props('dark filled')
                        m = ui.select({'manual': 'Manual', 'auto': 'Automático'}, value='manual', label='Modo') \
                              .classes('w-32').props('dark filled')
                        
                        ui.button('INICIAR PROCESO', on_click=lambda: self.start_workflow(n.value, i.value, m.value), color='blue-6') \
                          .classes('h-14 px-8 font-black rounded-xl shadow-xl shadow-blue-900/40 text-white')

    def build(self):
        ui.query('body').classes('bg-slate-950 font-sans selection:bg-blue-500/30')
        
        with ui.header().classes('bg-slate-900/80 backdrop-blur-md border-b border-slate-800 p-4 shrink-0'):
            with ui.row().classes('w-full items-center justify-between'):
                with ui.row().classes('items-center'):
                    with ui.element('div').classes('p-2 bg-blue-600 rounded-lg mr-3 shadow-lg shadow-blue-900/50'):
                        ui.icon('bolt', color='white', size='sm')
                    ui.label('SIOSAD').classes('text-xl font-black tracking-tighter text-white')
                    ui.label('CAPTURA PRO').classes('text-xl font-light tracking-widest text-slate-400 ml-1')
                
                # Rediseño del área de carga
                with ui.row().classes('items-center bg-slate-800/50 rounded-xl px-4 py-1 border border-slate-700 hover:border-blue-500 transition-all'):
                    ui.icon('upload_file', color='blue-400')
                    ui.upload(on_upload=self.handle_upload, label='Subir Excel', auto_upload=True) \
                      .props('flat color=black hide-upload-btn').classes('w-48')

        with ui.row().classes('w-full h-[calc(100vh-80px)] no-wrap gap-0'):
            # Sidebar
            self.sidebar = ui.scroll_area().classes('w-72 bg-slate-900/50 p-6 shrink-0 border-r border-slate-800')
            
            # Content
            with ui.column().classes('flex-1 h-full overflow-hidden'):
                self.sheet_config_card = ui.column().classes('w-full shrink-0')
                self.render_content()
                
                # Console / Logs
                with ui.column().classes('flex-1 w-full p-6 pt-0 overflow-hidden'):
                    with ui.column().classes('w-full h-full bg-slate-900/80 rounded-3xl border border-slate-800 overflow-hidden shadow-inner'):
                        with ui.row().classes('w-full bg-slate-800/80 px-6 py-3 items-center justify-between'):
                            with ui.row().classes('items-center gap-2'):
                                ui.element('div').classes('w-2 h-2 rounded-full bg-green-500 animate-pulse')
                                ui.label('CONSOLA DE EVENTOS').classes('text-[10px] font-black text-slate-500 tracking-[0.2em]')
                            ui.button(icon='delete_sweep', on_click=lambda: self.logs_area.clear()).props('flat round dense')
                        
                        self.logs_area = ui.scroll_area().classes('w-full h-full p-6')

        ui.timer(0.1, self.poll_queues)

app_logic = CapturadorApp()
app_logic.build()

ui.run(
    title="SIOSAD Capturador Pro",
    native=True,
    window_size=(1200, 900),
    dark=True,
    reload=False
)
