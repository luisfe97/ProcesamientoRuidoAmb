import sys
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, 
    QHBoxLayout, QPushButton, QLabel, QFileDialog, QComboBox, 
    QLineEdit, QFormLayout, QGroupBox, QTextEdit, QMessageBox,
    QProgressBar, QTableWidget, QTableWidgetItem, QHeaderView,
    QListWidget, QAbstractItemView, QCheckBox, QSplitter, QDialog,
    QDialogButtonBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt5.QtGui import QFont, QIcon

# Importando módulos del proyecto actual
# Estas importaciones hay que ajustarlas según la estructura real
# y considerar añadir la carpeta raíz al sys.path si es necesario
try:
    from data.constants import SHEETS_TO_PROCESS, ARCHIVO_EXCEL, OUTPUT_FOLDER
    from utils.file_utils import combine_excel_files
    from processing.acoustic import aplicar_Correccion
    from processing.meteorology import process_and_export_weather_data
    from processing.data_handler import (
        cargar_datos, procesar_tercios_octava, crear_tabla_procesada, 
        filtrar_por_periodos, procesar_diario
    )
    from processing.statistics import generar_resumenes, calcular_estadisticos, actualizar_resumen, calcular_L_Raseq_dn
    from processing.uncertainty_handler import calcular_incertidumbres
    from processing.compliance import (
        asignar_limites, asignar_limites_diarios, procesar_compliance_diurno, 
        procesar_compliance_nocturno, finalizar_agrupados
    )
    from export.excel import export_to_template
    from export.ruido_total import (procesar_excel_simple, combinar_excels)
    
    PROJECT_MODULES_IMPORTED = True
except ImportError as e:
    # Si falla la importación, lo capturamos para seguir con la interfaz
    # y mostrar una advertencia al usuario
    PROJECT_MODULES_IMPORTED = False
    print(f"Error al importar módulos del proyecto: {e}")


class MatplotlibCanvas(FigureCanvas):
    """Lienzo de Matplotlib para incluir gráficas en la interfaz"""
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = self.fig.add_subplot(111)
        super(MatplotlibCanvas, self).__init__(self.fig)


class ProcessingWorker(QThread):
    """Worker para ejecutar el procesamiento en segundo plano"""
    update_progress = pyqtSignal(int, str)
    finished_signal = pyqtSignal(dict)
    error_signal = pyqtSignal(str)
    
    def __init__(self, parameters):
        super().__init__()
        self.parameters = parameters
        self.running = True
        
    def run(self):
        try:
            if not PROJECT_MODULES_IMPORTED:
                raise ImportError("No se pudieron importar los módulos del proyecto.")
            
            # Extraer parámetros
            archivo_excel = self.parameters.get('input_file', ARCHIVO_EXCEL)
            sheets_to_process = self.parameters.get('sheets', SHEETS_TO_PROCESS)
            output_folder = self.parameters.get('output_folder', OUTPUT_FOLDER)
            template_path = self.parameters.get('template_file', "Plantilla/Plantilla_Macro.xlsx")
            
            # Crear carpeta de salida si no existe
            os.makedirs(output_folder, exist_ok=True)
            
            # Procesamiento por hojas
            pto = 1
            total_sheets = len(sheets_to_process)
            
            for idx, sheet in enumerate(sheets_to_process):
                if not self.running:
                    break
                
                # Actualizar progreso
                progress = int((idx / total_sheets) * 90)  # Reservamos 10% para el procesamiento final
                self.update_progress.emit(progress, f"Procesando hoja: {sheet}")
                
                # Procesar hoja
                try:
                    # Si integramos directamente con el código existente:
                    if PROJECT_MODULES_IMPORTED:
                        from main import procesar_hoja
                        pto = procesar_hoja(sheet, pto, archivo_excel, archivo_excel)
                    else:
                        # Simulamos el procesamiento para pruebas
                        import time
                        time.sleep(1)
                except Exception as e:
                    self.update_progress.emit(progress, f"Error en hoja {sheet}: {str(e)}")
            
            # Procesamiento final
            if self.running:
                self.update_progress.emit(90, "Combinando archivos Excel...")
                
                if PROJECT_MODULES_IMPORTED:
                    # Combinar archivos Excel
                    combine_excel_files(output_folder)
                    
                    # Procesar ruido total
                    ruta_excel = f"{output_folder}/Excel_Intercalado.xlsx"
                    dataframes = procesar_excel_simple(ruta_excel, output_folder)
                    
                    # Combinar resultados finales
                    archivo1 = os.path.join(output_folder, "RUIDO TOTAL.xlsx")
                    archivo2 = ruta_excel
                    archivo_salida = os.path.join(output_folder, "Plantilla Ruido Total (Ambiental).xlsx")
                    combinar_excels(archivo1, archivo2, archivo_salida)
                
                self.update_progress.emit(100, "Proceso completado con éxito.")
                
                # Recopilar resultados para mostrar en la interfaz
                results = {
                    "status": "success",
                    "output_folder": output_folder,
                    "processed_sheets": sheets_to_process,
                    "points": pto - 1
                }
                
                self.finished_signal.emit(results)
            
        except Exception as e:
            self.error_signal.emit(str(e))
    
    def stop(self):
        """Detener el procesamiento"""
        self.running = False


class SheetSelectorDialog(QDialog):
    """Diálogo para seleccionar hojas a procesar"""
    def __init__(self, excel_file, parent=None):
        super().__init__(parent)
        self.excel_file = excel_file
        self.selected_sheets = []
        self.initUI()
        
    def initUI(self):
        self.setWindowTitle("Seleccionar Hojas a Procesar")
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout()
        
        # Instrucciones
        info_label = QLabel("Selecciona las hojas del archivo Excel que deseas procesar:")
        layout.addWidget(info_label)
        
        # Lista de hojas
        self.sheet_list = QListWidget()
        self.sheet_list.setSelectionMode(QAbstractItemView.MultiSelection)
        
        # Cargar hojas del Excel
        try:
            if os.path.exists(self.excel_file):
                xls = pd.ExcelFile(self.excel_file)
                for sheet in xls.sheet_names:
                    self.sheet_list.addItem(sheet)
                
                # Preseleccionar las hojas definidas en SHEETS_TO_PROCESS
                for i in range(self.sheet_list.count()):
                    item = self.sheet_list.item(i)
                    if PROJECT_MODULES_IMPORTED and item.text() in SHEETS_TO_PROCESS:
                        item.setSelected(True)
            else:
                self.sheet_list.addItem("Error: No se pudo encontrar el archivo Excel")
        except Exception as e:
            self.sheet_list.addItem(f"Error al cargar hojas: {str(e)}")
        
        layout.addWidget(self.sheet_list)
        
        # Botones de acción
        self.select_all_btn = QPushButton("Seleccionar Todas")
        self.select_all_btn.clicked.connect(self.select_all_sheets)
        self.deselect_all_btn = QPushButton("Deseleccionar Todas")
        self.deselect_all_btn.clicked.connect(self.deselect_all_sheets)
        
        btn_layout = QHBoxLayout()
        btn_layout.addWidget(self.select_all_btn)
        btn_layout.addWidget(self.deselect_all_btn)
        layout.addLayout(btn_layout)
        
        # Botones estándar de diálogo
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
    
    def select_all_sheets(self):
        """Seleccionar todas las hojas"""
        for i in range(self.sheet_list.count()):
            self.sheet_list.item(i).setSelected(True)
    
    def deselect_all_sheets(self):
        """Deseleccionar todas las hojas"""
        for i in range(self.sheet_list.count()):
            self.sheet_list.item(i).setSelected(False)
    
    def get_selected_sheets(self):
        """Obtener las hojas seleccionadas"""
        selected_sheets = []
        for item in self.sheet_list.selectedItems():
            selected_sheets.append(item.text())
        return selected_sheets


class AcousticProcessingApp(QMainWindow):
    """Aplicación principal para procesamiento acústico"""
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Procesamiento Acústico y Meteorológico")
        self.setGeometry(100, 100, 1200, 800)
        
        # Variables para datos
        self.input_file = ""
        self.template_file = ""
        self.output_folder = ""
        self.selected_sheets = []
        self.results = None
        
        # Configurar la interfaz
        self.setup_ui()
        
        # Inicializar con valores por defecto si están disponibles
        if PROJECT_MODULES_IMPORTED:
            self.input_file = ARCHIVO_EXCEL
            self.input_file_edit.setText(ARCHIVO_EXCEL)
            self.output_folder = OUTPUT_FOLDER
            self.output_folder_edit.setText(OUTPUT_FOLDER)
            self.template_file = "Plantilla/Plantilla_Macro.xlsx"
            self.template_file_edit.setText(self.template_file)
        
    def setup_ui(self):
        """Configuración principal de la interfaz"""
        # Widget central y layout principal
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        main_layout = QVBoxLayout(self.central_widget)
        
        # Crear pestañas
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)
        
        # Pestaña de Configuración
        self.config_tab = QWidget()
        self.tab_widget.addTab(self.config_tab, "Configuración")
        self.setup_config_tab()
        
        # Pestaña de Procesamiento
        self.process_tab = QWidget()
        self.tab_widget.addTab(self.process_tab, "Procesamiento")
        self.setup_process_tab()
        
        # Pestaña de Resultados
        self.results_tab = QWidget()
        self.tab_widget.addTab(self.results_tab, "Resultados")
        self.setup_results_tab()
        
        # Pestaña de Visualización
        self.visualization_tab = QWidget()
        self.tab_widget.addTab(self.visualization_tab, "Visualización")
        self.setup_visualization_tab()
        
        # Barra de estado
        self.statusBar().showMessage("Listo")
    
    def setup_config_tab(self):
        """Configuración de la pestaña de configuración"""
        layout = QVBoxLayout(self.config_tab)
        
        # Grupo de archivos
        file_group = QGroupBox("Archivos y Carpetas")
        file_layout = QFormLayout(file_group)
        
        # Selección de archivo de entrada
        input_layout = QHBoxLayout()
        self.input_file_edit = QLineEdit()
        self.input_file_edit.setReadOnly(True)
        input_btn = QPushButton("Explorar...")
        input_btn.clicked.connect(self.select_input_file)
        input_layout.addWidget(self.input_file_edit)
        input_layout.addWidget(input_btn)
        file_layout.addRow("Archivo de datos:", input_layout)
        
        # Selección de plantilla
        template_layout = QHBoxLayout()
        self.template_file_edit = QLineEdit()
        self.template_file_edit.setReadOnly(True)
        template_btn = QPushButton("Explorar...")
        template_btn.clicked.connect(self.select_template_file)
        template_layout.addWidget(self.template_file_edit)
        template_layout.addWidget(template_btn)
        file_layout.addRow("Plantilla Excel:", template_layout)
        
        # Selección de carpeta de salida
        output_layout = QHBoxLayout()
        self.output_folder_edit = QLineEdit()
        self.output_folder_edit.setReadOnly(True)
        output_btn = QPushButton("Explorar...")
        output_btn.clicked.connect(self.select_output_folder)
        output_layout.addWidget(self.output_folder_edit)
        output_layout.addWidget(output_btn)
        file_layout.addRow("Carpeta de salida:", output_layout)
        
        layout.addWidget(file_group)
        
        # Grupo de hojas
        sheets_group = QGroupBox("Hojas a Procesar")
        sheets_layout = QVBoxLayout(sheets_group)
        
        sheets_info = QLabel("Selecciona las hojas del archivo Excel que se procesarán:")
        sheets_layout.addWidget(sheets_info)
        
        self.sheets_list = QListWidget()
        self.sheets_list.setSelectionMode(QAbstractItemView.MultiSelection)
        sheets_layout.addWidget(self.sheets_list)
        
        # Botones para selección de hojas
        sheets_buttons_layout = QHBoxLayout()
        select_sheets_btn = QPushButton("Seleccionar Hojas...")
        select_sheets_btn.clicked.connect(self.select_sheets)
        sheets_buttons_layout.addWidget(select_sheets_btn)
        
        sheets_layout.addLayout(sheets_buttons_layout)
        
        layout.addWidget(sheets_group)
        
        # Grupo de opciones avanzadas
        advanced_group = QGroupBox("Opciones Avanzadas")
        advanced_layout = QFormLayout(advanced_group)
        
        # Casillas de verificación para opciones
        self.combine_option = QCheckBox("Combinar archivos Excel al finalizar")
        self.combine_option.setChecked(True)
        advanced_layout.addRow(self.combine_option)
        
        self.process_total_option = QCheckBox("Procesar ruido total")
        self.process_total_option.setChecked(True)
        advanced_layout.addRow(self.process_total_option)
        
        layout.addWidget(advanced_group)
        
        # Botones de acción
        actions_layout = QHBoxLayout()
        save_config_btn = QPushButton("Guardar Configuración")
        save_config_btn.clicked.connect(self.save_configuration)
        load_config_btn = QPushButton("Cargar Configuración")
        load_config_btn.clicked.connect(self.load_configuration)
        
        actions_layout.addWidget(save_config_btn)
        actions_layout.addWidget(load_config_btn)
        layout.addLayout(actions_layout)
        
        # Espacio flexible al final
        layout.addStretch()
    
    def setup_process_tab(self):
        """Configuración de la pestaña de procesamiento"""
        layout = QVBoxLayout(self.process_tab)
        
        # Información del procesamiento
        info_group = QGroupBox("Información del procesamiento")
        info_layout = QFormLayout(info_group)
        
        self.input_file_label = QLabel("No seleccionado")
        info_layout.addRow("Archivo de entrada:", self.input_file_label)
        
        self.sheets_label = QLabel("Ninguna seleccionada")
        info_layout.addRow("Hojas a procesar:", self.sheets_label)
        
        self.output_folder_label = QLabel("No seleccionada")
        info_layout.addRow("Carpeta de salida:", self.output_folder_label)
        
        layout.addWidget(info_group)
        
        # Log de procesamiento
        log_group = QGroupBox("Log de procesamiento")
        log_layout = QVBoxLayout(log_group)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        layout.addWidget(log_group)
        
        # Barra de progreso
        progress_group = QGroupBox("Progreso")
        progress_layout = QVBoxLayout(progress_group)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)
        
        self.progress_label = QLabel("Esperando inicio...")
        progress_layout.addWidget(self.progress_label)
        
        layout.addWidget(progress_group)
        
        # Botones de acción
        actions_layout = QHBoxLayout()
        self.start_btn = QPushButton("Iniciar Procesamiento")
        self.start_btn.clicked.connect(self.start_processing)
        self.stop_btn = QPushButton("Detener")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_processing)
        
        actions_layout.addWidget(self.start_btn)
        actions_layout.addWidget(self.stop_btn)
        layout.addLayout(actions_layout)
    
    def setup_results_tab(self):
        """Configuración de la pestaña de resultados"""
        layout = QVBoxLayout(self.results_tab)
        
        # Resumen de resultados
        summary_group = QGroupBox("Resumen de resultados")
        summary_layout = QFormLayout(summary_group)
        
        self.status_label = QLabel("Sin procesar")
        summary_layout.addRow("Estado:", self.status_label)
        
        self.points_label = QLabel("0")
        summary_layout.addRow("Puntos procesados:", self.points_label)
        
        self.files_label = QLabel("Ninguno")
        summary_layout.addRow("Archivos generados:", self.files_label)
        
        layout.addWidget(summary_group)
        
        # Tabla de resultados
        table_group = QGroupBox("Archivos generados")
        table_layout = QVBoxLayout(table_group)
        
        self.files_table = QTableWidget()
        self.files_table.setColumnCount(3)
        self.files_table.setHorizontalHeaderLabels(["Archivo", "Tamaño", "Fecha"])
        self.files_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        table_layout.addWidget(self.files_table)
        
        # Botones para acciones con archivos
        files_actions_layout = QHBoxLayout()
        
        open_folder_btn = QPushButton("Abrir Carpeta de Resultados")
        open_folder_btn.clicked.connect(self.open_results_folder)
        
        open_file_btn = QPushButton("Abrir Archivo Seleccionado")
        open_file_btn.clicked.connect(self.open_selected_file)
        
        files_actions_layout.addWidget(open_folder_btn)
        files_actions_layout.addWidget(open_file_btn)
        
        table_layout.addLayout(files_actions_layout)
        
        layout.addWidget(table_group)
    
    def setup_visualization_tab(self):
        """Configuración de la pestaña de visualización"""
        layout = QVBoxLayout(self.visualization_tab)
        
        # Selector de archivo
        file_selector_group = QGroupBox("Selección de archivo para visualización")
        file_selector_layout = QFormLayout(file_selector_group)
        
        self.viz_file_combo = QComboBox()
        self.viz_file_combo.currentIndexChanged.connect(self.update_visualization)
        file_selector_layout.addRow("Archivo:", self.viz_file_combo)
        
        self.viz_sheet_combo = QComboBox()
        self.viz_sheet_combo.currentIndexChanged.connect(self.update_visualization)
        file_selector_layout.addRow("Hoja:", self.viz_sheet_combo)
        
        layout.addWidget(file_selector_group)
        
        # Visualización
        splitter = QSplitter(Qt.Horizontal)
        
        # Panel izquierdo: Gráfico
        chart_group = QGroupBox("Gráfico")
        chart_layout = QVBoxLayout(chart_group)
        
        self.chart_canvas = MatplotlibCanvas(width=6, height=5)
        chart_layout.addWidget(self.chart_canvas)
        
        # Selector de gráfico
        self.viz_type_combo = QComboBox()
        # Actualizar las opciones del combo box de visualización
        self.viz_type_combo.addItems(["Niveles acústicos diurnos", "Niveles acústicos nocturnos"])
        self.viz_type_combo.currentIndexChanged.connect(self.update_visualization)
        chart_layout.addWidget(QLabel("Tipo de visualización:"))
        chart_layout.addWidget(self.viz_type_combo)
        
        splitter.addWidget(chart_group)
        
        # Panel derecho: Tabla de datos
        data_group = QGroupBox("Datos")
        data_layout = QVBoxLayout(data_group)
        
        self.viz_table = QTableWidget()
        self.viz_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        data_layout.addWidget(self.viz_table)
        
        splitter.addWidget(data_group)
        
        layout.addWidget(splitter)
        
        # Botones de exportación
        export_layout = QHBoxLayout()
        
        export_chart_btn = QPushButton("Exportar Gráfico")
        export_chart_btn.clicked.connect(self.export_chart)
        
        export_data_btn = QPushButton("Exportar Datos")
        export_data_btn.clicked.connect(self.export_data)
        
        export_layout.addWidget(export_chart_btn)
        export_layout.addWidget(export_data_btn)
        
        layout.addLayout(export_layout)
    
    # Métodos para manejar eventos
    
    def select_input_file(self):
        """Seleccionar archivo de entrada"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar archivo de datos", "", 
            "Archivos Excel (*.xlsx *.xls);;Todos los archivos (*)"
        )
        if file_path:
            self.input_file = file_path
            self.input_file_edit.setText(file_path)
            self.input_file_label.setText(os.path.basename(file_path))
            
            # Intentar cargar las hojas
            try:
                xls = pd.ExcelFile(file_path)
                self.sheets_list.clear()
                
                for sheet in xls.sheet_names:
                    self.sheets_list.addItem(sheet)
                
                # Preseleccionar las hojas definidas en SHEETS_TO_PROCESS
                if PROJECT_MODULES_IMPORTED:
                    for i in range(self.sheets_list.count()):
                        item = self.sheets_list.item(i)
                        if item.text() in SHEETS_TO_PROCESS:
                            item.setSelected(True)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"No se pudieron cargar las hojas del Excel: {str(e)}")
    
    def select_template_file(self):
        """Seleccionar archivo de plantilla"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar plantilla", "", 
            "Archivos Excel (*.xlsx *.xls);;Todos los archivos (*)"
        )
        if file_path:
            self.template_file = file_path
            self.template_file_edit.setText(file_path)
    
    def select_output_folder(self):
        """Seleccionar directorio de salida"""
        dir_path = QFileDialog.getExistingDirectory(
            self, "Seleccionar carpeta de salida", ""
        )
        if dir_path:
            self.output_folder = dir_path
            self.output_folder_edit.setText(dir_path)
            self.output_folder_label.setText(dir_path)
    
    def select_sheets(self):
        """Seleccionar hojas a procesar mediante diálogo"""
        if not self.input_file:
            QMessageBox.warning(self, "Error", "Primero debe seleccionar un archivo Excel.")
            return
        
        dialog = SheetSelectorDialog(self.input_file, self)
        if dialog.exec_() == QDialog.Accepted:
            self.selected_sheets = dialog.get_selected_sheets()
            
            if self.selected_sheets:
                self.sheets_label.setText(", ".join(self.selected_sheets[:3]) + 
                                      (f" y {len(self.selected_sheets)-3} más" if len(self.selected_sheets) > 3 else ""))
            else:
                self.sheets_label.setText("Ninguna seleccionada")
    
    def save_configuration(self):
        """Guardar configuración actual a un archivo"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Guardar Configuración", "", 
            "Archivos de Configuración (*.json);;Todos los archivos (*)"
        )
        if file_path:
            try:
                import json
                
                # Recopilar configuración actual
                selected_sheets = [self.sheets_list.item(i).text() 
                                  for i in range(self.sheets_list.count()) 
                                  if self.sheets_list.item(i).isSelected()]
                
                config = {
                    "input_file": self.input_file,
                    "template_file": self.template_file,
                    "output_folder": self.output_folder,
                    "selected_sheets": selected_sheets,
                    "combine_files": self.combine_option.isChecked(),
                    "process_total": self.process_total_option.isChecked()
                }
                
                # Guardar a archivo
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(config, f, indent=4)
                
                QMessageBox.information(self, "Guardar Configuración", "Configuración guardada correctamente.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al guardar la configuración: {str(e)}")
    
    def load_configuration(self):
        """Cargar configuración desde un archivo"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Cargar Configuración", "", 
            "Archivos de Configuración (*.json);;Todos los archivos (*)"
        )
        if file_path:
            try:
                import json
                
                # Cargar desde archivo
                with open(file_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                # Aplicar configuración
                if "input_file" in config and os.path.exists(config["input_file"]):
                    self.input_file = config["input_file"]
                    self.input_file_edit.setText(config["input_file"])
                    self.input_file_label.setText(os.path.basename(config["input_file"]))
                    
                    # Cargar hojas
                    try:
                        xls = pd.ExcelFile(config["input_file"])
                        self.sheets_list.clear()
                        
                        for sheet in xls.sheet_names:
                            self.sheets_list.addItem(sheet)
                    except:
                        pass
                
                if "template_file" in config:
                    self.template_file = config["template_file"]
                    self.template_file_edit.setText(config["template_file"])
                
                if "output_folder" in config:
                    self.output_folder = config["output_folder"]
                    self.output_folder_edit.setText(config["output_folder"])
                    self.output_folder_label.setText(config["output_folder"])
                
                # Seleccionar hojas
                if "selected_sheets" in config:
                    self.selected_sheets = config["selected_sheets"]
                    
                    # Marcar hojas en la lista
                    for i in range(self.sheets_list.count()):
                        item = self.sheets_list.item(i)
                        if item.text() in self.selected_sheets:
                            item.setSelected(True)
                        else:
                            item.setSelected(False)
                            
                    if self.selected_sheets:
                        self.sheets_label.setText(", ".join(self.selected_sheets[:3]) + 
                                              (f" y {len(self.selected_sheets)-3} más" if len(self.selected_sheets) > 3 else ""))
                
                # Opciones adicionales
                if "combine_files" in config:
                    self.combine_option.setChecked(config["combine_files"])
                
                if "process_total" in config:
                    self.process_total_option.setChecked(config["process_total"])
                
                QMessageBox.information(self, "Cargar Configuración", "Configuración cargada correctamente.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al cargar la configuración: {str(e)}")
    
    def start_processing(self):
        """Iniciar el procesamiento de datos"""
        # Verificar que tenemos todos los datos necesarios
        if not self.input_file:
            QMessageBox.warning(self, "Error", "Debe seleccionar un archivo de entrada.")
            return
        
        if not self.template_file:
            QMessageBox.warning(self, "Error", "Debe seleccionar un archivo de plantilla.")
            return
        
        if not self.output_folder:
            QMessageBox.warning(self, "Error", "Debe seleccionar una carpeta de salida.")
            return
        
        # Obtener hojas seleccionadas
        if not self.selected_sheets:
            # Si no hay hojas explícitamente seleccionadas en el diálogo, usar las de la lista
            self.selected_sheets = [self.sheets_list.item(i).text() 
                                  for i in range(self.sheets_list.count()) 
                                  if self.sheets_list.item(i).isSelected()]
        
        if not self.selected_sheets:
            QMessageBox.warning(self, "Error", "Debe seleccionar al menos una hoja para procesar.")
            return
        
        # Actualizar la interfaz
        self.log_text.clear()
        self.progress_bar.setValue(0)
        self.progress_label.setText("Iniciando procesamiento...")
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        
        # Cambiar a la pestaña de procesamiento
        self.tab_widget.setCurrentWidget(self.process_tab)
        
        # Recopilar parámetros
        parameters = {
            'input_file': self.input_file,
            'template_file': self.template_file,
            'output_folder': self.output_folder,
            'sheets': self.selected_sheets,
            'combine_files': self.combine_option.isChecked(),
            'process_total': self.process_total_option.isChecked()
        }
        
        # Registrar el inicio en el log
        self.log_text.append(f"Iniciando procesamiento: {self.input_file}")
        self.log_text.append(f"Hojas a procesar: {', '.join(self.selected_sheets)}")
        self.log_text.append(f"Carpeta de salida: {self.output_folder}")
        self.log_text.append("-------------------------")
        
        # Crear y configurar el worker
        self.processing_worker = ProcessingWorker(parameters)
        self.processing_worker.update_progress.connect(self.update_progress)
        self.processing_worker.finished_signal.connect(self.processing_finished)
        self.processing_worker.error_signal.connect(self.processing_error)
        
        # Iniciar el procesamiento
        self.processing_worker.start()
    
    def stop_processing(self):
        """Detener el procesamiento"""
        if hasattr(self, 'processing_worker') and self.processing_worker.isRunning():
            self.processing_worker.stop()
            self.log_text.append("Solicitando detención del procesamiento...")
        
        # Restaurar la interfaz
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
    
    def update_progress(self, value, message):
        """Actualizar la barra de progreso y el mensaje"""
        self.progress_bar.setValue(value)
        self.progress_label.setText(message)
        self.log_text.append(message)
        
        # Mover el cursor al final del log
        self.log_text.moveCursor(self.log_text.textCursor().End)
    
    def processing_finished(self, results):
        """Manejar la finalización del procesamiento"""
        self.results = results
        self.log_text.append("-------------------------")
        self.log_text.append("Procesamiento completado con éxito.")
        
        # Actualizar la interfaz
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        
        # Actualizar pestaña de resultados
        self.status_label.setText("Procesado correctamente")
        self.points_label.setText(str(results.get("points", 0)))
        
        # Actualizar tabla de archivos
        self.update_files_table()
        
        # Actualizar selector de archivos para visualización
        self.update_viz_file_selector()
        
        # Cambiar a la pestaña de resultados
        self.tab_widget.setCurrentWidget(self.results_tab)
    
    def processing_error(self, error_msg):
        """Manejar errores de procesamiento"""
        self.log_text.append(f"ERROR: {error_msg}")
        QMessageBox.critical(self, "Error de procesamiento", f"Se produjo un error: {error_msg}")
        
        # Restaurar la interfaz
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
    
    def update_files_table(self):
        """Actualizar la tabla de archivos generados"""
        self.files_table.setRowCount(0)
        
        if not self.output_folder or not os.path.exists(self.output_folder):
            self.files_label.setText("Carpeta no encontrada")
            return
        
        try:
            files = [f for f in os.listdir(self.output_folder) if f.endswith(('.xlsx', '.xls'))]
            
            if files:
                self.files_label.setText(f"{len(files)} archivos")
                
                for idx, file in enumerate(files):
                    file_path = os.path.join(self.output_folder, file)
                    file_stat = os.stat(file_path)
                    
                    # Convertir tamaño a formato legible
                    size_bytes = file_stat.st_size
                    if size_bytes < 1024:
                        size_str = f"{size_bytes} B"
                    elif size_bytes < 1024 * 1024:
                        size_str = f"{size_bytes / 1024:.2f} KB"
                    else:
                        size_str = f"{size_bytes / (1024 * 1024):.2f} MB"
                    
                    # Fecha de modificación
                    import datetime
                    mod_time = datetime.datetime.fromtimestamp(file_stat.st_mtime)
                    date_str = mod_time.strftime("%d/%m/%Y %H:%M:%S")
                    
                    # Añadir a la tabla
                    self.files_table.insertRow(idx)
                    self.files_table.setItem(idx, 0, QTableWidgetItem(file))
                    self.files_table.setItem(idx, 1, QTableWidgetItem(size_str))
                    self.files_table.setItem(idx, 2, QTableWidgetItem(date_str))
            else:
                self.files_label.setText("Ningún archivo encontrado")
        except Exception as e:
            self.files_label.setText(f"Error: {str(e)}")
    
    def open_results_folder(self):
        """Abrir la carpeta de resultados en el explorador de archivos"""
        if not self.output_folder or not os.path.exists(self.output_folder):
            QMessageBox.warning(self, "Error", "La carpeta de resultados no existe.")
            return
        
        try:
            import subprocess
            import platform
            
            if platform.system() == "Windows":
                os.startfile(self.output_folder)
            elif platform.system() == "Darwin":  # macOS
                subprocess.Popen(["open", self.output_folder])
            else:  # Linux y otros
                subprocess.Popen(["xdg-open", self.output_folder])
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo abrir la carpeta: {str(e)}")
    
    def open_selected_file(self):
        """Abrir el archivo seleccionado en la tabla"""
        selected_items = self.files_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Error", "Debe seleccionar un archivo primero.")
            return
        
        # Obtener el nombre del archivo (primera columna)
        row = selected_items[0].row()
        file_name = self.files_table.item(row, 0).text()
        file_path = os.path.join(self.output_folder, file_name)
        
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "Error", f"El archivo {file_name} no existe.")
            return
        
        try:
            import subprocess
            import platform
            
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.Popen(["open", file_path])
            else:  # Linux y otros
                subprocess.Popen(["xdg-open", file_path])
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo abrir el archivo: {str(e)}")
    
    def update_viz_file_selector(self):
        """Actualizar el selector de archivos para visualización"""
        self.viz_file_combo.clear()
        
        if not self.output_folder or not os.path.exists(self.output_folder):
            return
        
        try:
            files = [f for f in os.listdir(self.output_folder) if f.endswith(('.xlsx', '.xls'))]
            
            if files:
                self.viz_file_combo.addItems(files)
                # Seleccionar el primer archivo y cargar sus hojas
                self.load_viz_file_sheets()
        except Exception as e:
            print(f"Error al actualizar selector de archivos: {str(e)}")
    
    def load_viz_file_sheets(self):
        """Cargar las hojas del archivo seleccionado para visualización"""
        self.viz_sheet_combo.clear()
        
        selected_file = self.viz_file_combo.currentText()
        if not selected_file:
            return
        
        file_path = os.path.join(self.output_folder, selected_file)
        
        try:
            if os.path.exists(file_path):
                xls = pd.ExcelFile(file_path)
                self.viz_sheet_combo.addItems(xls.sheet_names)
                
                # Seleccionar primera hoja y actualizar visualización
                self.update_visualization()
        except Exception as e:
            print(f"Error al cargar hojas: {str(e)}")
    
    def update_visualization(self):
        """Actualizar la visualización basada en el archivo y hoja seleccionados"""
        selected_file = self.viz_file_combo.currentText()
        selected_sheet = self.viz_sheet_combo.currentText()
        
        if not selected_file or not selected_sheet:
            return
        
        file_path = os.path.join(self.output_folder, selected_file)
        
        try:
            # Cargar datos
            df = pd.read_excel(file_path, sheet_name=selected_sheet)
            
            # Actualizar tabla
            self.update_viz_table(df)
            
            # Actualizar gráfico
            self.update_viz_chart(df)
        except Exception as e:
            self.chart_canvas.axes.clear()
            self.chart_canvas.axes.text(0.5, 0.5, f"Error al cargar datos: {str(e)}", 
                                         ha='center', va='center', fontsize=12)
            self.chart_canvas.draw()
            
            # Limpiar tabla
            self.viz_table.setRowCount(0)
            self.viz_table.setColumnCount(0)
    
    def update_viz_table(self, df):
        """Actualizar la tabla de visualización con los datos del DataFrame"""
        self.viz_table.setRowCount(0)
        
        # Determinar columnas y configurar la tabla
        if df.empty:
            self.viz_table.setColumnCount(0)
            return
        
        columns = df.columns.tolist()
        self.viz_table.setColumnCount(len(columns))
        self.viz_table.setHorizontalHeaderLabels(columns)
        
        # Llenar la tabla con datos
        for i, row in df.iterrows():
            row_position = self.viz_table.rowCount()
            self.viz_table.insertRow(row_position)
            
            for j, column in enumerate(columns):
                value = row[column]
                # Convertir a string para mostrar en la tabla
                if pd.isna(value):
                    value_str = ""
                elif isinstance(value, (int, float)):
                    value_str = f"{value:.4f}" if isinstance(value, float) else str(value)
                else:
                    value_str = str(value)
                
                self.viz_table.setItem(row_position, j, QTableWidgetItem(value_str))
    
    def update_viz_chart(self, df):
        """Actualizar el gráfico de visualización basado en los datos y tipo seleccionado"""
        # Limpiar gráfico actual
        self.chart_canvas.axes.clear()
        
        viz_type = self.viz_type_combo.currentText()
        
        # Verificar si tenemos datos suficientes
        if df.empty:
            self.chart_canvas.axes.text(0.5, 0.5, "No hay datos disponibles para visualizar", 
                                        ha='center', va='center', fontsize=12)
            self.chart_canvas.draw()
            return
        
        try:
            # Tomar la fila 8 (índice 7) como encabezados
            df_headers = df.iloc[7:8].reset_index(drop=True)
            # Tomar los datos a partir de la fila 9 (índice 8)
            df_clean = df.iloc[8:].reset_index(drop=True)
            
            # Definir manualmente los nombres de columnas
            column_names = ['Fecha corregida', 'LASeq,i', 'LAIeq,i', 'KI,i', 'KT,i', 'Bandas', 'LRASeq,i',
                        'col8', 'Fecha_Dia', 'TipoDia_Dia', 'Nm,1d_Dia', 'LASeq,1d_Dia', 'LAIeq,1d_Dia',
                        'KI,i_Dia', 'KT,i_Dia', 'Bandas_Dia', 'LRASeq,1d_Dia', 'TU_Dia', 'k_Dia', 'U_Dia',
                        'E_Dia', 'w_Dia', 'AU_Dia', 'z_Dia', 'Rp*=Pc_Dia', 'Rc_Dia', 'Declaración_Dia',
                        'col28', 'Fecha_Noche', 'TipoDia_Noche', 'Nm,1d_Noche', 'LASeq,1d_Noche',
                        'LAIeq,1d_Noche', 'KI,i_Noche', 'KT,i_Noche', 'Bandas_Noche', 'LRASeq,1d_Noche',
                        'TU_Noche', 'k_Noche', 'U_Noche', 'E_Noche', 'w_Noche', 'AU_Noche', 'z_Noche',
                        'Rp*=Pc_Noche', 'Rc_Noche', 'Declaración_Noche', 'col48', 'col49', 'Nm,k',
                        'LASeq,k', 'LAIeq,k', 'LRASeq,k', 'sk2_1', 'sk_1', 'TU_1', 'k_1', 'U_1', 'E_1',
                        'w_1', 'AU_1', 'z_1', 'Rp*=Pc_1', 'Rc_1', 'Declaración_1', 'col66', 'col67',
                        'Nm,k,n', 'LASeq,k,n', 'LAIeq,k,n', 'LRASeq,k,n', 'sk2_2', 'sk_2', 'TU_2', 'k_2',
                        'U_2', 'E_2', 'w_2', 'AU_2', 'z_2', 'Rp*=Pc_2', 'Rc_2', 'Declaración_2', 'col84',
                        'col85', 'Nm,dn', 'LASeq', 'LRASeq', 'LAIeq']
            
            # Asignar nombres de columnas solo si el número de columnas coincide
            if len(df_clean.columns) == len(column_names):
                df_clean.columns = column_names
            else:
                print(f"Advertencia: No coincide el número de columnas. DataFrame tiene {len(df_clean.columns)} columnas, pero se proporcionaron {len(column_names)} nombres.")
                # En caso de que no coincida, usar los encabezados originales
                df_clean.columns = df_headers.iloc[0].values
            
            print("Columnas disponibles:", df_clean.columns.tolist())
            
            # Seleccionar las columnas basadas en el tipo de visualización
            if viz_type == "Niveles acústicos diurnos":
                fecha_col = 'Fecha_Dia'
                lraseq_col = 'LRASeq,1d_Dia'
                tu_col = 'TU_Dia'
                periodo = "diurno"
            elif viz_type == "Niveles acústicos nocturnos":
                fecha_col = 'Fecha_Noche'
                lraseq_col = 'LRASeq,1d_Noche'
                tu_col = 'TU_Noche'
                periodo = "nocturno"
            else:
                # Por defecto, usar datos diurnos
                fecha_col = 'Fecha_Dia'
                lraseq_col = 'LRASeq,1d_Dia'
                tu_col = 'TU_Dia'
                periodo = "diurno"
            
            # Verificar que las columnas existen
            if fecha_col not in df_clean.columns or lraseq_col not in df_clean.columns or tu_col not in df_clean.columns:
                message = f"No se encontraron las columnas necesarias para período {periodo}.\nColumnas disponibles: {', '.join(map(str, df_clean.columns[:20]))}"
                self.chart_canvas.axes.text(0.5, 0.5, message, 
                                            ha='center', va='center', fontsize=10)
                self.chart_canvas.draw()
                return
            
            print(f"Utilizando columnas para período {periodo}: {fecha_col}, {lraseq_col} y {tu_col}")
            
            # Asegurarse que los datos sean del tipo correcto
            try:
                # Convertir la columna de fecha a datetime
                if not pd.api.types.is_datetime64_any_dtype(df_clean[fecha_col]):
                    df_clean[fecha_col] = pd.to_datetime(df_clean[fecha_col], errors='coerce')
                
                # Convertir las columnas numéricas
                if not pd.api.types.is_numeric_dtype(df_clean[lraseq_col]):
                    df_clean[lraseq_col] = pd.to_numeric(df_clean[lraseq_col], errors='coerce')
                    
                if not pd.api.types.is_numeric_dtype(df_clean[tu_col]):
                    df_clean[tu_col] = pd.to_numeric(df_clean[tu_col], errors='coerce')
                    
                # Eliminar filas con valores nulos después de la conversión
                df_clean = df_clean.dropna(subset=[fecha_col, lraseq_col, tu_col])
                
                if df_clean.empty:
                    self.chart_canvas.axes.text(0.5, 0.5, "No hay datos válidos después de procesar", 
                                                ha='center', va='center', fontsize=12)
                    self.chart_canvas.draw()
                    return
            except Exception as e:
                self.chart_canvas.axes.text(0.5, 0.5, f"Error al procesar datos: {str(e)}", 
                                            ha='center', va='center', fontsize=12)
                self.chart_canvas.draw()
                return
            
            # Crear gráfico
            self.chart_canvas.axes.plot(df_clean[fecha_col], df_clean[lraseq_col], 
                                        'o-', linewidth=2, color='blue', label=lraseq_col)
            
            self.chart_canvas.axes.plot(df_clean[fecha_col], df_clean[tu_col], 
                                        linewidth=2, color='red', label="Límite NCh 0627")
            
            self.chart_canvas.axes.set_title(f"Niveles Acústicos {periodo.capitalize()}: {lraseq_col}")
            self.chart_canvas.axes.set_xlabel("Fecha")
            self.chart_canvas.axes.set_ylabel("Nivel (dB)")
            self.chart_canvas.axes.legend()
            
            # Rotar las etiquetas de fecha para mejor legibilidad
            import matplotlib.pyplot as plt
            plt.setp(self.chart_canvas.axes.get_xticklabels(), rotation=45, ha='right')
            
            # Formatear el eje x para mostrar fechas de manera adecuada
            import matplotlib.dates as mdates
            import matplotlib.ticker as mticker
            self.chart_canvas.axes.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
            
            # Ajustar la cantidad de ticks en el eje x para evitar sobrecarga
            if len(df_clean) > 20:
                self.chart_canvas.axes.xaxis.set_major_locator(mticker.MaxNLocator(10))
            
            # Añadir cuadrícula para mejor lectura
            self.chart_canvas.axes.grid(True, linestyle='--', alpha=0.7)
            
            # Configuración final y mostrar
            self.chart_canvas.fig.tight_layout()
            self.chart_canvas.draw()
            
        except Exception as e:
            import traceback
            error_msg = str(e) + "\n" + traceback.format_exc()
            print(f"Error completo: {error_msg}")
            
            self.chart_canvas.axes.clear()
            self.chart_canvas.axes.text(0.5, 0.5, f"Error al crear gráfico: {str(e)}", 
                                        ha='center', va='center', fontsize=12)
            self.chart_canvas.draw()
                
    def export_chart(self):
        """Exportar el gráfico actual como imagen"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Guardar Gráfico", "", 
            "Imágenes PNG (*.png);;Imágenes JPG (*.jpg);;Todos los archivos (*)"
        )
        if file_path:
            try:
                self.chart_canvas.fig.savefig(file_path, dpi=300, bbox_inches='tight')
                QMessageBox.information(self, "Exportar Gráfico", "Gráfico exportado correctamente.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al exportar el gráfico: {str(e)}")
    
    def export_data(self):
        """Exportar los datos de la tabla de visualización"""
        selected_file = self.viz_file_combo.currentText()
        selected_sheet = self.viz_sheet_combo.currentText()
        
        if not selected_file or not selected_sheet:
            QMessageBox.warning(self, "Error", "No hay datos para exportar.")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Guardar Datos", "", 
            "Archivos CSV (*.csv);;Archivos Excel (*.xlsx);;Todos los archivos (*)"
        )
        if file_path:
            try:
                # Construir DataFrame desde la tabla
                rows = self.viz_table.rowCount()
                cols = self.viz_table.columnCount()
                headers = [self.viz_table.horizontalHeaderItem(i).text() for i in range(cols)]
                
                df = pd.DataFrame(columns=headers)
                
                for row in range(rows):
                    data = {}
                    for col in range(cols):
                        item = self.viz_table.item(row, col)
                        if item is not None:
                            data[headers[col]] = item.text()
                        else:
                            data[headers[col]] = ""
                    df = df.append(data, ignore_index=True)
                
                # Exportar según la extensión
                if file_path.endswith('.csv'):
                    df.to_csv(file_path, index=False)
                elif file_path.endswith('.xlsx'):
                    df.to_excel(file_path, index=False)
                else:
                    # Añadir extensión por defecto
                    file_path += '.csv'
                    df.to_csv(file_path, index=False)
                
                QMessageBox.information(self, "Exportar Datos", "Datos exportados correctamente.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al exportar los datos: {str(e)}")


def main():
    """Función principal para iniciar la aplicación"""
    app = QApplication(sys.argv)
    window = AcousticProcessingApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()