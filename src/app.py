import os
import sys
import ctypes
import locale
import importlib
import requests
from pathlib import Path
import pandas as pd

from PyQt6.QtCore import QMutex, QObject, QThread, pyqtSignal, QUrl, Qt
from PyQt6.QtGui import QAction, QIcon, QDoubleValidator, QDesktopServices
from PyQt6.QtWidgets import (
    QApplication,
    QLabel,
    QMainWindow,
    QPushButton,
    QVBoxLayout,
    QHBoxLayout,
    QWidget,
    QScrollArea,
    QStyle,
    QMessageBox,
    QProgressBar,
    QLineEdit,
    QCheckBox,
    QTextEdit,
    QFileDialog
)
from openpyxl.utils.exceptions import InvalidFileException
from export_excel import ExportExcel

if '_PYIBoot_SPLASH' in os.environ and importlib.util.find_spec("pyi_splash"):
    import pyi_splash
    pyi_splash.update_text('Loading...')
    pyi_splash.close()


if importlib.util.find_spec("win32com"):
    from win32com.client import *
    def get_version_number(file_path):
        information_parser = Dispatch("Scripting.FileSystemObject")
        version = information_parser.GetFileVersion(file_path)
        return version
    VERSION = get_version_number(sys.argv[0])
else:
    VERSION = 'DEV VERSION'

if os.name == 'nt':
    windll = ctypes.windll.kernel32
    windll.GetUserDefaultUILanguage()
    LANGUAGE = locale.windows_locale[ windll.GetUserDefaultUILanguage() ]
    try:
        from ctypes import windll  # Only exists on Windows.
        APPID = 'joeklein.fk.plugindepth.1'
        windll.shell32.SetCurrentProcessExplicitAppUserModelID(APPID)
    except ImportError:
        pass

if os.name == 'posix':
    LANGUAGE = locale.getlocale()[0]


mutex = QMutex()
pd.set_option('display.max_columns', None)
basedir = os.path.dirname(__file__)


class MeasurmentTask(QObject):
    finished = pyqtSignal()
    broken = pyqtSignal()
    addData = pyqtSignal(list)
    addMinMaxData = pyqtSignal(list)
    terminated = pyqtSignal(bool)

    def __init__(self, info_text, cycle_filter, measurments_folder_path, infusion_filter_entry, injection_filter_entry, output):
        super().__init__()
        self.info_text = info_text
        self.cycle_filter = cycle_filter
        self.measurments_folder_path = measurments_folder_path
        self.infusion_filter_entry = infusion_filter_entry
        self.injection_filter_entry = injection_filter_entry
        self.output = output
        self.df_output = []
        self.file = ''
        self.output_file_data = ()

    def calc_cycles(self, df1, df2):
        df1.drop(index=df1.index[:20], inplace=True)
        df1.dropna(inplace=True)
        df1.reset_index(drop=True, inplace=True)

        # Find the positions where a new cycle starts (value < 0.01)
        cycle_start1 = df1[df1 < float(self.cycle_filter.text().replace(',','.'))].index.tolist()

        # Default value if df2 is None
        cycle_start = cycle_start1

        if df2 is not None:
            df2.drop(index=df2.index[:20], inplace=True)
            df2.dropna(inplace=True)
            df2.reset_index(drop=True, inplace=True)

            # Find the positions where a new cycle starts (value < 0.01)
            cycle_start2 = df2[df2 < float(self.cycle_filter.text().replace(',','.'))].index.tolist()

            # Use the shortest cycle start points to get equal cycles
            cycle_start = cycle_start2
            if len(cycle_start1) < len(cycle_start2):
                cycle_start = cycle_start1

        # Find the length of the shortest cycle amount
        cycle_length = len(cycle_start)

        if cycle_length == len(df1.index):
            return None, None

        # Divide the data into cycles
        cycles1 = [df1[cycle_start[i]:cycle_start[i+1]] if i < cycle_length - 1 else df1[cycle_start[i]:] for i in range(cycle_length)]
        cycles2 = [ None for _ in range(cycle_length)]

        if df2 is not None:
            cycles2 = [df2[cycle_start[i]:cycle_start[i+1]] if i < cycle_length - 1 else df2[cycle_start[i]:] for i in range(cycle_length)]
        return cycles1, cycles2

    def get_limits(self, df, name):
        min_min = df.iloc[0][name]
        max_max = df.iloc[1][name]
        return min_min, max_max

    def evaluation(self, file):
        col_infusion = []
        col_injection = []
        col_error_infusion = []
        col_error_injection = []
        oneport = False

        try:
            file_path = os.path.join(self.measurments_folder_path, file[0])
            df = pd.read_excel(file_path, engine='openpyxl')

            if not set(['Infusion']).issubset(df.columns):
                self.info_text.insertPlainText(f'This file ({file[0]}) is beyond my capabilities.\nSorry it\'s my first day coding.! 👍\n')
                self.broken.emit()
                return

            if not set(['Injection']).issubset(df.columns):
                self.info_text.insertPlainText(f'One Port file ({file[0]}) detected.\n')
                oneport = True

            limits = [self.get_limits(df, 'Infusion'), self.get_limits(df, 'Injection') if not oneport else (0, 0)]
            infusion, injection = self.calc_cycles(df['Infusion'], df['Injection'] if not oneport else None)

            if not infusion or (not injection and not oneport):
                self.info_text.insertPlainText('Phu you filtered the shi* out of the Cycle values.\n')
                self.broken.emit()
                return

            # Add maximum values to result array
            for (infusion, injection) in zip(infusion, injection):
                infusion_max = max(infusion)
                append_infusion = infusion_max >= float(self.infusion_filter_entry.text().replace(',','.'))

                append_injection = False
                injection_max = 0
                if injection is not None:
                    injection_max = max(injection)
                    append_injection = injection_max >= float(self.injection_filter_entry.text().replace(',','.'))

                if append_infusion or append_injection:
                    col_infusion.append(infusion_max)
                    col_injection.append(injection_max)

                if append_infusion and not append_injection:
                    col_error_infusion.append(False)
                    col_error_injection.append(True)
                elif append_injection and not append_infusion:
                    col_error_injection.append(False)
                    col_error_infusion.append(True)
                elif append_infusion and append_injection:
                    col_error_infusion.append(False)
                    col_error_injection.append(False)

            if len(col_infusion) == 0:
                self.info_text.insertPlainText('Phu you filtered the shi* out of the Infusion values.\n')
                self.broken.emit()
                return

            if len(col_injection) == 0:
                self.info_text.insertPlainText('Phu you filtered the shi* out of the Injection values.\n')
                self.broken.emit()
                return

            if all(col_error_injection) is True and not oneport:
                self.info_text.insertPlainText(f'\nWarning: No injection measurements found. One Port file ({file[0]}) detected.\n\n')
                oneport = True

            if len(col_infusion) > 15:
                self.info_text.insertPlainText(f"\nWarning: More than 15 measurements found on file \"{file[0]}\".\nThe handover from 3T to 2F could have triggered a new cycle. Adjust the filters to remove unwanted measurements.\n\n")

            self.df_output = pd.DataFrame(list(zip(col_infusion, col_injection, col_error_infusion, col_error_injection)))
            self.df_output.columns =['Infusion', 'Injection', 'Error Infusion', 'Error Injection']
            self.output_file_data = (file[0], self.df_output, limits, oneport)

            mutex.lock()
            self.addData.emit(self.output_file_data)
            mutex.unlock()
            self.finished.emit()
        
        except InvalidFileException:
            self.info_text.insertPlainText(f'\nError: File "{file[0]}" is not an Excel file or cant be read!\n\n')
            mutex.lock()
            self.terminated.emit(True)
            self.broken.emit()
            mutex.unlock()
        except Exception as e:
            if str(e) == 'File is not a zip file':
                self.info_text.insertPlainText(f'\nError: File "{file[0]}" is not an Excel file or cant be read!\n\n')
            else:
                self.info_text.insertPlainText(f'\nError: File "{file[0]}" abort with exception:\n{e}\n\n')
            mutex.lock()
            self.terminated.emit(True)
            self.broken.emit()
            mutex.unlock()


class Window(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi()
        self.threads = []
        self.output_file = ''
        self.output_file_data = []
        self.min_max_data = []
        self.measurment_files = []
        self.measurments_folder_path = None
        self.count_terminated = None

    def select_folder(self):
        self.measurments_folder_path = QFileDialog.getExistingDirectory(self, 'Measurements', options=QFileDialog.Option.DontUseNativeDialog)
        if self.measurments_folder_path:
            self.measurments_folder_entry.setText(self.measurments_folder_path)

    def select_output_file(self):
        file_path = QFileDialog.getSaveFileName(self, 'Export F:xile', '2F_plugin_depth_result', 'Excel files (*.xlsx)', options=QFileDialog.Option.DontUseNativeDialog)
        self.output_file = file_path[0]
        if file_path[0]:
            self.output_file_entry.setText(file_path[0])
            self.startThreads()

    def empty_output_file(self):
        self.output_file_entry.setText('')
        self.output_file = self.output_file_entry.text()

    def update_measurments_folder(self):
        self.measurments_folder_path = self.measurments_folder_entry.text()
        if os.path.exists(self.measurments_folder_path):
            self.show_measurment_files()

    def update_output_file(self):
        self.output_file = self.output_file_entry.text()

    def show_measurment_files(self):
        if not self.measurments_folder_path:
            self.info_text.insertPlainText('No source folder selected!\n')
            return

        for file in self.measurment_files:
            file[1].setParent(None)
        self.measurment_files = []

        try:
            for file in os.listdir(path=str(self.measurments_folder_path)):
                if Path(file).suffix == '.xlsx' and 'result' not in file.lower() and file[:1] != '.':
                    self.file_area.show()
                    self.select_none.show()
                    self.select_all.show()
                    self.setMinimumHeight(500)
                    checkbox = QCheckBox(file)
                    self.measurment_files.append((file, checkbox))
                    self.file_layout.addWidget(checkbox)
        except Exception:
            # Dir does not exist
            pass

        if not self.measurment_files:
            self.select_none.hide()
            self.select_all.hide()
            self.file_area.hide()
            self.setMinimumHeight(250)
            self.resize(800,250)

    def select_measurment_files(self, state):
        for file in self.measurment_files:
            file[1].setChecked(state)

    def select_all_measurment_files(self):
        self.select_measurment_files(True)

    def select_none_measurment_files(self):
        self.select_measurment_files(False)

    def write_to_excel(self):
        writer = ExportExcel(self.output_file_data, f'{self.output_file}.xlsx', self.info_text, LANGUAGE)
        writer.write_to_excel()
        self.prog_bar.setValue(100)
        self.info_text.insertPlainText(f'File was saved at\n{self.output_file}.xlsx')

        title = 'Export successfully!'
        text = 'Export successfully!'
        buttonText = 'Open Export'
        icon = QMessageBox.Icon.Information
        if self.count_terminated != 0:
            title = 'Export successfully with error!'
            text = 'Export successfully with error!'
            buttonText= 'Open export anyway'
            icon = QMessageBox.Icon.Warning

        self.msg_box(title=title,
                     text=text,
                     buttonText=buttonText,
                     icon=icon,
                     buttonClick=lambda _, path=QUrl.fromLocalFile(f'{self.output_file}.xlsx'): QDesktopServices.openUrl(path))

    def show_measurments(self):
        for name, df, _, oneport in self.output_file_data:
            df.index += 1
            if oneport:
                df.drop(['Injection', 'Error Injection'], axis=1, inplace=True)
            self.info_text.insertHtml(f"<p style='font-family:Verdana;font-weight:bold;'>{name}</p>\n")
            self.info_text.insertPlainText(f'\n{df}\n')
    
    def add_data_output(self, data=None):
        if data:
            self.output_file_data.append(data)

        step = 100/len(self.threads)/3*2
        step += self.prog_bar.value()
        self.prog_bar.setValue(int(step))
        self.prog_bar.show()

        if self.count_terminated != 0:
            self.info_text.setMaximumSize(1920,850)
            self.resize(800,650)

        if len(self.threads) - self.count_terminated == len(self.output_file_data):
            if self.output_file:
                self.write_to_excel()
            else:
                self.show_measurments()
            self.prog_bar.hide()
            self.evaluate_button.show()

    def startThreads(self):
        self.threads.clear()
        self.output_file_data = []
        self.info_text.setPlainText('')
        self.count_terminated = 0

        if not self.measurments_folder_path:
            self.info_text.insertPlainText('No source folder selected!\n')
            self.evaluate_button.show()
            return

        if sum([1 for file in self.measurment_files if file[1].isChecked()]) == 0:
            self.info_text.insertPlainText('No data source selected!\n')
            self.evaluate_button.show()
            return
        
        self.info_text.setPlainText('')
        output = False
        if self.output_file:
            output = True
            self.info_text.setMaximumSize(1920,60)
            self.resize(800,450)
        else:
            self.info_text.setMaximumSize(1920,850)
            self.resize(800,650)

        self.threads = [
            self.createThread(file, output)
            for file in self.measurment_files if file[1].isChecked()
        ]

        step = 100/len(self.threads)/3*len(self.threads)
        for thread in self.threads:
            self.evaluate_button.hide()
            self.prog_bar.setValue(0)
            thread.start()
            self.prog_bar.setValue(int(step))
            self.prog_bar.show()

    def createThread(self, file, output):
        thread = QThread()
        worker = MeasurmentTask(self.info_text,
                                self.cycle_filter_entry,
                                self.measurments_folder_path,
                                self.infusion_filter_entry,
                                self.injection_filter_entry,
                                output
                                )
        worker.moveToThread(thread)
        thread.started.connect(lambda: worker.evaluation(file))
        worker.broken.connect(lambda: self.kill_thread(thread))
        worker.terminated.connect(self.count_terminated_threads)
        worker.addData.connect(self.add_data_output)
        worker.finished.connect(thread.terminate)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        return thread
    
    def count_terminated_threads(self, status):
        if status:
            self.count_terminated += 1
            self.add_data_output()
    
    def kill_thread(self, thread):
        if thread in self.threads:  # Check if the thread is in the list of threads
            thread.terminate()

    def show_info(self):
        self.msg_box('About', 
                     f'Version {VERSION}\nJoel Klein (jkfres)\n',
                     buttonText='Open Github',
                     buttonClick=lambda: QDesktopServices.openUrl(QUrl('https://github.com/jkfres/2F-plug-in-depth-evaluation/'))
                    )

    def msg_box(self, title, text, icon: QMessageBox.Icon=QMessageBox.Icon.Information, buttonText=None, buttonClick=None):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle(title)
        msg_box.setIcon(icon)
        msg_box.setText(text)
        msg_box.addButton(QMessageBox.StandardButton.Ok)
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        if buttonText and buttonClick:
            custom_button = msg_box.addButton(buttonText, QMessageBox.ButtonRole.ActionRole)
            custom_button.clicked.connect(buttonClick)
        msg_box.exec()

    def setupUi(self):
        self.setWindowTitle(f'2F plug in depth evaluation v{VERSION}')
        self.setMinimumWidth(800)
        self.setMinimumHeight(250)

        self.measurments_folder_path = []
        self.measurment_files = []

        self.validator_float = QDoubleValidator(0.000, 10.000, 3)
        self.validator_float.setNotation(QDoubleValidator.Notation.StandardNotation)


        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        layout = QVBoxLayout()
        self.central_widget.setLayout(layout)

        upper_layout = QVBoxLayout()
        upper_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        layout.addLayout(upper_layout)
        upper_layout.addSpacing(5)

        layout.addStretch()

        lower_layout = QVBoxLayout()
        upper_layout.setAlignment(Qt.AlignmentFlag.AlignBottom)
        layout.addLayout(lower_layout)
        lower_layout.addSpacing(15)

        select_measurments_action = QAction("Select Measurments", self)
        select_measurments_action.setShortcut("Ctrl+m")
        select_measurments_action.setStatusTip("Select Measurments Folder")
        select_measurments_action.triggered.connect(self.select_folder)

        reload_measurments_action = QAction("Reload Mesasurment Files", self)
        reload_measurments_action.setShortcut("Ctrl+r")
        reload_measurments_action.setStatusTip("Reload Mesasurment Files")
        reload_measurments_action.triggered.connect(self.show_measurment_files)

        select_output_action = QAction("Select Output", self)
        select_output_action.setShortcut("Ctrl+x")
        select_output_action.setStatusTip("Select Output File")
        select_output_action.triggered.connect(self.select_output_file)

        evaluate_action = QAction("Evaluate", self)
        evaluate_action.setShortcut("Ctrl+Return")
        evaluate_action.setStatusTip("Evaluate")
        evaluate_action.triggered.connect(self.startThreads)


        about_action = QAction('About', self)
        about_action.setStatusTip('Show info')
        about_action.triggered.connect(self.show_info)

        exit_action = QAction("Exit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.setStatusTip("Exit application")
        exit_action.triggered.connect(self.close) # type: ignore

        menu = self.menuBar()
        file_menu = menu.addMenu("&Help")
        file_menu.addAction(select_measurments_action)
        file_menu.addAction(reload_measurments_action)
        file_menu.addSeparator()
        file_menu.addAction(select_output_action)
        file_menu.addAction(evaluate_action)
        file_menu.addSeparator()
        file_menu.addAction(about_action)
        file_menu.addAction(exit_action)

        # BoxLayout Measurments Folder selection
        layout_measurments_folder = QHBoxLayout()
        upper_layout.addLayout(layout_measurments_folder)

        folder_label = QLabel('Measurements:')
        layout_measurments_folder.addWidget(folder_label)

        self.measurments_folder_entry = QLineEdit()
        layout_measurments_folder.addWidget(self.measurments_folder_entry)
        self.measurments_folder_entry.textChanged[str].connect(self.update_measurments_folder)

        self.folder_button = QPushButton('Select', clicked=self.select_folder) # type: ignore
        layout_measurments_folder.addWidget(self.folder_button)

        self.folder_scan_button = QPushButton('Scan', clicked=self.show_measurment_files) # type: ignore
        layout_measurments_folder.addWidget(self.folder_scan_button)

        # ScrollArea Measurments selection
        self.file_area = QScrollArea()
        self.file_area.setWidgetResizable(True)
        self.file_area.setMinimumSize(300,200)
        self.file_area.setMaximumSize(1920,550)
        self.file_area.hide()

        file_widget = QWidget()
        self.file_area.setWidget(file_widget)
        upper_layout.addWidget(self.file_area)

        self.file_layout = QVBoxLayout()
        file_widget.setLayout(self.file_layout)

        self.file_select_buttons_layout = QHBoxLayout()
        self.file_select_buttons_layout.addSpacing(20)
        self.file_select_buttons_layout.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.file_select_buttons_layout.setContentsMargins(0,0,4,2)
        upper_layout.addLayout(self.file_select_buttons_layout)

        pixmapi = QStyle.StandardPixmap.SP_DialogApplyButton
        icon = self.style().standardIcon(pixmapi)
        self.select_all = QPushButton('All', clicked=self.select_all_measurment_files) # type: ignore
        self.select_all.setIcon(icon)
        self.select_all.setFixedSize(60, 25)
        self.file_select_buttons_layout.addWidget(self.select_all, alignment=Qt.AlignmentFlag.AlignRight)
        self.select_all.hide()

        pixmapi = QStyle.StandardPixmap.SP_DialogCloseButton
        icon = self.style().standardIcon(pixmapi)
        self.select_none = QPushButton('None', clicked=self.select_none_measurment_files) # type: ignore
        self.select_none.setIcon(icon)
        self.select_none.setFixedSize(60, 25)
        self.file_select_buttons_layout.addWidget(self.select_none, alignment=Qt.AlignmentFlag.AlignRight)
        self.select_none.hide()

        filter_layout = QHBoxLayout()
        upper_layout.addLayout(filter_layout)

        cycle_filter_label = QLabel('Cycle filter:')
        filter_layout.addWidget(cycle_filter_label)

        self.cycle_filter_entry = QLineEdit()
        filter_layout.addWidget(self.cycle_filter_entry)
        self.cycle_filter_entry.setValidator(self.validator_float)
        self.cycle_filter_entry.setText('0,01')

        infusion_filter_label = QLabel('Infusion filter:')
        filter_layout.addWidget(infusion_filter_label)

        self.infusion_filter_entry = QLineEdit()
        filter_layout.addWidget(self.infusion_filter_entry)
        self.infusion_filter_entry.setValidator(self.validator_float)
        self.infusion_filter_entry.setText('0,2')

        injection_layout = QHBoxLayout()
        layout.addLayout(injection_layout)

        injection_filter_label = QLabel('Injection filter:')
        filter_layout.addWidget(injection_filter_label)

        self.injection_filter_entry = QLineEdit()
        filter_layout.addWidget(self.injection_filter_entry)
        self.injection_filter_entry.setValidator(self.validator_float)
        self.injection_filter_entry.setText('0,1')

        output_layout = QHBoxLayout()
        lower_layout.addLayout(output_layout)

        self.output_file_label = QLabel('Result:')
        output_layout.addWidget(self.output_file_label)

        self.output_file_entry = QLineEdit()
        output_layout.addWidget(self.output_file_entry)
        self.output_file_entry.textChanged[str].connect(self.update_output_file)

        self.output_file_button = QPushButton('Select', clicked=self.select_output_file) # type: ignore
        output_layout.addWidget(self.output_file_button)

        self.empty_output_file_button = QPushButton('Empty', clicked=self.empty_output_file) # type: ignore
        output_layout.addWidget(self.empty_output_file_button) 

        self.evaluate_button = QPushButton('Evaluate', clicked=self.startThreads) # type: ignore
        lower_layout.addWidget(self.evaluate_button)

        self.prog_bar = QProgressBar(self)
        self.prog_bar.setGeometry(50, 100, 250, 30)
        self.prog_bar.setValue(0)
        lower_layout.addWidget(self.prog_bar)
        self.prog_bar.hide()

        self.info_text = QTextEdit()
        self.info_text.setMinimumSize(300,60)
        self.info_text.setMaximumSize(1920,60)
        lower_layout.addWidget(self.info_text)

        self.check_for_update()

    def check_for_update(self):
        try:
            response = requests.get('https://api.github.com/repos/jkfres/2F-plug-in-depth-evaluation/releases')
            if response.status_code != 200:
                return

            releases = response.json()
            if not releases:
                return

            if releases[0]['tag_name'] != VERSION and not 'DEV VERSION':
                self.msg_box(title='Eine neue Version ist verfügbar!', text='Update verfügbar', icon=QMessageBox.Icon.Information, buttonText='Update herunterladen', buttonClick=lambda: QDesktopServices.openUrl(QUrl('https://github.com/jkfres/2F-plug-in-depth-evaluation/releases')))
        except Exception:
            pass


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(os.path.join(basedir,'files','icon.ico')))
    window = Window()
    window.show()
    sys.exit(app.exec())
