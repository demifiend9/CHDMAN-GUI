import sys
import os
import subprocess
import re
import signal
import uuid
import platform
from datetime import datetime
from pathlib import Path

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QFileDialog, QTableWidget,
                             QTableWidgetItem, QProgressBar, QLabel,
                             QSpinBox, QMessageBox, QTextEdit, QCheckBox,
                             QLineEdit, QSplitter, QStyle, QHeaderView, QGroupBox, QMenu)
from PyQt6.QtCore import QRunnable, QThreadPool, pyqtSignal, QObject, pyqtSlot, Qt, QSettings
from PyQt6.QtGui import QShortcut, QKeySequence

PROGRESS_REGEX = re.compile(r"(\d+(?:\.\d+)?)\s*%")

class KeepAwake:
    """Handles OS-level calls to prevent the system from going to sleep during batch processing."""
    def __init__(self):
        self.os_type = platform.system()
        self.mac_process = None

    def prevent_sleep(self):
        if self.os_type == 'Windows':
            import ctypes
            # ES_CONTINUOUS | ES_SYSTEM_REQUIRED
            ctypes.windll.kernel32.SetThreadExecutionState(0x80000000 | 0x00000001)
        elif self.os_type == 'Darwin': # macOS
            self.mac_process = subprocess.Popen(['caffeinate', '-i'])

    def allow_sleep(self):
        if self.os_type == 'Windows':
            import ctypes
            # ES_CONTINUOUS (Clears the system required flag)
            ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)
        elif self.os_type == 'Darwin' and self.mac_process:
            self.mac_process.terminate()
            self.mac_process = None

def get_chdman_path():
    """Smart resolution: Look next to the executable first, then fallback to system PATH."""
    exe_name = "chdman.exe" if os.name == 'nt' else "chdman"

    if getattr(sys, 'frozen', False):
        base_dir = Path(sys.executable).parent
    else:
        base_dir = Path(__file__).parent

    local_path = base_dir / exe_name
    if local_path.exists():
        return str(local_path)

    return exe_name

class WorkerSignals(QObject):
    progress = pyqtSignal(str, int)
    finished = pyqtSignal(str, str)
    error = pyqtSignal(str, str)
    log = pyqtSignal(str)

class ChdmanWorker(QRunnable):
    def __init__(self, job_id, input_file, output_file, cmd_type, chdman_exe):
        super().__init__()
        self.job_id = job_id
        self.input_file = input_file
        self.output_file = output_file
        self.temp_output = f"{self.output_file}.partial"
        self.cmd_type = cmd_type
        self.chdman_exe = chdman_exe
        self.signals = WorkerSignals()

        self.is_cancelled = False
        self.process = None

    @pyqtSlot()
    def run(self):
        filename = Path(self.input_file).name
        self.signals.log.emit(f"--- Starting: {filename} ({self.cmd_type}) ---")

        cmd = [self.chdman_exe, self.cmd_type, "-i", self.input_file, "-o", self.temp_output]

        try:
            self.process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                universal_newlines=True,
                bufsize=1
            )

            buffer = ""
            while True:
                if self.is_cancelled:
                    self.process.terminate()
                    self.cleanup_partial()
                    self.signals.error.emit(self.job_id, "Cancelled")
                    self.signals.log.emit(f"--- Cancelled: {filename} ---")
                    return

                char = self.process.stdout.read(1)
                if not char:
                    break

                buffer += char
                if char in ('\r', '\n'):
                    match = PROGRESS_REGEX.search(buffer)
                    if match:
                        percent = int(float(match.group(1)))
                        self.signals.progress.emit(self.job_id, percent)
                    else:
                        clean_line = buffer.strip()
                        if clean_line:
                            self.signals.log.emit(f"[{filename}] {clean_line}")
                    buffer = ""

            self.process.wait()

            if self.process.returncode == 0:
                if os.path.exists(self.output_file):
                    os.remove(self.output_file)
                os.rename(self.temp_output, self.output_file)

                self.signals.finished.emit(self.job_id, "Done")
                self.signals.log.emit(f"--- Finished: {filename} ---")
            else:
                self.cleanup_partial()
                err_msg = f"Failed (Code {self.process.returncode})"
                self.signals.error.emit(self.job_id, err_msg)
                self.signals.log.emit(f"--- Error on {filename}: Exit code {self.process.returncode} ---")

        except FileNotFoundError:
            self.signals.error.emit(self.job_id, "chdman not found")
            self.signals.log.emit(f"ERROR: '{self.chdman_exe}' not found next to app or in system PATH.")
        except Exception as e:
            self.cleanup_partial()
            self.signals.error.emit(self.job_id, "Exception Occurred")
            self.signals.log.emit(f"ERROR on {filename}: {str(e)}")

    def cleanup_partial(self):
        if os.path.exists(self.temp_output):
            try:
                os.remove(self.temp_output)
            except OSError:
                pass

    def cancel(self):
        self.is_cancelled = True

    def toggle_pause(self, pause: bool):
        if os.name == 'posix' and self.process and self.process.pid:
            try:
                if pause:
                    os.kill(self.process.pid, signal.SIGSTOP)
                else:
                    os.kill(self.process.pid, signal.SIGCONT)
            except ProcessLookupError:
                pass

class ChmanMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CHDMAN-GUI")
        self.resize(1100, 750)
        self.setAcceptDrops(True)

        cd_icon = self.style().standardIcon(QStyle.StandardPixmap.SP_DriveCDIcon)
        self.setWindowIcon(cd_icon)

        self.threadpool = QThreadPool()
        self.workers = []
        self.active_jobs = 0
        self.is_paused = False
        self.chdman_path = get_chdman_path()
        self.keep_awake = KeepAwake()

        self.settings = QSettings("OpenSource", "CHDMAN-GUI")

        self.init_ui()
        self.load_settings()

    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)

        top_panels_layout = QHBoxLayout()

        # --- Input Group ---
        grp_input = QGroupBox("Input")
        grp_input_layout = QVBoxLayout()

        row1_layout = QHBoxLayout()

        btn_add_layout = QVBoxLayout()
        self.btn_add_files = QPushButton(" Add files")
        self.btn_add_files.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon))
        self.btn_add_files.clicked.connect(self.browse_files)

        self.lbl_drag = QLabel("(or drag and drop)")
        self.lbl_drag.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_drag.setStyleSheet("color: gray; font-size: 11px;")

        btn_add_layout.addWidget(self.btn_add_files)
        btn_add_layout.addWidget(self.lbl_drag)
        btn_add_layout.setAlignment(Qt.AlignmentFlag.AlignTop)

        row1_layout.addLayout(btn_add_layout)
        row1_layout.addStretch()

        lbl_threads = QLabel("Concurrent Threads:")
        lbl_threads.setToolTip("WARNING: Disk I/O is the bottleneck. High values on HDDs will cause severe slowdowns.")
        row1_layout.addWidget(lbl_threads)

        self.thread_spinner = QSpinBox()
        self.thread_spinner.setToolTip(lbl_threads.toolTip())
        self.thread_spinner.setMinimum(1)
        self.thread_spinner.setMaximum(os.cpu_count() or 4)
        self.thread_spinner.setValue(1)
        row1_layout.addWidget(self.thread_spinner)

        grp_input_layout.addLayout(row1_layout)
        grp_input.setLayout(grp_input_layout)
        top_panels_layout.addWidget(grp_input)

        # --- Output Group ---
        grp_output = QGroupBox("Output")
        grp_output_layout = QVBoxLayout()

        row2_layout = QHBoxLayout()
        self.chk_same_dir = QCheckBox("Output same as input directory")
        self.chk_same_dir.setChecked(True)
        self.chk_same_dir.toggled.connect(self.toggle_custom_outdir)
        row2_layout.addWidget(self.chk_same_dir)

        self.txt_outdir = QLineEdit()
        self.txt_outdir.setEnabled(False)
        self.txt_outdir.setPlaceholderText("Select custom output directory...")
        row2_layout.addWidget(self.txt_outdir)

        self.btn_browse_outdir = QPushButton("Browse")
        self.btn_browse_outdir.setEnabled(False)
        self.btn_browse_outdir.clicked.connect(self.browse_outdir)
        row2_layout.addWidget(self.btn_browse_outdir)

        grp_output_layout.addLayout(row2_layout)

        row3_layout = QHBoxLayout()
        self.chk_save_log = QCheckBox("Save .log to output directory")
        self.chk_save_log.setChecked(False)
        row3_layout.addWidget(self.chk_save_log)

        self.chk_play_sound = QCheckBox("Play sound on finish")
        self.chk_play_sound.setChecked(False)
        row3_layout.addWidget(self.chk_play_sound)

        grp_output_layout.addLayout(row3_layout)
        grp_output.setLayout(grp_output_layout)
        top_panels_layout.addWidget(grp_output)

        layout.addLayout(top_panels_layout)

        # --- Middle Section (Table & Console) ---
        splitter = QSplitter(Qt.Orientation.Vertical)

        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["Filename", "Input Folder", "Output Format", "Progress", "Status"])

        self.table.setSortingEnabled(True)
        self.table.horizontalHeader().setSortIndicator(-1, Qt.SortOrder.AscendingOrder)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(2, 130)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(3, 130)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(4, 130)

        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)

        self.del_shortcut = QShortcut(QKeySequence(Qt.Key.Key_Delete), self.table)
        self.del_shortcut.activated.connect(self.remove_selected_rows)

        splitter.addWidget(self.table)

        self.console = QTextEdit()
        self.console.setReadOnly(True)
        self.console.setStyleSheet("background-color: #1e1e1e; color: #00ff00; font-family: monospace;")
        splitter.addWidget(self.console)

        splitter.setSizes([450, 150])
        layout.addWidget(splitter)

        # --- Bottom Controls ---
        bottom_layout = QHBoxLayout()

        self.btn_start = QPushButton("Start Batch")
        self.btn_start.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay))
        self.btn_start.clicked.connect(self.start_conversion)
        self.btn_start.setStyleSheet("background-color: #2e7d32; color: white; font-weight: bold; padding: 8px;")

        self.btn_pause = QPushButton("Pause")
        self.btn_pause.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPause))
        self.btn_pause.clicked.connect(self.toggle_pause)
        self.btn_pause.setStyleSheet("font-weight: bold; padding: 8px;")

        if os.name != 'posix':
            self.btn_pause.hide()
        else:
            self.btn_pause.setEnabled(False)

        self.btn_stop = QPushButton("Stop/Cancel")
        self.btn_stop.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaStop))
        self.btn_stop.clicked.connect(self.stop_conversion)
        self.btn_stop.setStyleSheet("background-color: #c62828; color: white; font-weight: bold; padding: 8px;")

        self.btn_quit = QPushButton("Quit")
        self.btn_quit.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogCloseButton))
        self.btn_quit.clicked.connect(self.close)
        self.btn_quit.setStyleSheet("font-weight: bold; padding: 8px;")

        bottom_layout.addWidget(self.btn_start)
        if os.name == 'posix':
            bottom_layout.addWidget(self.btn_pause)
        bottom_layout.addWidget(self.btn_stop)
        bottom_layout.addWidget(self.btn_quit)
        layout.addLayout(bottom_layout)

    # --- Settings Handlers ---
    def load_settings(self):
        self.thread_spinner.setValue(int(self.settings.value("threads", 1)))

        same_dir = self.settings.value("same_dir", True, type=bool)
        self.chk_same_dir.setChecked(same_dir)
        self.toggle_custom_outdir(same_dir)

        self.txt_outdir.setText(self.settings.value("custom_outdir", ""))
        self.chk_save_log.setChecked(self.settings.value("save_log", False, type=bool))
        self.chk_play_sound.setChecked(self.settings.value("play_sound", False, type=bool))
        self.last_input_dir = self.settings.value("last_input_dir", os.path.expanduser("~"))

    def save_settings(self):
        self.settings.setValue("threads", self.thread_spinner.value())
        self.settings.setValue("same_dir", self.chk_same_dir.isChecked())
        self.settings.setValue("custom_outdir", self.txt_outdir.text())
        self.settings.setValue("save_log", self.chk_save_log.isChecked())
        self.settings.setValue("play_sound", self.chk_play_sound.isChecked())
        self.settings.setValue("last_input_dir", self.last_input_dir)

    # --- UI Helpers ---
    def find_row_by_id(self, job_id):
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item and item.data(Qt.ItemDataRole.UserRole)['id'] == job_id:
                return row
        return -1

    def toggle_custom_outdir(self, checked):
        self.txt_outdir.setEnabled(not checked)
        self.btn_browse_outdir.setEnabled(not checked)

        if not checked and not self.txt_outdir.text():
            if self.table.rowCount() > 0:
                self.txt_outdir.setText(self.table.item(0, 1).text())
            else:
                self.txt_outdir.setText(os.path.expanduser("~"))

    def browse_outdir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "Select Output Directory", self.txt_outdir.text())
        if dir_path:
            self.txt_outdir.setText(dir_path)

    def show_context_menu(self, pos):
        if self.table.selectionModel().hasSelection():
            menu = QMenu(self)
            remove_action = menu.addAction("Remove Selected Files")
            action = menu.exec(self.table.viewport().mapToGlobal(pos))
            if action == remove_action:
                self.remove_selected_rows()

    def remove_selected_rows(self):
        rows_to_delete = sorted([index.row() for index in self.table.selectedIndexes()], reverse=True)
        rows_to_delete = list(dict.fromkeys(rows_to_delete))

        for row in rows_to_delete:
            status = self.table.item(row, 4).text()
            if status not in ["Processing...", "Paused"]:
                self.table.removeRow(row)

    # --- Input Handlers ---
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        files = []
        valid_exts = ('.cue', '.iso', '.chd', '.gdi', '.toc')
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if os.path.isfile(path) and path.lower().endswith(valid_exts):
                files.append(path)
        if files:
            self.process_added_files(files)

    def browse_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Files", self.last_input_dir,
            "Supported Files (*.cue *.iso *.chd *.gdi *.toc);;All Files (*)"
        )
        if files:
            self.last_input_dir = str(Path(files[0]).parent)
            self.process_added_files(files)

    def probe_chd(self, filepath):
        """Intelligently detects if a CHD should be extracted to CUE or ISO based on MAME metadata tags."""
        try:
            # Capture both stdout and stderr (chdman prints info to stderr!)
            result = subprocess.run([self.chdman_path, "info", "-i", filepath],
                                    capture_output=True, text=True, timeout=2)
            output = (result.stdout + result.stderr).upper()

            # Dreamcast GD-ROMs use extractcd
            if "TAG='CHGD'" in output:
                return "cd"

            # Multiple tracks definitively means CD (.cue)
            if "TRACK:2" in output:
                return "cd"

            # RAW track modes = CD formats (PS1, Sega CD, Saturn)
            if "TYPE:MODE1_RAW" in output or "TYPE:MODE2_RAW" in output:
                return "cd"

            # Standard MODE1 (non-raw) = ISO format (PS2 DVD, PSP UMD, PC, GameCube)
            if "TYPE:MODE1 SUBTYPE:" in output:
                return "dvd"

            # Fallback safely to CD
            return "cd"
        except Exception:
            return "cd"

    def process_added_files(self, files):
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        self.table.setSortingEnabled(False)

        existing_inputs = set()
        for i in range(self.table.rowCount()):
            job_data = self.table.item(i, 0).data(Qt.ItemDataRole.UserRole)
            if job_data:
                existing_inputs.add(job_data['input_path'])

        for file in files:
            input_path = Path(file)

            if str(input_path) in existing_inputs:
                continue

            filename = input_path.name
            ext = input_path.suffix.lower()

            if ext in ['.cue', '.gdi', '.toc']:
                cmd_type = 'createcd'
                out_ext = '.chd'
            elif ext == '.iso':
                cmd_type = 'createdvd'
                out_ext = '.chd'
            elif ext == '.chd':
                chd_type = self.probe_chd(str(input_path))
                if chd_type == "cd":
                    cmd_type = 'extractcd'
                    out_ext = '.cue'
                else:
                    cmd_type = 'extractdvd'
                    out_ext = '.iso'
            else:
                continue

            row = self.table.rowCount()
            self.table.insertRow(row)

            job_id = str(uuid.uuid4())
            job_data = {
                'id': job_id,
                'input_path': str(input_path),
                'cmd_type': cmd_type,
                'out_ext': out_ext
            }
            item_filename = QTableWidgetItem(filename)
            item_filename.setData(Qt.ItemDataRole.UserRole, job_data)

            self.table.setItem(row, 0, item_filename)
            self.table.setItem(row, 1, QTableWidgetItem(str(input_path.parent)))

            format_item = QTableWidgetItem(f"-> {out_ext}")
            format_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(row, 2, format_item)

            progress_bar = QProgressBar()
            progress_bar.setValue(0)
            self.table.setCellWidget(row, 3, progress_bar)

            self.table.setItem(row, 4, QTableWidgetItem("Pending"))

        self.table.setSortingEnabled(True)
        QApplication.restoreOverrideCursor()

    # --- Processing Logic ---
    def append_log(self, text):
        self.console.append(text)
        scrollbar = self.console.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def start_conversion(self):
        self.threadpool.setMaxThreadCount(self.thread_spinner.value())
        self.workers.clear()
        self.is_paused = False

        if os.name == 'posix':
            self.btn_pause.setText("Pause")
            self.btn_pause.setEnabled(True)

        self.btn_quit.setEnabled(False)
        self.console.clear()
        self.append_log(f"--- Batch started at {datetime.now().strftime('%H:%M:%S')} ---")
        self.active_jobs = 0

        use_same_dir = self.chk_same_dir.isChecked()
        custom_out_dir = self.txt_outdir.text()

        for row in range(self.table.rowCount()):
            status = self.table.item(row, 4).text()

            if status == "Pending" or "Error" in status or "Cancelled" in status:
                job_data = self.table.item(row, 0).data(Qt.ItemDataRole.UserRole)
                job_id = job_data['id']
                input_file = job_data['input_path']
                cmd_type = job_data['cmd_type']
                out_ext = job_data['out_ext']

                if use_same_dir or not custom_out_dir:
                    output_file = str(Path(input_file).with_suffix(out_ext))
                else:
                    output_file = str(Path(custom_out_dir) / Path(input_file).with_suffix(out_ext).name)

                if Path(output_file).exists():
                    self.table.item(row, 4).setText("Error: Output exists")
                    self.append_log(f"Skipped {Path(input_file).name}: Target file already exists.")
                    continue

                self.table.item(row, 4).setText("Queued")
                self.table.cellWidget(row, 3).setValue(0)

                worker = ChdmanWorker(job_id, input_file, output_file, cmd_type, self.chdman_path)
                worker.signals.progress.connect(self.update_progress)
                worker.signals.finished.connect(self.job_finished)
                worker.signals.error.connect(self.job_error)
                worker.signals.log.connect(self.append_log)

                self.workers.append(worker)
                self.threadpool.start(worker)
                self.active_jobs += 1

        if self.active_jobs > 0:
            self.keep_awake.prevent_sleep()
        else:
            if os.name == 'posix':
                self.btn_pause.setEnabled(False)
            self.btn_quit.setEnabled(True)
            self.append_log("--- No valid jobs to process ---")

    def toggle_pause(self):
        self.is_paused = not self.is_paused

        if self.is_paused:
            self.btn_pause.setText("Resume")
            self.btn_pause.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay))
            self.append_log("--- BATCH PAUSED ---")
        else:
            self.btn_pause.setText("Pause")
            self.btn_pause.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPause))
            self.append_log("--- BATCH RESUMED ---")

        for worker in self.workers:
            worker.toggle_pause(self.is_paused)

    def stop_conversion(self):
        for worker in self.workers:
            if self.is_paused:
                worker.toggle_pause(False)
            worker.cancel()

        if os.name == 'posix':
            self.btn_pause.setEnabled(False)

        self.btn_quit.setEnabled(True)
        self.keep_awake.allow_sleep()
        self.append_log("--- BATCH CANCELLED BY USER ---")

    def check_batch_completion(self):
        self.active_jobs -= 1
        if self.active_jobs <= 0:
            self.active_jobs = 0

            if os.name == 'posix':
                self.btn_pause.setEnabled(False)

            self.btn_quit.setEnabled(True)
            self.keep_awake.allow_sleep()
            self.append_log(f"--- Batch completed at {datetime.now().strftime('%H:%M:%S')} ---")

            if self.chk_play_sound.isChecked():
                QApplication.beep()

            if self.chk_save_log.isChecked():
                self.save_log_file()

    def save_log_file(self):
        if self.chk_same_dir.isChecked() and self.table.rowCount() > 0:
            job_data = self.table.item(0, 0).data(Qt.ItemDataRole.UserRole)
            out_dir = Path(job_data['input_path']).parent
        else:
            out_dir = Path(self.txt_outdir.text() or os.getcwd())

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = out_dir / f"chdman_log_{timestamp}.log"

        try:
            with open(log_path, "w") as f:
                f.write(self.console.toPlainText())
            self.append_log(f"Log saved successfully to: {log_path}")
        except Exception as e:
            self.append_log(f"Failed to save log: {e}")

    @pyqtSlot(str, int)
    def update_progress(self, job_id, percentage):
        row = self.find_row_by_id(job_id)
        if row != -1:
            progress_bar = self.table.cellWidget(row, 3)
            if progress_bar:
                if self.is_paused:
                    self.table.item(row, 4).setText("Paused")
                else:
                    progress_bar.setValue(percentage)
                    self.table.item(row, 4).setText("Processing...")

    @pyqtSlot(str, str)
    def job_finished(self, job_id, status):
        row = self.find_row_by_id(job_id)
        if row != -1:
            progress_bar = self.table.cellWidget(row, 3)
            if progress_bar:
                progress_bar.setValue(100)
            self.table.item(row, 4).setText(status)
        self.check_batch_completion()

    @pyqtSlot(str, str)
    def job_error(self, job_id, error_msg):
        row = self.find_row_by_id(job_id)
        if row != -1:
            self.table.item(row, 4).setText(f"Error: {error_msg}")
        self.check_batch_completion()

    def closeEvent(self, event):
        self.save_settings()
        self.keep_awake.allow_sleep()

        if self.active_jobs > 0:
            reply = QMessageBox.question(self, 'Quit',
                                         'Jobs are still running. Are you sure you want to quit?',
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                         QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                self.stop_conversion()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = ChmanMainWindow()
    window.show()
    sys.exit(app.exec())
