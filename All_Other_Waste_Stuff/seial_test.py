import os, configparser, datetime, sys, time, openpyxl, serial
import pyvisa as visa
from PyQt5 import uic
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

from popup import MyDialog  # need to remove later


##################################################################################################################################
class WorkerThread(QThread):
    process_completed = pyqtSignal()
    result_signal = pyqtSignal(str)
    response_signal = pyqtSignal(str)
    def __init__(self, commands, serial_port):
        super().__init__()
        self.commands = commands
        self.serial_port = serial_port
    def run(self):
        try:
            try:
                if self.serial_port.is_open:
                    self.serial_port.timeout = 5
                    self.next_command_index = 0
                    for command in self.commands:
                        if self.next_command_index < (len(self.commands)+1):
                            self.serial_port.write(command.encode())
                            # self.textBrowser.append('Processed Command: '+command + ' ')
                            response = self.serial_port.readline().decode('ascii')
                            self.result_signal.emit((f"Response: {response}"))
                            self.next_command_index += 1
                            current_date = datetime.datetime.now()
                            decimal_date = int(current_date.strftime('%Y%m%d'))
                            hex_date = hex(decimal_date)[2:].upper().zfill(8)
                            if command == self.commands[1]:
                                ading = response.split(':')[3][:-1]
                                self.commands[2] = self.commands[2]+ading+'01'+hex_date+'2A0101030000FFFF'
                                self.response_signal.emit(ading)
                            if command == self.commands[2]:
                                start_time = time.time()
                                while time.time() - start_time < 5:
                                    if response:
                                        break
                                    response = self.serial_port.readline().decode('ascii')
                                    self.result_signal.emit(response)
                    self.serial_port.close()
                    self.process_completed.emit()
                else:
                    self.result_signal.emit('Serial Port Closed')
            except IndexError as e:
                self.result_signal.emit(str(e))
        except AttributeError as e:
            self.result_signal.emit(str(e))
    def on_button_clicked(self):
        self.result_signal.emit("Button clicked")
##################################################################################################################################
class SerialPortThread(QThread):
    com_port_found = pyqtSignal(str)

    def run(self):
        for i in range(256):
            com_port = f'COM{i}'
            try:
                serial_port = serial.Serial(com_port, baudrate=115200, timeout=1)
                response = self.check_response(serial_port)
                serial_port.close()
                if 'i2c:devs:5373' in response or response.startswith('i2c:devs:'):
                    self.com_port_found.emit(com_port)
                    break
            except serial.SerialException:
                pass

    def check_response(self, serial_port):
        try:
            serial_port.write('i2c:scan'.encode())
            response = serial_port.readline().decode('ascii').strip()
            return response
        except Exception as e:
            return str(e)
##################################################################################################################################
class App(QMainWindow):
    def __init__(self):
        super(App, self).__init__()
        uic.loadUi("UI/Test.ui", self)
        self.setWindowIcon(QIcon('images_/icons/Moewe.jpg'))
        # self.setFixedSize()
        self.setStatusTip('Moewe Optik Gmbh. 2023')
        self.show()
        self.selected_com_port = None  # New variable to store the selected COM port

        self.serial_port = None
        self.thread = None
        self.worker_thread = SerialPortThread()
        self.worker_thread.com_port_found.connect(self.on_com_port_found)
        self.worker_thread.start()
        self.commands = ['i2c:scan', 'i2c:read:53:04:FC', 'i2c:write:53:', 'i2c:read:53:20:00', 'i2c:write:73:04', 'i2c:scan','i2c:write:21:0300','i2c:write:21:0100','i2c:write:21:01FF', 'i2c:write:73:01',
                    'i2c:scan', 'i2c:write:4F:06990918', 'i2c:write:4F:01F8', 'i2c:read:4F:1E:00']
        self.start_button.clicked.connect(self.start_process)

    def on_com_port_found(self, com_port):
        self.textBrowser.append(f"Found COM Port: {com_port}")
        if self.selected_com_port is None:
            self.selected_com_port = com_port

    def on_widget_button_clicked(self, message):
        self.textBrowser.append(message)

    def update_lineinsert(self, response):
        self.id_Edit.setText(response)

    def start_process(self):
        if self.start_button.text() == 'START':
            if self.selected_com_port:
                try:
                    self.serial_port = serial.Serial(self.selected_com_port, 115200, timeout=1)
                    self.textBrowser.append('Serial Communication Connected')
                    self.start_button.setText('STOP')
                    self.start()
                except serial.SerialException as e:
                    QMessageBox.warning(self, "Serial Port Error", f"Error opening serial port: {e}")
        elif self.start_button.text() == 'STOP':
            try:
                self.serial_port.close()
            except serial.SerialException as e:
                QMessageBox.warning(self, "Serial Port Error", f"Error closing serial port: {e}")
            self.textBrowser.append('Communication Disconnected')
            self.start_button.setText('START')

    def start(self):
        if self.thread is None or not self.thread.isRunning():
            self.info_label.setText("Process Started")
            self.start_button.setEnabled(False)  # Disable the button while the process is running
            self.thread = WorkerThread(self.commands, self.serial_port)
            self.thread.result_signal.connect(self.on_widget_button_clicked)
            self.thread.process_completed.connect(self.process_completed)
            self.thread.response_signal.connect(self.update_lineinsert)
            self.thread.finished.connect(self.process_finished)  # Connect the finished signal to enable the button
            self.thread.start()
        else:
            QMessageBox.warning(self, "Process In Progress", "Process is already running.")

    def process_completed(self):
        self.info_label.setText("Process has been completed.")

    def process_finished(self):
        self.start_button.setEnabled(True)  # Enable the button when the process is finished

        
def main():
    app = QApplication(sys.argv)
    Window = App()
    sys.exit(app.exec_())
if __name__ == '__main__':
    main()
