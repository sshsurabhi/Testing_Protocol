import configparser, datetime, os, sys, time, openpyxl, serial, threading
import pyvisa as visa
from PyQt5 import uic
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

# Import the necessary modules for the new thread
from PyQt5.QtCore import QThread, pyqtSignal

class DCVoltageThread(QThread):
    voltage_readings_signal = pyqtSignal(list)
    def __init__(self, multimeter):
        super().__init__()
        self.multimeter = multimeter
        print(self.multimeter)
    def run(self):
        try:
            time.sleep(2)
            self.voltage_readings = []
            for _ in range(20):
                self.multimeter.write('CONF:VOLT:DC 10')
                voltage_reading = self.multimeter.query('READ?')
                self.voltage_readings.append(float(voltage_reading))
            self.voltage_readings_signal.emit(self.voltage_readings)
        except Exception as e:
            print(f"Error in DCVoltageThread: {str(e)}")

class ACVoltageThread(QThread):
    voltage_signal = pyqtSignal(list)
    def __init__(self, multimeter):
        super().__init__()
        self.multimeter = multimeter
    def run(self):
        voltage_readings = []
        for _ in range(20):
            self.multimeter.write('CONF:VOLT:AC 1')
            voltage_reading = self.multimeter.query('READ?')
            voltage_readings.append(float(voltage_reading))
        self.voltage_signal.emit('Hi'+ voltage_readings)

class MyDialog(QMainWindow):
    def __init__(self, parent=None):
        super(MyDialog, self).__init__(parent)
        uic.loadUi("UI/untitled.ui", self)
        self.setWindowIcon(QIcon('2.jpg'))
        self.setFixedSize(self.size())
        self.show()

        self.test_images = ['images_/images/R700.jpg','images_/images/R709_before_jumper.jpg','images_/images/R700_DC.jpg', 'images_/images/PP2.png','images_/images/C443.jpg','images_/images/C442.jpg','images_/images/C441.jpg','images_/images/C412.jpg','images_/images/C430.jpg','images_/images/C443_1.jpg','images_/images/C442_1.jpg','images_/images/C441_1.jpg','images_/images/C412_1.jpg','images_/images/C430_1.jpg', 'images_/images/PP.jpg']
        self.pushButton.clicked.connect(self.change_image)
        self.pushButton_2.clicked.connect(self.change_power)
        self.test_index = 0

        self.DC_event = threading.Event()
        self.AC_event = threading.Event()
        self.DC_event.clear()
        self.AC_event.clear()

        self.rm = visa.ResourceManager()
        self.multimeter = self.rm.open_resource('TCPIP0::192.168.222.207::INSTR')
        print(self.multimeter.query('*IDN?'))
        self.powersupply = self.rm.open_resource('TCPIP0::192.168.222.141::INSTR')
        print(self.powersupply.query('*IDN?'))
        self.powersupply.write('OUTPut CH1,ON')
        self.on_button_click('images_/images/R709.jpg')

        self.DC_readings = []


    def on_widget_button_clicked(self, message):
            self.DC_readings = message
            message_1 = max(message)

            self.textBrowser.append(str(message_1))

    def on_button_click(self, file_path):
        if file_path:
            pixmap = QPixmap(file_path)
            self.label.setPixmap(pixmap)
            self.label.setScaledContents(True)
            self.label.setFixedSize(pixmap.width(), pixmap.height())

    def change_power(self):
        # Wait for 2 seconds
        time.sleep(2)

        # Start the DCVoltageThread
        self.change_image()

    def change_image(self):
        self.voltage_thread = DCVoltageThread(self.multimeter)
        self.voltage_thread.voltage_readings_signal.connect(self.on_widget_button_clicked)

        # Wait for 2 seconds
        time.sleep(2)

        # Start the DCVoltageThread
        if not self.voltage_thread.isRunning():
            self.voltage_thread.start()
            

        # # Wait for DC_voltage readings to reach the desired length (20)
        # while len(self.voltage_thread.voltage_readings_signal) < 20:
        #     print('DC readings', self.DC_readings)
        #     time.sleep(1)

        # # Set the DC_event to True
        # self.DC_event.set()

        # # Wait for 2 seconds
        # time.sleep(2)

        # # Start the ACVoltageThread if DC_event is True
        # if self.DC_event.is_set():
        #     self.change_ac_voltage()

    def change_ac_voltage(self):
        self.ac_voltage_thread = ACVoltageThread(self.multimeter)
        self.ac_voltage_thread.voltage_signal.connect(self.on_widget_button_clicked)

        # Wait for AC_voltage readings to reach the desired length (20)
        while len(self.ac_voltage_thread.voltage_readings) < 20:
            time.sleep(1)

        # Set the AC_event to True
        self.AC_event.set()

    def closeEvent(self, event):
        if self.voltage_thread.isRunning():
            self.voltage_thread.quit()
            self.voltage_thread.wait()
        event.accept()

def main():
    app = QApplication([])
    Window = MyDialog()
    app.exec_()

if __name__ == '__main__':
    main()
