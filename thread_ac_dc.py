import pyvisa as visa
# rm = visa.ResourceManager()
# multimeter = rm.open_resource('TCPIP0::192.168.222.207::INSTR')

# multimeter.write(":MEASure:VOLTage:AC? (@104)")
# acv_measurement_channel_4 = multimeter.read()
# print(f"AC Voltage (Channel 4): {acv_measurement_channel_4}")
# multimeter.write(":MEASure:VOLTage:DC? (@105)")
# dcv_measurement_channel_5 = multimeter.read()
# print(f"DC Voltage (Channel 5): {dcv_measurement_channel_5}")
# print(multimeter.query('ROUT:CHAN? 4'))
# # print(multimeter.query('ROUTe:STARt?'))
# print(multimeter.query('ROUTe:FUNCtion?'))
# print(multimeter.query('ROUTe:DELay?'))
# print(multimeter.query('ROUTe:COUNt?'))
# print(multimeter.write("ROUTe:SCAN ON"))
# print(multimeter.write("ROUT:CHAN 1,ACV"))
# print(multimeter.query("ROUTe:STARt?"))
# print(multimeter.query("ROUT:CHAN? 5"))

# import pyvisa as visa
# import threading
# import time

# # Function to continuously read and print measurements from the specified channel
# def read_and_print(channel, measurement_type):
#     while True:
#         measurement = multimeter.query(f":MEASure:VOLTage:{measurement_type} {channel}")
#         print(f"Channel {channel} ({measurement_type}): {measurement}")
#         time.sleep(1)  # Adjust the sleep interval as needed

# # Initialize the VISA resource manager
# rm = visa.ResourceManager()

# # Connect to the instrument
# instrument_address = 'TCPIP0::192.168.222.207::inst0::INSTR'  # Replace with your instrument's VISA address
# multimeter = rm.open_resource(instrument_address)

# # Configure Channel 1 for DCV measurements
# multimeter.write("ROUTe:CHANnel 1,DCV")

# # Configure Channel 2 for ACV measurements
# multimeter.write("ROUTe:CHANnel 2,ACV")

# # Create two threads for continuous measurements
# thread_dcv = threading.Thread(target=read_and_print, args=(1, "DC"))
# thread_acv = threading.Thread(target=read_and_print, args=(2, "AC"))

# # Start the threads
# thread_dcv.start()
# thread_acv.start()

# # Join the threads to keep the program running
# thread_dcv.join()
# thread_acv.join()

# from pyvisa import ResourceManager

# # Connect to the card reader
# rm = ResourceManager()
# card_reader = rm.open_resource('TCPIP0::192.168.222.207::INSTR')

# # Configure the card reader for continuous reading
# card_reader.write('*RST')  # reset the card reader
# card_reader.write(':*OPC?')  # configure for continuous reading

# while True:
#     try:
#         # Read card data from the card reader
#         card_data = card_reader.read()
#         print(f"Card data: {card_data}")
#     except KeyboardInterrupt:
#         # Clean up the connection to the card reader when the script is interrupted
#         card_reader.close()
#         break

##################################################################################################################################################

# import sys, time
# from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QTextBrowser, QVBoxLayout, QWidget
# from PyQt5.QtCore import QThread, pyqtSignal

# class MultimeterThread(QThread):
#     measurement_completed = pyqtSignal(str)
#     dcv_volatge = pyqtSignal(str)
#     acv_volatge = pyqtSignal(str)
#     def __init__(self, resource_name):
#         super().__init__()
#         self.multimeter = resource_name
#         self.running = True
#     def run(self):
#         try:
#             scan_ON = self.multimeter.query('ROUTe:STATe?')
#             self.multimeter.write('ROUTe:SCAN 1')
#             self.multimeter.write('ROUTe:START ON')
#             self.multimeter.write('ROUTe:START ON')
#             self.multimeter.write('ROUTe:LIMI:HIGH 3')
#             self.multimeter.write('ROUTe:LIMI:LOW 2')
#             # channel_low = self.multimeter.query('ROUTe:LIMI:LOW?')
#             # time.sleep(2)
#             # func_doing = self.multimeter.query('ROUTe:FUNC?')
#             self.multimeter.write('ROUT:CHAN 2,ON,ACV,AUTO,FAST')
#             self.multimeter.write('ROUT:RELA ACV,ON')
#             self.multimeter.write('ROUT:CHAN 3,ON,DCV,AUTO,FAST')
#             self.multimeter.write('ROUT:RELA DCV,ON')
#             while self.running:
#                 ACV_Volt = self.multimeter.query('ROUTe:DATA? 2')
#                 DCV_Volt = self.multimeter.query('ROUTe:DATA? 3')
#                 self.dcv_volatge.emit(DCV_Volt)
#                 self.acv_volatge.emit(ACV_Volt)
#                 time.sleep(2)
#             self.measurement_completed.emit(scan_ON)
#         except AttributeError:
#             print('error occured')
    
#     def stop(self):
#         self.running = False

#     # def __del__(self):
#     #     self.quit()
#     #     self.wait()
# class MultimeterApp(QMainWindow):
#     def __init__(self):
#         super().__init__()
#         self.setWindowTitle("Multimeter Data Logger")
#         self.setGeometry(100, 100, 600, 400)

#         self.initUI()

#     def initUI(self):
#         layout = QVBoxLayout()

#         self.btn_connect = QPushButton("Connect and Read")
#         self.btn_connect.clicked.connect(self.start_measurements)

#         self.text_browser_ac = QTextBrowser()
#         self.text_browser_dc = QTextBrowser()

#         layout.addWidget(self.btn_connect)
#         layout.addWidget(self.text_browser_ac)
#         layout.addWidget(self.text_browser_dc)

#         container = QWidget()
#         container.setLayout(layout)
#         self.setCentralWidget(container)


#         self.rm = visa.ResourceManager()
#         self.multimeter = self.rm.open_resource("TCPIP0::192.168.222.207::INSTR")
        
#     def start_measurements(self):
#         th = MultimeterThread(self.multimeter)
#         th.dcv_volatge.connect(self.on_dc_measurement_completed)
#         th.acv_volatge.connect(self.on_ac_measurement_completed)
#         th.start()
#         self.multimeter_thread = th

#     def on_dc_measurement_completed(self, result):
#         self.text_browser_dc.append(result)

#     def on_ac_measurement_completed(self, result):
#         self.text_browser_ac.append(result)

#     def closeEvent(self, event):
#         # Stop the thread when the application is closed.
#         if hasattr(self, 'multimeter_thread'):
#             self.multimeter_thread.stop()
#             self.multimeter_thread.wait()


# def main():
#     app = QApplication(sys.argv)
#     window = MultimeterApp()
#     window.show()
#     sys.exit(app.exec_())

# if __name__ == '__main__':
#     main()

# ##########################################################################################################################################
# import time
# rm = visa.ResourceManager()

# multimeter = rm.open_resource('TCPIP0::192.168.222.207::INSTR')

# # print(multimeter.query('*IDN?'))
# # multimeter.read_termination = '\n'
# # multimeter.write_termination = '\n'


# multimeter.write('ROUTe:SCAN 1')
# print(multimeter.query('ROUTe:SCAN?'))
# multimeter.write('ROUTe:START ON')
# print(multimeter.write('ROUTe:FUNC SCAN'))
# multimeter.write('ROUTe:LIMI:HIGH 3')
# multimeter.write('ROUTe:LIMI:LOW 2')

# print(multimeter.query('ROUTe:LIMI:LOW?'))

# time.sleep(2)
# print(multimeter.query('ROUTe:FUNC?'))

# multimeter.write('ROUT:CHAN 2,ON,ACV,AUTO,FAST')

# multimeter.write('ROUT:RELA ACV,ON')

# print(multimeter.query('ROUTe:CHANnel? 1'))

# multimeter.write('ROUT:CHAN 3,ON,DCV,AUTO,FAST')
# multimeter.write('ROUT:RELA DCV,ON')
# print(multimeter.query('ROUTe:CHANnel? 2'))

# # print(multimeter.write('ROUTe:SCAN ON'))


# # print(multimeter.query('ROUTe:START?'))

# print(multimeter.query('ROUTe:DATA? 2'))
# print(multimeter.query('ROUTe:DATA? 3'))


# multimeter.write('ROUTe:SCAN 0')
# multimeter.close()

import pyvisa as visa
import sys
import time
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QThread, pyqtSignal, QTimer, QtMsgType
from PyQt5.QtGui import *

class MultimeterThread(QThread):
    measurement_completed = pyqtSignal(str)
    dcv_volatge = pyqtSignal(str)
    acv_volatge = pyqtSignal(str)

    def __init__(self, resource_name):
        super().__init__()
        self.multimeter = resource_name
        self.running = True

    def run(self):
        try:
            self.multimeter.write('ROUTe:SCAN 1')
            self.multimeter.write('ROUTe:START ON')
            self.multimeter.write('ROUTe:FUNC SCAN')
            self.multimeter.write('ROUTe:LIMI:HIGH 3')
            self.multimeter.write('ROUTe:LIMI:LOW 2')
            time.sleep(2)
            self.multimeter.write('ROUT:CHAN 2,ON,ACV,AUTO,FAST')
            self.multimeter.write('ROUT:CHAN 3,ON,DCV,AUTO,FAST')
            while self.running:
                ACV_Volt = self.multimeter.query('ROUTe:DATA? 2')
                DCV_Volt = self.multimeter.query('ROUTe:DATA? 3')
                self.dcv_volatge.emit(DCV_Volt)
                self.acv_volatge.emit(ACV_Volt)
                time.sleep(1)  # Adjust the sleep duration to control update frequency

            scan_OFF = self.multimeter.query('ROUTe:STATe?')
            self.multimeter.write('ROUTe:SCAN 0')
            self.measurement_completed.emit(scan_OFF)
        except AttributeError:
            print('error occurred')

    def stop(self):
        self.running = False

class MultimeterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Multimeter Data Logger")
        self.setGeometry(100, 100, 800, 600)

        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.btn_connect = QPushButton("Connect and Read")
        self.btn_connect.clicked.connect(self.start_measurements)

        self.text_browser_ac = QTextBrowser()
        self.text_browser_dc = QTextBrowser()
        self.image_label = QLabel()
        self.image_index = 0

        # Add a new button to display images
        self.btn_display_images = QPushButton("Display Images")
        self.btn_display_images.clicked.connect(self.start_image_display)

        self.dc_voltlabel = QLabel()
        self.ac_voltlabel = QLabel()

        layout.addWidget(self.btn_connect)
        layout.addWidget(self.text_browser_ac)
        layout.addWidget(self.text_browser_dc)
        layout.addWidget(self.image_label)
        layout.addWidget(self.btn_display_images)
        layout.addWidget(self.dc_voltlabel)
        layout.addWidget(self.ac_voltlabel)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.rm = visa.ResourceManager()
        self.multimeter = self.rm.open_resource("TCPIP0::192.168.222.207::INSTR")
        self.multimeter.read_termination = '\n'
        self.multimeter.write_termination = '\n'
        self.test_images = [
            'images_/images/R700.jpg',
            'images_/images/R709_before_jumper.jpg',
            'images_/images/R700_DC.jpg',
            'images_/images/PP2.png',
            'images_/images/C443.jpg',
            'images_/images/C442.jpg',
            'images_/images/C441.jpg',
            'images_/images/C412.jpg',
            'images_/images/C430.jpg',
            'images_/images/C443_1.jpg',
            'images_/images/C442_1.jpg',
            'images_/images/C441_1.jpg',
            'images_/images/C412_1.jpg',
            'images_/images/C430_1.jpg',
            'images_/images/PP.jpg'
        ]

        self.image_timer = QTimer(self)
        self.image_timer.timeout.connect(self.display_next_image)
        # self.image_timer.start(10000)  # Display each image for 10 seconds

    def start_measurements(self):
        th = MultimeterThread(self.multimeter)
        th.dcv_volatge.connect(self.on_dc_measurement_completed)
        th.acv_volatge.connect(self.on_ac_measurement_completed)
        th.start()
        self.multimeter_thread = th

    def on_dc_measurement_completed(self, result):
        self.text_browser_dc.append(result)
        self.dc_voltlabel.setText(result)

    def on_ac_measurement_completed(self, result):
        self.text_browser_ac.append(result)
        res = result.split(' ')
        print(float(res[0]))
        # if 0.13 < float(res[0]) < 0.15:
        #     self.ac_voltlabel.setStyleSheet("background-color: red;")
        self.ac_voltlabel.setText(result)
        # else:
        #     self.ac_voltlabel.setStyleSheet("background-color: green;")
        # self.ac_voltlabel.setText(result)

    def display_next_image(self):
        if self.image_index < len(self.test_images):
            image_path = self.test_images[self.image_index]
            pixmap = QPixmap(image_path)
            self.image_label.setPixmap(pixmap)
            if self.image_index == 0:
                dcv_val = self.dc_voltlabel.text()
                dcv_value = dcv_val.split(' ')
                print('DCV 0', float(dcv_value[0]))
                dcv_value_ = float(dcv_value[0])
                if not (0.03 < dcv_value_ < 0.04):
                    QMessageBox.information(self, "DC Voltage Out of Range", "step 1 ---- index 0.")
                    return
            elif self.image_index == 1:
                dcv_val = self.dc_voltlabel.text()
                dcv_value = dcv_val.split(' ')
                print('DCV 1',float(dcv_value[0]))
                dcv_value_ = float(dcv_value[0])
                if (1.13 <= dcv_value_ <= 1.16):
                    QMessageBox.information(self, "DC Voltage Out of Range", "step 2 ---- index 1.")
                    return
            elif self.image_index == 2:
                acv_val = self.ac_voltlabel.text()
                QMessageBox.information(self, "DC Voltage Out of Range", "step 3 ---- index 2.")

            elif self.image_index == 3:
                acv_val = self.ac_voltlabel.text()
                acv_value = acv_val.split(' ')
                print('ACV 1', float(acv_value[0]))
                acv_value_ = float(acv_value[0])
                if not (0.013 <= acv_value_ <= 0.016):
                    QMessageBox.information(self, "DC Voltage Out of Range", "step 4 ---- index 3.")
                    return



            self.image_index += 1
        else:
            self.image_timer.stop()

    def start_image_display(self):
        self.image_index = 0
        self.image_timer.start(10000)  # Display each image for 10 seconds

    def closeEvent(self, event):
        if hasattr(self, 'multimeter_thread'):
            self.multimeter_thread.stop()
            self.multimeter_thread.wait()

def main():
    app = QApplication(sys.argv)
    window = MultimeterApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
