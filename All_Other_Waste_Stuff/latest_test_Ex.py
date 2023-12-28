# import pyvisa as visa
# import time

# # Initialize the VISA resource manager and open the multimeter connection
# rm = visa.ResourceManager()
# multimeter = rm.open_resource('TCPIP0::192.168.222.207::inst0::INSTR')  # Replace with your multimeter's VISA address
# def read_voltage_and_check_condition():
#     # Configure the multimeter settings as needed (e.g., voltage range, measurement type)
#     multimeter.write('CONF:VOLT:DC 5.5')  # Example configuration for DC voltage measurement

#     # Wait for a certain amount of time (e.g., 1 second) before taking the reading
#     time.sleep(1)

#     voltage_reading = float(multimeter.query('READ?'))

#     if 3.25 < voltage_reading < 5.05:
#         print('DC OK')
#     else:
#         print('DC Not OK')

#     time.sleep(1)
#     multimeter.write('CONF:VOLT:AC 5')
#     # Read the voltage value from the multimeter
#     voltage_reading = float(multimeter.query('READ?'))
#     print(voltage_reading)
#     # Check if the voltage meets your condition
#     if float(voltage_reading) < 0.01:
#         return 'AC OK'
#     else:
#         return 'AC Not OK'
# num_components = 14  # Number of components to test

# for component_number in range(1, num_components + 1):
#     print(f"Testing Component {component_number}: {read_voltage_and_check_condition()}")
# multimeter.close()

# print(min([1,5,9,3,76,34,54,41]))

# import time
# import pyvisa as visa

# def main():
#     rm = visa.ResourceManager()
    
#     # Replace 'GPIB0::1::INSTR' with your multimeter's VISA address
#     multimeter = rm.open_resource('TCPIP0::192.168.222.207::inst0::INSTR')
    
#     # Set up the multimeter for voltage measurements
#     multimeter.write('CONF:VOLT:DC 10')  # Set to DC voltage measurement with a 10V range
    
#     # Wait for the multimeter to settle (adjust as needed)
#     time.sleep(2)
    
    
#     while True:
#         voltage = float(multimeter.query('READ?'))
#         if 3.25 < voltage < 3.35:
#             print('OK')
#         else:
#             print('Not OK')


# if __name__ == "__main__":
#     main()


############################################################################################################
# from PyQt5 import uic
# from PyQt5.QtWidgets import *
# from PyQt5.QtCore import *
# from PyQt5.QtGui import *
# import pyvisa as visa
# import time


# class MyDialog(QMainWindow):
#     def __init__(self, parent=None):
#     #def __init__(self, parent=None):
#         super(MyDialog, self).__init__(parent)
#         uic.loadUi("UI/untitled.ui", self)
#         self.setWindowIcon(QIcon('2.jpg'))
#         self.setFixedSize(self.size())
#         self.show()

#         self.test_images = ['images_/images/R700.jpg','images_/images/R709_before_jumper.jpg','images_/images/R700_DC.jpg', 'images_/images/PP2.png','images_/images/C443.jpg','images_/images/C442.jpg','images_/images/C441.jpg','images_/images/C412.jpg',
#                     'images_/images/C430.jpg','images_/images/C443_1.jpg','images_/images/C442_1.jpg','images_/images/C441_1.jpg','images_/images/C412_1.jpg','images_/images/C430_1.jpg', 'images_/images/PP.jpg']
#         print('length', len(self.test_images))
#         self.pushButton.clicked.connect(self.change_image)
#         self.pushButton_2.clicked.connect(self.change_power)
#         self.test_index = 0


#     def on_button_click(self, file_path):
#         if file_path:
#             pixmap = QPixmap(file_path)
#             self.label.setPixmap(pixmap)
#             self.label.setScaledContents(True)
#             self.label.setFixedSize(pixmap.width(), pixmap.height())

#     def change_power(self):
#         self.rm = visa.ResourceManager()
#         self.multimeter = self.rm.open_resource('TCPIP0::192.168.222.207::INSTR')
#         print(self.multimeter.query('*IDN?'))
#         self.powersupply = self.rm.open_resource('TCPIP0::192.168.222.141::INSTR')
#         print(self.powersupply.query('*IDN?'))
#         self.powersupply.write('OUTPut CH1,ON')
#         self.on_button_click('images_/images/R709.jpg')

   
#     def change_image(self):
#         while self.test_index < len(self.test_images):
#             time.sleep(0.5)


#             if self.test_index == 0:
#                 voltage_readings = self.measure_voltage()  # Run measure_voltage function
#                 if 3.25 < max(voltage_readings) < 3.35:
#                     self.on_button_click(self.test_images[self.test_index])
#                     self.lineEdit.setText(str(max(voltage_readings)))
#                     QMessageBox.information(self, "Important", "Voltage at R709 is in correct range.")
#                 else:
#                     response = QMessageBox.question(self, "Retry", "Voltage reading is out of range. Retry?",
#                                                     QMessageBox.Yes | QMessageBox.No)
#                     if response == QMessageBox.Yes:
#                         self.on_button_click('images_/images/R709.jpg')
#                         self.test_index = 0
#                         self.change_image()  # Recursive call to retry
#                     elif response == QMessageBox.No:
#                         self.on_button_click(self.test_images[self.test_index])
#                         QMessageBox.information(self, "Important", "Dont Proceed.")
#                         print('OK R709')


#                 if self.test_index == 1:
#                     voltage_reading = self.measure_voltage()  # Run measure_voltage function
#                     if (4.95 < max(voltage_reading) < 5.05):                
#                         self.on_button_click(self.test_images[self.test_index])
#                         self.lineEdit.setText(str(max(voltage_reading)))
                        
#                         QMessageBox.information(self, "Important", "Voltage at R700 is in correct range.")
#                     else:
#                         response = QMessageBox.question(self, "Retry", "Voltage reading is out of range. Retry?", QMessageBox.Yes | QMessageBox.No)
#                         if response == QMessageBox.No:
#                             print('OK')
#         self.test_index += 1


#     def measure_voltage(self):
#         # time.sleep(5)
#         voltage_readings = []
#         for _ in range(25):
#             self.multimeter.write('CONF:VOLT:DC? 5')
#             voltage_reading = self.multimeter.query('READ?')
#             voltage_readings.append(float(voltage_reading))
#             # print(voltage_readings)
#         return voltage_readings
        


# def main():
#     app = QApplication([])
#     Window = MyDialog()
#     app.exec_()
# if __name__=='__main__':
#     main()

##########################################################################################################################

# import sys
# from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QLabel

# class NumberPrinterApp(QMainWindow):
#     def __init__(self):
#         super().__init__()

#         self.setWindowTitle("Number Printer")
#         self.setGeometry(100, 100, 400, 200)

#         self.central_widget = QWidget()
#         self.setCentralWidget(self.central_widget)

#         self.layout = QVBoxLayout(self.central_widget)

#         self.print_button = QPushButton("Print Numbers")
#         self.print_button.clicked.connect(self.print_numbers)
#         self.layout.addWidget(self.print_button)

#         self.status_label = QLabel("Status: Not Printing")
#         self.layout.addWidget(self.status_label)

#         self.number_start = 200
#         self.number_end = 300
#         self.printing = False

#     def print_numbers(self):
#         if not self.printing:
#             self.print_button.setText("Stop Printing")
#             self.status_label.setText("Status: Printing (Green)")
#             self.print_button.setStyleSheet("background-color: green;")
#             self.printing = True

#             # Simulate printing numbers from 200 to 300
#             for number in range(self.number_start, self.number_end + 1):
#                 print(number)

#             self.print_button.setText("Print Numbers")
#             self.status_label.setText("Status: Printing Complete (Red)")
#             self.print_button.setStyleSheet("background-color: red;")
#             self.printing = False

# def main():
#     app = QApplication(sys.argv)
#     window = NumberPrinterApp()
#     window.show()
#     sys.exit(app.exec_())

# if __name__ == "__main__":
#     main()


import sys
from PyQt5.QtWidgets import QApplication, QMessageBox

def show_yes_no_dialog():
    app = QApplication(sys.argv)
    
    # Create a QMessageBox
    msgBox = QMessageBox()
    
    # Set the text and type of the message box
    msgBox.setText("Do you want to proceed?")
    msgBox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    
    # Apply a stylesheet to customize the YES and NO buttons
    msgBox.setStyleSheet("""
        QPushButton#qt_msgbox_button:default { /* Styling for default button (NO in this case) */
            background-color: red;
            color: white;
        }
        
        QPushButton#qt_msgbox_button {
            background-color: green; /* Styling for YES button */
            color: white;
        }
    """)
    
    # Show the message box and get the user's response
    response = msgBox.exec()
    
    # Check the user's response
    if response == QMessageBox.Yes:
        print("User clicked YES")
    elif response == QMessageBox.No:
        print("User clicked NO")
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    show_yes_no_dialog()



