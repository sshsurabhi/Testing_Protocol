# import openpyxl

# def read_excel(file_path, sheet_name):
#     try:
#         workbook = openpyxl.load_workbook(file_path)
#         sheet = workbook[sheet_name]
        
#         for row in sheet.iter_rows(values_only=True):
#             for cell in row:
#                 print(cell, end='\t')
#             print()
            
#     except Exception as e:
#         print(f"An error occurred: {e}")

# if __name__ == "__main__":
#     excel_file = "configs/Test_result.xlsx"
#     sheet_name = "Sheet1"  # Change this to the name of your sheet
    
#     read_excel(excel_file, sheet_name)

# import openpyxl

# def modify_excel(file_path, sheet_name, modifications):
#     try:
#         workbook = openpyxl.load_workbook(file_path)
#         sheet = workbook[sheet_name]

#         for row_index, (cell_value, new_value) in enumerate(modifications, start=2):
#             sheet.cell(row=row_index, column=2, value=new_value)
        
#         workbook.save(file_path)
#         print("Modifications saved successfully.")

#     except Exception as e:
#         print(f"An error occurred: {e}")

# if __name__ == "__main__":
#     excel_file = "configs/Test_result.xlsx"
#     sheet_name = "Sheet1"  # Change this to the name of your sheet
    
#     modifications = [
#         ("Voltage Set", 1),
#         ("Current Set", 2),
#         ("Read Current before Soldering", 3),
#         ("Voltage read before Jumper Close", 4),
#         ("Current read after Jumper close", 5),
#         ("DCV b/w GND - R709", 6),
#         ("DCV b/w GND - R700", 7),
#         ("ACV b/w GND - R709", 8),
#         ("ACV b/w GND - R700", 9),
#         ("DCV b/w ISOGND - C443", 10),
#         ("DCV b/w ISOGND - C442", 11),
#         ("DCV b/w ISOGND - C441", 12),
#         ("DCV b/w ISOGND - C412", 13),
#         ("DCV 2.048V b/w ISOGND - C430", 14),
#         ("ACV b/w ISOGND - C443", 15),
#         ("ACV b/w ISOGND - C442", 16),
#         ("ACV b/w ISOGND - C441", 17),
#         ("ACV b/w ISOGND - C412", 18),
#         ("ACV 2.048 b/w ISOGND - C430", 19),
#         ("Current Read after SDIO board Fix", 20),
#         ("Current Read after FPGA & SD-Card Fix", 21),
#         ("UID", 22),
#         ("IC704 result registers Reading", 23),
#     ]
    
#     modify_excel(excel_file, sheet_name, modifications)

# import sys
# from PyQt5.QtWidgets import QApplication, QWidget, QMessageBox, QVBoxLayout, QPushButton
# from PyQt5.QtCore import QTimer

# class MyApp(QWidget):
#     def __init__(self):
#         super().__init__()
#         self.initUI()

#     def initUI(self):
#         layout = QVBoxLayout()

#         self.show_message_button = QPushButton('Show Message')
#         self.show_message_button.clicked.connect(self.show_message)

#         layout.addWidget(self.show_message_button)

#         self.setLayout(layout)
#         self.setWindowTitle('PyQt5 Message Dialog Example')

#     def show_message(self):
#         msg_box = QMessageBox(self)
#         msg_box.setWindowTitle('Information')
#         msg_box.setIcon(QMessageBox.Information)
#         msg_box.setText('This is a message that will stay visible for a minimum of 5 seconds.')
#         msg_box.setStandardButtons(QMessageBox.Ok)
#         msg_box.setDefaultButton(QMessageBox.Ok)

#         ok_button = msg_box.button(QMessageBox.Ok)
#         ok_button.setEnabled(False)  # Disable the OK button initially

#         timer = QTimer(self)
#         timer.timeout.connect(lambda: ok_button.setEnabled(True))
#         timer.start(5000)  # 5000 ms (5 seconds)

#         msg_box.exec_()

# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     window = MyApp()
#     window.show()
#     sys.exit(app.exec_())

# import time

# def measure_dc_voltage(channel):
#     # Simulate measuring DC voltage on the specified channel
#     voltage = 5.0  # Replace with actual measurement code
#     return voltage

# def measure_ac_voltage(channel):
#     # Simulate measuring AC voltage on the specified channel
#     voltage = 220.0  # Replace with actual measurement code
#     return voltage

# def main():
#     num_dc_measurements = 4
#     num_ac_measurements = 4

#     # First set of DC voltage measurements
#     for i in range(num_dc_measurements):
#         dc_voltage = measure_dc_voltage(i + 1)
#         print(f"DC Voltage Measurement {i+1}: {dc_voltage} V")
#         time.sleep(10)

#     # First set of AC voltage measurements
#     for i in range(num_ac_measurements):
#         ac_voltage = measure_ac_voltage(i + 1)
#         print(f"AC Voltage Measurement {i+1}: {ac_voltage} V")
#         time.sleep(10)

#     print("Arranging wire settings...")
#     time.sleep(15)

#     # Second set of DC voltage measurements
#     for i in range(3):
#         dc_voltage = measure_dc_voltage(i + 1)
#         print(f"DC Voltage Measurement {i+5}: {dc_voltage} V")
#         time.sleep(10)

#     # Second set of AC voltage measurements
#     for i in range(3):
#         ac_voltage = measure_ac_voltage(i + 1)
#         print(f"AC Voltage Measurement {i+5}: {ac_voltage} V")
#         time.sleep(10)

# if __name__ == "__main__":
#     main()


# import time

# # Simulate the function to measure DC voltage
# def measure_dc_voltage(component_number):
#     # Replace this with your actual code to measure DC voltage using the multimeter
#     print(f"Measuring DC Voltage for component {component_number}")

# # Simulate the function to measure AC voltage
# def measure_ac_voltage(component_number):
#     # Replace this with your actual code to measure AC voltage using the multimeter
#     print(f"Measuring AC Voltage for component {component_number}")

# # Main function to perform measurements
# def main():
#     for component in range(1, 15):
#         time.sleep(10)
#         if component <= 4:
#             measure_dc_voltage(component)
#         elif 4 < component <= 8:
#             measure_ac_voltage(component)
        
#         # time.sleep(10)  # 10 seconds time difference between measurements

#         if component == 4:
#             print("Arrange small wire settings...")
#             time.sleep(15)  # 15 seconds for arranging wire settings
        
#         if 8 < component <= 11:
#             measure_dc_voltage(component)
#         elif 11 < component <= 14:
#             measure_ac_voltage(component)

#         time.sleep(10)  # 10 seconds time difference between measurements

# # Call the main function when the button is pressed in your PyQt5 app
# if __name__ == "__main__":
#     main()
# from serial import Serial
# from datetime import time
# import datetime
# import time
# port = 'COM4'
# baud = 115200
# commands = ['i2c:scan', 'i2c:read:53:04:FC', 'i2c:write:53:', 'i2c:read:53:20:00', 'i2c:write:73:04', 'i2c:scan','i2c:write:21:0300','i2c:write:21:0100','i2c:write:21:01FF', 'i2c:write:73:01',
#             'i2c:scan', 'i2c:write:4F:06990918', 'i2c:write:4F:01F8', 'i2c:read:4F:1E:00']
# serial_port = Serial(port, baud, timeout=1) 
# if serial_port.is_open:
#     serial_port.timeout = 5
#     next_command_index = 0
#     for command in commands:
#         if next_command_index < (len(commands)+1):
#             serial_port.write(command.encode())
#             # self.textBrowser.append('Processed Command: '+command + ' ')
#             response = serial_port.readline().decode('ascii')
#             print('response', response)
#             next_command_index += 1
#             current_date = datetime.datetime.now()
#             decimal_date = int(current_date.strftime('%Y%m%d'))
#             hex_date = hex(decimal_date)[2:].upper().zfill(8)
#             if command == commands[1]:
#                 ading = response.split(':')[3][:-1]
#                 commands[2] = commands[2]+ading+'01'+hex_date+'2A0101030000FFFF'
#                 print('ading', ading)
#             if command == commands[2]:
#                 start_time = time.time()
#                 while time.time() - start_time < 5:
#                     if response:
#                         break
#                     response = serial_port.readline().decode('ascii')
#                     print('response', response)
#     serial_port.close()

import sys
import time
import random
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QVBoxLayout, QWidget, QMessageBox
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import *
import pyvisa  as visa

class VoltageMeasurementApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Voltage Measurement App")
        self.setGeometry(100, 100, 400, 300)

        self.start_button = QPushButton("Start Measurements", self)
        self.start_button.clicked.connect(self.start_measurements)
        self.rm = visa.ResourceManager()
        self.multimeter = self.rm.open_resource('TCPIP0::192.168.222.207::INSTR')
        self.power_supply = self.rm.open_resource('TCPIP0::192.168.222.141::INSTR')
        self.image_label = QLabel(self)
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_list = ['images_/images/R700.jpg','images_/images/R709_before_jumper.jpg','images_/images/R700_DC.jpg', 'images_/images/PP2.png','images_/images/C443.jpg','images_/images/C442.jpg','images_/images/C441.jpg','images_/images/C412.jpg',
                            'images_/images/C430.jpg','images_/images/C443_1.jpg','images_/images/C442_1.jpg','images_/images/C441_1.jpg','images_/images/C412_1.jpg','images_/images/C430_1.jpg', 'images_/images/PP.jpg',]

        self.vlayout = QVBoxLayout()
        self.vlayout.addWidget(self.start_button)
        self.vlayout.addWidget(self.image_label)

        self.central_widget = QWidget()
        self.central_widget.setLayout(self.vlayout)
        self.setCentralWidget(self.central_widget)

        self.desired_voltage_range = (4.95, 5.05)  # Replace with your desired voltage range

    def start_measurements(self):
        for i in range(10):
            # Display the current image
            image_path = self.image_list[i]
            pixmap = QPixmap(image_path)
            self.image_label.setPixmap(pixmap)
            self.image_label.repaint()

            # Perform voltage measurement
            voltage = self.measure_voltage()
            
            if not self.is_voltage_in_range(voltage):
                retry = self.ask_retry_measurement(i + 1)
                if retry:
                    i -= 1  # Retry the current measurement
                else:
                    break

            time.sleep(5)  # Wait for 5 seconds between measurements

        self.show_measurement_result()

    def measure_voltage(self):
        # Implement code to measure voltage using the multimeter
        # Example: voltage = self.multimeter.measure_voltage()
        # You will need to consult the Siglent multimeter documentation for this part.
        voltage = random.uniform(4.0, 6.0)  # Placeholder code for testing
        return voltage

    def is_voltage_in_range(self, voltage):
        return self.desired_voltage_range[0] <= voltage <= self.desired_voltage_range[1]

    def ask_retry_measurement(self, component_number):
        msg = QMessageBox.question(self, "Retry Measurement",
                                   f"Measurement for component {component_number} is not in the desired range. "
                                   "Do you want to retry?",
                                   QMessageBox.Yes | QMessageBox.No)

        return msg == QMessageBox.Yes

    def show_measurement_result(self):
        QMessageBox.information(self, "Measurement Complete", "Voltage measurements for all components are complete.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = VoltageMeasurementApp()
    window.show()
    sys.exit(app.exec_())
