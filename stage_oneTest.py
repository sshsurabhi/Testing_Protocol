from PyQt5 import uic
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import configparser
import pyvisa as visa


# import sys
# from PyQt5.QtWidgets import QApplication, QMainWindow, QLineEdit, QVBoxLayout, QWidget
# from PyQt5.QtCore import Qt

# class MyMainWindow(QMainWindow):
#     def __init__(self):
#         super().__init__()
#         self.setWindowTitle("Value Range Example")
#         self.central_widget = QWidget()
#         self.layout = QVBoxLayout()

#         # Create the QLineEdit for the value input.
#         self.value_input = QLineEdit(self)
#         self.value_input.setPlaceholderText("Enter a value between 4.2 and 4.6")

#         # Connect the textChanged signal to update the background color.
#         self.value_input.textChanged.connect(self.update_background_color)

#         self.layout.addWidget(self.value_input)
#         self.central_widget.setLayout(self.layout)
#         self.setCentralWidget(self.central_widget)

#         # Set the initial background color based on the placeholder text.
#         self.update_background_color()

#     def update_background_color(self):
#         # Get the current text from the QLineEdit.
#         text = self.value_input.text()

#         # Check if the text is a valid float within the desired range.
#         try:
#             value = float(text)
#             if 4.2 >= value <= 4.6:
#                 # Set the background color to green if the value is within the range.
#                 self.value_input.setStyleSheet("QLineEdit { background-color: #c1ffc1; }")
#             else:
#                 # Set the background color to red if the value is outside the range.
#                 self.value_input.setStyleSheet("QLineEdit { background-color: #ff6f6f; }")
#         except ValueError:
#             # Set the background color to white if the input is not a valid float.
#             self.value_input.setStyleSheet("QLineEdit { background-color: #ffffff; }")

# if __name__ == "__main__":
#     app = QApplication([])
#     window = MyMainWindow()
#     window.show()
#     sys.exit(app.exec_())



# import visa
# Connect to the power supply
rm = visa.ResourceManager()
ps = rm.open_resource('TCPIP0::192.168.222.207::INSTR')
# ps.read_termination = '\n'
# ps.write_termination = '\n'
# Get the instrument ID

print(ps.query('*IDN?'))
# Set the output channel to CH1
# ps.write('INSTrument CH1')
# # Set the output voltage to 30V
# ps.write('CH1:VOLTage 30')
# Set the output current to 0.5A
# ps.write('CH1:CURRent 0.5')
# # Turn on the output
# ps.write('OUTPut CH1,ON')
# Turn off the output
print(ps.query('ROUTe:STATe?'))
print(ps.write('ROUTe:SCAN 1'))
print(ps.query('ROUTe:SCAN?'))
print(ps.query('ROUTe:CHANnel？'))
# print(ps.write('ROUT:CHAN? 1'))


print(ps.query('MEAS:VOLT:DC?'))
# print(ps.query('MEASure:CURRent? CH1'))
# ps.write('OUTPut CH1,OFF')
# Close the connection to the power supply
ps.close()

try:
    #response = ps.query('OUTPut CH1,ON')
    # print(ps)
    response = ps.query('OUTPut CH1,OFF')
    print(ps)
    # Process the response as needed
except visa.errors.VisaIOError as e:
    print(f"An error occurred during communication: {e}")


class MyDialog(QMainWindow):
    button_clicked = pyqtSignal(str)
    #def __init__(self, rm, multimeter, parent=None):
    def __init__(self, parent=None):
        super(MyDialog, self).__init__(parent)
        uic.loadUi("UI/ascha.ui", self)
        self.setWindowIcon(QIcon('images_/1.jpg'))
        self.setFixedSize(self.size())
        self.show()

        # self.rm = rm
        # self.multimeter = multimeter
        self.multimeter = None
        self.powersupply = None
        self.rm = visa.ResourceManager()
        self.AC_DC_box.addItems(['<select>', 'DCV', 'ACV'])
        self.AC_DC_box.currentTextChanged.connect(self.handleSelectionChange)
        
        self.test_button.clicked.connect(self.on_cal_voltage_current)



        self.MM_button.clicked.connect(self.connect_multimeter)
        self.PS_button.clicked.connect(self.connect_powersupply)

        self.test_button.setEnabled(False)
        self.AC_DC_box.setEnabled(False)
        self.PS_button.setEnabled(False)

        self.selected_command = None        
        self.DC_values = []
        self.AC_values = []

        self.info_label.setText('Press Multi ON button.')
        self.button_labels = []

    def on_button_clicked(self):
        self.button_clicked.emit("Button clicked")

    def image_change(self, file_path):
        if file_path:
            pixmap = QPixmap(file_path)
            self.image_label.setPixmap(pixmap)
            self.image_label.setScaledContents(True)
            self.image_label.setFixedSize(pixmap.width(), pixmap.height())

            # if self.start_button.text() == 'Step4':
            #     reply = self.show_good_message()
            #     if reply == QMessageBox.Yes:
            #         # Continue to the next step
            #         self.start_button.setText('Step5')
            #     else:
            #         # Go back to Step4
            #         self.on_button_click('images_/Start4.png')

    def handleSelectionChange(self, text):
        if text == 'DCV':
            self.selected_command = 'MEAS:VOLT:DC?'
            self.test_button.setEnabled(True)
            self.image_change('images_/3.jpg')
            self.test_button.setText('R709')
            self.info_label.setText('Press R709')
            self.AC_DC_box.setEnabled(False)  # Disable AC_DC_box after selecting 'DCV'
        elif text == 'ACV':
            self.selected_command = 'MEAS:VOLT:AC?'
            self.test_button.setEnabled(True)
            self.test_button.setText('GO')
            self.AC_DC_box.setEnabled(True)  # Enable AC_DC_box after selecting 'ACV'
            self.info_label.setText('Ready to proceed with AC voltage measurement.')
        else:
            self.selected_command = None
            self.test_button.setEnabled(False)
            self.info_label.setText('Select DC in the combobox to proceed further.')



    def on_cal_voltage_current(self):
        if self.AC_DC_box.currentText() == 'DCV' and self.test_button.text() == 'GO':
            self.multimeter.query('*IDN?')
            self.info_label.setText('Press R709')
            
        elif self.AC_DC_box.currentText() == 'DCV' and self.test_button.text() == 'R709':
            self.result_edit.setText(self.multimeter.query('MEASure:VOLTage:DC?'))
            self.textBrowser.setText(str(float(self.powersupply.query('MEASure:VOLTage? CH1'))))

        elif self.AC_DC_box.currentText() == 'ACV' and self.test_button.text() == 'R709':
            self.result_edit.setText(self.multimeter.query('MEAS:VOLT:AC?'))


    # def calcVoltage(self):
    #     if self.selected_command == 'MEAS:VOLT:DC?' and self.test_button.text() == 'R709':
    #         voltage_str = self.getSelectedVoltage()
    #         try:
    #             voltage = float(voltage_str)
    #             if 3.28 <= voltage <= 3.38:
    #                 self.result_edit.setText(f"{voltage} V")
    #             else:
    #                 self.result_edit.clear()
    #         except ValueError:
    #             self.result_edit.clear()

    #         print(voltage_str)
    #         self.info_label.setText(f"Voltage at R709 \n: {voltage_str} V")
    #         self.test_button.setText('R700')
    #     elif self.selected_command == 'MEAS:VOLT:DC?' and self.test_button.text() == 'R700':

    #         # Perform any additional operations or measurement for R700 (if needed)
    #         voltage = self.getSelectedVoltage()

    #         self.info_label.setText(f"Voltage at R700: {voltage_str} V")
    #     elif self.selected_command == 'MEAS:VOLT:AC?' and self.test_button.text() == 'GO':
    #         voltage = self.getSelectedVoltage()
    #         self.info_label.setText(f"AC Voltage: {voltage_str} V")
            
            
    # def getSelectedVoltage(self):
    #     if self.selected_command == 'MEAS:VOLT:DC?':
    #         # voltage = self.multimeter.query('MEAS:VOLT:DC?')
    #         # self.multimeter.read_termination = '\n'
    #         # self.multimeter.write_termination = '\n'
    #         self.multimeter.read()
    #         voltage = self.multimeter.query('MEAS:VOLT:DC?')
    #         self.button_clicked.emit(str(voltage))
    #         self.info_label.setText(str(voltage))
    #         return voltage
    #     elif self.selected_command == 'MEAS:VOLT:AC?':
    #         voltage = self.multimeter.query('MEAS:VOLT:AC?')
    #         self.button_clicked.emit(str(voltage))
    #         return voltage
    #     else:
    #         return 0.0


    def get_lineinser1_value(self):
        pass

    def connect_multimeter(self):
        
        if not self.multimeter:
            try:
                self.multimeter = self.rm.open_resource('TCPIP0::192.168.222.207::5024::SOCKET')
                self.multimeter.read_termination = '\n'
                self.multimeter.write_termination = '\n'
                print(self.multimeter)
                self.textBrowser.setText(self.multimeter.read())
                self.MM_button.setText('MM OFF')
                self.MM_button.setEnabled(False)
                self.PS_button.setEnabled(True)
                self.info_label.setText('Press NETZTEIL ON button')
            except visa.errors.VisaIOError:
                QMessageBox.information(self, "Multimeter Connection", "Multimeter is not present at the given IP address.")
                self.textBrowser.setText('Multimeter has not been presented')
        else:
            
            self.multimeter.close()
            self.multimeter = None
            self.MM_button.setText('MM ON')
            self.textBrowser.setText('Multimeter Disconnected')

    def connect_powersupply(self):
        if not self.powersupply:
            try:
                self.powersupply = self.rm.open_resource('TCPIP0::192.168.222.141::INSTR')
                self.powersupply.read_termination = '\n'
                self.powersupply.write_termination = '\n'
                self.textBrowser.setText(self.powersupply.query('SYSTem:VERSion?'))
                self.info_label.setText('Select DCV from dropdownbox')
                self.PS_button.setText('PS OFF')
            except visa.errors.VisaIOError:
                QMessageBox.information(self, "PowerSupply Connection", "PowerSupply is not present at the given IP Address.")
                self.textBrowser.setText('Powersupply has not been presented.')
        else:
            self.powersupply.close()
            self.powersupply = None
            self.PS_button.setText('PS ON')
            self.textBrowser.setText('Netzteil Disconnected')
        self.AC_DC_box.setEnabled(True)
        self.PS_button.setEnabled(False)
        self.MM_button.setEnabled(False)
############################################################
        
# def main():
#     app = QApplication([])
#     Window = MyDialog()
#     app.exec_()
# if __name__=='__main__':
#     main()

    
