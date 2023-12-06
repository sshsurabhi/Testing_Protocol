import sys, time
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import *
import pyvisa as visa

class MyGUI(QWidget):
    def __init__(self):
        super().__init__()

        # Widgets
        self.label_image = QLabel(self)
        self.label_image.setPixmap(QPixmap('images_/images/Welcome.jpg'))

        self.button_connect_multimeter = QPushButton('Connect Multimeter', self)
        self.button_connect_powersupply = QPushButton('Connect Power Supply', self)
        self.button_measure = QPushButton('Measure', self)

        self.info_label = QLabel('Result:', self)
        self.textBrowser = QTextBrowser(self)
        self.line_edit = QLineEdit(self)

        # Layout
        layout = QVBoxLayout(self)
        layout.addWidget(self.label_image)
        layout.addWidget(self.button_connect_multimeter)
        layout.addWidget(self.button_connect_powersupply)
        layout.addWidget(self.button_measure)
        layout.addWidget(self.info_label)
        layout.addWidget(self.textBrowser)
        layout.addWidget(self.line_edit)

        self.button_connect_multimeter.clicked.connect(self.connect_multimeter)
        self.button_connect_powersupply.clicked.connect(self.connect_powersupply)
        self.button_measure.clicked.connect(self.start_continuous_measurement)

        self.rm = visa.ResourceManager()
        self.multimeter = None
        self.powersupply = None

        self.measure_timer = QTimer(self)
        self.measure_count = 0
        self.attempt_count = 0
        self.dc_measure_count = 0
        self.ac_measure_count = 0

        self.DCV_Results = [0,0,0,0,0,0,0]
        self.ACV_Results = [0,0,0,0,0,0,0]

        self.test_images = ['images_/images_/R700.jpg','images_/images_/ISOGND.jpg', 'images_/images_/C443.jpg','images_/images_/C442.jpg','images_/images_/C441.jpg','images_/images_/C412.jpg','images_/images_/C430.jpg',
                            'images_/Images/ISOGND.jpg']


    def connect_multimeter(self):
        if not self.multimeter:
            try:
                self.multimeter = self.rm.open_resource('TCPIP0::192.168.222.207::INSTR')
                self.textBrowser.append(self.multimeter.query('*IDN?'))
                self.info_label.setText('Press POWER ON button.\n \n It connects the powersupply...!' )
            except visa.errors.VisaIOError:
                self.textBrowser.append('Multimeter has not been presented')
        else:
            self.multimeter.close()
            self.multimeter = None
            self.textBrowser.append(self.multimeter.query('*IDN?'))

    def connect_powersupply(self):
        self.show_message('Power Supply connected!')

    def start_continuous_measurement(self):
        self.textBrowser.append('Measure according to the image..')
        self.label_image.setPixmap(QPixmap('images_/images/R709_.jpg'))
        self.show_message('measurement start.', QMessageBox.Information)
        self.attempt_count = 0
        self.measure_count = 0
        self.dc_measure_count = 0
        self.ac_measure_count = 0
        self.is_ac_measurement = False
        time.sleep(5)
        self.measure()

    def measure(self):
        try:            
            if not self.is_ac_measurement:
                dc_voltage = float(self.multimeter.query('MEAS:VOLT:DC?'))
                self.textBrowser.append(f'DC Voltage (Measurement {self.measure_count + 1}, Attempt {self.attempt_count + 1}): {dc_voltage}')
                if self.dc_measure_count==0:# and self.attempt_count<=3:
                    self.line_edit.setText(str(dc_voltage))
                    if 3.25 <= dc_voltage <= 3.35:
                        self.line_edit.setStyleSheet('background-color: green;')
                        self.label_image.setPixmap(QPixmap('images_/images/R709_.jpg'))
                        self.DCV_Results[0] = dc_voltage
                        self.dc_measure_count += 1
                        self.measure_count += 1
                        self.is_ac_measurement = True
                        print(f'R709 DC : {dc_voltage}\n')
                    else:
                        self.line_edit.setStyleSheet('background-color: red;')
                        self.label_image.setPixmap(QPixmap('images_/images/R709_.jpg'))
                        self.DCV_Results[0] = dc_voltage
                        print(f'DC R709 -ve: {dc_voltage}\n')
                        self.attempt_count += 1

                elif self.dc_measure_count == 1:
                    # dc_voltage = float(self.multimeter.query('MEAS:VOLT:DC?'))
                    # self.textBrowser.append(f'DC Voltage (Measurement {self.measure_count + 1}, Attempt {self.attempt_count + 1}): {dc_voltage}')
                    self.line_edit.setText(str(dc_voltage))
                    if 4.95 <= dc_voltage <= 5.05:
                        self.line_edit.setStyleSheet('background-color: green;')
                        self.label_image.setPixmap(QPixmap('images_/images/R700.jpg'))
                        self.DCV_Results[1] = dc_voltage
                        self.dc_measure_count += 1
                        self.measure_count += 1
                        self.is_ac_measurement = True
                        print(f'R700 DC  : {dc_voltage}\n')
                    else:
                        self.line_edit.setStyleSheet('background-color: red;')
                        self.label_image.setPixmap(QPixmap('images_/images/R700.jpg'))
                        self.DCV_Results[1] = dc_voltage
                        print(f'DC R700 -ve: {dc_voltage}\n')
                        self.attempt_count += 1

                # elif self.dc_measure_count == 2:
                #     # self.show_message('Change the Groung to ISOGND.', QMessageBox.Information)
                #     # self.label_image.setPixmap(QPixmap('images_/images/C443.jpg'))
                #     self.dc_measure_count += 1
                #     self.measure_count += 1
                #     self.is_ac_measurement = True


                elif self.dc_measure_count == 2:
                    # dc_voltage = float(self.multimeter.query('MEAS:VOLT:DC?'))
                    # self.textBrowser.append(f'DC Voltage (Measurement {self.measure_count - 1}, Attempt {self.attempt_count + 1}): {dc_voltage}')
                    self.line_edit.setText(str(dc_voltage))
                    if 11.95 <= dc_voltage <= 12.05:
                        self.line_edit.setStyleSheet('background-color: green;')
                        self.label_image.setPixmap(QPixmap('images_/images/C443.jpg'))
                        self.DCV_Results[2] = dc_voltage
                        self.dc_measure_count += 1
                        self.measure_count += 1
                        self.is_ac_measurement = True
                        print(f'C443 : {dc_voltage}\n')
                    else:
                        self.line_edit.setStyleSheet('background-color: red;')
                        self.label_image.setPixmap(QPixmap('images_/images/C443.jpg'))
                        self.DCV_Results[2] = dc_voltage
                        print(f'C443 -ve: {dc_voltage}\n')
                        self.attempt_count += 1

                elif self.dc_measure_count == 3:
                    # dc_voltage = float(self.multimeter.query('MEAS:VOLT:DC?'))
                    # self.textBrowser.append(f'DC Voltage (Measurement {self.measure_count - 1}, Attempt {self.attempt_count + 1}): {dc_voltage}')
                    self.line_edit.setText(str(dc_voltage))
                    if 4.95 <= dc_voltage <= 5.05:
                        self.line_edit.setStyleSheet('background-color: green;')
                        self.label_image.setPixmap(QPixmap('images_/images/C442.jpg'))
                        self.DCV_Results[3] = dc_voltage
                        self.dc_measure_count += 1
                        self.measure_count += 1
                        self.is_ac_measurement = True
                        print(f'C442: {dc_voltage}\n')
                    else:
                        self.line_edit.setStyleSheet('background-color: red;')
                        self.label_image.setPixmap(QPixmap('images_/images/C442.jpg'))
                        self.DCV_Results[3] = dc_voltage
                        print(f'C442 -ve: {dc_voltage}\n')
                        self.attempt_count += 1

                elif self.dc_measure_count == 4:
                    # dc_voltage = float(self.multimeter.query('MEAS:VOLT:DC?'))
                    # self.textBrowser.append(f'DC Voltage (Measurement {self.measure_count - 1}, Attempt {self.attempt_count + 1}): {dc_voltage}')
                    self.line_edit.setText(str(dc_voltage))
                    if 4.95 <= dc_voltage <= 5.05:
                        self.line_edit.setStyleSheet('background-color: green;')
                        self.label_image.setPixmap(QPixmap('images_/images/C441.jpg'))
                        self.DCV_Results[4] = dc_voltage
                        self.dc_measure_count += 1
                        self.measure_count += 1
                        self.is_ac_measurement = True
                        print(f'C441: {dc_voltage}\n')
                    else:
                        self.line_edit.setStyleSheet('background-color: red;')
                        self.label_image.setPixmap(QPixmap('images_/images/C441.jpg'))
                        self.DCV_Results[4] = dc_voltage
                        print(f'C441 -ve: {dc_voltage}\n')
                        self.attempt_count += 1

                elif self.dc_measure_count == 5:
                    # dc_voltage = float(self.multimeter.query('MEAS:VOLT:DC?'))
                    # self.textBrowser.append(f'DC Voltage (Measurement {self.measure_count - 1}, Attempt {self.attempt_count + 1}): {dc_voltage}')
                    self.line_edit.setText(str(dc_voltage))
                    if 4.98 <= dc_voltage <= 5.02:
                        self.line_edit.setStyleSheet('background-color: green;')
                        self.label_image.setPixmap(QPixmap('images_/images/C412.jpg'))
                        self.DCV_Results[5] = dc_voltage
                        self.dc_measure_count += 1
                        self.measure_count += 1
                        self.is_ac_measurement = True
                        print(f'C412: {dc_voltage}\n')
                    else:
                        self.line_edit.setStyleSheet('background-color: red;')
                        self.label_image.setPixmap(QPixmap('images_/images/C412.jpg'))
                        self.DCV_Results[5] = dc_voltage
                        self.attempt_count += 1
                        print(f'C412 -ve: {dc_voltage}\n')

                elif self.dc_measure_count == 6:
                    # dc_voltage = float(self.multimeter.query('MEAS:VOLT:DC?'))
                    # self.textBrowser.append(f'DC Voltage (Measurement {self.measure_count - 1}, Attempt {self.attempt_count + 1}): {dc_voltage}')
                    self.line_edit.setText(str(dc_voltage))
                    if 2.046 <= dc_voltage <= 2.050:
                        self.line_edit.setStyleSheet('background-color: green;')
                        self.label_image.setPixmap(QPixmap('images_/images/C430.jpg'))
                        self.DCV_Results[6] = dc_voltage
                        self.dc_measure_count += 1
                        self.measure_count += 1
                        self.is_ac_measurement = True
                        print(f'C430: {dc_voltage}\n')
                    else:
                        self.line_edit.setStyleSheet('background-color: red;')
                        self.label_image.setPixmap(QPixmap('images_/images/C430.jpg'))
                        self.DCV_Results[6] = dc_voltage
                        print(f'C430 -ve: {dc_voltage}\n')
                        self.attempt_count += 1
                        
                else:
                    self.line_edit.setStyleSheet('background-color: red;')
                    self.label_image.setPixmap(QPixmap('images_/images/PP1.jpg'))
                    self.measure_count += 1

                if self.attempt_count == 3:
                    self.save_measurement(dc_voltage)
                    self.attempt_count = 0
                    self.dc_measure_count += 1
                    self.is_ac_measurement = True
                    self.measure_count += 1

            else:
                ac_voltage = float(self.multimeter.query('MEAS:VOLT:AC?'))
                self.textBrowser.append(f'AC Voltage (Measurement {self.measure_count}, Attempt {self.attempt_count + 1}): {ac_voltage}')
                self.line_edit.setText(str(ac_voltage))

                if self.ac_measure_count == 0:
                    if ac_voltage < 0.01:
                        self.line_edit.setStyleSheet('background-color: green;')
                        self.label_image.setPixmap(QPixmap('images_/images/R700.jpg'))
                        self.ACV_Results[0] = ac_voltage
                        self.ac_measure_count += 1
                        self.is_ac_measurement = False
                        print(f'R709 AC Volt: {ac_voltage}\n')
                        # self.measure_count += 1
                    else:
                        self.line_edit.setStyleSheet('background-color: red;')
                        self.label_image.setPixmap(QPixmap('images_/images/R709_.jpg'))
                        self.attempt_count += 1
                        self.ACV_Results[0] = ac_voltage
                        print(f'R709 AC Volt -ve: {ac_voltage}\n')

                elif self.ac_measure_count == 1:
                    if ac_voltage < 0.01:
                        self.line_edit.setStyleSheet('background-color: green;')
                        self.label_image.setPixmap(QPixmap('images_/images/ISOGND.jpg'))######
                        self.ACV_Results[1] = ac_voltage
                        self.ac_measure_count += 1
                        self.is_ac_measurement = False
                        print(f'R700 AC Volt: {ac_voltage}\n')
                        self.show_message('Change the Groung to ISOGND.', QMessageBox.Information)
                        self.label_image.setPixmap(QPixmap('images_/images/C443.jpg'))
                    else:
                        self.line_edit.setStyleSheet('background-color: red;')
                        self.label_image.setPixmap(QPixmap('images_/images/R700.jpg'))
                        self.attempt_count += 1
                        self.ACV_Results[1] = ac_voltage
                        print(f'R700 AC Volt -ve: {ac_voltage}\n')

                # elif self.ac_measure_count == 2:
                #     # self.show_message('Change the Groung to ISOGND.', QMessageBox.Information)
                #     # self.label_image.setPixmap(QPixmap('images_/images/C443.jpg'))
                #     self.ac_measure_count += 1
                #     # self.measure_count += 1
                #     self.is_ac_measurement = True
                #     print(f'Normal AC Volt: {ac_voltage}\n')

                elif self.ac_measure_count == 2:
                    if ac_voltage < 0.01:
                        self.line_edit.setStyleSheet('background-color: green;')
                        self.label_image.setPixmap(QPixmap('images_/images/C442.jpg'))
                        self.ACV_Results[2] = ac_voltage
                        self.ac_measure_count += 1
                        self.is_ac_measurement = False
                        print(f'C443 AC Volt: {ac_voltage}\n')
                    else:
                        self.line_edit.setStyleSheet('background-color: red;')
                        self.label_image.setPixmap(QPixmap('images_/images/C443.jpg'))
                        self.ACV_Results[2] = ac_voltage
                        print(f'C443 AC Volt -ve: {ac_voltage}\n')
                        self.attempt_count += 1

                elif self.ac_measure_count == 3:
                    if ac_voltage < 0.005:
                        self.line_edit.setStyleSheet('background-color: green;')
                        self.label_image.setPixmap(QPixmap('images_/images/C441.jpg'))
                        self.ACV_Results[3] = ac_voltage
                        self.ac_measure_count += 1
                        self.is_ac_measurement = False
                        print(f'C442 AC Volt: {ac_voltage}\n')
                    else:
                        self.line_edit.setStyleSheet('background-color: red;')
                        self.label_image.setPixmap(QPixmap('images_/images/C442.jpg'))
                        self.ACV_Results[3] = ac_voltage
                        print(f'C442 AC Volt -ve: {ac_voltage}\n')
                        self.attempt_count += 1

                elif self.ac_measure_count == 4:
                    if ac_voltage < 0.005:
                        self.line_edit.setStyleSheet('background-color: green;')
                        self.label_image.setPixmap(QPixmap('images_/images/C412.jpg'))
                        self.ACV_Results[4] = ac_voltage
                        self.ac_measure_count += 1
                        self.is_ac_measurement = False
                        print(f'C441 AC Volt: {ac_voltage}\n')
                    else:
                        self.line_edit.setStyleSheet('background-color: red;')
                        self.label_image.setPixmap(QPixmap('images_/images/C441.jpg'))
                        self.ACV_Results[4] = ac_voltage
                        print(f'C441 AC Volt -ve: {ac_voltage}\n')
                        self.attempt_count += 1

                elif self.ac_measure_count == 5:
                    if ac_voltage <= 0.001:
                        self.line_edit.setStyleSheet('background-color: green;')
                        self.label_image.setPixmap(QPixmap('images_/images/C430.jpg'))
                        self.ACV_Results[5] = ac_voltage
                        self.ac_measure_count += 1
                        self.is_ac_measurement = False
                        print(f'C412 AC Volt: {ac_voltage}\n')
                    else:
                        self.line_edit.setStyleSheet('background-color: red;')
                        self.label_image.setPixmap(QPixmap('images_/images/C412.jpg'))
                        self.ACV_Results[5] = ac_voltage
                        print(f'C412 AC Volt -ve: {ac_voltage}\n')
                        self.attempt_count += 1

                elif self.ac_measure_count == 6:
                    if ac_voltage <= 0.001:
                        self.line_edit.setStyleSheet('background-color: green;')
                        self.label_image.setPixmap(QPixmap('images_/images/FPFPGA.jpg'))
                        self.ACV_Results[6] = ac_voltage
                        self.ac_measure_count += 1
                        self.is_ac_measurement = False
                        print(f'C430 AC Volt: {ac_voltage}\n')
                    else:
                        self.line_edit.setStyleSheet('background-color: red;')
                        self.label_image.setPixmap(QPixmap('images_/images/C430.jpg'))
                        self.ACV_Results[6] = ac_voltage
                        print(f'C430 AC Volt -ve: {ac_voltage}\n')
                        self.attempt_count += 1
                else:
                    self.attempt_count += 1


                if self.attempt_count == 3:
                    if self.ac_measure_count  ==1:
                        self.attempt_count = 0
                        self.label_image.setPixmap(QPixmap(self.test_images[self.ac_measure_count]))
                        self.save_measurement(ac_voltage)
                        self.ac_measure_count += 1
                        self.is_ac_measurement = False
                        self.show_message('Change the Groung to ISOGND.', QMessageBox.Information)
                        self.label_image.setPixmap(QPixmap('images_/images/C443.jpg'))
                        time.sleep(2)
                    else:
                        self.attempt_count = 0
                        self.label_image.setPixmap(QPixmap(self.test_images[self.ac_measure_count]))
                        self.save_measurement(ac_voltage)
                        self.ac_measure_count += 1
                        self.is_ac_measurement = False
                        # self.measure_count += 1
                        time.sleep(2)


        except visa.errors.VisaIOError as e:
            error_message = f'Measurement error: {str(e)}'
            self.textBrowser.append(error_message)
            self.line_edit.setStyleSheet('background-color: red;')
            self.show_message(error_message, QMessageBox.Critical)

        if not (self.ac_measure_count  == 7) and self.measure_count < 8 :
            print('self.ac_measure_count',self.ac_measure_count)
            QTimer.singleShot(3000, self.measure)
        else:
            self.show_message('Continuous measurement completed.', QMessageBox.Information)
            print('resulted values are \n', self.ACV_Results ,'\n', self.DCV_Results)

    def save_measurement(self, value):
        print(f'Measurement saved: {value}')













































































































































    def connect_powersupply(self):
        if not self.powersupply:
            try:
                self.powersupply = self.rm.open_resource('TCPIP0::192.168.222.141::INSTR')
                self.textBrowser.append(self.powersupply.query('*IDN?'))
                self.line_edit.setStyleSheet("background-color: lightyellow;")
                self.info_label.setText('Write CH1 in the Yellow Box (Highlighted)\n \n next to CH \n\n Press "ENTER"')
            except visa.errors.VisaIOError:
                QMessageBox.information(self, "PowerSupply Connection", "PowerSupply is not present at the given IP Address.")
                self.textBrowser.setText('Powersupply has not been presented.')
        else:
            self.powersupply.close()
            self.powersupply = None
            self.textBrowser.setText('Netzteil Disconnected')

    def show_message(self, message, icon=QMessageBox.Information):
        msg_box = QMessageBox(self)
        msg_box.setIcon(icon)
        msg_box.setText(message)
        msg_box.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        return msg_box.exec_()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MyGUI()
    window.show()
    sys.exit(app.exec_())
