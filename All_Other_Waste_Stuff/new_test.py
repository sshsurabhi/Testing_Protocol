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
    def on_button_clicked(self):
        self.result_signal.emit("Button clicked")
##################################################################################################################################
class SerialPortThread(QThread):
    com_ports_available = pyqtSignal(list)
    def run(self):
        com_ports = []
        for i in range(256):
            try:
                s = serial.Serial('COM'+str(i))
                com_ports.append('COM'+str(i))
                s.close()
            except serial.SerialException:
                pass
        self.com_ports_available.emit(com_ports)
##################################################################################################################################
class App(QMainWindow):
    def __init__(self):
        super(App, self).__init__()
        uic.loadUi("UI/Test_App.ui", self)
        self.setWindowIcon(QIcon('images_/icons/Moewe.jpg'))
        # self.setFixedSize()
        self.setStatusTip('Moewe Optik Gmbh. 2023')
        self.show()
        ########################################################################################################
        self.serial_port = None
        self.thread = None
        self.serial_thread = SerialPortThread()
        self.serial_thread.com_ports_available.connect(self.update_com_ports)
        self.serial_thread.start()
        self.baudrate_box.addItems(['9600','57600','115200'])
        self.baudrate_box.setCurrentText('115200')
        self.connect_button.clicked.connect(self.connect_or_disconnect_serial_port)
        self.refresh_button.clicked.connect(self.refresh_connect)
        ########################################################################################################
        self.commands = ['i2c:scan', 'i2c:read:53:04:FC', 'i2c:write:53:', 'i2c:read:53:20:00', 'i2c:write:73:04', 'i2c:scan','i2c:write:21:0300','i2c:write:21:0100','i2c:write:21:01FF', 'i2c:write:73:01',
                    'i2c:scan', 'i2c:write:4F:06990918', 'i2c:write:4F:01F8', 'i2c:read:4F:1E:00']
        self.start_button.clicked.connect(self.connect)
        ########################################################################################################
        # self.timer = QTimer(self)
        # self.timer.timeout.connect(self.update_time_label)
        # self.timer.start(1000) 
        ########################################################################################################
        self.rm = visa.ResourceManager()
        self.multimeter = None
        self.powersupply = None
        self.test_button.clicked.connect(self.on_cal_voltage_current)
        ########################################################################################################
        self.config_file = configparser.ConfigParser()
        self.config_file.read('conf_igg.ini')
        self.PS_channel = self.config_file.get('Power Supplies', 'Channel_set')
        self.max_voltage = self.config_file.get('Power Supplies', 'Voltage_set')
        self.max_current = self.config_file.get('Power Supplies', 'Current_set')
        self.value_edit.returnPressed.connect(self.load_voltage_current)
        ########################################################################################################

        # self.firstMessage()
        self.on_button_click('images_/images/PP1.jpg')
        self.info_label.setText('Welcome \n \n Press START')
        ########################################################################################################


        # self.save_button.clicked.connect(self.create_ini_file)


        self.test_images = ['images_/images/R700.jpg','images_/images/R709_before_jumper.jpg','images_/images/R700_DC.jpg', 'images_/images/PP2.png','images_/images/C443.jpg','images_/images/C442.jpg','images_/images/C441.jpg','images_/images/C412.jpg',
                            'images_/images/C430.jpg','images_/images/C443_1.jpg','images_/images/C442_1.jpg','images_/images/C441_1.jpg','images_/images/C412_1.jpg','images_/images/C430_1.jpg', 'images_/images/PP.jpg',]
        self.test_index = 0
        self.DCV_readings = [0,0,0,0,0,0,0]
        self.ACV_readings = [0,0,0,0,0,0,0]
        #########################################################################self.rm.open_resource('TCPIP0::192.168.222.207::INSTR')###############################
    def firstMessage(self):
        msgBox = QMessageBox()
        msgBox.setWindowIcon(QIcon('images_/icons/icon.png'))
        msgBox.setText("Welcome to the testing world.")
        msgBox.setInformativeText("Press OK if you are ready.")
        msgBox.setWindowTitle("Message")
        self.on_button_click('images_/images/Welcome.jpg')
        msgBox.setStandardButtons(QMessageBox.Ok)
        ret_value = msgBox.exec_()
        if ret_value == QMessageBox.Ok:
            self.secondMessage()
    ########################################################################################################
    def secondMessage(self):
        msgBox = QMessageBox()
        msgBox.setWindowIcon(QIcon('images_/icons/icon.png'))
        msgBox.setText("Press the PowerON buttons of PowerSupply and Multimeter to avoid delay.\n\nSet all the environment as shown in the image.\n\nRead the information carefully everytime.\n")
        msgBox.setInformativeText("Then press the button.")
        msgBox.setWindowTitle("Message")
        msgBox.setStandardButtons(QMessageBox.Ok)
        ret_value = msgBox.exec_()
        if ret_value == QMessageBox.Ok:
            self.title_label.setText('Preparation Test')
            self.info_label.setText('Welcome \n \n Press START')
            self.on_button_click('images_/images/PP1.jpg')
    ########################################################################################################
    def on_button_click(self, file_path):
        if file_path:
            pixmap = QPixmap(file_path)
            self.image_label.setPixmap(pixmap)
            self.image_label.setScaledContents(True)
            self.image_label.setFixedSize(pixmap.width(), pixmap.height())

            if self.start_button.text() == 'STEP4':
                reply = self.show_good_message('Check All your environment Correct. We cannot change later')                
                if reply == QMessageBox.Yes:                    
                    self.start_button.setText('NEXT')
                    self.on_button_click('images_/images/On_Devices.jpg')
                    self.info_label.setText('Power ON of the Powersupply and also Multimeter...\n and wait for 12 seconds.\n Press "NEXT"')                    
                else:
                    self.on_button_click('images_/icons/next_1.jpg')

    def connect(self):
        if self.start_button.text() == 'START':
            self.on_button_click('images_/images/board_on_mat_.jpg')
            self.info_label.setText("Legen Sie die Platine auf die ESD-Matte\n\n(siehe Abbildung rechts).\n\nÜberprüfen Sie die gesamte Umgebung anhand der Abbildung.\n\nPrüfen Sie alle Anschlüsse.\n\n Drücken Sie 'STEP1'.")
            self.start_button.setText("STEP1")
        elif self.start_button.text() == 'STEP1':
            self.on_button_click('images_/images/PP4_.jpg')
            self.info_label.setText('Überprüfen Sie alle "4" Schrauben der Platine (siehe Abbildung).\n\nMontieren Sie alle 4 Schrauben (4x M2,5x5 Torx)\n\nDrücken Sie "STEP2".')
            self.start_button.setText("STEP2")
        elif self.start_button.text() == 'STEP2':
            self.on_button_click('images_/images/board_on_mat.jpg')
            self.info_label.setText('AFter fitting the 4 Screws, \n\n Place back the Board on ESD Matte. \\n then, Press "STEP3"')
            self.start_button.setText("STEP3")
        elif self.start_button.text()=='STEP3':
            self.on_button_click('images_/images/board_with_cabels.jpg')
            self.info_label.setText('Connect the Power Cables to the board (see image).\n\n\n Press STEP4.')
            self.start_button.setText("STEP4")
        elif self.start_button.text()=='STEP4':
            self.on_button_click('images_/icons/next.jpg')
        elif self.start_button.text()=='NEXT':
            self.info_label.setText("You can see MULTIMETER Name on TextBox.\n\nwait for 10-15 seconds.\n\nPress MULTI ON.")
            self.start_button.setEnabled(False)
            self.show_good_message('Wait for 10 seconds. Untill the Powersupply and Multimeter get SET')
            self.start_button.setText('MULTI ON')            
            self.on_button_click('images_/images/PP9.jpg')
        elif self.start_button.text()=='MULTI ON':
            self.connect_multimeter()
        elif self.start_button.text()=='POWER ON':
            self.connect_powersupply()
        elif self.start_button.text()=='STROM-I':
            time.sleep(2)
            self.calc_voltage_before_jumper()
        elif self.start_button.text()=='SPANNUNG':
            self.start_button.setEnabled(False)
            self.voltage_before_jumper = float(self.multimeter.query('MEAS:VOLT:DC?'))
            self.result_label.setText('Voltage before Jumper\n\n'+str(float(self.voltage_before_jumper))+'V')
            time.sleep(2)
            if 3.28 < self.voltage_before_jumper < 3.38:
                QMessageBox.information(self, 'Information', 'Proceed Next')
            else:
                QMessageBox.information(self, 'Information', 'Wrong Voltage. Check Connections again.')
            self.start_button.setEnabled(True)
            self.start_button.setText('POWER OFF')            
            self.info_label.setText('Press "POWER OFF" button')
            self.on_button_click('images_/images/Start2.png')
        elif self.start_button.text()=='POWER OFF':
            self.powersupply.write('OUTPut '+self.PS_channel+',OFF')
            self.info_label.setText('press "Close J" button\n and close the JUMPER with Soldering \n wait 10 seconds')
            self.start_button.setText('Close J')
            self.on_button_click('images_/images/close_jumper.jpg')
        elif self.start_button.text()=='Close J':
            reply = self.show_good_message('CLOSE the Jumper with Soldering. \n If You Close then Press YES')
            if reply == QMessageBox.Yes:                    
                self.start_button.setText('STROM')
                self.on_button_click('images_/images/PP8.jpg')
                self.powersupply.write('OUTPut '+self.PS_channel+',ON')
                self.info_label.setText('Press STROM button...\n and and Calculate the supply current\n after closed the JUMPER')
            else:
                self.on_button_click('images_/images/close_jumper.jpg')
        elif self.start_button.text()=='STROM':
            time.sleep(2)
            self.calc_voltage_before_jumper()
            
        elif self.start_button.text() == 'NEXTT':
            self.on_button_click('images_/images/Start2.png')
            self.info_label.setText('Press FPGA...\n Switch OFF the PowerSupply and Multimeter.')
            self.start_button.setText('FPGA')
        elif self.start_button.text()=='FPGA':
            self.on_button_click('images_/images/Fix_Bolts.jpg')
            self.info_label.setText('Press FPGA> \n Take the board and fix the bolts (refer to Image).')
            self.start_button.setText('FPGA>')
        elif self.start_button.text() == 'FPGA>':
            self.start_button.setText('FPGA>>')
            self.on_button_click('images_/images/Fix_Dip_Switch.jpg')
            self.info_label.setText("See the Image and Do the same.\n\n Press FPGA>>")
        elif self.start_button.text() == 'FPGA>>':
            self.start_button.setText('FPGA>>>')
            self.on_button_click('images_/images/FPGA_.jpg')
            self.info_label.setText("See the Image and Do the same\n\n Press FPGA>>>")
        elif self.start_button.text() == 'FPGA>>>':
            self.start_button.setText('FPGA>>>>')
            self.on_button_click('images_/images/FPGA_2.jpg')
            self.info_label.setText("See the Image and Do the same\n\n Press FPGA>>>>")
        elif self.start_button.text() == 'FPGA>>>>':
            self.start_button.setText('AUTO_Test')
            self.on_button_click('images_/images/FPGA_B.jpg')
            self.info_label.setText("See the Image and Do the same\n\n Press FPGA>>>>>")
        elif self.start_button.text() == 'AUTO_Test':
            self.on_button_click('images_/icons/next.jpg')
            self.info_label.setText("Press Connect Button")
        elif self.start_button.text()=='SERIAL TEST':
            self.start_process()
            self.info_label.setText('Wait 10 seconds ')
            self.on_button_click('images_/images/Start2.png')
            QMessageBox.information(self, "Important", "Read and watch the image clearly and do the process carefully.")

        ########################################################################################################
    def load_voltage_current(self):
        if self.vals_button.text() == 'CH':
            self.PS_channel = str(self.value_edit.text())
            self.powersupply.write('INSTrument '+self.PS_channel)
            self.config_file.set('Power Supplies', 'Channel_set', self.PS_channel)
            self.vals_button.setText('V')
            self.value_edit.setText(self.max_voltage)
            self.info_label.setText('Write 30 in the Yellow Box next to "V" \n \n Press "Enter"\n You can check the value in the powersupply.')
            self.on_button_click('images_/images/PP7_1.jpg')
            self.value_edit.setValidator(QRegExpValidator(QRegExp(r'^\d+(\.\d+)?$')))
        elif self.vals_button.text() == 'V':
            self.max_voltage =  str(self.value_edit.text())
            self.config_file.set('Power Supplies', 'Voltage_set', self.max_voltage)
            self.powersupply.write(self.PS_channel+':VOLTage ' + self.max_voltage)
            self.vals_button.setText('I')
            self.value_edit.setText(self.max_current)
            self.info_label.setText('Enter 0.5 in the box next to I\n\n Press "Enter".\n Check the value change in the Powersupply.')
            self.on_button_click('images_/images/PP7_2.jpg')
        elif self.vals_button.text() == 'I':
            self.max_current = self.value_edit.text()
            self.powersupply.write('CH1:CURRent ' + self.max_current)
            # self.textBrowser.append(self.powersupply.query(self.PS_channel+':CURRent?'))
            self.info_label.setText('Enter 0.05 in the box next to I')
            self.powersupply.write('OUTPut '+self.PS_channel+',ON')
            self.info_label.setText('Press STROM-I\n Check the "CUrrent" Value.') # modify here'
            self.start_button.setText('STROM-I')
            self.value_edit.setStyleSheet("")
            self.value_edit.clear()
            self.on_button_click('images_/images/PP8.jpg')
        else:
            self.textBrowser.append('Wrong Input')

    def calc_voltage_before_jumper(self):
        current = float(self.powersupply.query('MEASure:CURRent? '+self.PS_channel))
        if self.start_button.text() == 'STROM-I':
            self.result_label.setText('Current before Jumper\n'+str(current)+'A')
            if 0.04 <= current <= 0.06:
                self.start_button.setText('SPANNUNG')
                self.current_before_jumper = current
                self.on_button_click('images_/images/R709_before_jumper.jpg')
                self.info_label.setText('Press SPANNUNG to Calculate initial VOLTAGE at R709.\n \n Calculate Voltage at the Component \n Shown in the figure.')
            else:
                self.current_before_jumper = current
                self.start_button.setText('SPANNUNG')
                self.info_label.setText('Press SPANNUNG to Calculate initial VOLTAGE at R709.\n \n Calculate Voltage at the Component \n Shown in the figure.')
                self.on_button_click('images_/images/R709_before_jumper.jpg')
        elif self.start_button.text() == 'STROM':
            self.result_label.setText('Current After Jumper\n'+str(current)+'A')
            if 0.09 <= current <= 0.15:
                self.current_after_jumper = current
                QMessageBox.information(self, "Information", "Now Everything is perfect. Please be care full with each and every step from here.")
                self.on_button_click('images_/images/R709.jpg')
                self.info_label.setText('\n \n Press TEST V Button to run the Voltage Tests. Be careful.')
                self.test_button.setText('TEST-V')                
            else:
                QMessageBox.information(self, 'Information', 'Supplying Current is either more or less. So please Swith OFF the PowerSupply, and Put back all the Euipment back.')

    def connect_multimeter(self):
        if not self.multimeter:
            try:
                self.multimeter = self.rm.open_resource('TCPIP0::192.168.222.207::INSTR')
                self.textBrowser.append(self.multimeter.query('*IDN?'))
                self.on_button_click('images_/images/Power_ON_PS.jpg')
                self.start_button.setText('POWER ON')
                self.info_label.setText('Press POWER ON button.\n \n It connects the powersupply...!' )
            except visa.errors.VisaIOError:
                self.textBrowser.append('Multimeter has not been presented')
        else:
            self.multimeter.close()
            self.multimeter = None
            self.textBrowser.append(self.multimeter.query('*IDN?'))

    def connect_powersupply(self):
        if not self.powersupply:
            try:
                self.powersupply = self.rm.open_resource('TCPIP0::192.168.222.141::INSTR')
                self.textBrowser.append(self.powersupply.query('*IDN?'))
                self.value_edit.setStyleSheet("background-color: lightyellow;")
                self.info_label.setText('Write CH1 in the Yellow Box (Highlighted)\n \n next to CH \n\n Press "ENTER"')
                self.vals_button.setText('CH')
                self.value_edit.setText(self.PS_channel)
                self.on_button_click('images_/images/PP7.jpg')
            except visa.errors.VisaIOError:
                QMessageBox.information(self, "PowerSupply Connection", "PowerSupply is not present at the given IP Address.")
                self.textBrowser.setText('Powersupply has not been presented.')
        else:
            self.powersupply.close()
            self.powersupply = None
            self.PS_button.setText('PS ON')
            self.textBrowser.setText('Netzteil Disconnected')

    def show_good_message(self, message):
        self.timer1 = QTimer()
        self.timer1.timeout.connect(self.enable_button)
        self.timer1.start(10000)
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Question)
        msgBox.setText(message)
        msgBox.setWindowTitle("Congratulations!")
        self.title_label.setText('Powersupply Test')
        msgBox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)        
        return msgBox.exec_()
    


    def on_cal_voltage_current(self):
        self.test_button.setEnabled(False)
        QMessageBox.information(self, "Important", "Careful with the Test Leads.")
        self.image_timer = QTimer(self)
        self.image_timer.timeout.connect(self.change_image)
        self.image_timer.start(5000)

    def change_image(self):
        if self.test_index < len(self.test_images):
            image_name = self.test_images[self.test_index]
            self.on_button_click(image_name)
            if self.test_index == 0:                
                time.sleep(5)
                attempt = 0
                self.DCV_readings[0] = self.DC_voltage_R709()
                while attempt < 2 and not (3.25 < self.DCV_readings[0] < 3.35):
                    attempt += 1
                    time.sleep(5)
                    self.DCV_readings[0] = self.DC_voltage_R709()
                if 3.25 < self.DCV_readings[0] < 3.35:
                    self.info_label.setText('Calculate voltage at R700')
                else:
                    self.info_label.setText('Mesurement is not in range.')
                self.info_label.setText('Calculate Voltage at this R700 component. Wait 5 seconds')

            elif self.test_index == 1:
                time.sleep(5)
                attempt = 0
                self.DCV_readings[1] = self.DC_voltage_R700()
                while attempt < 2 and not (4.95 < self.DCV_readings[1] < 5.05):
                    attempt += 1
                    time.sleep(5)
                    self.DCV_readings[1] = self.DC_voltage_R700()
                if 4.95 < self.DCV_readings[1] < 5.05:
                    self.info_label.setText('Measure at R709 Component. wait 5 seconds to get the readings.')
                else:
                    self.info_label.setText('Measurement is not in exected range. Do another time ot pack it back to Box.')



            elif self.test_index == 2:
                time.sleep(2)
                self.multimeter.query('MEAS:VOLT:AC?')

                time.sleep(5)
                attempt = 0
                self.ACV_readings[0] = self.AC_voltage_R709_R700()
                while attempt < 2 and not (self.ACV_readings[0] <= 0.01):
                    attempt += 1
                    time.sleep(5)
                    self.ACV_readings[0] = self.AC_voltage_R709_R700()
                if self.ACV_readings[0] <= 0.01:
                    self.info_label.setText('Measure AC Voltage at R700. \n Careful.')
                else:
                    self.info_label.setText('Measurement is not in exected range. Do another time ot pack it back to Box.')



            elif self.test_index == 3:
                time.sleep(5)
                attempt = 0
                self.ACV_readings[1] = self.AC_voltage_R709_R700()
                while attempt < 2 and not (self.ACV_readings[1] <= 0.01):
                    attempt += 1
                    time.sleep(5)
                    self.ACV_readings[1] = self.AC_voltage_R709_R700()
                if self.ACV_readings[1] <= 0.01:
                    self.info_label.setText('Measure DC voltage at C443.. \n check the image for component placement.')
                else:
                    self.info_label.setText('Measurement is not in exected range. Do another time ot pack it back to Box.')



            elif self.test_index == 4:
                self.show_good_message('Change the Ground Connection according to the Image.')


            elif self.test_index == 5:
                time.sleep(5)
                attempt = 0
                self.DCV_readings[2] = self.DC_voltage_C443()
                while attempt < 2 and not (11.95 <= self.DCV_readings[2] <= 12.05):
                    attempt += 1
                    time.sleep(5)
                    self.DCV_readings[2] =  self.DC_voltage_C443()
                if 11.95 <= self.DCV_readings[2] <= 12.05:
                    self.info_label.setText('Measure DC voltage at C442.. \n ')
                else:
                    self.info_label.setText('Measurement is not in exected range. Do another time ot pack it back to Box.')



            elif self.test_index == 6:
                time.sleep(5)
                attempt = 0
                self.DCV_readings[3] = self.DC_voltage_C442_C441()
                while attempt < 2 and not (4.95 < self.DCV_readings[3] < 5.05):
                    attempt += 1
                    time.sleep(5)
                    self.DCV_readings[3] =  self.DC_voltage_C442_C441()
                if 4.95 < self.DCV_readings[3] < 5.05:
                    self.info_label.setText('Measure AC Voltage at C441 Component')
                else:
                    self.info_label.setText('Measurement is not in exected range. Do another time ot pack it back to Box.')



            elif self.test_index == 7:
                time.sleep(5)
                attempt = 0
                self.DCV_readings[4] = self.DC_voltage_C442_C441()
                while attempt < 2 and not (4.95 < self.DCV_readings[4] < 5.05):
                    attempt += 1
                    time.sleep(5)
                    self.DCV_readings[4] =  self.DC_voltage_C442_C441()
                if 4.95 < self.DCV_readings[4] < 5.05:
                    self.info_label.setText('Measure AC Voltage at C412 Component')
                else:
                    self.info_label.setText('Measurement is not in exected range. Do another time ot pack it back to Box.')


            elif self.test_index == 8:
                time.sleep(5)
                attempt = 0
                self.DCV_readings[5] = self.DC_voltage_C412()
                while attempt < 2 and not (4.98 < self.DCV_readings[5] < 5.02):
                    attempt += 1
                    time.sleep(5)
                    self.DCV_readings[5] =  self.DC_voltage_C412()
                if 4.98 < self.DCV_readings[5] < 5.02:
                    self.info_label.setText('Measure AC Voltage at C430 Component')
                else:
                    self.info_label.setText('Measurement is not in exected range. Do another time ot pack it back to Box.')



            elif self.test_index == 9:
                time.sleep(5)
                attempt = 0
                self.DCV_readings[6] = self.DC_voltage_C430()
                while attempt < 2 and not (2.028 < self.DCV_readings[6] < 2.068):
                    attempt += 1
                    time.sleep(5)
                    self.DCV_readings[6] = self.DC_voltage_C430()
                if 2.028 < self.DCV_readings[6] < 2.068:
                    self.info_label.setText('Measure AC voltage at C443.. \n ')
                else:
                    self.info_label.setText('Measurement is not in exected range. Do another time ot pack it back to Box.')




            elif self.test_index == 10:
                time.sleep(5)
                attempt = 0
                self.ACV_readings[2] = self.AC_voltage_C443()
                while attempt < 2 and not (self.ACV_readings[2] <= 0.01):
                    attempt += 1
                    time.sleep(5)
                    self.ACV_readings[2] = self.AC_voltage_C443()
                if self.ACV_readings[2] <= 0.01:
                    self.info_label.setText('Measure AC voltage at C442.. \n ')
                else:
                    self.info_label.setText('Measurement is not in exected range. Do another time ot pack it back to Box.')



            elif self.test_index == 11:
                time.sleep(5)
                attempt = 0
                self.ACV_readings[3] = self.AC_voltage_C442_C441()
                while attempt < 2 and not (self.ACV_readings[3] <= 0.05):
                    attempt += 1
                    time.sleep(5)
                    self.ACV_readings[3] = self.AC_voltage_C442_C441()
                if self.ACV_readings[3] <= 0.05:
                    self.info_label.setText('Measure AC voltage at C441.. \n ')
                else:
                    self.info_label.setText('Measurement is not in exected range. Do another time ot pack it back to Box.')



            elif self.test_index == 12:
                time.sleep(5)
                attempt = 0
                self.ACV_readings[4] = self.AC_voltage_C442_C441()
                while attempt < 2 and not (self.ACV_readings[4] <= 0.05):
                    attempt += 1
                    time.sleep(5)
                    self.ACV_readings[4] = self.AC_voltage_C442_C441()
                if self.ACV_readings[4] <= 0.05:
                    self.info_label.setText('Measure AC voltage at C412.. \n ')
                else:
                    self.info_label.setText('Measurement is not in exected range. Do another time ot pack it back to Box.')



            elif self.test_index == 13:
                time.sleep(5)
                attempt = 0
                self.ACV_readings[5] = self.AC_voltage_C412_C430()
                while attempt < 2 and not (self.ACV_readings[5] <= 0.001):
                    attempt += 1
                    time.sleep(5)
                    self.ACV_readings[5] = self.AC_voltage_C412_C430()
                if self.ACV_readings[5] <= 0.001:
                    self.info_label.setText('Measure AC voltage at C430.. \n ')
                else:
                    self.info_label.setText('Measurement is not in exected range. Do another time ot pack it back to Box.')



            elif self.test_index == 14:
                time.sleep(5)
                attempt = 0
                self.ACV_readings[6] = self.AC_voltage_C412_C430()
                while attempt < 2 and not (self.ACV_readings[6] <= 0.001):
                    attempt += 1
                    time.sleep(5)
                    self.ACV_readings[6] = self.AC_voltage_C412_C430()
                if self.ACV_readings[6] <= 0.001:
                    self.info_label.setText('All Measurements Complete... \n ')
                else:
                    self.info_label.setText('Measurement is not in exected range. Do another time ot pack it back to Box.')
                ####################################################################################################################################
                self.start_button.setText('NEXTT')
            self.test_index += 1

        else:            
            self.image_timer.stop()
    ############################################################################################################################################
    def DC_voltage_R709(self):
        voltage = float(self.multimeter.query('MEAS:VOLT:DC?'))
        if 3.25 <= voltage <= 3.35:
            self.textBrowser.append('DC Voltage at R709 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: green;")
            self.result_label.setText('DC VOltage @R709\n'+str(voltage))
        else:
            self.textBrowser.append('DC Voltage at R709 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: red;")
            self.result_label.setText('DC VOltage @R709\n'+str(voltage))
            QMessageBox.information(self, "Information", "Value wrong. Will check another time")
        QMessageBox.information(self, "Information", "Check the Image to measure component.")
        return voltage
    
    
    def DC_voltage_R700(self):
        voltage = float(self.multimeter.query('MEAS:VOLT:DC?'))
        if 4.95 <= voltage <= 5.05:
            self.textBrowser.append('DC Voltage at R700 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: green;")
            self.result_label.setText('DC VOltage @R700\n'+str(voltage))
        else:
            self.textBrowser.append('DC Voltage at R700 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: red;")
            self.result_label.setText('DC VOltage @R700\n'+str(voltage))
            QMessageBox.information(self, "Information", "Result has been different with Estimated. Please try all results one more time.")
        QMessageBox.information(self, "Information", "Check the Image to measure component.")
        return voltage

    def AC_voltage_R709_R700(self):
        voltage = float(self.multimeter.query('MEAS:VOLT:AC?'))
        if voltage <= 0.01:
            self.textBrowser.append('DC Voltage at R709/R700 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: green;")
            self.result_label.setText('DC VOltage @R709\n'+str(voltage))
        else:
            self.textBrowser.append('DC Voltage at R709/R700 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: red;")
            self.result_label.setText('DC VOltage @R709 or @R700\n'+str(voltage))
            QMessageBox.information(self, "Information", "Now Everything is perfect. Please be care full with each and every step from here.")   
        QMessageBox.information(self, "Information", "Check the Image to measure component.")
        return voltage

    def DC_voltage_C443(self):
        voltage = float(self.multimeter.query('MEAS:VOLT:DC?'))
        if 11.95 <= voltage <= 12.05:
            self.textBrowser.append('DC Voltage at C442 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: green;")
            self.result_label.setText('DC VOltage @C442\n'+str(voltage))
        else:
            self.textBrowser.append('DC Voltage at C442 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: red;")
            self.result_label.setText('DC VOltage @C442\n'+str(voltage))
            QMessageBox.information(self, "Information", "Now Everything is perfect. Please be care full with each and every step from here.")
        QMessageBox.information(self, "Information", "Check the Image to measure component.")
        return voltage

    def DC_voltage_C442_C441(self):
        voltage = float(self.multimeter.query('MEAS:VOLT:DC?'))
        if 4.95 <= voltage <= 5.05:
            self.textBrowser.append('DC Voltage at C441 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: green;")
            self.result_label.setText('DC VOltage @R709\n'+str(voltage))
        else:
            self.textBrowser.append('DC Voltage at C441 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: red;")
            self.result_label.setText('DC VOltage @C441\n'+str(voltage))
            QMessageBox.information(self, "Information", "Now Everything is perfect. Please be care full with each and every step from here.")
        QMessageBox.information(self, "Information", "Check the Image to measure component.")
        return voltage
    
    def DC_voltage_C412(self):
        voltage = float(self.multimeter.query('MEAS:VOLT:DC?'))
        if 4.98 <= voltage <= 5.02:
            self.textBrowser.append('DC Voltage at C412 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: green;")
            self.result_label.setText('DC VOltage @R709\n'+str(voltage))
        else:
            self.textBrowser.append('DC Voltage at C412 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: red;")
            self.result_label.setText('DC VOltage @C412\n'+str(voltage))
            QMessageBox.information(self, "Information", "Now Everything is perfect. Please be care full with each and every step from here.")
        QMessageBox.information(self, "Information", "Check the Image to measure component.")
        return voltage

    def DC_voltage_C430(self):
        voltage = float(self.multimeter.query('MEAS:VOLT:DC?'))
        if 2.028 <= voltage <= 2.068:
            self.textBrowser.append('DC Voltage at C430 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: green;")
            self.result_label.setText('DC VOltage @C430\n'+str(voltage))
        else:
            self.textBrowser.append('DC Voltage at C430 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: red;")
            self.result_label.setText('DC VOltage @C430\n'+str(voltage))
            QMessageBox.information(self, "Information", "Now Everything is perfect. Please be care full with each and every step from here.")
        QMessageBox.information(self, "Information", "Check the Image to measure component.")
        return voltage

    def AC_voltage_C443(self):
        voltage = float(self.multimeter.query('MEAS:VOLT:AC?'))
        if voltage <= 0.01:
            self.textBrowser.append('AC_voltage_C443 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: green;")
            self.result_label.setText('VOltage @R709\n'+str(voltage))
        else:
            self.textBrowser.append('AC_voltage_C443 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: red;")
            self.result_label.setText('AC_voltage_C443\n'+str(voltage))
            QMessageBox.information(self, "Information", "Now Everything is perfect. Please be care full with each and every step from here.")   
        QMessageBox.information(self, "Information", "Check the Image to measure component.")
        return voltage

    def AC_voltage_C442_C441(self):
        voltage = float(self.multimeter.query('MEAS:VOLT:AC?'))
        if voltage <= 0.05:
            self.textBrowser.append('AC_voltage_C442_C441 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: green;")
            self.result_label.setText('AC_voltage_C442_C441'+str(voltage))
        else:
            self.textBrowser.append('AC_voltage_C442_C441 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: red;")
            self.result_label.setText('AC_voltage_C442_C441 \n'+str(voltage))
            QMessageBox.information(self, "Information", "Now Everything is perfect. Please be care full with each and every step from here.")   
        QMessageBox.information(self, "Information", "Check the Image to measure component.")
        return voltage

    def AC_voltage_C412_C430(self):
        voltage = float(self.multimeter.query('MEAS:VOLT:AC?'))
        if voltage <= 0.001:
            self.textBrowser.append('DC Voltage at AC_voltage_C412_C430 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: green;")
            self.result_label.setText('Voltage\n'+str(voltage))
        else:
            self.textBrowser.append('AC_voltage_C412_C430 Component'+str(voltage))
            self.result_label.setStyleSheet("background-color: red;")
            self.result_label.setText('AC_voltage_C412_C430\n'+str(voltage))
            QMessageBox.information(self, "Information", "Now Everything is perfect. Please be care full with each and every step from here.")   
        QMessageBox.information(self, "Information", "Check the Image to measure component.")
        return voltage

    ########################################################################################################
    def update_time_label(self):
        current_time = QTime.currentTime().toString(Qt.DefaultLocaleLongDate)
        current_date = QDate.currentDate().toString(Qt.DefaultLocaleLongDate)
        self.time_label.setText(f"{current_time} - {current_date}")
        # return current_date, current_time
    ########################################################################################################
    def update_com_ports(self, com_ports):
        self.port_box.clear()
        self.port_box.addItems(com_ports)
    ########################################################################################################
    def connect_or_disconnect_serial_port(self):
        if self.serial_port is None:
            com_port = self.port_box.currentText()            # Get the selected com port and baud rate
            baud_rate = int(self.baudrate_box.currentText())
            self.serial_port = serial.Serial(com_port, baud_rate, timeout=1)            # Create a new serial port object and open it
            self.port_box.setEnabled(False)  # Disable the combo boxes and change the button text
            self.start_button.setText('SERIAL TEST')
            self.baudrate_box.setEnabled(False)
            self.connect_button.setText('Disconnect')
            self.textBrowser.append('Serial Communication Connected')
            # self.powersupply = self.rm.open_resource('TCPIP0::192.168.222.141::INSTR')
            # time.sleep(5)
            # self.powersupply.write('OUTPut '+self.PS_channel+',ON')
            # self.rm.open_resource('TCPIP0::192.168.222.207::INSTR')
            self.refresh_button.setEnabled(False)
        else:
            self.serial_port.close()            # Close the serial port
            self.serial_port = None
            self.connect_button.setEnabled(True)
            self.port_box.setEnabled(True)            # Enable the combo boxes and change the button text
            self.baudrate_box.setEnabled(True)
            self.refresh_button.setEnabled(True)
            self.connect_button.setText('Connect')
            self.textBrowser.append('Communication Disconnected')
    ########################################################################################################
    def refresh_connect(self):
        self.serial_thread.quit()
        self.serial_thread.wait()
        self.serial_thread.start()
    ########################################################################################################
    def on_widget_button_clicked(self, message):
        self.textBrowser.append(message)
    def update_lineinsert(self, response):
            self.id_Edit.setText(response)
    ########################################################################################################
    def start_process(self):
        if self.thread is None or not self.thread.isRunning():
            QMessageBox.information(self, "Process Started", "Process has been started.")
            self.thread = WorkerThread(self.commands, self.serial_port)
            self.thread.result_signal.connect(self.on_widget_button_clicked)
            self.thread.process_completed.connect(self.process_completed)
            self.thread.response_signal.connect(self.update_lineinsert)
            self.thread.start()
        else:
            QMessageBox.warning(self, "Process In Progress", "Process is already running.")
    ########################################################################################################
    def process_completed(self):
        QMessageBox.information(self, "Process Completed", "Process has been completed.")

    def enable_button(self):
        self.timer1.stop()
        self.start_button.setEnabled(True)


    def closeEvent(self, event):
        if not self.powersupply is None:
            try:
                response = self.powersupply.query(f"OUTP? CH{self.PS_channel}")
                print(response)
                if not response == None:
                    self.powersupply.write('OUTPut '+self.PS_channel+',OFF')
                else:
                    pass
            except visa.errors.InvalidSession:
                print('Error in PS')
        else:
            pass
        event.accept()
    ########################################################################################################
    def create_ini_file(self):
        config = configparser.ConfigParser()
        config.add_section('Powersupply Test')
        config.add_section('I2C Test')
        power_supply_values = {
            "Supply Voltage" : self.max_voltage,
            "Supply Current" : self.max_current,
            "current_before_jumper" : self.current_before_jumper,
            "voltage_before_jumper" : self.voltage_before_jumper,
            "current_after_jumper" : self.current_after_jumper,
            "dcv_bw_gnd_n_r709" : self.DCV_readings[0],
            "dcv_bw_gnd_n_r700" : self.DCV_readings[1],
            "acv_bw_gnd_n_r709" : self.ACV_readings[0],
            "acv_bw_gnd_n_r700" : self.ACV_readings[1],
            "dcv_bw_gnd_n_c443" : self.DCV_readings[2],
            "dcv_bw_gnd_n_c442" : self.DCV_readings[3],
            "dcv_bw_gnd_n_c441" : self.DCV_readings[4],
            "dcv_bw_gnd_n_c412" : self.DCV_readings[5],
            "dcv_bw_gnd_n_c430" : self.DCV_readings[6],
            "acv_bw_gnd_n_c443" : self.ACV_readings[2],
            "acv_bw_gnd_n_c442" : self.ACV_readings[3],
            "acv_bw_gnd_n_c441" : self.ACV_readings[4],
            "acv_bw_gnd_n_c412" : self.ACV_readings[5],
            "acv_bw_gnd_n_c430" : self.ACV_readings[6]
        }
        I2C_Values = {
            "UID" : self.uid,
            "ic704_register_reading" : self.ic704_register_reading
        }
        for key, value in power_supply_values.items():
            config.set('Powersupply Test', key, str(value))
        for key, value in I2C_Values.items():
            config.set('I2C Test', key, str(value))
        with open('conf_igg.ini', 'w') as configfile:
            config.write(configfile)


def main():
    app = QApplication(sys.argv)
    Window = App()
    sys.exit(app.exec_())
if __name__ == '__main__':
    main()
