import configparser, datetime, os, sys, time, openpyxl, serial
import pyvisa as visa
from PyQt5 import uic
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

class WorkerThread(QThread):
    process_completed = pyqtSignal()
    result_signal = pyqtSignal(str)
    response_signal = pyqtSignal(str)
    final_result_signal = pyqtSignal(str)
    def __init__(self, commands, serial_port):
        super().__init__()
        self.commands = commands
        self.serial_port = serial_port
    def run(self):
        try:
            if self.serial_port.is_open:
                self.serial_port.timeout = 5
                self.next_command_index = 0
                for command in self.commands:
                    if self.next_command_index < (len(self.commands)+1):
                        self.serial_port.write(command.encode())
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
                self.final_result_signal.emit(response)
                self.serial_port.close()
                self.process_completed.emit()
            else:
                self.result_signal.emit('Serial Port Closed')
        except IndexError:
            self.result_signal.emit('FPGA is not perfectly set')

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
        self.setFixedSize(self.size())
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

        self.test_button.clicked.connect(self.start_continuous_measurement)
        self.actionLoad.triggered.connect(self.showDialog)
        ########################################################################################################
        self.config_file = configparser.ConfigParser()
        self.config_file.read('conf_igg.ini')
        self.PS_channel = self.config_file.get('Power Supplies', 'Channel_set')
        self.max_voltage = self.config_file.get('Power Supplies', 'Voltage_set')
        self.max_current = self.config_file.get('Power Supplies', 'Current_set')
        self.channel_combobox.activated.connect(self.channelSet)
        self.value_edit.returnPressed.connect(self.load_voltage_current)


        # self.final_config_file = configparser.ConfigParser()
        # self.final_config_file.read(self.current_config_file)
        # self.current_before_jumper = self.final_config_file.
        ########################################################################################################

        self.vals_button.setVisible(False)
        self.version_button.setVisible(False)
        self.value_edit.setVisible(False)
        self.version_edit.setVisible(False)
        self.test_button.setVisible(False)

        self.port_label.setVisible(False)
        self.baudrate_label.setVisible(False)
        self.port_box.setVisible(False)
        self.baudrate_box.setVisible(False)
        self.connect_button.setVisible(False)
        self.refresh_button.setVisible(False)
    
        self.result_label.setVisible(False)
        # self.save_button.setVisible(False)
        self.id_Edit.setVisible(False)
        self.Final_result_box.setVisible(False)
        self.ch_button.setVisible(False)
        self.channel_combobox.setVisible(False)

        # self.firstMessage()
        self.on_button_click('images_/images/PP1.jpg')
        self.info_label.setText('Welcome \n \n Drücken Sie die Taste START.')
        ########################################################################################################


        # self.save_button.clicked.connect(self.create_ini_file)


        self.test_images = ['images_/images_/R700.jpg','images_/images_/ISOGND.jpg', 'images_/images_/C443.jpg','images_/images_/C442.jpg','images_/images_/C441.jpg','images_/images_/C412.jpg','images_/images_/C430.jpg',
                            'images_/Images/ISOGND.jpg']
        self.DCV_Results = [0,0,0,0,0,0,0]
        self.ACV_Results = [0,0,0,0,0,0,0]
        #########################################################################self.rm.open_resource('TCPIP0::192.168.222.207::INSTR')###############################
    def firstMessage(self):
        msgBox = QMessageBox()
        msgBox.setWindowIcon(QIcon('images_/icons/icon.png'))
        msgBox.setText("Willkommen bei den Tests..")
        msgBox.setInformativeText("Drücken Sie OK, wenn Sie bereit sind..")
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
        msgBox.setText("Drücken Sie die 'PowerON'-Tasten des Netzteils und des Multimeters, um Verzögerungen zu vermeiden.\nStellen Sie alle Umgebungen wie auf dem Bild gezeigt ein.\n Lesen Sie die Informationen jedes Mal sorgfältig durch.\n")
        msgBox.setInformativeText("drücken Sie die Taste..")
        msgBox.setWindowTitle("Nachricht")
        msgBox.setStandardButtons(QMessageBox.Ok)
        ret_value = msgBox.exec_()
        if ret_value == QMessageBox.Ok:
            self.title_label.setText('Preparation Test')
            self.info_label.setText("Drücken Sie die Taste 'START'.")
            self.on_button_click('images_/images/PP1.jpg')
########################################################################################################

    def update_line_edit_color(self,color):
        palette = QPalette()
        palette.setColor(QPalette.Base, QColor(color))
        self.result_label.setPalette(palette)


    def showDialog(self):
        fname, _ = QFileDialog.getOpenFileName(self, 'Open file', '/', 'INI Files (*.ini);;All Files (*)')
        if fname:
            self.loadConfigFile(fname)
    def loadConfigFile(self, filename):
        config = configparser.ConfigParser()
        config.read(filename)
        missing_attributes = []
        for section in config.sections():
            for key, value in config.items(section):
                if not value.strip():
                    missing_attributes.append(f"{key} in section {section}")
        if missing_attributes:
            message = "The following attributes are missing or empty in the config file:\n"
            message += "\n".join(missing_attributes)
            message += "\n\nPlease fill in the missing values and try again."
            QMessageBox.warning(self, "Missing Attributes", message, QMessageBox.Ok)
        else:
            self.model_num = config.get('Settings', 'Model Num')
            self.version_num = config.get('Settings', 'Version Num')
            self.serial_num = config.get('Settings', 'Serial Num')
            self.supply_voltage = config.get('Power Supply', 'Supply_Voltage')
            print(self.model_num, self.version_num, self.serial_num, self.supply_voltage)

    ########################################################################################################
    def on_button_click(self, file_path):
        if file_path:
            pixmap = QPixmap(file_path)
            self.image_label.setPixmap(pixmap)
            self.image_label.setScaledContents(True)
            self.image_label.setFixedSize(pixmap.width(), pixmap.height())

            if self.start_button.text() == 'STEP4':
                reply = self.show_good_message('Prüfen Sie Ihre gesamte Umgebung korrekt und sorgfältig. Wir können sie später nicht mehr ändern.')                
                if reply == QMessageBox.Yes:                    
                    self.start_button.setText('NEXT')
                    self.on_button_click('images_/images/On_Devices.jpg')
                    self.info_label.setText('Drücken Sie den "Power ON"-Schalter am \n\n Netzteil und auch am Multimeter. Siehe die Knöpfe auf der \n\n nebenstehenden Abbildung. Warten Sie 10 bis 12 Sekunden,\n\n num diese Geräte einzuschalten. Sie können die\n\nTaste "MULTI ON" später sehen. Drücken Sie die Taste "NEXT".')                    
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
            self.info_label.setText('Nachdem Sie die 4 Schrauben angebracht haben,\n\n legen Sie die Platine wieder auf die ESD-Matte und\n\n drücken Sie dann "STEP3".')
            self.start_button.setText("STEP3")
        elif self.start_button.text()=='STEP3':
            self.on_button_click('images_/images/board_with_cabels.jpg')
            self.info_label.setText('Schließen Sie die Stromkabel an die Platine an (siehe Abbildung).\n\n\nDrücken Sie STEP4.')
            self.start_button.setText("STEP4")
            item = QTableWidgetItem('STEP4')
            self.tableWidget.setItem(0, 1, item)

        elif self.start_button.text()=='STEP4':
            self.on_button_click('images_/icons/next.jpg')
        elif self.start_button.text()=='NEXT':
            self.info_label.setText("Sie können MULTIMETER Name auf TextBox sehen.\n\nDrücken Sie die Taste MULTI ON, wenn sie erscheint.")
            self.start_button.setEnabled(False)
            self.show_good_message('Warten Sie 10 Sekunden lang. Bis das Netzgerät und das Multimeter SET sind.')
            self.start_button.setText('MULTI ON')            
            self.on_button_click('images_/images/PP9.jpg')
            self.show_input_dialog()
        elif self.start_button.text()=='MULTI ON':
            self.connect_multimeter()
        elif self.start_button.text()=='POWER ON':
            self.connect_powersupply()
        elif self.start_button.text()=='STROM-I':
            self.result_label.setVisible(True)
            time.sleep(1)
            # self.calc_voltage_before_jumper()
            # self.result_label.setStyleSheet("")
            current = float(self.powersupply.query('MEASure:CURRent? '+self.PS_channel))
            if 0.04 <= current <= 0.06:
                self.update_line_edit_color('green')
                self.result_label.setText(str(current))

                print('correct',current)
            else:
                self.update_line_edit_color('red')
                self.result_label.setText(str(current))
                print(current)
            self.start_button.setText('SPANNUNG')

        elif self.start_button.text()=='SPANNUNG':
            self.start_button.setEnabled(False)
            time.sleep(0.5)
            self.voltage_before_jumper = self.voltage_find_before_jumper()
            atmpt = 0
            while not ( 3.25 < self.voltage_before_jumper < 3.35):
                self.result_label.setStyleSheet("background-color: red;")
                self.result_label.setText('Voltage before Jumper\n\n'+str(self.voltage_before_jumper)+'V')
                QMessageBox.information(self, 'Information', 'wrong Voltage.')
                atmpt += 1
                time.sleep(0.5)
                self.voltage_before_jumper = self.voltage_find_before_jumper()
            if 3.25 < self.voltage_before_jumper < 3.35:
                self.result_label.setStyleSheet("background-color: green;")
                self.start_button.setText('SPANNUNG')
                self.result_label.setText('Voltage before Jumper\n\n'+str(self.voltage_before_jumper)+'V')
                QMessageBox.information(self, 'Information', 'Die Spannung zwischen GND und R709 liegt zwischen 3,25 und 3,35. Dies ist ein guter Wert. Fahren Sie fort, indem Sie die Taste NEXT drücken.')
                
            self.start_button.setEnabled(True)
            self.result_label.setStyleSheet("")
            self.start_button.setText('POWER OFF')
            self.info_label.setText('Drücken Sie die Taste "POWER OFF", um das Netzteil auszuschalten.')
            self.on_button_click('images_/images/PowerOFF.jpg')

        elif self.start_button.text()=='POWER OFF':
            self.result_label.setVisible(False)
            time.sleep(1)
            self.powersupply.write('OUTPut '+self.PS_channel+',OFF')
            self.info_label.setText('press "Close J" button\n and close the JUMPER with Soldering \n wait 10 seconds')
            self.start_button.setText('Close J')
            self.on_button_click('images_/images/close_jumper.jpg')
        elif self.start_button.text()=='Close J':
            reply = self.show_good_message('CLOSE the Jumper with Soldering. \n If You Close then Press YES')
            if reply == QMessageBox.Yes:                    
                self.start_button.setText('STROM')
                self.on_button_click('images_/images/PP8.jpg')
                time.sleep(1)
                self.powersupply.write('OUTPut '+self.PS_channel+',ON')
                self.info_label.setText('Press STROM button...\n and and Calculate the supply current\n after closed the JUMPER')
            else:
                self.on_button_click('images_/images/close_jumper.jpg')
        elif self.start_button.text()=='STROM':
            time.sleep(1)
            self.calc_voltage_before_jumper()
            self.result_label.setStyleSheet("")
        elif self.start_button.text() == 'NEXTT':
            self.result_label.setVisible(False)
            self.on_button_click('images_/images/PP/_2.jpg')
            self.powersupply.write('OUTPut '+self.PS_channel+',OFF')
            self.multimeter.close()
            self.info_label.setText('Press FPGA...\n')
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
            self.start_button.setText('TEST-I')
            self.on_button_click('images_/images/PP7.jpg')
            self.info_label.setText("Check current. It should be 0.09 and 0.15\n\n Press TEST-I")
        elif self.start_button.text()== 'TEST-I':
            # print('channel',self.PS_channel)
            # self.powersupply = self.rm.open_resource('TCPIP0::192.168.222.141::INSTR')

            # self.powersupply.write('OUTPut '+self.PS_channel+',ON')
            # self.rm.open_resource('TCPIP0::192.168.222.207::INSTR')
            # time.sleep(3)
            self.try_power_multi()
            self.calc_voltage_before_jumper()
            self.start_button.setText('CHECK_L')
            self.on_button_click('images_/images/LED.jpg')
            self.result_label.setVisible(False)
            self.info_label.setText("Check left 2 green lights\n\n Press CHECK_L")
        elif self.start_button.text()=='CHECK_L':
            self.start_button.setText('AUTO_Test')
            self.on_button_click('images_/images/FPGA_B.jpg')
            self.info_label.setText("Check STROM \n\n Press AUTO_Test")
        elif self.start_button.text() == 'AUTO_Test':
            self.on_button_click('images_/icons/next.jpg')
            self.info_label.setText("Press Connect Button")
            self.start_button.setVisible(False)
            self.port_label.setVisible(True)
            self.baudrate_label.setVisible(True)
            self.port_box.setVisible(True)
            self.baudrate_box.setVisible(True)
            self.connect_button.setVisible(True)
            self.refresh_button.setVisible(True)
        elif self.start_button.text()=='SERIAL TEST':
            self.start_process()
            self.info_label.setText('Wait 10 seconds ')
            self.on_button_click('images_/images/LED.png')
            QMessageBox.information(self, "Important", "Watch the image clearly and do the process carefully.")

        ########################################################################################################

    def try_power_multi(self):
        try:
            if not self.powersupply:
                self.powersupply = self.rm.open_resource('TCPIP0::192.168.222.141::INSTR')
                self.powersupply.write('OUTPut '+self.PS_channel+',ON')
            elif not self.multimeter:
                self.rm.open_resource('TCPIP0::192.168.222.207::INSTR')
            else:
                self.powersupply.write('OUTPut '+self.PS_channel+',ON')
                self.rm.open_resource('TCPIP0::192.168.222.207::INSTR')
        except AttributeError:
            self.textBrowser.append('Erro in connecting PowerSupply and Multimeter')

    def voltage_find_before_jumper(self):
        self.multimeter.write('CONF:VOLT:DC 5')
        voltage = float(self.multimeter.query('READ?'))
        time.sleep(2)
        return voltage

    def channelSet(self):
        self.PS_channel = self.channel_combobox.currentText()
        self.powersupply.write('INSTrument '+self.PS_channel)
        self.config_file.set('Power Supplies', 'Channel_set', self.PS_channel)
        self.value_edit.setVisible(True)
        self.vals_button.setVisible(True)
        self.vals_button.setText('V')
        self.value_edit.setText(self.max_voltage)
        self.info_label.setText('Write 30 in the Yellow Box next to "V" \n \n Press "Enter"\n You can check the value in the powersupply.')
        self.on_button_click('images_/images/PP7_1.jpg')
        self.value_edit.setValidator(QRegExpValidator(QRegExp(r'^\d+(\.\d+)?$')))

    def load_voltage_current(self):
        if self.vals_button.text() == 'V':
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
            self.info_label.setText('Enter 0.05 in the box next to I')
            self.powersupply.write('OUTPut '+self.PS_channel+',ON')
            self.value_edit.setVisible(False)
            self.start_button.setVisible(True)
            self.info_label.setText('Press STROM-I\n Check the "Current" Value.') # modify here'
            self.start_button.setText('STROM-I')
            self.value_edit.setStyleSheet("")
            self.value_edit.clear()
            self.vals_button.setVisible(False)
            self.ch_button.setVisible(False)
            self.channel_combobox.setVisible(False)
            self.on_button_click('images_/images/PP8.jpg')
        else:
            self.textBrowser.append('Wrong Input')
########################################################################################################
    def calc_voltage_before_jumper(self):
        current = float(self.powersupply.query('MEASure:CURRent? '+self.PS_channel))
        self.result_label.setVisible(True)
        if self.start_button.text() == 'STROM-I':
            self.result_label.setText('Current before Jumper\n'+str(current)+'A')
            if 0.04 <= current <= 0.06:
                self.update_line_edit_color('green')
                self.start_button.setText('SPANNUNG')
                # self.start_button.setVisible(False)
                self.current_before_jumper = current
                self.on_button_click('images_/images/R709_before_jumper.jpg')
                self.info_label.setText('Press SPANNUNG to Calculate initial VOLTAGE at R709.\n \n Calculate Voltage at the Component \n Shown in the figure.')
            else:
                self.update_line_edit_color('red')
                self.current_before_jumper = current
                self.start_button.setText('SPANNUNG')
                self.info_label.setText('Press SPANNUNG to Calculate initial VOLTAGE at R709.\n \n Calculate Voltage at the Component \n Shown in the figure.')
                self.on_button_click('images_/images/R709_before_jumper.jpg')
                self.update_line_edit_color('red')
                QMessageBox.information(self, 'Information', 'Supplying Current is either more or less. So please Swith OFF the PowerSupply, and Put back all the Euipment back.')
            # self.result_label.setVisible(False)
        elif self.start_button.text() == 'STROM':
            self.result_label.setText('Current After Jumper\n'+str(current)+'A')
            if 0.09 <= current <= 0.15:
                self.update_line_edit_color('green')
                self.current_after_jumper = current
                QMessageBox.information(self, "Information", "Now Everything is perfect. Please be care full with each and every step from here.")
                self.on_button_click('images_/images/R709.jpg')
                self.info_label.setText('\n \n Press TEST V Button to run the Voltage Tests. Be careful.')
                self.test_button.setText('TEST-V')
                self.test_button.setVisible(True)
                self.start_button.setVisible(False)
            else:
                self.update_line_edit_color('red')
                QMessageBox.information(self, 'Information', 'Supplying Current is either more or less. So please Swith OFF the PowerSupply, and Put back all the Euipment back.')       
        elif self.start_button.text() == 'TEST-I':
            self.result_label.setText('Current After FPGA\n'+str(current)+'A')
            if 0.095 <= current <= 0.155:
                self.update_line_edit_color('green')
                self.current_after_FPGA = current
                QMessageBox.information(self, "Information", "Perfect. Please be care full with each and every step from here.")
                self.on_button_click('images_/images/R709.jpg')
                self.info_label.setText('\n \n Press TEST V Button to run the Voltage Tests. Be careful.')
                self.test_button.setText('TEST-V')
            else:
                self.update_line_edit_color('red')
                QMessageBox.information(self, 'Information', 'Supplying Current is either more or less. So please Swith OFF the PowerSupply, and Put back all the Equipment back.')          
########################################################################################################
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
########################################################################################################
    def connect_powersupply(self):
        if not self.powersupply:
            try:
                self.powersupply = self.rm.open_resource('TCPIP0::192.168.222.141::INSTR')
                self.textBrowser.append(self.powersupply.query('*IDN?'))
                # self.textBrowser.append(self.powersupply.resource_name)
                self.start_button.setVisible(False)
                self.value_edit.setStyleSheet("background-color: lightyellow;")
                self.info_label.setText('Write CH1 in the Yellow Box (Highlighted)\n \n next to CH \n\n Press "ENTER"')
                self.ch_button.setVisible(True)
                self.channel_combobox.setVisible(True)
                self.on_button_click('images_/images/PP7.jpg')
            except visa.errors.VisaIOError:
                QMessageBox.information(self, "PowerSupply Connection", "PowerSupply is not present at the given IP Address.")
                self.textBrowser.append('Powersupply has not been presented.')
        else:
            self.powersupply.close()
            self.powersupply = None
            self.PS_button.setText('PS ON')
            self.textBrowser.setText('Netzteil Disconnected')
########################################################################################################
    def start_continuous_measurement(self):
        self.textBrowser.append('Measure according to the image..')
        self.image_label.setPixmap(QPixmap('images_/images/R709_.jpg'))
        self.show_good_message('measurement start.')
        self.attempt_count = 0
        self.measure_count = 0
        self.dc_measure_count = 0
        self.ac_measure_count = 0
        self.is_ac_measurement = False
        self.measure()
########################################################################################################
    def show_message(self, message, icon=QMessageBox.Information):
        msg_box = QMessageBox(self)
        msg_box.setIcon(icon)
        msg_box.setText(message)
        msg_box.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        return msg_box.exec_()
########################################################################################################
    def measure(self):
        try:
            if not self.is_ac_measurement:
                dc_voltage = float(self.multimeter.query('MEAS:VOLT:DC?'))
                self.textBrowser.append(f'DC Voltage (Measurement {self.measure_count + 1}, Attempt {self.attempt_count + 1}): {dc_voltage}')
                self.result_label.setText(str(dc_voltage))
                if self.dc_measure_count==0:
                    if 3.25 <= dc_voltage <= 3.35:
                        self.result_label.setStyleSheet('background-color: green;')
                        self.image_label.setPixmap(QPixmap('images_/images/R709_.jpg'))
                        self.DCV_Results[0] = dc_voltage
                        self.dc_measure_count += 1
                        self.measure_count += 1
                        self.is_ac_measurement = True
                        self.update_config_file(self.current_config_file, {'Power Supply': {'voltage_before_jumper': dc_voltage}})
                    else:
                        self.result_label.setStyleSheet('background-color: red;')
                        self.image_label.setPixmap(QPixmap('images_/images/R709_.jpg'))
                        self.DCV_Results[0] = dc_voltage
                        self.attempt_count += 1

                elif self.dc_measure_count == 1:
                    if 4.95 <= dc_voltage <= 5.05:
                        self.result_label.setStyleSheet('background-color: green;')
                        self.image_label.setPixmap(QPixmap('images_/images/R700.jpg'))
                        self.DCV_Results[1] = dc_voltage
                        self.dc_measure_count += 1
                        self.measure_count += 1
                        self.is_ac_measurement = True
                    else:
                        self.result_label.setStyleSheet('background-color: red;')
                        self.image_label.setPixmap(QPixmap('images_/images/R700.jpg'))
                        self.DCV_Results[1] = dc_voltage
                        self.attempt_count += 1

                elif self.dc_measure_count == 2:
                    if 11.90 <= dc_voltage <= 12.10:
                        self.result_label.setStyleSheet('background-color: green;')
                        self.image_label.setPixmap(QPixmap('images_/images/C443.jpg'))
                        self.DCV_Results[2] = dc_voltage
                        self.dc_measure_count += 1
                        self.measure_count += 1
                        self.is_ac_measurement = True
                    else:
                        self.result_label.setStyleSheet('background-color: red;')
                        self.image_label.setPixmap(QPixmap('images_/images/C443.jpg'))
                        self.DCV_Results[2] = dc_voltage
                        self.attempt_count += 1

                elif self.dc_measure_count == 3:
                    if 4.95 <= dc_voltage <= 5.05:
                        self.result_label.setStyleSheet('background-color: green;')
                        self.image_label.setPixmap(QPixmap('images_/images/C442.jpg'))
                        self.DCV_Results[3] = dc_voltage
                        self.dc_measure_count += 1
                        self.measure_count += 1
                        self.is_ac_measurement = True
                    else:
                        self.result_label.setStyleSheet('background-color: red;')
                        self.image_label.setPixmap(QPixmap('images_/images/C442.jpg'))
                        self.DCV_Results[3] = dc_voltage
                        self.attempt_count += 1

                elif self.dc_measure_count == 4:
                    dc_voltage = float(self.multimeter.query('MEAS:VOLT:DC?'))
                    self.textBrowser.append(f'DC Voltage (Measurement {self.measure_count + 1}, Attempt {self.attempt_count + 1}): {dc_voltage}')
                    self.result_label.setText(str(dc_voltage))
                    if 4.95 <= dc_voltage <= 5.05:
                        self.result_label.setStyleSheet('background-color: green;')
                        self.image_label.setPixmap(QPixmap('images_/images/C441.jpg'))
                        self.DCV_Results[4] = dc_voltage
                        self.dc_measure_count += 1
                        self.measure_count += 1
                        self.is_ac_measurement = True
                    else:
                        self.result_label.setStyleSheet('background-color: red;')
                        self.image_label.setPixmap(QPixmap('images_/images/C441.jpg'))
                        self.DCV_Results[4] = dc_voltage
                        self.attempt_count += 1

                elif self.dc_measure_count == 5:
                    dc_voltage = float(self.multimeter.query('MEAS:VOLT:DC?'))
                    self.textBrowser.append(f'DC Voltage (Measurement {self.measure_count + 1}, Attempt {self.attempt_count + 1}): {dc_voltage}')
                    self.result_label.setText(str(dc_voltage))
                    if 4.95 <= dc_voltage <= 5.05:
                        self.result_label.setStyleSheet('background-color: green;')
                        self.image_label.setPixmap(QPixmap('images_/images/C412.jpg'))
                        self.DCV_Results[5] = dc_voltage
                        self.dc_measure_count += 1
                        self.measure_count += 1
                        self.is_ac_measurement = True
                    else:
                        self.result_label.setStyleSheet('background-color: red;')
                        self.image_label.setPixmap(QPixmap('images_/images/C412.jpg'))
                        self.DCV_Results[5] = dc_voltage
                        self.attempt_count += 1

                elif self.dc_measure_count == 6:
                    dc_voltage = float(self.multimeter.query('MEAS:VOLT:DC?'))
                    self.textBrowser.append(f'DC Voltage (Measurement {self.measure_count + 1}, Attempt {self.attempt_count + 1}): {dc_voltage}')
                    self.result_label.setText(str(dc_voltage))
                    if 2.046 <= dc_voltage <= 2.050:
                        self.result_label.setStyleSheet('background-color: green;')
                        self.image_label.setPixmap(QPixmap('images_/images/C430.jpg'))
                        self.DCV_Results[6] = dc_voltage
                        self.dc_measure_count += 1
                        self.measure_count += 1
                        self.is_ac_measurement = True
                    else:
                        self.result_label.setStyleSheet('background-color: red;')
                        self.image_label.setPixmap(QPixmap('images_/images/C430.jpg'))
                        self.DCV_Results[6] = dc_voltage
                        self.attempt_count += 1
                        
                else:
                    self.result_label.setStyleSheet('background-color: red;')
                    self.image_label.setPixmap(QPixmap('images_/images/PP1.jpg'))
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
                self.result_label.setText(str(ac_voltage))
                if self.ac_measure_count == 0:
                    if ac_voltage < 0.01:
                        self.result_label.setStyleSheet('background-color: green;')
                        self.image_label.setPixmap(QPixmap('images_/images/R700.jpg'))
                        self.ACV_Results[0] = ac_voltage
                        self.ac_measure_count += 1
                        self.is_ac_measurement = False
                        # self.measure_count += 1
                    else:
                        self.result_label.setStyleSheet('background-color: red;')
                        self.image_label.setPixmap(QPixmap('images_/images/R709_.jpg'))
                        self.attempt_count += 1
                        self.ACV_Results[0] = ac_voltage

                elif self.ac_measure_count == 1:
                    if ac_voltage < 0.01:
                        self.result_label.setStyleSheet('background-color: green;')
                        self.image_label.setPixmap(QPixmap('images_/images/ISOGND.jpg'))
                        self.ACV_Results[1] = ac_voltage
                        self.ac_measure_count += 1
                        self.is_ac_measurement = False
                        self.show_message('Change the Groung to ISOGND.', QMessageBox.Information)
                        self.image_label.setPixmap(QPixmap('images_/images/C443.jpg'))
                    else:
                        self.result_label.setStyleSheet('background-color: red;')
                        self.image_label.setPixmap(QPixmap('images_/images/R700.jpg'))
                        self.attempt_count += 1
                        self.ACV_Results[1] = ac_voltage

                elif self.ac_measure_count == 2:
                    if ac_voltage < 0.01:
                        self.result_label.setStyleSheet('background-color: green;')
                        self.image_label.setPixmap(QPixmap('images_/images/C442.jpg'))
                        self.ACV_Results[2] = ac_voltage
                        self.ac_measure_count += 1
                        self.is_ac_measurement = False
                    else:
                        self.result_label.setStyleSheet('background-color: red;')
                        self.image_label.setPixmap(QPixmap('images_/images/C443.jpg'))
                        self.ACV_Results[2] = ac_voltage
                        self.attempt_count += 1

                elif self.ac_measure_count == 3:
                    if ac_voltage < 0.005:
                        self.result_label.setStyleSheet('background-color: green;')
                        self.image_label.setPixmap(QPixmap('images_/images/C441.jpg'))
                        self.ACV_Results[3] = ac_voltage
                        self.ac_measure_count += 1
                        self.is_ac_measurement = False
                    else:
                        self.result_label.setStyleSheet('background-color: red;')
                        self.image_label.setPixmap(QPixmap('images_/images/C442.jpg'))
                        self.ACV_Results[3] = ac_voltage
                        self.attempt_count += 1

                elif self.ac_measure_count == 4:
                    if ac_voltage < 0.005:
                        self.result_label.setStyleSheet('background-color: green;')
                        self.image_label.setPixmap(QPixmap('images_/images/C412.jpg'))
                        self.ACV_Results[4] = ac_voltage
                        self.ac_measure_count += 1
                        self.is_ac_measurement = False
                    else:
                        self.result_label.setStyleSheet('background-color: red;')
                        self.image_label.setPixmap(QPixmap('images_/images/C441.jpg'))
                        self.ACV_Results[4] = ac_voltage
                        self.attempt_count += 1

                elif self.ac_measure_count == 5:
                    if ac_voltage <= 0.001:
                        self.result_label.setStyleSheet('background-color: green;')
                        self.image_label.setPixmap(QPixmap('images_/images/C430.jpg'))
                        self.ACV_Results[5] = ac_voltage
                        self.ac_measure_count += 1
                        self.is_ac_measurement = False
                    else:
                        self.result_label.setStyleSheet('background-color: red;')
                        self.image_label.setPixmap(QPixmap('images_/images/C412.jpg'))
                        self.ACV_Results[5] = ac_voltage
                        self.attempt_count += 1

                elif self.ac_measure_count == 6:
                    if ac_voltage <= 0.001:
                        self.result_label.setStyleSheet('background-color: green;')
                        self.image_label.setPixmap(QPixmap('images_/images/FPFPGA.jpg'))
                        self.ACV_Results[6] = ac_voltage
                        self.ac_measure_count += 1
                        self.is_ac_measurement = False
                    else:
                        self.result_label.setStyleSheet('background-color: red;')
                        self.image_label.setPixmap(QPixmap('images_/images/C430.jpg'))
                        self.ACV_Results[6] = ac_voltage
                        self.attempt_count += 1
                else:
                    self.attempt_count += 1


                if self.attempt_count == 3:
                    if self.ac_measure_count == 1:
                        self.attempt_count = 0
                        self.image_label.setPixmap(QPixmap(self.test_images[self.ac_measure_count]))
                        self.save_measurement(ac_voltage)
                        self.ac_measure_count += 1
                        self.is_ac_measurement = False
                        self.show_message('Change the Groung to ISOGND.', QMessageBox.Information)
                        self.image_label.setPixmap(QPixmap('images_/images/C443.jpg'))
                        time.sleep(2)
                    else:
                        self.attempt_count = 0
                        self.image_label.setPixmap(QPixmap(self.test_images[self.ac_measure_count]))
                        self.save_measurement(ac_voltage)
                        self.ac_measure_count += 1
                        self.is_ac_measurement = False
                        time.sleep(2)


        except visa.errors.VisaIOError as e:
            error_message = f'Measurement error: {str(e)}'
            self.textBrowser.append(error_message)
            self.result_label.setStyleSheet('background-color: red;')
            self.show_message(error_message, QMessageBox.Critical)

        if not (self.ac_measure_count  == 7) and self.measure_count < 8 :
            print('self.ac_measure_count',self.ac_measure_count)
            QTimer.singleShot(5000, self.measure)
        else:
            self.show_message('Continuous measurement completed.', QMessageBox.Information)
            print('resulted values are \n', self.ACV_Results ,'\n', self.DCV_Results)
            self.test_button.setVisible(False)
            self.start_button.setText('NEXTT')
            self.start_button.setVisible(True)

    def save_measurement(self, value):
        print(f'Measurement saved: {value}')

        ###########################################################################################################################
    def show_good_message(self, message):
        self.timer1 = QTimer()
        self.timer1.timeout.connect(self.enable_button)
        self.timer1.start(5000)
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Question)
        msgBox.setText(message)
        msgBox.setWindowTitle("Congratulations!")
        self.title_label.setText('Powersupply Test')
        msgBox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)        
        return msgBox.exec_()    
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
            self.start_button.setVisible(True)
            self.connect_button.setVisible(True)
            self.baudrate_box.setEnabled(False)
            self.connect_button.setText('Disconnect')
            self.textBrowser.append('Serial Communication Connected')
            self.refresh_button.setEnabled(False)
        else:
            self.serial_port.close()            # Close the serial port
            self.serial_port = None
            self.connect_button.setEnabled(True)
            self.port_box.setEnabled(True)            # Enable the combo boxes and change the button text
            self.baudrate_box.setEnabled(True)
            self.refresh_button.setEnabled(True)
            self.connect_button.setText('Connect')
            self.start_button.setVisible(False)
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
        self.id_Edit.setVisible(True)
        self.id_Edit.setText(response)

    def update_finalresult(self, response):
        self.Final_result_box.setVisible(True)
        self.Final_result_box.setText(response)
    ########################################################################################################
    def start_process(self):
        if self.thread is None or not self.thread.isRunning():
            QMessageBox.information(self, "Process Started", "Process has been started.")
            self.thread = WorkerThread(self.commands, self.serial_port)
            self.thread.result_signal.connect(self.on_widget_button_clicked)
            self.thread.process_completed.connect(self.process_completed)
            self.thread.response_signal.connect(self.update_lineinsert)
            self.thread.final_result_signal.connect(self.update_finalresult)
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
#################################################################################################################################################################################
    def show_input_dialog(self):
        while True:
            name, ok = QInputDialog.getText(self, 'Enter Name', 'Enter a name for the INI file:')
            if not ok:
                break
            if not name:
                continue  # Ask again if an empty name is provided

            if self.ini_file_exists(name):
                reply = QMessageBox.question(self, 'File Exists', 'File with this name already exists. Choose a new name', QMessageBox.Yes | QMessageBox.No)
                if reply == QMessageBox.No:
                    self.start_button.setText('START')
                    self.start_button.setEnabled(True)
                    break  # Ask again for a new name
                elif reply == QMessageBox.Yes:
                    # self.create_ini_file(name)
                    continue
            else:
                self.create_config_file(name)
                break

    def ini_file_exists(self, name):
        return os.path.exists(name + '.ini')

    def create_config_file(self, name):
        # Check if the file already exists
        ini_file_path = name + '.ini'

        if os.path.exists(ini_file_path):
            reply = QMessageBox.question(self, 'File Exists', 'File with this name already exists. Do you want to overwrite it?', QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.No:
                return
            
        # Get input for 'Model Num', 'Version Num', and 'Serial Num'
        model_num, ok = QInputDialog.getText(self, 'Enter Model Num', 'Enter Model Num:')
        if not ok:
            return
        version_num, ok = QInputDialog.getText(self, 'Enter Version Num', 'Enter Version Num:')
        if not ok:
            return        
        # Default values for power supply attributes
        power_supply_values = {
            "Supply_Voltage": "",
            "Supply_Current": "",
            "current_before_jumper": "",
            "voltage_before_jumper": "",
            "current_after_jumper": "",
            "current_after_FPGA": "",
            "DCV_R709": "",
            "DCV_R700": "",
            "ACV_R709": "",
            "ACV_R700": "",
            "DCV_C443": "",
            "DCV_C442": "",
            "DCV_C441": "",
            "DCV_C412": "",
            "DCV_C430": "",
            "ACV_C443": "",
            "ACV_C442": "",
            "ACV_C441": "",
            "ACV_C412": "",
            "ACV_C430": ""
        }
        I2C_test_values = {
            "UiD": "",
            "Voltage_F": "",
            "Vendor_iD": "",
            "Result": ""
        }
        # Create or update the INI file
        config = configparser.ConfigParser()
        config['Settings'] = {'Model Num': model_num,
                              'Version Num': version_num}
        config['Power Supply'] = power_supply_values
        config['I2C_Test'] = I2C_test_values

        with open(ini_file_path, 'w') as configfile:
            config.write(configfile)
        self.current_config_file = ini_file_path
        print('config file created', self.current_config_file)
        #QMessageBox.information(self, 'Success', 'INI File created successfully.')

############################################################################################
    def update_config_file(self, name, updated_values):
        ini_file_path = name + '.ini'

        if not os.path.exists(ini_file_path):
            QMessageBox.warning(self, 'File Not Found', f'INI file "{ini_file_path}" not found.')
            return

        config = configparser.ConfigParser()
        config.read(ini_file_path)

        # Update the config file with new values
        for section, values in updated_values.items():
            for key, value in values.items():
                config.set(section, key, str(value))

        with open(ini_file_path, 'w') as configfile:
            config.write(configfile)
#########################################################################################################################################################################################

def main():
    app = QApplication(sys.argv)
    Window = App()
    sys.exit(app.exec_())
if __name__ == '__main__':
    main()
