import configparser, datetime, os, sys, time, openpyxl, serial
import pyvisa as visa
from PyQt5 import uic
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *



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
        # self.test_button.clicked.connect(self.on_cal_voltage_current)
        self.test_button.clicked.connect(self.start_measurements)
        ########################################################################################################
        self.config_file = configparser.ConfigParser()
        self.config_file.read('conf_igg.ini')
        self.PS_channel = self.config_file.get('Power Supplies', 'Channel_set')
        self.max_voltage = self.config_file.get('Power Supplies', 'Voltage_set')
        self.max_current = self.config_file.get('Power Supplies', 'Current_set')
        self.channel_combobox.setCurrentText(self.PS_channel)
        self.channel_combobox.activated.connect(self.channelSet)
        self.value_edit.returnPressed.connect(self.load_voltage_current)
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


        self.image_timer = QTimer(self)
        self.image_timer.timeout.connect(self.display_next_image)
        # self.btn_display_images.clicked.connect(self.start_image_display)
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
        elif self.start_button.text()=='MULTI ON':
            self.connect_multimeter()
        elif self.start_button.text()=='POWER ON':
            self.connect_powersupply()
        elif self.start_button.text()=='STROM-I':
            time.sleep(1)
            self.calc_voltage_before_jumper()
            self.result_label.setStyleSheet("")

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
                self.result_label.setText('Voltage before Jumper\n\n'+str(self.voltage_before_jumper)+'V')
                QMessageBox.information(self, 'Information', 'Die Spannung zwischen GND und R709 liegt zwischen 3,25 und 3,35. Dies ist ein guter Wert. Fahren Sie fort, indem Sie die Taste NEXT drücken.')
                self.start_button.setText('SPANNUNG')

            self.start_button.setEnabled(True)
            self.result_label.setStyleSheet("")
            self.start_button.setText('POWER OFF')
            self.info_label.setText('Drücken Sie die Taste "POWER OFF", um das Netzteil auszuschalten.')
            self.on_button_click('images_/images/Start2.png')
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
            self.start_button.setText('TEST-I')
            self.on_button_click('images_/images/PP7.jpg')
            self.info_label.setText("Check current. It should be 0.09 and 0.15\n\n Press TEST-I")
        elif self.start_button.text()== 'TEST-I':
            print('channel',self.PS_channel)
            time.sleep(1)
            try:
                self.powersupply = self.rm.open_resource('TCPIP0::192.168.222.141::INSTR')
                time.sleep(1)
                self.powersupply.write('OUTPut '+self.PS_channel+',ON')
                self.rm.open_resource('TCPIP0::192.168.222.207::INSTR')
                time.sleep(1)
            except visa.errors.VisaIOError:
                self.textBrowser.append('on the powersupply')
            self.calc_voltage_before_jumper()
            self.start_button.setText('CHECK_L')
            self.on_button_click('images_/images/Welcome.jpg')
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
            self.on_button_click('images_/images/Start2.png')
            QMessageBox.information(self, "Important", "Read and watch the image clearly and do the process carefully.")

#####################################################################################################################################################
    def start_measurements(self):
        th = MultimeterThread(self.multimeter)
        th.dcv_volatge.connect(self.on_dc_measurement_completed)
        th.acv_volatge.connect(self.on_ac_measurement_completed)
        th.start()
        self.multimeter_thread = th

    def on_dc_measurement_completed(self, result):
        self.textBrowser.append('DC Volt '+result)
        self.dc_voltlabel.setText(result)

    def on_ac_measurement_completed(self, result):
        self.textBrowser.append('AC Volt '+result)
        # res = result.split(' ')
        # # print(float(res[0]))
        # # if 0.13 < float(res[0]) < 0.15:
        # #     self.ac_voltlabel.setStyleSheet("background-color: red;")
        self.ac_voltlabel.setText(result)
        # # else:
        # #     self.ac_voltlabel.setStyleSheet("background-color: green;")
        # #     self.ac_voltlabel.setText(result)

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
                if not (3.25 < dcv_value_ < 3.35):
                    QMessageBox.information(self, "DC Voltage Out of Range", "step 1 ---- index 0.")
                    return
                
            elif self.image_index == 1:
                dcv_val = self.dc_voltlabel.text()
                dcv_value = dcv_val.split(' ')
                print('DCV 1',float(dcv_value[0]))
                dcv_value_ = float(dcv_value[0])
                if not (4.95 <= dcv_value_ <= 5.05):
                    QMessageBox.information(self, "DC Voltage Out of Range", "step 2 ---- index 1.")
                    return
            elif self.image_index == 2:
                acv_val = self.ac_voltlabel.text()
                acv_value = acv_val.split(' ')
                print('ACV 0', float(acv_value[0]))
                acv_value_ = float(acv_value[0])
                if not (acv_value_ < 0.01):
                    QMessageBox.information(self, "DC Voltage Out of Range", "step 3 ---- index 2.")
                    return
            elif self.image_index == 3:
                acv_val = self.ac_voltlabel.text()
                acv_value = acv_val.split(' ')
                print('ACV 1', float(acv_value[0]))
                acv_value_ = float(acv_value[0])
                if not (acv_value_ < 0.01):
                    QMessageBox.information(self, "DC Voltage Out of Range", "step 4 ---- index 3.")
                    return
                
            elif self.image_index == 4:
                dcv_val = self.dc_voltlabel.text()
                dcv_value = dcv_val.split(' ')
                print('DCV 1',float(dcv_value[0]))
                dcv_value_ = float(dcv_value[0])
                QMessageBox.information(self, "DC Voltage Out of Range", "step 5 ---- index 4.")
                

            elif self.image_index == 5:
                dcv_val = self.dc_voltlabel.text()
                dcv_value = dcv_val.split(' ')
                print('DCV 1',float(dcv_value[0]))
                dcv_value_ = float(dcv_value[0])
                if not (11.95 <= dcv_value_ <= 12.05):
                    QMessageBox.information(self, "DC Voltage Out of Range", "step 5 ---- index 4.")
                    return
                
            elif self.image_index == 6:
                dcv_val = self.dc_voltlabel.text()
                dcv_value = dcv_val.split(' ')
                print('DCV 1',float(dcv_value[0]))
                dcv_value_ = float(dcv_value[0])
                if not (4.95 <= dcv_value_ <= 5.05):
                    QMessageBox.information(self, "DC Voltage Out of Range", "step 6 ---- index 5.")
                    return
            elif self.image_index == 7:
                dcv_val = self.dc_voltlabel.text()
                dcv_value = dcv_val.split(' ')
                print('DCV 1',float(dcv_value[0]))
                dcv_value_ = float(dcv_value[0])
                if not (4.95 <= dcv_value_ <= 5.05):
                    QMessageBox.information(self, "DC Voltage Out of Range", "step 7 ---- index 6.")
                    return
            elif self.image_index == 8:
                dcv_val = self.dc_voltlabel.text()
                dcv_value = dcv_val.split(' ')
                print('DCV 1',float(dcv_value[0]))
                dcv_value_ = float(dcv_value[0])
                if not (4.98 <= dcv_value_ <= 5.02):
                    QMessageBox.information(self, "DC Voltage Out of Range", "step 8 ---- index 7.")
                    return

            elif self.image_index == 9:
                dcv_val = self.dc_voltlabel.text()
                dcv_value = dcv_val.split(' ')
                print('DCV 1',float(dcv_value[0]))
                dcv_value_ = float(dcv_value[0])
                if not (2.046 <= dcv_value_ <= 2.050):
                    QMessageBox.information(self, "DC Voltage Out of Range", "step 9 ---- index 8.")
                    return
            elif self.image_index == 10:
                acv_val = self.ac_voltlabel.text()
                acv_value = acv_val.split(' ')
                print('ACV 1', float(acv_value[0]))
                acv_value_ = float(acv_value[0])
                if not (acv_value_ < 0.01):
                    QMessageBox.information(self, "DC Voltage Out of Range", "step 10 ---- index 9.")
                    return
            elif self.image_index == 11:
                acv_val = self.ac_voltlabel.text()
                acv_value = acv_val.split(' ')
                print('ACV 1', float(acv_value[0]))
                acv_value_ = float(acv_value[0])
                if not (acv_value_ < 0.005):
                    QMessageBox.information(self, "DC Voltage Out of Range", "step 11 ---- index 10.")
                    return
            elif self.image_index == 12:
                acv_val = self.ac_voltlabel.text()
                acv_value = acv_val.split(' ')
                print('ACV 1', float(acv_value[0]))
                acv_value_ = float(acv_value[0])
                if not (acv_value_ < 0.005):
                    QMessageBox.information(self, "DC Voltage Out of Range", "step 12 ---- index 11.")
                    return
            elif self.image_index == 13:
                acv_val = self.ac_voltlabel.text()
                acv_value = acv_val.split(' ')
                print('ACV 1', float(acv_value[0]))
                acv_value_ = float(acv_value[0])
                if not (acv_value_ < 0.001):
                    QMessageBox.information(self, "DC Voltage Out of Range", "step 13 ---- index 12.")
                    return
            elif self.image_index == 14:
                acv_val = self.ac_voltlabel.text()
                acv_value = acv_val.split(' ')
                print('ACV 1', float(acv_value[0]))
                acv_value_ = float(acv_value[0])
                if not (acv_value_ < 0.001):
                    QMessageBox.information(self, "DC Voltage Out of Range", "step 14 ---- index 13.")
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



###################################################################################################################################################################
    def voltage_find_before_jumper(self):
        self.multimeter.write('CONF:VOLT:DC 5')
        voltage = float(self.multimeter.query('READ?'))
        time.sleep(2)
        # if 3.25 < voltage < 3.35:
        #     self.result_label.setStyleSheet("background-color: green;")
        #     self.result_label.setText('Voltage before Jumper\n\n'+str(voltage)+'V')
        # else:
        #     self.result_label.setStyleSheet("background-color: red;")
        #     self.result_label.setText('Voltage before Jumper\n\n'+str(voltage)+'V')
        #     # QMessageBox.information(self, 'Information', 'Wrong Voltage. Check Connections again.')
        #     print(voltage)
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
            # self.textBrowser.append(self.powersupply.query(self.PS_channel+':CURRent?'))
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

    def calc_voltage_before_jumper(self):
        current = float(self.powersupply.query('MEASure:CURRent? '+self.PS_channel))
        self.result_label.setVisible(True)
        if self.start_button.text() == 'STROM-I':
            self.result_label.setText('Current before Jumper\n'+str(current)+'A')
            if 0.04 <= current <= 0.06:
                self.result_label.setStyleSheet("background-color: green;")
                self.start_button.setText('SPANNUNG')
                # self.start_button.setVisible(False)
                self.current_before_jumper = current
                self.on_button_click('images_/images/R709_before_jumper.jpg')
                self.info_label.setText('Press SPANNUNG to Calculate initial VOLTAGE at R709.\n \n Calculate Voltage at the Component \n Shown in the figure.')
            else:
                self.result_label.setStyleSheet("background-color: red;")
                self.current_before_jumper = current
                self.start_button.setText('SPANNUNG')
                self.info_label.setText('Press SPANNUNG to Calculate initial VOLTAGE at R709.\n \n Calculate Voltage at the Component \n Shown in the figure.')
                self.on_button_click('images_/images/R709_before_jumper.jpg')
                self.result_label.setStyleSheet("background-color: red;")
                QMessageBox.information(self, 'Information', 'Supplying Current is either more or less. So please Swith OFF the PowerSupply, and Put back all the Euipment back.')
            # self.result_label.setVisible(False)
        elif self.start_button.text() == 'STROM':
            self.result_label.setText('Current After Jumper\n'+str(current)+'A')
            if 0.09 <= current <= 0.15:
                self.result_label.setStyleSheet("background-color: green;")
                self.current_after_jumper = current
                QMessageBox.information(self, "Information", "Now Everything is perfect. Please be care full with each and every step from here.")
                self.on_button_click('images_/images/R709.jpg')
                self.info_label.setText('\n \n Press TEST V Button to run the Voltage Tests. Be careful.')
                self.test_button.setText('TEST-V')
                self.test_button.setVisible(True)
                self.start_button.setVisible(False)
            else:
                self.result_label.setStyleSheet("background-color: red;")
                QMessageBox.information(self, 'Information', 'Supplying Current is either more or less. So please Swith OFF the PowerSupply, and Put back all the Euipment back.')       
        elif self.start_button.text() == 'TEST-I':
            self.result_label.setText('Current After FPGA\n'+str(current)+'A')
            if 0.095 <= current <= 0.155:
                self.result_label.setStyleSheet("background-color: green;")
                self.current_after_FPGA = current
                QMessageBox.information(self, "Information", "Now Everything is perfect. Please be care full with each and every step from here.")
                self.on_button_click('images_/images/R709.jpg')
                self.info_label.setText('\n \n Press TEST V Button to run the Voltage Tests. Be careful.')
                self.test_button.setText('TEST-V')
            else:
                self.result_label.setStyleSheet("background-color: red;")
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
                # self.textBrowser.append(self.powersupply.resource_name)
                self.start_button.setVisible(False)
                self.value_edit.setStyleSheet("background-color: lightyellow;")
                self.info_label.setText('Write CH1 in the Yellow Box (Highlighted)\n \n next to CH \n\n Press "ENTER"')
                self.ch_button.setVisible(True)
                self.channel_combobox.setVisible(True)
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
        self.start_button.setVisible(False)
        QMessageBox.information(self, "Important", "Careful with the Test Leads.")
        self.image_timer = QTimer(self)
        self.image_timer.timeout.connect(self.change_image)
        self.image_timer.start(3000)


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
        self.save_button.setVisible(True)

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
        try:
            config = configparser.ConfigParser()
            config.add_section('Powersupply Test')
            config.add_section('I2C Test')
            power_supply_values = {
                "Supply Voltage" : self.max_voltage,
                "Supply Current" : self.max_current,
                "current_before_jumper" : self.current_before_jumper,
                "voltage_before_jumper" : self.voltage_before_jumper,
                "current_after_jumper" : self.current_after_jumper,
                "current after FPGA" : self.current_after_FPGA,
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
        except AttributeError:
            self.textBrowser.append("Take all the values")

def main():
    app = QApplication(sys.argv)
    Window = App()
    sys.exit(app.exec_())
if __name__ == '__main__':
    main()
