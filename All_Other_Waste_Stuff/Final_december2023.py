import configparser, datetime, os, sys, time, openpyxl, serial
import pyvisa as visa
from PyQt5 import uic
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *


class App(QMainWindow):
    def __init__(self):
        super(App, self).__init__()
        uic.loadUi("UI/Test_App.ui", self)
        self.setWindowIcon(QIcon('images_/icons/Moewe.jpg'))
        self.setFixedSize(self.size())
        self.setStatusTip('Moewe Optik Gmbh. 2023')
        self.show()


        self.on_button_click('images_/images/PP1.jpg')
        self.info_label.setText('Welcome \n \n Drücken Sie die Taste START.')
        self.start_button.clicked.connect(self.connect)
        self.rm = visa.ResourceManager()
        self.multimeter = None
        self.powersupply = None
        self.actionLoad.triggered.connect(self.showDialog)


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
########################################################################################################################
    def show_good_message(self, message):
        self.timer1 = QTimer()
        self.timer1.timeout.connect(self.enable_button)
        self.timer1.start(5000)
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Question)
        msgBox.setText(message)
        msgBox.setWindowTitle("Congratulations!")
        msgBox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)        
        return msgBox.exec_()
    
    def enable_button(self):
        self.timer1.stop()
        self.start_button.setEnabled(True)
########################################################################################################################
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
                self.start_button.setText('SPANNUNG')
                self.result_label.setText('Voltage before Jumper\n\n'+str(self.voltage_before_jumper)+'V')
                #QMessageBox.information(self, 'Information', 'Die Spannung zwischen GND und R709 liegt zwischen 3,25 und 3,35. Dies ist ein guter Wert. Fahren Sie fort, indem Sie die Taste NEXT drücken.')
                
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
                self.textBrowser.append('Powersupply has not been presented.')
        else:
            self.powersupply.close()
            self.powersupply = None
            self.PS_button.setText('PS ON')
            self.textBrowser.setText('Netzteil Disconnected')

    def voltage_find_before_jumper(self):
        self.multimeter.write('CONF:VOLT:DC 5')
        voltage = float(self.multimeter.query('READ?'))
        time.sleep(2)
        return voltage
#########################################################################################################################
    def process_completed(self):
        QMessageBox.information(self, "Process Completed", "Process has been completed.")
        self.save_button.setVisible(True)

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
#########################################################################################################################
    def show_input_dialog(self):
        while True:
            name, ok = QInputDialog.getText(self, 'Enter Name', 'Enter a name for the INI file:')
            if not ok:
                break
            if not name:
                continue 
            if self.ini_file_exists(name):
                reply = QMessageBox.question(self, 'File Exists', 'File with this name already exists. Choose a new name', QMessageBox.Yes | QMessageBox.No)
                if reply == QMessageBox.No:
                    self.start_button.setText('START')
                    self.start_button.setEnabled(True)
                    break
                elif reply == QMessageBox.Yes:
                    continue
            else:
                self.create_config_file(name)
                break
    def ini_file_exists(self, name):
        return os.path.exists(name + '.ini')
    def create_config_file(self, name):
        ini_file_path = name + '.ini'
        if os.path.exists(ini_file_path):
            reply = QMessageBox.question(self, 'File Exists', 'File with this name already exists. Do you want to overwrite it?', QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.No:
                return
        model_num, ok = QInputDialog.getText(self, 'Enter Model Num', 'Enter Model Num:')
        if not ok:
            return
        version_num, ok = QInputDialog.getText(self, 'Enter Version Num', 'Enter Version Num:')
        if not ok:
            return
        serial_num, ok = QInputDialog.getText(self, 'Enter Serial Num', 'Enter Serial Num:')
        if not ok:
            return
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
        config = configparser.ConfigParser()
        config['Settings'] = {'Model Num': model_num,
                              'Version Num': version_num,
                              'Serial Num': serial_num}
        config['Power Supply'] = power_supply_values
        config['I2C_Test'] = I2C_test_values
        with open(ini_file_path, 'w') as configfile:
            config.write(configfile)
        QMessageBox.information(self, 'Success', 'INI File created successfully.')
##################################################################################################################################
def main():
    app = QApplication(sys.argv)
    Window = App()
    sys.exit(app.exec_())
if __name__ == '__main__':
    main()