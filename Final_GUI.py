import os, configparser, datetime, sys, time, openpyxl, serial
import pyvisa as visa
from PyQt5 import uic
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *


class MultimeterThread(QThread):
    connected = pyqtSignal()  # Add a signal

    def __init__(self, app):
        super().__init__()
        self.app = app

    def run(self):
        while self.app.multimeter is None:
            self.app.connect_multimeter()
            time.sleep(1)
        if self.app.multimeter:
            while self.app.powersupply is None:    
                self.app.connect_powersupply()
                # if self.app.powersupply:
                self.connected.emit()

class App(QMainWindow):
    def __init__(self):
        super(App, self).__init__()
        uic.loadUi("UI/New_Final_GUI.ui", self)
        self.setFixedSize(self.size())
        self.setStatusTip('Moewe Optik Gmbh. 2023')
        self.show()

        self.rm = visa.ResourceManager()
        self.multimeter = None
        self.powersupply = None
        self.on_button_click("Images/PP1.jpg")
        self.info_label.setText("Willkommen bei den Tests..\n\nDrücken Sie die Taste 'START'.")

        self.multimeter_thread = MultimeterThread(self)
        self.multimeter_thread.connected.connect(self.multimeter_connected)

        self.start_button.clicked.connect(self.connect_)

        self.button_click = 0
        self.click_event = 0
        self.show_input_dialog()

        self.actionExit.triggered.connect(self.close)



    def on_button_click(self, file_path):
        if file_path:
            pixmap = QPixmap(file_path)
            self.image_label.setPixmap(pixmap)
            self.image_label.setScaledContents(True)
            self.image_label.setFixedSize(pixmap.width(), pixmap.height())

    def update_line_edit_color(self, color):
        palette = QPalette()
        palette.setColor(QPalette.Base, QColor(color))
        self.result_label.setPalette(palette)

    def show_good_message(self, message):
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Question)
        msgBox.setText(message)
        msgBox.setWindowTitle("Congratulations!")
        self.title_label.setText('Powersupply Test')
        msgBox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        return msgBox.exec_()

    def connect_(self):
        # if self.start_button.text() == 'TEST' and self.click_event == 0:
        #     self.on_button_click("Images/board_on_mat_.jpg")
        #     self.info_label.setText("Legen Sie die Platine auf die ESD-Matte\n\n(siehe Abbildung rechts).\n\nÜberprüfen Sie die gesamte Umgebung anhand der Abbildung.\n\nPrüfen Sie alle Anschlüsse.\n\n Drücken Sie 'STEP1'.")
            
        # elif self.start_button.text() == 'TEST' and self.click_event == 1:
        #     self.on_button_click('Images/PP4_.jpg')
        #     self.info_label.setText('Überprüfen Sie alle "4" Schrauben der Platine (siehe Abbildung).\n\nMontieren Sie alle 4 Schrauben (4x M2,5x5 Torx)\n\nDrücken Sie "STEP2".')
        # elif self.start_button.text() == 'TEST' and self.click_event == 2:
        #     self.on_button_click("Images/board_on_mat_.jpg")
        #     self.info_label.setText('Nachdem Sie die 4 Schrauben angebracht haben,\n\n legen Sie die Platine wieder auf die ESD-Matte und\n\n drücken Sie dann "STEP3".')
        # elif self.start_button.text() == 'TEST' and self.click_event == 3:
        #     self.on_button_click("Images/board_with_cabels.jpg")
        #     self.info_label.setText('Schließen Sie die Stromkabel an die Platine an (siehe Abbildung).\n\n\nDrücken Sie STEP4.')

        # elif self.start_button.text() == 'TEST' and self.click_event == 4:
        #     self.on_button_click('Images/next.jpg')
        #     reply = self.show_good_message('Prüfen Sie Ihre gesamte Umgebung korrekt und sorgfältig. Wir können sie später nicht mehr ändern.')
        #     if reply == QMessageBox.Yes:
        #         self.start_button.setText('ON')
        #         self.on_button_click('Images/On_Devices.jpg')
        #         self.info_label.setText('')
        #     else:
        #         self.on_button_click('Images/PP1.jpg')
        #         self.click_event = 0
        #         # self.start_button.setText("START")
        #         self.start_button.setText("TEST")
        if self.start_button.text() == 'ON':
            self.start_button.setEnabled(False)
            self.multimeter_thread.start()
        elif self.start_button.text() == 'POWER ON':
            self.connect_power()

    def connect_power(self):
        try:
            cur_config = configparser.ConfigParser()
            cur_config.read_file(self.current_config_file)
            print(cur_config.read_file(self.current_config_file))
            self.PS_channel = cur_config.get('Power Supply', 'powersupply_channel')

        except ValueError as e:
            print('channel', e, self.PS_channel)
        # self.powersupply.write('OUTPut '+self.PS_channel+',ON')

        




    def multimeter_connected(self):
        if self.powersupply:
            self.start_button.setEnabled(True)
            self.start_button.setText("POWER ON")
            # try:
            #     self.PS_channel = self.channel_combobox.currentText()
            #     self.powersupply.write(self.PS_channel+':VOLTage 30')
            #     self.powersupply.write(self.PS_channel+':CURRent 0.5')
            #     self.start_button.setText("POWER ON")
            # except visa.errors.VisaIOError as e:
            #     self.textBrowser.append(str(e))

    def connect_multimeter(self):
        try:
            self.multimeter = self.rm.open_resource('TCPIP0::192.168.222.207::INSTR')
            self.textBrowser.append(self.multimeter.query('*IDN?'))
        except visa.errors.VisaIOError:
            self.textBrowser.append('Multimeter has not been presented')

    def connect_powersupply(self):
        try:
            self.powersupply = self.rm.open_resource('TCPIP0::192.168.222.141::INSTR')
            self.textBrowser.append(self.powersupply.query('*IDN?'))
            self.channel_combobox.setVisible(True)
        except visa.errors.VisaIOError:
            self.textBrowser.append("PowerSupply is not present at the given IP Address.")



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
            "powersupply_channel": "",
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
            self.current_config_file = configfile
        QMessageBox.information(self, 'Success', 'INI File created successfully.')


def main():
    app = QApplication(sys.argv)
    Window = App()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
