import configparser, datetime, os, sys, time, openpyxl, serial
import pyvisa as visa
from PyQt5 import uic
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *


class WorkerThread(QThread):
    measurement_finished = pyqtSignal()

    def run(self):
        # Simulate the continuous measurement
        for i in range(5):
            print(f"Measurement {i + 1}")
            time.sleep(1)

        self.measurement_finished.emit()



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
        fname, _ = QFileDialog.getOpenFileName(self, 'Open file', '/home', 'INI Files (*.ini);;All Files (*)')

        if fname:
            self.loadConfigFile(fname)

    def loadConfigFile(self, filename):
        config = configparser.ConfigParser()
        config.read(filename)

        # Define a list to store missing attributes
        missing_attributes = []

        # Read values from the config file
        for section in config.sections():
            for key, value in config.items(section):
                if not value.strip():  # Check if the value is empty or whitespace
                    missing_attributes.append(f"{key} in section {section}")

        # If there are missing attributes, prompt the user
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

    def show_good_message(self, message):
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Question)
        msgBox.setText(message)
        msgBox.setWindowTitle("Congratulations!")
        msgBox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)        
        return msgBox.exec_()  
    

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
            self.start_button.setText('MULTI ON')
            self.on_button_click('images_/images/PP9.jpg')
            self.show_input_dialog()            
        elif self.start_button.text()=='MULTI ON':
            self.connect_multimeter()

        elif self.start_button.text()=='POWER ON':
            self.connect_powersupply()


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
        # Check if the file already exists
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
def main():
    app = QApplication(sys.argv)
    Window = App()
    sys.exit(app.exec_())
if __name__ == '__main__':
    main()