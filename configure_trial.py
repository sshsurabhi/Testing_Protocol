import sys, os,configparser, random
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
class MyGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.config = None  # Initialize config attribute
        self.label_image = QLabel(self)
        self.label_image.setPixmap(QPixmap('images_/images/Welcome.jpg'))
        self.button_connect_multimeter = QLineEdit(self)
        self.button_connect_powersupply = QLineEdit(self)
        self.button_measure = QPushButton('Measure', self)
        self.info_label = QLabel('Result:', self)
        self.textBrowser = QTextBrowser(self)
        self.line_edit = QLineEdit(self)
        layout = QVBoxLayout(self)
        H_Layout1 = QHBoxLayout()
        H_Layout2 = QHBoxLayout()
        layout.addWidget(self.label_image)
        self.lower_label = QLabel('Lower', self)
        H_Layout1.addWidget(self.lower_label)
        H_Layout1.addWidget(self.button_connect_multimeter)
        self.upper_label = QLabel('Upper', self)
        H_Layout1.addWidget(self.upper_label)
        H_Layout1.addWidget(self.button_connect_powersupply)
        self.multi_label = QPushButton('Multi')
        self.add_label = QPushButton('Add')
        layout.addWidget(self.line_edit)
        H_Layout2.addWidget(self.multi_label)
        H_Layout2.addWidget(self.add_label)
        layout.addLayout(H_Layout1)
        layout.addLayout(H_Layout2)
        layout.addWidget(self.button_measure)
        layout.addWidget(self.info_label)
        layout.addWidget(self.textBrowser)
        
        self.show_input_dialog()
        self.button_measure.clicked.connect(self.connect_button)




        self.model_num_limit = float(self.config['Settings']['Model Num Limit'])
        # self.multimeter_value = float(self.button_connect_multimeter.text())
        # self.powersupply_value = float(self.button_connect_powersupply.text())

    def show_input_dialog(self):
        while True:
            name, ok = QInputDialog.getText(self, 'Enter Name', 'Enter a name for the INI file:')
            if not ok:
                break
            if not name:
                continue
            if self.ini_file_exists(name):
                reply = QMessageBox.question(self, 'File Exists', 'File with this name already exists. Do you want to overwrite it?', QMessageBox.Yes | QMessageBox.No)
                if reply == QMessageBox.Yes:
                    self.update_existing_config(name)
                    return
                elif reply == QMessageBox.No:
                    continue
            else:
                self.create_config_file(name)
                break

    def update_existing_config(self, name):
        existing_config = configparser.ConfigParser()
        existing_config.read(name + '.ini')
        random_num_limit, ok = QInputDialog.getText(self, 'Enter Model Num Limit', 'Enter Model Num Limit:')
        if not ok:
            return
        existing_config['Settings']['Model Num Limit'] = random_num_limit
        existing_config['Random Numbers'] = {}
        existing_config['Power Supply'] = {
            f"Random Number {i+1}": "" for i in range(10)
        }
        with open(name + '.ini', 'w') as configfile:
            existing_config.write(configfile)
        self.current_config_file = name + '.ini'
        self.config = existing_config
        print('Config file updated', self.current_config_file)

    def ini_file_exists(self, name):
        return name and os.path.exists(name + '.ini')

    def create_config_file(self, name):
        ini_file_path = name + '.ini'
        if os.path.exists(ini_file_path):
            reply = QMessageBox.question(self, 'File Exists', 'File with this name already exists. Do you want to overwrite it?', QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.No:
                self.update_existing_config(name)
                return
        random_num_limit, ok = QInputDialog.getText(self, 'Enter Model Num Limit', 'Enter Model Num Limit:')
        if not ok:
            return
        power_supply_values = {
            f"Random Number {i+1}": "" for i in range(10)
        }
        I2C_test_values = {
            "Add+Multi": "",
            "Add*Multi": "",
            "Sub*Multi": "",
            "Sub+Add": ""
        }
        config = configparser.ConfigParser()
        config['Settings'] = {'Model Num Limit': random_num_limit}
        config['Power Supply'] = power_supply_values
        config['I2C_Test'] = I2C_test_values
        config['Random Numbers'] = {}
        with open(ini_file_path, 'w') as configfile:
            config.write(configfile)
        self.current_config_file = ini_file_path
        self.config = config
        print('Config file created', self.current_config_file)


    def addition(self, num1, num2):
        res = num1 + num2
        print('number add',res)
        return res
    def multiply(self, num1, num2):
        res = num1 * num2
        print('number multiply', res)
        return res

    def update_line_edit_color(self,color):
        palette = QPalette()
        palette.setColor(QPalette.Base, QColor(color))
        self.line_edit.setPalette(palette)

    def connect_button(self):
        if self.button_measure.text()=='Measure':
            y = self.addition(float(self.button_connect_multimeter.text()), float(self.button_connect_powersupply.text()))
            if y < self.model_num_limit:
                self.update_line_edit_color('green')
                self.line_edit.setText(str(y))
            else:
                self.update_line_edit_color('red')
                self.line_edit.setText(str(y))
            self.button_measure.setText('Measure2')

        elif self.button_measure.text()=='Measure2':
            x = self.perform_measurement()

        elif self.button_measure.text()=='Measure3':
            z = self.multiply(float(self.button_connect_multimeter.text()), float(self.button_connect_powersupply.text()))
            if z > self.model_num_limit:
                self.line_edit.setText(str(z))
                self.update_line_edit_color('green')
            else:
                self.line_edit.setText(str(z))
                self.update_line_edit_color('red')
            self.button_measure.setText('Measure')

    def perform_measurement(self):
        try:
            while True:
                random_num = random.uniform(float(self.button_connect_multimeter.text()), float(self.button_connect_powersupply.text()))

                if random_num > self.model_num_limit:
                    self.line_edit.setText(str(random_num))
                    self.update_line_edit_color('red')
                    print('Hi')
                    QTimer.singleShot(5000, lambda: self.update_line_edit_color('white'))
                    print('Hello')
                else:
                    self.line_edit.setText(str(random_num))
                    self.update_line_edit_color('green')
                    self.config['Power Supply'][f'Random Number {1}'] = str(random_num)
                    self.button_measure.setText('Measure3')
                    with open(self.current_config_file, 'w') as configfile:
                        self.config.write(configfile)
                    break
                print('How r U')
                QTimer.singleShot(2000, lambda: self.update_line_edit_color('white'))
                print('I dont know')
                

        except ValueError:
            QMessageBox.warning(self, 'Invalid Input', 'Please enter valid numerical values for Multimeter and Power Supply.')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MyGUI()
    window.show()
    sys.exit(app.exec_())
