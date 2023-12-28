import configparser, datetime, os, sys, time, openpyxl, serial
import pyvisa as visa
from PyQt5 import uic
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        main_layout = QVBoxLayout()
        self.setFixedSize(1423, 840)
        self.title_label = QLabel("TEST MODULE")
        self.title_label.setStyleSheet("font: bold 19pt;")
        main_layout.addWidget(self.title_label)
        self.title_label.setAlignment(Qt.AlignCenter)

        self.time_label = QLabel()
        self.time_label.setStyleSheet("font: bold 11pt;")
        self.title_label.setAlignment(Qt.AlignCenter)
        main_H_layout = QHBoxLayout()#main_ layout 

        layout_V1 = QVBoxLayout() #visible box 1
        
        Layout_H1 = QHBoxLayout()
        self.start_button = QPushButton("START")
        self.start_button.setStyleSheet("font: bold 12pt; QPushButton { qproperty-translatable: true; }")
        self.start_button.setFixedSize(231, 121)
        self.channel_label = QLabel("Channel")
        self.channel_label.setFixedWidth(41)
        self.channel_combobox = QComboBox()
        self.test_button = QPushButton("Test V")
        self.test_button.setStyleSheet("font: bold 12pt; QPushButton { qproperty-translatable: true; }")
        self.test_button.setFixedSize(231,121)
        Layout_H1.addWidget(self.start_button)
        Layout_H1.addWidget(self.channel_label)
        Layout_H1.addWidget(self.channel_combobox)
        Layout_H1.addWidget(self.test_button)
        layout_V1.addLayout(Layout_H1)
        self.info_label = QLabel()
        self.info_label.setStyleSheet("font-family: Cambria; font-weight: bold; font-size: 15pt; border: 1px solid black; padding: 5px;")
        self.info_label.setAlignment(Qt.AlignCenter | Qt.AlignVCenter)
        self.info_label.setFixedSize(641,331)
        layout_V1.addWidget(self.info_label)
        Layout_H2 = QHBoxLayout()
        Layout_V3 = QVBoxLayout()
        Layout_H3 = QHBoxLayout()
        self.port_label = QLabel("Port")
        self.port_box = QComboBox()
        self.port_box.setFixedWidth(150)
        self.connect_button = QPushButton("Connect")
        self.connect_button.setFixedWidth(71)
        Layout_H3.addWidget(self.port_label)
        self.port_label.setFixedWidth(71)
        Layout_H3.addWidget(self.port_box)
        Layout_H3.addWidget(self.connect_button)
        Layout_H4 = QHBoxLayout()
        self.baudrate_label = QLabel("Baudrate")
        self.baudrate_label.setFixedWidth(71)
        self.baudrate_box = QComboBox()
        self.baudrate_box.setFixedWidth(150)
        self.refresh_button = QPushButton("Refresh")
        self.refresh_button.setFixedWidth(71)
        Layout_H4.addWidget(self.baudrate_label)
        Layout_H4.addWidget(self.baudrate_box)
        Layout_H4.addWidget(self.refresh_button)
        Layout_V3.addLayout(Layout_H3)
        Layout_V3.addLayout(Layout_H4)
        self.result_label = QLineEdit()
        self.result_label.setFixedSize(332,51)
        Layout_H2.addLayout(Layout_V3)
        layout_V1.addLayout(Layout_H2)
        Layout_H2.addWidget(self.result_label)


        self.textBrowser = QTextBrowser()
        self.textBrowser.setFixedSize(641,161)
        layout_V1.addWidget(self.textBrowser)

        Layout_H5 = QHBoxLayout()
        self.id_edit = QLineEdit()
        self.id_edit.setFixedWidth(151)
        self.Final_result_box = QLineEdit()
        self.Final_result_box.setFixedWidth(481)
        Layout_H5.addWidget(self.id_edit)
        Layout_H5.addWidget(self.Final_result_box)
        layout_V1.addLayout(Layout_H5)
        layout_V2 = QVBoxLayout() # visible box 2
        self.image_label = QLabel()
        # self.image_label.setStyleSheet("border: 1px solid black; padding: 5px;")
        frame_style = QFrame.Box | QFrame.Raised  # Combine Panel and Raised styles
        self.image_label.setFrameStyle(frame_style)
        self.image_label.setLineWidth(4)
        self.image_label.setFixedSize(761,751)
        layout_V2.addWidget(self.image_label)
        main_H_layout.addLayout(layout_V1)
        main_H_layout.addLayout(layout_V2)
        main_layout.addLayout(main_H_layout)
        self.central_widget = QWidget(self)
        self.central_widget.setLayout(main_layout)
        self.setCentralWidget(self.central_widget)

        menubar = self.menuBar()
        file_menu = menubar.addMenu('File')
        edit_menu = menubar.addMenu('Edit')
        tools_menu = menubar.addMenu('Tools')
        help_menu = menubar.addMenu('Help')

        exit_action = QAction('Exit', self)
        exit_action.triggered.connect(self.close)  # Close the application when 'Exit' is triggered
        file_menu.addAction(exit_action)
################################################################################################################################################################

        self.rm = visa.ResourceManager()
        self.multimeter = None
        self.powersupply = None
        self.on_button_click("Images/PP1.jpg")
        self.info_label.setText("Willkommen bei den Tests..\n\nDrücken Sie die Taste 'START'.")
        self.start_button.clicked.connect(self.connect)




        self.click_event = 0

    def on_button_click(self, file_path):
        if file_path:
            pixmap = QPixmap(file_path)
            self.image_label.setPixmap(pixmap)
            self.image_label.setScaledContents(True)
            self.image_label.setFixedSize(pixmap.width(), pixmap.height())



    def update_line_edit_color(self,color):
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

    def connect(self):
        if self.start_button.text() == 'START':
            self.on_button_click("Images/board_on_mat_.jpg")
            self.info_label.setText("Legen Sie die Platine auf die ESD-Matte\n\n(siehe Abbildung rechts).\n\nÜberprüfen Sie die gesamte Umgebung anhand der Abbildung.\n\nPrüfen Sie alle Anschlüsse.\n\n Drücken Sie 'STEP1'.")
            self.start_button.setText("Next")
        elif self.start_button.text() == 'Next' and self.click_event==1:
            self.on_button_click('Images/PP4_.jpg')
            self.info_label.setText('Überprüfen Sie alle "4" Schrauben der Platine (siehe Abbildung).\n\nMontieren Sie alle 4 Schrauben (4x M2,5x5 Torx)\n\nDrücken Sie "STEP2".')
        elif self.click_event==2:
            self.on_button_click("Images/board_on_mat_.jpg")
            self.info_label.setText('Nachdem Sie die 4 Schrauben angebracht haben,\n\n legen Sie die Platine wieder auf die ESD-Matte und\n\n drücken Sie dann "STEP3".')
        elif self.click_event==3:
            self.on_button_click("Images/board_with_cabels.jpg")
            self.info_label.setText('Schließen Sie die Stromkabel an die Platine an (siehe Abbildung).\n\n\nDrücken Sie STEP4.')
        elif self.click_event == 4:
            self.on_button_click('Images/next.jpg')
            reply = self.show_good_message('Prüfen Sie Ihre gesamte Umgebung korrekt und sorgfältig. Wir können sie später nicht mehr ändern.')
            if reply == QMessageBox.Yes:                    
                self.start_button.setText('ON')
                self.on_button_click('Images/On_Devices.jpg')
                self.info_label.setText('Drücken Sie den "Power ON"-Schalter am \n\n Netzteil und auch am Multimeter. Siehe die Knöpfe auf der \n\n nebenstehenden Abbildung. Warten Sie 10 bis 12 Sekunden,\n\n num diese Geräte einzuschalten. Sie können die\n\nTaste "MULTI ON" später sehen. Drücken Sie die Taste "NEXT".')                    
            else:
                self.on_button_click('images_/icons/next_1.jpg')
                self.click_event = 0
                self.start_button.setText("START")

        

        self.click_event += 1


    def connect_multimeter(self):        
        try:
            self.multimeter = self.rm.open_resource('TCPIP0::192.168.222.207::INSTR')
            self.textBrowser.append(self.multimeter.query('*IDN?'))
            self.start_button.setText('POWER ON')
            self.info_label.setText('Press POWER ON button.\n \n It connects the powersupply...!' )
        except visa.errors.VisaIOError:
            self.textBrowser.append('Multimeter has not been presented')
    def connect_powersupply(self):
        try:
            self.powersupply = self.rm.open_resource('TCPIP0::192.168.222.141::INSTR')
            self.textBrowser.append(self.powersupply.query('*IDN?'))
            self.info_label.setText('Write CH1 in the Yellow Box (Highlighted)\n \n next to CH \n\n Press "ENTER"')
            self.channel_combobox.setVisible(True)
        except visa.errors.VisaIOError:
            QMessageBox.information(self, "PowerSupply Connection", "PowerSupply is not present at the given IP Address.")
    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
