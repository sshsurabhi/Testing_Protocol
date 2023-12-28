import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QTextBrowser, QLabel, QMainWindow, QAction, QMenuBar

class ClickCounterApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.click_count = 0
        self.button_click_count = 0

        # Create widgets
        self.button = QPushButton('Click me!')
        self.label = QLabel('Click count: 0')
        self.text_browser = QTextBrowser()
        self.reset_button = QPushButton('Reset Click Count')

        # Set up the layout
        layout = QVBoxLayout()
        layout.addWidget(self.button)
        layout.addWidget(self.label)
        layout.addWidget(self.text_browser)
        layout.addWidget(self.reset_button)

        # Set the layout for the central widget
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # Create menu bar and actions
        menubar = self.menuBar()
        file_menu = menubar.addMenu('File')

        start_action = QAction('Start', self)
        start_action.setShortcut('Ctrl+S')  # Set a shortcut for the action
        start_action.triggered.connect(self.resetClickCount)
        file_menu.addAction(start_action)

        # Connect the button click event to the buttonClicked method
        self.button.clicked.connect(self.buttonClicked)
        self.reset_button.clicked.connect(self.additionalButtonClicked)

        # Set up the main window
        self.setGeometry(300, 300, 300, 200)
        self.setWindowTitle('Click Counter App')
        self.show()

    def buttonClicked(self):
        self.click_count += 1

        # Check for milestones
        if self.click_count == 4:
            self.button.setText('Step 4')

        # Update label text with the current click count
        self.label.setText(f'Click count: {self.click_count}')

        # Check for additional milestones
        if self.click_count > 5:
            self.reset_button.setEnabled(True)
            if self.click_count == 6:
                self.button.setText('Button')
                self.button.clicked.connect(self.additionalButtonClicked)

        if self.click_count == 10:
            # Append text to the text browser
            self.text_browser.append('10 clicks completed')

    def additionalButtonClicked(self):
        self.button_click_count += 1
        if self.button_click_count == 2:
            self.text_browser.append('button has been clicked 2 times.')

    def resetClickCount(self):
        if self.click_count < 5:
            self.click_count = 0
        elif self.click_count >=5:
            self.text_browser.append('click the reset button 2 times')
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ClickCounterApp()
    sys.exit(app.exec_())
