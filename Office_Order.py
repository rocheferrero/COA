import sqlite3
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QFont, QTextDocument, QTextCursor, QColor, QTextTableCellFormat, QBrush, QPixmap
from PyQt5.QtPrintSupport import QPrintPreviewDialog
from PyQt5.QtSql import QSqlDatabase, QSqlQuery
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QGridLayout, QTableWidget, QTableWidgetItem, QSizePolicy, QLabel, QLineEdit, \
                            QMessageBox, QComboBox, QInputDialog, QFrame, QVBoxLayout, QGroupBox, QHeaderView, QDesktopWidget
import qtawesome
import csv
import pandas as pd

class DatabaseWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.showNormal()
        # Set up the GUI layout
        layout = QGridLayout()

                # Create a new container widget for the login form
        self.login_container = QFrame(self)
        self.login_container.setFrameStyle(QFrame.Panel | QFrame.Raised)
        self.login_container.setStyleSheet("background-color: white")
        self.login_container.setFixedWidth(300)
        self.login_container.setFixedHeight(220)
        self.login_container_layout = QVBoxLayout(self.login_container)
        self.login_container_layout.setContentsMargins(20, 20, 20, 20)

        # Add the login form widgets to the container widget
        
        self.title_label = QLabel(self.login_container)
        self.title_label.setText("COA Login")
        self.title_label.setAlignment(Qt.AlignCenter)  # set alignment to center
        font = QFont()
        font.setPointSize(16)  # set font size to 16
        self.title_label.setFont(font)
        self.login_container_layout.addWidget(self.title_label)

        self.username_label = QLabel(self.login_container)
        self.username_label.setText("Username:")
        self.login_container_layout.addWidget(self.username_label)

        self.username_edit = QLineEdit(self.login_container)
        self.login_container_layout.addWidget(self.username_edit)

        self.password_label = QLabel(self.login_container)
        self.password_label.setText("Password:")
        self.login_container_layout.addWidget(self.password_label)

        self.password_edit = QLineEdit(self.login_container)
        self.password_edit.setEchoMode(QLineEdit.Password)
        self.login_container_layout.addWidget(self.password_edit)

        self.login_button = QPushButton(self.login_container)
        self.login_button.setText("Log In")
        self.login_button.clicked.connect(self.login)
        self.login_container_layout.addWidget(self.login_button)

        # Connect the returnPressed signal of the password_edit field to the login method
        self.password_edit.returnPressed.connect(self.login)

        # Center the login container in the main widget
        layout.addWidget(self.login_container, 2, 0, 3, 7, alignment=Qt.AlignCenter)

        # Hide the login form initially
        self.login_container.setVisible(True)

        # Create database connection
        db = QSqlDatabase.addDatabase("QSQLITE")
        db.setDatabaseName("./officeorder_1.db")
        db.open()

        self.setWindowTitle("COA Office Order")
        self.setWindowIcon(QIcon('logo.png'))
        
        # Create a combo box to select the table
        self.table_selector = QComboBox()
        self.table_selector.currentIndexChanged.connect(self.show_table_data)
        layout.addWidget(self.table_selector, 6, 6)
        self.table_selector.setFixedWidth(210 )
        self.table_selector.setVisible(False) 

        # Total Rows Counts
        self.row_count_label1 = QLabel(self)
        layout.addWidget(self.row_count_label1, 10, 0, alignment=Qt.AlignLeft)
        self.row_count_label1.setVisible(False)
        self.row_count_label2 = QLabel(self)
        layout.addWidget(self.row_count_label2, 10, 2, alignment=Qt.AlignCenter)
        self.row_count_label2.setVisible(False)
        self.row_count_label3 = QLabel(self)
        layout.addWidget(self.row_count_label3, 10, 3, alignment=Qt.AlignLeft)
        self.row_count_label3.setVisible(False)

        # Create a table widget to display the data
        self.table = QTableWidget()

        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.table.setStyleSheet("QTableWidget { border: 1px solid black; }")

        self.table.setVisible(False)
        
        layout.addWidget(self.table, 8, 0, 1, 7)

        # Set up the database connection
        self.conn = sqlite3.connect('./officeorder_1.db')
        self.current_table = None

        # Create a group box
        self.group_box = QGroupBox()

        # Set the border width and color
        self.group_box.setStyleSheet("background-color: white; border: 1px solid black;")

        # Create a layout for the group box
        self.group_box_layout = QGridLayout()
        self.group_box_layout.setSpacing(10) # reduce spacing between widgets

        # Label for table selection
        self.insert_label = QLabel("<b>INSERT NEW PERSONNEL</b>")
        self.insert_label.setAlignment(Qt.AlignCenter)
        self.group_box_layout.addWidget(self.insert_label, 0, 0, 1, 2, alignment=Qt.AlignHCenter)
        self.insert_label.setFixedWidth(250)
        font = QFont()
        font.setPointSize(10) # set the font size to 16
        self.insert_label.setFont(font)
        self.insert_label.setStyleSheet("border: none;")

        # Label for table selection
        self.table_label = QLabel("Select table:")
        self.group_box_layout.addWidget(self.table_label, 1, 0,)
        self.table_label.setFixedWidth(150)
        self.table_label.setStyleSheet("border: none;")

        # Option menu for table selection
        self.selected_table = QComboBox()
        self.group_box_layout.addWidget(self.selected_table, 1, 1)
        self.selected_table.setFixedWidth(300)
        self.selected_table.setVisible(False)
        self.table_label.setVisible(False)

        # Populate the combo box with table names from the database
        cursor = self.conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = [row[0] for row in cursor.fetchall()]
        self.selected_table.addItems(["Select Table"] + tables)

        # Populate the combo box with the table names
        cursor = self.conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = cursor.fetchall()
        for table in tables:
            self.table_selector.addItem(table[0])

        # Label and entry for name
        self.name_label = QLabel("Enter Name:")
        self.group_box_layout.addWidget(self.name_label, 2, 0)
        self.name_label.setStyleSheet("border: none;")
        self.name_entry = QLineEdit()
        self.name_entry.setPlaceholderText("Enter Name")
        self.group_box_layout.addWidget(self.name_entry, 2, 1)
        self.name_entry.setFixedWidth(300)
        self.name_entry.setVisible(False)
        self.name_label.setVisible(False)

        # Create query to retrieve department names from oo_2022_749 table
        query = QSqlQuery("SELECT DISTINCT department FROM oo_2022_749", db)

        # Create combo box and add department names as options
        self.department_label = QLabel("Select Department:")
        self.group_box_layout.addWidget(self.department_label, 3, 0)
        self.department_label.setStyleSheet("border: none;")
        self.department_box = QComboBox()
        self.department_box.insertItem(0, "Select Department")
        while query.next():
            department_name = query.value(0)
            self.department_box.addItem(department_name)
        self.group_box_layout.addWidget(self.department_box, 3, 1)
        self.department_box.setFixedWidth(300)
        self.department_box.setVisible(False)
        self.department_label.setVisible(False)

        # Create query to retrieve department from oo_2022_749 table
        query = QSqlQuery("SELECT DISTINCT region FROM oo_2022_749", db)

        # Create combo box and add region as options
        self.region_label = QLabel("Select Region:")
        self.group_box_layout.addWidget(self.region_label, 4, 0)
        self.region_label.setStyleSheet("border: none;")
        self.region_box = QComboBox()
        self.region_box.insertItem(0, "Select Region")
        while query.next():
            region_name = query.value(0)
            self.region_box.addItem(region_name)
        self.group_box_layout.addWidget(self.region_box, 4, 1)
        self.region_box.setFixedWidth(300)
        self.region_box.setVisible(False)
        self.region_label.setVisible(False)

        # Label for table turnover
        self.turnover_label = QLabel("Select Turnover:")
        self.group_box_layout.addWidget(self.turnover_label, 5, 0)
        self.turnover_label.setStyleSheet("border: none;")
        # Option menu for turnover selection
        turnover = ["Select Turnover","yes", "no"]
        self.selected_turnover = QComboBox()
        self.selected_turnover.addItems(turnover)
        self.group_box_layout.addWidget(self.selected_turnover, 5, 1)
        self.selected_turnover.setFixedWidth(300)
        self.selected_turnover.setVisible(False)
        self.turnover_label.setVisible(False)
        
        # Add an insert button and connect it to the insert method
        self.insert_button = QPushButton("Insert data")
        self.insert_button.clicked.connect(lambda: self.insert_data())
        self.group_box_layout.addWidget(self.insert_button, 6, 0, 1, 2,alignment=Qt.AlignLeft)
        self.insert_button.setFixedWidth(515)
        self.insert_button.setStyleSheet("""QPushButton:hover {background-color: lightgray;}""")
        self.insert_button.setVisible(False)

        # Set the layout for the group box
        self.group_box.setLayout(self.group_box_layout)

        # Add the group box to the main layout
        layout.addWidget(self.group_box, 1, 0, 6, 2)
        self.group_box.setFixedWidth(600)
        
        self.group_box.setVisible(False)

        # Create a label for the image
        self.image_label = QLabel(self)
        # Load the image file
        pixmap = QPixmap('logo.png')
        # Scale the image to a specific size
        scaled_pixmap = pixmap.scaled(250, 250, Qt.KeepAspectRatio)
        # Set the scaled pixmap to the label
        self.image_label.setPixmap(scaled_pixmap)
        # Add the label to the layout
        layout.addWidget(self.image_label, 1, 3, 6, 1)

        self.image_label.setVisible(False)

        # Create a button to print the table
        self.print_button = QPushButton()
        self.print_button.setIcon(qtawesome.icon('fa.print'))
        self.print_button.setText("Print Table")  # add text to the button                   
        self.print_button.clicked.connect(self.print_table)
        layout.addWidget(self.print_button, 5, 6, alignment=Qt.AlignRight)
        self.print_button.setFixedWidth(100)
        self.print_button.setVisible(False)

        # Create the export button with an icon
        self.export_button = QPushButton(self)
        self.export_button.setIcon(qtawesome.icon('fa.save'))
        self.export_button.setText("Export Table")  # add text to the button                   
        self.export_button.clicked.connect(lambda: self.export_table_to_csv('OfficeOrder.csv'))
        layout.addWidget(self.export_button, 5, 6, alignment=Qt.AlignLeft)
        self.export_button.setFixedWidth(100)
        self.export_button.setVisible(False)

        # Create the logout button with an icon
        self.logout_button = QPushButton(self)
        self.logout_button.setIcon(qtawesome.icon('fa.sign-out'))
        self.logout_button.setText("Log Out")  # add text to the button                   
        self.logout_button.clicked.connect(self.log_out)
        layout.addWidget(self.logout_button, 1, 6, alignment=Qt.AlignRight)
        self.logout_button.setFixedWidth(100)
        self.logout_button.setVisible(False)

        icon = QIcon("table.png")
        # Create a button to add a new table
        self.add_table_button = QPushButton(self)
        self.add_table_button.setIcon(icon)   
        self.add_table_button.setText("Add Table")  # add text to the button                           
        self.add_table_button.clicked.connect(self.add_table)
        layout.addWidget(self.add_table_button, 6, 5, alignment=Qt.AlignRight)
        self.add_table_button.setFixedWidth(100)
        self.add_table_button.setVisible(False)

        # Create a Refresh button
        self.refresh_button = QPushButton(self)
        self.refresh_button.setIcon(qtawesome.icon('fa.repeat'))
        self.refresh_button.setText("Refresh")  # add text to the button                           
        self.refresh_button.clicked.connect(self.refresh_table)
        layout.addWidget(self.refresh_button, 10, 6, alignment=Qt.AlignRight)
        self.refresh_button.setFixedWidth(210)
        self.refresh_button.setVisible(False)

        self.delete_button = QPushButton("Delete Row")
        self.delete_button.clicked.connect(lambda: self.delete_data(self.current_table))
        layout.addWidget(self.delete_button, 11, 5)
        self.delete_button.setVisible(False)

        self.search_box1 = QLineEdit()
        self.search_box1.setPlaceholderText("Search in the Table")
        layout.addWidget(self.search_box1, 11, 1, 1, 4)
        self.search_box1.setStyleSheet("QLineEdit { border: 1px solid black; }")
        self.search_box1.setVisible(False)

             # Connect search box signals to search functions
        self.search_box1.textChanged.connect(lambda text: self.filter_table(self.table, text))

                # Create a QComboBox widget to select the type of sortation
        self.sort_box = QComboBox()
        self.sort_box.addItem("Sort by name")
        self.sort_box.addItem("Sort by department")
        self.sort_box.addItem("Sort by region")
        self.sort_box.addItem("Sort by turnover")
        self.sort_box.currentIndexChanged.connect(self.sort_table)
        layout.addWidget(self.sort_box, 11, 0)
        self.sort_box.setVisible(False)

        # Add an update button and connect it to the update method
        self.update_button = QPushButton("Update")
        self.update_button.clicked.connect(lambda: self.update_data(self.current_table))
        layout.addWidget(self.update_button, 11, 6)
        self.update_button.setVisible(False)

        # Set the default username and password
        self.default_username = "admin"
        self.default_password = "admin"

        # Set the layout for the window
        self.setLayout(layout)
    
    def refresh_table(self):
        # Refresh the table by calling the show_table_data method
        self.show_table_data()
    
    def export_table_to_csv(self, filename):
        # Get the selected table name from the combo box
        table_name = self.table_selector.currentText()

        # Select all data from the selected table
        cursor = self.conn.cursor()
        cursor.execute(f'SELECT name, department, region, turnover, province FROM "{table_name}"')
        data = cursor.fetchall()

        # Open the file for writing
        with open(filename, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)

            # Write the header row
            writer.writerow(['Name', 'Department', 'Region', 'Turnover', 'Agency Assignments'])

            # Write the data rows
            for row in data:
                writer.writerow(row)

        # Display a success message
        QMessageBox.information(self, "Export", f"The data has been exported to {filename}.")
    
    def print_table(self):
        # Get the data from the visible rows of the table
        data = []
        for row in range(self.table.rowCount()):
            if not self.table.isRowHidden(row):
                row_data = []
                for column in range(self.table.columnCount()):
                    item = self.table.item(row, column)
                    if item is not None and item.text() != "None" and item.text():
                        row_data.append(item.text())
                    else:
                        row_data.append("")
                data.append(row_data) 

        # Create a QTextDocument and set its default font
        document = QTextDocument()          
        font = QFont("Arial", 9)
        document.setDefaultFont(font)

        # Create a QTextCursor to insert content into the document
        cursor = QTextCursor(document)

        # Create a QTextTable and set its properties
        table = cursor.insertTable(len(data)+3, len(data[0]))  # Add 1 row for the header and 1 row for "Office Order"
        table_format = table.format()
        table_format.setHeaderRowCount(2) # Increase the number of header rows to 2
        table_format.setAlignment(Qt.AlignHCenter)

        # Insert table name into the first row, second column
        cursor = table.cellAt(0, 1).firstCursorPosition()
        format = cursor.charFormat()
        format.setFontWeight(QFont.Bold)
        format.setForeground(QBrush(QColor("blue")))
        cursor.setCharFormat(format)
        cursor.insertText(self.table_selector.currentText())

        # Set the cell padding
        cell_format = QTextTableCellFormat()
        cell_format.setPadding(5)
        for i in range(table.rows()):
            for j in range(table.columns()):
                cell = table.cellAt(i, j)
                cell.setFormat(cell_format)

        # Insert "Office Order" into the first row
        cursor = table.cellAt(0, 0).firstCursorPosition()
        format = cursor.charFormat()
        format.setFontWeight(QFont.Bold)
        cursor.setCharFormat(format)
        cursor.insertText("Office Order:")
        
        # Insert headers into the second row and set them to bold and center align
        header_row = ["Name","Department","Region","Turnover","Agency Assignments"]  # Replace with your actual header names
        for column, cell_data in enumerate(header_row):
            cursor = table.cellAt(1, column).firstCursorPosition() # Start inserting headers at row 1
            format = cursor.charFormat()
            format.setFontWeight(QFont.Bold)
            cursor.setCharFormat(format)
            
            # Align text within the cell to center
            cell_format = cursor.blockFormat()
            cell_format.setAlignment(Qt.AlignHCenter)
            cursor.setBlockFormat(cell_format)
            
            cursor.insertText(str(cell_data))


                                # Insert data into the table
        for row, row_data in enumerate(data):
            for column, cell_data in enumerate(row_data):
                cursor = table.cellAt(row+2, column).firstCursorPosition() # Start inserting data at row 2

                if column == 3:  # Check if the column is column 3
                    cell_format = cursor.blockFormat()  # Get the format of the current cell
                    cell_format.setAlignment(Qt.AlignHCenter)  # Set text alignment to center
                    cursor.setBlockFormat(cell_format)  # Apply the updated format to the current cell
            

                if column == 4:  # Check if the column is Province column
                    provinces = cell_data.split(",")
                    for i, province in enumerate(provinces):
                        if province.strip():  # Check if the province value is not empty
                            cursor.insertText(f"{i+1}. {province.strip()}")
                            cursor.insertBlock()

                elif cell_data == "yes" or cell_data == "Yes":  # Replace "yes" with green checkmark icon
                    check_mark = chr(0x2713)  # Unicode character for checkmark
                    check_format = cursor.charFormat()  # Get the format of the current character
                    check_format.setForeground(Qt.green)  # Set color to green
                    cursor.setCharFormat(check_format)  # Apply the updated format to the current character
                    cursor.insertText(check_mark)  # Insert the green checkmark

                elif cell_data == "no" or cell_data == "No":  # Replace "no" with red x symbol
                    x_symbol = chr(0x2717)  # Unicode character for x symbol
                    x_format = cursor.charFormat()  # Get the format of the current character
                    x_format.setForeground(Qt.red)  # Set color to red
                    cursor.setCharFormat(x_format)  # Apply the updated format to the current character
                    cursor.insertText(x_symbol)  # Insert the red x symbol

                else:
                    cursor.insertText(str(cell_data))

        preview = QPrintPreviewDialog()
        preview.paintRequested.connect(lambda printer: document.print_(printer))
        preview.exec_()
    
    def sort_table(self, index):
        column_names = ["name", "department", "region", "turnover"]
        column_name = column_names[index]
        self.table.sortItems(column_names.index(column_name))
    
    def get_province(self, department):

        c = self.conn.cursor()
        c.execute("SELECT province FROM oo_2022_749 WHERE department = ?", (department,))
        result = c.fetchone()

        if result:
            return result[0]
        else:
            return None

        # Function for inserting data
    
    def login(self):
        # Check if the username and password are correct
        if self.username_edit.text() == self.default_username and self.password_edit.text() == self.default_password:
            # Enable the update button
            self.update_button.setVisible(True)
            self.insert_button.setVisible(True)
            self.selected_table.setVisible(True)
            self.name_entry.setVisible(True)
            self.department_box.setVisible(True)
            self.region_box.setVisible(True)
            self.selected_turnover.setVisible(True)
            self.table_label.setVisible(True)
            self.name_label.setVisible(True)
            self.department_box.setVisible(True)
            self.department_label.setVisible(True)
            self.region_label.setVisible(True)
            self.region_box.setVisible(True)
            self.turnover_label.setVisible(True)
            self.table.setVisible(True)
            self.search_box1.setVisible(True)
            self.sort_box.setVisible(True)
            self.table_selector.setVisible(True)
            self.add_table_button.setVisible(True)
            self.print_button.setVisible(True)
            self.delete_button.setVisible(True)
            self.export_button.setVisible(True)
            self.row_count_label1.setVisible(True)
            self.row_count_label1.setVisible(True)
            self.row_count_label2.setVisible(True)
            self.row_count_label2.setVisible(True)
            self.row_count_label3.setVisible(True)
            self.row_count_label3.setVisible(True)
            self.logout_button.setVisible(True)
            self.group_box.setVisible(True)
            self.refresh_button.setVisible(True)
            self.image_label.setVisible(True)

            self.showFullScreen()


            # Disable the login form
            self.title_label.setVisible(False)
            self.username_label.setVisible(False)
            self.username_edit.setVisible(False)
            self.password_label.setVisible(False)
            self.password_edit.setVisible(False)
            self.login_button.setVisible(False)
            self.login_container.setVisible(False)

        else:
            # Display an error message
            QMessageBox.critical(self, "Error", "Incorrect username or password")
    
    def log_out(self):
        # Display a confirmation message before logging out
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Question)
        msg_box.setText("Are you sure you want to log out?")
        msg_box.setWindowTitle("Log out")
        msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg_box.setDefaultButton(QMessageBox.No)
        button_clicked = msg_box.exec()

        if button_clicked == QMessageBox.Yes:
            # Disable all
            self.update_button.setVisible(False)
            self.insert_button.setVisible(False)
            self.selected_table.setVisible(False)
            self.name_entry.setVisible(False)
            self.department_box.setVisible(False)
            self.region_box.setVisible(False)
            self.selected_turnover.setVisible(False)
            self.table_label.setVisible(False)
            self.name_label.setVisible(False)
            self.department_box.setVisible(False)
            self.department_label.setVisible(False)
            self.region_label.setVisible(False)
            self.region_box.setVisible(False)
            self.turnover_label.setVisible(False)
            self.table.setVisible(False)
            self.search_box1.setVisible(False)
            self.sort_box.setVisible(False)
            self.table_selector.setVisible(False)
            self.add_table_button.setVisible(False)
            self.print_button.setVisible(False)
            self.delete_button.setVisible(False)
            self.export_button.setVisible(False)
            self.row_count_label1.setVisible(False)
            self.row_count_label1.setVisible(False)
            self.row_count_label2.setVisible(False)
            self.row_count_label2.setVisible(False)
            self.row_count_label3.setVisible(False)
            self.row_count_label3.setVisible(False)
            self.logout_button.setVisible(False)
            self.group_box.setVisible(False)
            self.refresh_button.setVisible(False)
            self.image_label.setVisible(False)


            # Enable the login form.
            self.title_label.setVisible(True)
            self.username_label.setVisible(True)
            self.username_edit.clear() 
            self.username_edit.setVisible(True)
            self.password_label.setVisible(True)
            self.password_edit.clear() 
            self.password_edit.setVisible(True)
            self.login_button.setVisible(True)
            self.login_container.setVisible(True)
            self.close()
    
    def filter_table(self, table, search_text):

        num_visible_rows = 0
        for row in range(table.rowCount()):
            row_hidden = True
            for column in range(table.columnCount()):
                item = table.item(row, column)
                if item and search_text.lower() in item.text().lower():
                    row_hidden = False
                    break
            table.setRowHidden(row, row_hidden)
            if not row_hidden:
                num_visible_rows += 1
                
        yes_count = 0
        none_count = 0
        for row in range(table.rowCount()):
            if not table.isRowHidden(row):
                item = table.item(row, 3)
                if item and item.text().lower() == "yes":
                    yes_count += 1
                elif not item or item.text().lower() in ["none", "no"]:
                    none_count += 1
        
        self.row_count_label1.setText(f"Personnels Total: <b>{num_visible_rows}</b>")
        self.row_count_label2.setText(f"Turned Over/Yes: <b>{yes_count}</b>")
        self.row_count_label3.setText(f"None: <b>{none_count}</b>")
    
    def show_table_data(self):
        # Get the selected table name from the combo box
        table_name = self.table_selector.currentText()

        # Select all data from the selected table
        cursor = self.conn.cursor()
        cursor.execute(f'SELECT name, department, region, turnover, province FROM "{table_name}"')
        data = cursor.fetchall()

        # Clear the table widget
        self.table.clear()

        # Set the table headers with bold font
        headers = ["Name", "Department", "Region", "Turnover", "Agency Assignments"]
        header_font = QFont()
        header_font.setBold(True)
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.horizontalHeader().setFont(header_font)
        self.table.horizontalHeader().setSectionsClickable(QHeaderView.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Fixed)

        # Add a border to the header
        self.table.horizontalHeader().setStyleSheet("""QHeaderView::section {
                border-left: 1px solid black;
                border-right: 1px solid black;
                border-bottom: 1px solid black;
                background-color: white;
            }""")

        # Set the column widths
        self.table.setColumnWidth(0, 234)
        self.table.setColumnWidth(1, 400)
        self.table.setColumnWidth(2, 100)
        self.table.setColumnWidth(3, 100)
        self.table.setColumnWidth(4, 1000)

        # Set the number of rows in the table widget to the number of rows in the data
        self.table.setRowCount(len(data))

        # Populate the table widget with data
        for row_idx, row_data in enumerate(data):
            for col_idx, cell_data in enumerate(row_data):
                if cell_data == "None":
                    cell_data = ""
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(cell_data)))

        # Store the current table name
        self.current_table = table_name

        # Get the data from the database and populate the table
        query = f"SELECT name, department, region, turnover, province FROM {table_name}"
        data = pd.read_sql_query(query, self.conn)
        self.table.setRowCount(len(data))
        self.table.setColumnCount(len(data.columns))

        yes_count = 0
        none_count = 0

        for i in range(len(data)):
            has_none = False
                        
        for i in range(len(data)):
            has_none = False
                                
            for j in range(len(data.columns)):
                item = QTableWidgetItem(str(data.iloc[i, j]))
                if item.text().lower() == "none":
                    item.setText("")
                self.table.setItem(i, j, item)
                if j == 3 and item.text().lower() in ["none", "no"]:
                    has_none = True
                elif j == 3 and item.text().lower() == "yes":
                    # Highlight the row with light blue color
                    for k in range(self.table.columnCount()):
                        self.table.item(i, k).setBackground(QColor(176, 226, 255))
                    yes_count += 1

            if has_none:
                # Highlight the row with light red color
                for k in range(self.table.columnCount()):
                    self.table.item(i, k).setBackground(QColor(255, 200, 200))
                none_count += 1

            # Highlight the whole row with light blue color
            if data.iloc[i, 3] == "yes":
                for k in range(self.table.columnCount()):
                    self.table.item(i, k).setBackground(QColor(176, 226, 255))

        # Update the row count label
        row_count = self.table.rowCount()
        self.row_count_label1.setText(f"Personnels Total: <b>{row_count}</b>")
        self.row_count_label2.setText(f"Turned Over/Yes: <b>{yes_count}</b>")
        self.row_count_label3.setText(f"None: <b>{none_count}</b>") 
    
    def insert_data(self):
        c = self.conn.cursor()
        table = self.selected_table.currentText()
        name = self.name_entry.text().strip()
        department = self.department_box.currentText()
        region = self.region_box.currentText()
        turnover = self.selected_turnover.currentText()
        province = self.get_province(department)

        # Replace "None" with an empty string
        if turnover == "None":
            turnover = ""

        if turnover != "Select Turnover" and name and department != "Select Department" and region != "Select Region" and table != "Select Table":
            c.execute(f"INSERT INTO {table} (name, department, region, turnover, province) VALUES (?, ?, ?, ?, ?)", (name, department, region, turnover, province))
            self.conn.commit()
            self.name_entry.setText("")
            self.department_box.setCurrentIndex(0)
            self.region_box.setCurrentIndex(0)
            self.selected_turnover.setCurrentIndex(0)
            QMessageBox.information(self, "Success", "New data inserted successfully")
        else:
            # display an error message or handle the missing data in some other way
            print("Please fill in all required fields.")
            QMessageBox.warning(self, "Missing fields", "Please fill in all required fields.")
    
    def update_data(self, table_name):
        # Update the data in the specified table from the table widget
        cursor = self.conn.cursor()
        for row in range(self.table.rowCount()):
            name = self.table.item(row, 0).text()
            department = self.table.item(row, 1).text()
            region = self.table.item(row, 2).text()
            turnover = self.table.item(row, 3).text()
            province = self.table.item(row, 4).text()

            # Check the name variable for invalid characters
            if any(char in name for char in ["'", '"']):
                name = name.replace("'", "''").replace('"', '""')

            cursor.execute(f'UPDATE "{table_name}" SET department=?, region=?, turnover=?, province=? WHERE name="{name}"', (department, region, turnover, province))
        self.conn.commit()
        
        QMessageBox.information(self, "Update Success", "Row Updated")  
    
    def add_table(self):
        # Prompt the user for the new table name
        table_name, ok = QInputDialog.getText(self, 'Add New Table', 'Enter the name of the new table (must start with "oo_"):')
        if ok:
            if not table_name.startswith("oo_"):
                QMessageBox.critical(self, "Invalid Table Name", "Table name must start with 'oo_'")
                return
            # Create a new table with the same columns as oo_2022_749 plus id and name
            cursor = self.conn.cursor()
            cursor.execute(f'CREATE TABLE "{table_name}" AS SELECT id, name, * FROM oo_2022_749 WHERE 1=0')
            
            # Insert the id and name values into the new table
            cursor.execute(f'INSERT INTO "{table_name}" (id, name) SELECT id, name FROM oo_2022_749')
            self.conn.commit()

            # Add the new table name to the combo box
            self.table_selector.addItem(table_name)
            self.selected_table.addItem(table_name)
    
    def delete_data(self, table):
        # Get the selected row(s)
        selected_rows = self.table.selectionModel().selectedRows()
            
        # Make sure at least one row is selected
        if len(selected_rows) == 0:
            QMessageBox.warning(self, "Warning", "Please select at least one row to delete.")
            return
            
        # Get the IDs of the selected row(s)
        ids = [self.table.item(row.row(), 0).text() for row in selected_rows]
        
        # Ask user for confirmation before deleting
        confirm = QMessageBox.question(self, "Delete rows", f"Are you sure you want to delete row?", QMessageBox.Yes | QMessageBox.No)
        if confirm == QMessageBox.No:
            return
        
        # Delete the selected row(s) from the database
        c = self.conn.cursor()
        c.execute(f"DELETE FROM {table} WHERE name IN ({','.join('?' for _ in ids)})", ids)
        self.conn.commit()
            
        # Delete the selected row(s) from the QTableWidget object
        for row in reversed(sorted(selected_rows, key=lambda x: x.row())):
            self.table.removeRow(row.row())
        QMessageBox.information(self, "Delete Success", "Row Deleted")

if __name__ == '__main__':
    # Create the application and window
    app = QApplication([])

    app.setWindowIcon(QIcon('logo.png'))

    window = DatabaseWindow()
    window.show()

    # Start the event loop
    app.exec_()
