import os
import re
import json
from PyQt5.QtWidgets import QWidget, QTableWidget, QTableWidgetItem, QVBoxLayout, QHBoxLayout, QPushButton, QHeaderView, QFrame, QMainWindow, QLabel, QAbstractItemView
from PyQt5.QtCore import Qt, pyqtSignal, QThread, QItemSelectionModel

from contact_sync import ContactSync


class SyncStatusWindow(QMainWindow):
    """Window to display the sync process status."""

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        '''This method sets up the user interface. It creates the widgets and layouts and adds them to the main window.'''
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)
        layout = QVBoxLayout(self.central_widget)

        self.status_label = QLabel("Starting sync...")
        layout.addWidget(self.status_label)

        self.setWindowTitle('Sync Status')
        self.resize(400, 200)

    def update_status(self, message):
        '''This method updates the status label with the given message.'''
        self.status_label.setText(message)


class SyncThread(QThread):
    """Thread to run the sync process."""
    update_signal = pyqtSignal(str)

    def __init__(self, sync_function, contacts, users, *args, **kwargs):
        '''This method initializes the thread with the sync function, contacts, and users. It also stores any additional arguments and keyword arguments.'''
        super().__init__()
        self.sync_function = sync_function
        self.contacts = contacts
        self.users = users
        self.args = args
        self.kwargs = kwargs

    def run(self):
        '''This method is called when the thread is started. It calls the sync function and emits a signal to update the status.'''
        self.update_signal.emit(f"Starting sync for {len(self.contacts)} contacts and {
                                len(self.users)} users...")  # Emit a signal to update the status
        self.sync_function(self.contacts, self.users)  # Call the sync function
        # Emit a signal to update the status
        self.update_signal.emit("Sync complete.")


class UserSyncGUI(QWidget):
    '''This class represents the main window of the application. It is a subclass of QWidget and is used to display the user interface.'''

    def __init__(self, csv_file_path, client_id, client_secret, tenant_id):
        super().__init__()
        self.csv_file_path = csv_file_path
        self.checked_users_file = 'checked_users.json'
        self.checked_users = []
        self.contact_sync = ContactSync(client_id, client_secret, tenant_id)
        self.users = self.contact_sync.filter_users()
        self.contacts = self.contact_sync.read_csv_file(csv_file_path)
        self.load_checked_states()
        self.initUI()

    def initUI(self):
        '''This method sets up the user interface. It creates the widgets and layouts and adds them to the main window.'''
        self.layout = QHBoxLayout(self)
        self.contact_table_layout = QVBoxLayout()
        # Contacts Header
        self.contacts_header = QLabel("Contacts to Sync")
        self.contacts_header.setAlignment(Qt.AlignCenter)
        self.contact_table_layout.addWidget(self.contacts_header)

        self.contact_table = QTableWidget(self)
        self.contact_table.setColumnCount(8)
        self.contact_table.setHorizontalHeaderLabels([
            'Given Name', 'Surname', 'Email', 'Business Phone', 'Mobile', 'Department', 'Job Title', 'Office Location'])
        self.contact_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.contact_table.verticalHeader().setVisible(False)
        self.contact_table.setAlternatingRowColors(True)
        self.contact_table.setStyleSheet(
            "alternate-background-color: lightgrey; background-color: white;")
        self.contact_table.setSortingEnabled(True)
        self.populate_contact_table()
        self.contact_table.setEditTriggers(QTableWidget.AllEditTriggers)
        # Create a button for adding a new contact
        self.save_list_btn = QPushButton('Save Contact List', self)
        # Connect the button to the add_contact method
        self.save_list_btn.clicked.connect(self.save_contact_list)

        self.update_contact_list_btn = QPushButton('Update Contact List', self)
        self.update_contact_list_btn.clicked.connect(self.update_contact_list)
        self.contact_table_layout.addWidget(self.contact_table)
        self.refresh_list_btn = QPushButton('Refresh List', self)
        self.refresh_list_btn.clicked.connect(self.refresh_contact_list)
        self.contact_table_layout.addWidget(self.refresh_list_btn)
        # Create a new layout for buttons
        self.contact_table_buttons_layout = QHBoxLayout()
        self.contact_table_buttons_layout.addWidget(
            self.update_contact_list_btn)
        self.contact_table_buttons_layout.addWidget(self.save_list_btn)
        self.contact_table_layout.addLayout(self.contact_table_buttons_layout)

        self.layout.addLayout(self.contact_table_layout)
        self.user_table_layout = QVBoxLayout()
        # Users Header
        self.users_header = QLabel("User List")
        self.users_header.setAlignment(Qt.AlignCenter)
        self.user_table_layout.addWidget(self.users_header)

        self.user_table = QTableWidget(self)
        self.user_table.setColumnCount(9)
        self.user_table.setHorizontalHeaderLabels(
            ['Check', 'Given Name', 'Surname', 'Email', 'Business Phone', 'Mobile', 'Department', 'Job Title', 'Office Location'])
        self.user_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.user_table.verticalHeader().setVisible(False)
        self.user_table.setAlternatingRowColors(True)
        self.user_table.setStyleSheet(
            "alternate-background-color: lightgrey; background-color: white;")
        self.user_table.setSortingEnabled(True)
        self.populate_user_table()

        self.uncheck_all_btn_user = QPushButton('Uncheck All', self)
        self.check_all_btn_user = QPushButton('Check All', self)
        self.uncheck_all_btn_user.clicked.connect(self.uncheck_all_user_table)
        self.check_all_btn_user.clicked.connect(self.check_all_user_table)

        self.sync_contacts_btn = QPushButton('Sync Contacts', self)
        self.sync_contacts_btn.clicked.connect(self.start_sync_process)

        self.user_table_layout.addWidget(self.user_table)
        self.user_table_buttons_layout = QHBoxLayout()
        self.user_table_buttons_layout.addWidget(self.check_all_btn_user)
        self.user_table_buttons_layout.addWidget(self.uncheck_all_btn_user)
        self.user_table_layout.addLayout(self.user_table_buttons_layout)
        self.user_table_layout.addWidget(self.sync_contacts_btn)
        self.user_table.setSelectionMode(QAbstractItemView.MultiSelection)
        self.user_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.layout.addLayout(self.user_table_layout)

        divider = QFrame()
        divider.setFrameShape(QFrame.VLine)
        divider.setFrameShadow(QFrame.Sunken)
        self.layout.addWidget(divider)

        self.setLayout(self.layout)
        self.setWindowTitle('Contact Sync')
        self.setMinimumWidth(2100)
        self.setMinimumHeight(800)

    def populate_contact_table(self):
        '''This method populates the contact table with the contacts from the CSV file.'''
        self.contact_table.setRowCount(len(self.contacts))

        for i, contact in enumerate(self.contacts):
            self.contact_table.setItem(i, 0, QTableWidgetItem(
                str(contact.get('givenName', 'NONE'))))
            self.contact_table.setItem(i, 1, QTableWidgetItem(
                str(contact.get('surname', 'NONE'))))
            self.contact_table.setItem(i, 2, QTableWidgetItem(
                str(contact.get('mail', 'NONE'))))
            self.contact_table.setItem(i, 3, QTableWidgetItem(
                str(contact.get('businessPhones', 'NONE'))))
            self.contact_table.setItem(i, 4, QTableWidgetItem(
                str(contact.get('mobilePhone', 'NONE'))))
            self.contact_table.setItem(i, 5, QTableWidgetItem(
                str(contact.get('department', 'NONE'))))
            self.contact_table.setItem(i, 6, QTableWidgetItem(
                str(contact.get('jobTitle', 'NONE'))))
            self.contact_table.setItem(i, 7, QTableWidgetItem(
                str(contact.get('officeLocation', 'NONE'))))

    def populate_user_table(self):
        '''This method populates the user table with the users from the Graph API.'''
        # Set the number of rows in the user table to the number of users
        self.user_table.setRowCount(len(self.users))

        # Iterate over the users
        for i, user in enumerate(self.users):
            # Create a checkbox item for each user
            chkBoxItem = QTableWidgetItem()

            # Set the flags of the checkbox item to be user checkable and enabled
            chkBoxItem.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)

            # If the user's ID is in the list of checked users, set the checkbox to checked; otherwise, set it to unchecked
            chkBoxItem.setCheckState(Qt.Checked if user.get(
                'id') in self.checked_users else Qt.Unchecked)

            # Add the checkbox item to the first column of the user table
            self.user_table.setItem(i, 0, chkBoxItem)

            # For each user, set the item in each column of the user table to the corresponding attribute of the user
            # If the attribute does not exist for the user, set the item to 'NONE'
            self.user_table.setItem(i, 1, QTableWidgetItem(
                str(user.get('givenName', 'NONE'))))
            self.user_table.setItem(i, 2, QTableWidgetItem(
                str(user.get('surname', 'NONE'))))
            self.user_table.setItem(i, 3, QTableWidgetItem(
                str(user.get('mail', 'NONE'))))
            self.user_table.setItem(i, 4, QTableWidgetItem(
                str(user.get('businessPhones', 'NONE'))))
            self.user_table.setItem(i, 5, QTableWidgetItem(
                str(user.get('mobilePhone', 'NONE'))))
            self.user_table.setItem(i, 6, QTableWidgetItem(
                str(user.get('department', 'NONE'))))
            self.user_table.setItem(i, 7, QTableWidgetItem(
                str(user.get('jobTitle', 'NONE'))))
            self.user_table.setItem(i, 8, QTableWidgetItem(
                str(user.get('officeLocation', 'NONE'))))

            # add the user id as a hidden item
            self.user_table.setItem(i, 9, QTableWidgetItem(
                str(user.get('id', 'NONE'))))

    def closeEvent(self, event):
        '''This method is called when the window is about to close. It saves the checked states of the users to a file and accepts the close event.'''
        # Save the checked states of the users
        self.save_checked_states()
        # Accept the close event
        event.accept()

    def save_checked_states(self):
        '''This method saves the checked states of the users to a file.'''
        # Create a list of the IDs of the users whose checkboxes are checked
        checked_users = [self.users[i].get('id') for i in range(
            self.user_table.rowCount()) if self.user_table.item(i, 0).checkState() == Qt.Checked]

        # Open the file for writing with explicit encoding
        with open(self.checked_users_file, 'w', encoding='utf-8') as f:
            # Dump the list of checked user IDs to the file in JSON format
            json.dump(checked_users, f)

    def load_checked_states(self):
        '''This method loads the checked states of the users from a file.'''
        # Initialize the list of checked user IDs
        self.checked_users = []

        # Check if the file exists and is not empty
        if os.path.exists(self.checked_users_file) and os.path.getsize(self.checked_users_file) > 0:
            try:
                # Open the file for reading
                with open(self.checked_users_file, 'r', encoding='utf-8') as f:
                    # Load the list of checked user IDs from the file
                    self.checked_users = json.load(f)
            except json.JSONDecodeError as e:
                # Print an error message and exit the method if the file could not be read
                print(f"Error loading checked states: {e}")
                return

    def check_all_user_table(self):
        '''This method checks all the checkboxes in the user table'''
        # Iterate over the rows in the user table
        for i in range(self.user_table.rowCount()):
            # Set the check state of the checkbox in the first column of the current row to checked
            self.user_table.item(i, 0).setCheckState(Qt.Checked)

    def uncheck_all_user_table(self):
        '''This method unchecks all the checkboxes in the user table'''

        # Iterate over the rows in the user table
        for i in range(self.user_table.rowCount()):
            # Set the check state of the checkbox in the first column of the current row to unchecked
            self.user_table.item(i, 0).setCheckState(Qt.Unchecked)

    def save_contact_list(self):
        """This method saves the contact list to the CSV file"""

        # Update the contact list with the data in the contact table
        updated_contacts = []
        headers = ['givenName', 'surname', 'mail', 'businessPhones',
                   'mobilePhone', 'department', 'jobTitle', 'officeLocation']
        for row in range(self.contact_table.rowCount()):
            contact = {}
            is_empty_row = True
            for col in range(self.contact_table.columnCount()):
                value = self.contact_table.item(row, col).text()
                if value:
                    is_empty_row = False
                if headers[col] == 'businessPhones':
                    # Remove the square brackets and convert to string
                    value = value.strip("['']")
                    # Format the phone number
                    value = self.format_phone_number(value)
                    if not value:
                        value = None
                elif headers[col] == 'mobilePhone':
                    # Format the phone number
                    value = self.format_phone_number(value)
                contact[headers[col]] = value
            if not is_empty_row:
                updated_contacts.append(contact)

        # Use ContactSync to write the updated contacts to the CSV file
        self.contact_sync.write_to_csv(updated_contacts, self.csv_file_path)

    def format_phone_number(self, phone_number):
        """This method formats the phone number to the format (XXX) XXX-XXXX. If the phone number is empty, it returns None."""
        # Remove any non-digit characters from the phone number
        phone_number = re.sub(r'\D', '', phone_number)
        # If the phone number is empty, return None
        if not phone_number:
            return None

        # Format the phone number
        formatted_phone_number = f"({phone_number[0:3]}) {
            phone_number[3:6]}-{phone_number[6:]}"

        return formatted_phone_number

    def update_contact_list(self):
        """
        This method updates the contact list by pulling the latest users from the Graph API and filtering them.
        """
        self.gal_window = GALWindow(self.contact_sync, self.contact_table)
        self.gal_window.show()

    def refresh_contact_list(self):
        '''This method refreshes the contact list by reloading the contacts from the CSV file and repopulating the contact table.'''
        # Reload the contacts from the CSV file
        self.contacts = self.contact_sync.read_csv_file(self.csv_file_path)

        # Repopulate the contact table with the reloaded contacts
        self.populate_contact_table()

    def get_contacts_from_table(self):
        '''This method retrieves the contacts from the contact table and returns them as a list of dictionaries.'''
        contacts = []
        for row in range(self.contact_table.rowCount()):
            contact = {}
            # Assuming 'email' is in the third column
            email = self.contact_table.item(row, 2).text()
            # Assuming 'phone' is in the fourth column
            mobile = self.contact_table.item(row, 3).text()
            # Assuming 'department' is in the fifth column
            business_phone = self.contact_table.item(row, 4).text()

            # Assuming 'name' is in the second column
            contact['givenName'] = self.contact_table.item(row, 0).text()
            contact['surname'] = self.contact_table.item(row, 1).text()
            contact['displayName'] = contact['givenName'] + \
                ' ' + contact['surname']
            contact['emailAddresses'] = [
                {'name': contact['givenName'] + ' ' + contact['surname'], 'address': email}]
            contact['mobilePhone'] = mobile
            contact['businessPhones'] = [business_phone]
            contact['department'] = self.contact_table.item(row, 5).text()

            contact['officeLocation'] = self.contact_table.item(row, 7).text()

            contact['JobTitle'] = self.contact_table.item(row, 6).text()

            contacts.append(contact)
        return contacts

    def get_selected_users_from_table(self):
        '''This method retrieves the selected users from the user table and returns them as a list of dictionaries.'''
        selected_users = []
        for row in range(self.user_table.rowCount()):
            # If the checkbox is checked
            if self.user_table.item(row, 0).checkState() == Qt.Checked:
                # Assuming email is displayed in the 3rd column
                email = self.user_table.item(row, 3).text()
                # Use 'mail' field instead of 'email' to match with Graph API user data structure
                user_info = next(
                    (user for user in self.users if user.get('mail') == email), None)
                if user_info:
                    # Modify the user_info dictionary to match the desired format
                    modified_user_info = {
                        'id': user_info.get('id'),
                        'givenName': user_info.get('givenName'),
                        'surname': user_info.get('surname'),
                        'displayName': f"{user_info.get('givenName')} {user_info.get('surname')}",
                        'emailAddresses': [{'name': f"{user_info.get('givenName')} {user_info.get('surname')}", 'address': user_info.get('mail')}],
                        'mobilePhone': user_info.get('mobilePhone'),
                        'businessPhones': user_info.get('businessPhones'),
                        'department': user_info.get('department'),
                        'officeLocation': user_info.get('officeLocation'),
                        'jobTitle': user_info.get('jobTitle')
                    }
                    selected_users.append(modified_user_info)
        # print(selected_users) # Debugging
        return selected_users

    def start_sync_process(self):
        '''This method starts the sync process by creating a SyncThread and starting it. It also creates a SyncStatusWindow to display the sync status.'''
        self.sync_status_window = SyncStatusWindow()
        self.sync_status_window.show()

        # Fetch contacts and users data before starting the thread
        contacts = self.contact_sync.format_contact_list(
            self.contact_sync.read_csv_file(self.csv_file_path))
        # de-serialize the contacts prior to passing to the sync function
        selected_users = self.get_selected_users_from_table()
        # Pass the contacts and selected users to the SyncThread
        self.sync_thread = SyncThread(
            self.sync_contacts_button, contacts, selected_users)
        self.sync_thread.update_signal.connect(
            self.sync_status_window.update_status)
        self.sync_thread.start()

    def sync_contacts_button(self, contacts, users):
        '''This method is called when the sync button is clicked. It processes the contacts and users to sync them.'''
        self.contact_sync.process_users_concurrently(users, contacts)


class GALWindow(QMainWindow):
    '''This class represents the Global Address List window. It is a subclass of QMainWindow and is used to display the Global Address List.'''

    def __init__(self, contact_sync, contact_table):
        super().__init__()
        self.contact_sync = contact_sync
        self.contact_table = contact_table
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Global Address List")  # Set the window title
        self.setGeometry(100, 100, 875, 400)  # Set the window geometry

        self.central_widget = QWidget(self)  # Create a central widget
        self.setCentralWidget(self.central_widget)  # Set the central widget
        layout = QVBoxLayout(self.central_widget)  # Create a vertical layout

        self.gal_table = QTableWidget()  # Create a table widget
        layout.addWidget(self.gal_table)  # Add the table to the layout

        self.gal_table.setSelectionMode(
            QAbstractItemView.MultiSelection)  # Set the selection mode
        self.gal_table.setSelectionBehavior(
            QAbstractItemView.SelectRows)  # Set the selection behavior

        self.gal_table.setColumnCount(8)  # Set the number of columns
        self.gal_table.setHorizontalHeaderLabels(  # Set the column headers
            ['Given Name', 'Surname', 'Email', 'Business Phone', 'Mobile', 'Department', 'Job Title', 'Office Location'])

        # Populate the table with all users (dummy function, implement fetching logic)
        self.populate_gal_table()

        # Buttons
        buttons_layout = QHBoxLayout()  # Create a horizontal layout
        # Create a button to select all users
        select_all_btn = QPushButton('Select All')
        # Connect the button to the select_all method
        select_all_btn.clicked.connect(self.select_all)
        # Create a button to clear the selected users
        clear_selected_btn = QPushButton('Clear Selected')
        # Connect the button to the clear_selected method
        clear_selected_btn.clicked.connect(self.clear_selected)
        # Create a button to add the selected users to the contact list
        add_to_list_btn = QPushButton('Add to Contact List')
        # Connect the button to the add_to_contact_list method
        add_to_list_btn.clicked.connect(self.add_to_contact_list)

        # Add the select all button to the layout
        buttons_layout.addWidget(select_all_btn)
        # Add the clear selected button to the layout
        buttons_layout.addWidget(clear_selected_btn)
        # Add the add to contact list button to the layout
        buttons_layout.addWidget(add_to_list_btn)

        # Add the buttons layout to the main layout
        layout.addLayout(buttons_layout)

    def populate_gal_table(self):
        '''This method populates the Global Address List table with the users from the Graph API.'''
        users = self.contact_sync.filter_users()
        self.gal_table.setRowCount(len(users))
        for i, user in enumerate(users):
            self.gal_table.setItem(i, 0, QTableWidgetItem(
                str(user.get('givenName', 'NONE'))))
            self.gal_table.setItem(i, 1, QTableWidgetItem(
                str(user.get('surname', 'NONE'))))
            self.gal_table.setItem(i, 2, QTableWidgetItem(
                str(user.get('mail', 'NONE'))))
            self.gal_table.setItem(i, 3, QTableWidgetItem(
                str(user.get('businessPhones', 'NONE'))))
            self.gal_table.setItem(i, 4, QTableWidgetItem(
                str(user.get('mobilePhone', 'NONE'))))
            self.gal_table.setItem(i, 5, QTableWidgetItem(
                str(user.get('department', 'NONE'))))
            self.gal_table.setItem(i, 6, QTableWidgetItem(
                str(user.get('jobTitle', 'NONE'))))
            self.gal_table.setItem(i, 7, QTableWidgetItem(
                str(user.get('officeLocation', 'NONE'))))

    def select_all(self):
        '''This method selects all users in the Global Address List table.'''
        # Get the table's selection model
        selectionModel = self.gal_table.selectionModel()

        # Clear current selection for a clean state
        selectionModel.clearSelection()

        # Create a selection that includes all rows
        for row in range(self.gal_table.rowCount()):
            # Get index of first column in each row
            index = self.gal_table.model().index(row, 0)
            selectionModel.select(
                index, QItemSelectionModel.Select | QItemSelectionModel.Rows)

        # Ensure the table has focus to show the selection highlight
        self.gal_table.setFocus()

    def clear_selected(self):
        '''This method clears the selected users in the Global Address List table.'''
        # Clear selection of all items in the table
        self.gal_table.clearSelection()

    def add_to_contact_list(self):
        '''This method adds the selected users in the Global Address List table to the main window's contact table.'''
        # Add selected users to the main window's contact table
        selected_rows = self.gal_table.selectionModel().selectedRows()

        # clear the contact table
        self.contact_table.setRowCount(0)

        for row in selected_rows:
            # Get the user data from the selected row
            user_data = [self.gal_table.item(row.row(), col).text()
                         for col in range(self.gal_table.columnCount())]

            # Modify the user data to get the first value from businessPhones list
            user_data[3] = user_data[3].split(',')[0]

            # Add the user data to the contact table
            self.contact_table.insertRow(self.contact_table.rowCount())
            for col, data in enumerate(user_data):
                self.contact_table.setItem(
                    self.contact_table.rowCount() - 1, col, QTableWidgetItem(data))

        self.close()
