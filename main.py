import sys
import os

from PyQt5.QtWidgets import QApplication
from dotenv import load_dotenv

from graph import GraphAPIContactManager
from contact_sync import ContactSync
from gui import UserSyncGUI

# Load environment variables from .env file
load_dotenv()

# Get the values from environment variables
contact_folder_name = os.getenv('CONTACT_FOLDER_NAME')
client_id = os.getenv('CLIENT_ID')
tenant_id = os.getenv('TENANT_ID')
client_secret = os.getenv('CLIENT_SECRET')
csv_file_path = os.getenv('CSV_FILE_PATH')

# Create instances of ContactSync and GraphAPIContactManager
sync = ContactSync(client_id, client_secret, tenant_id)
gcm = GraphAPIContactManager(client_id, client_secret, tenant_id)


def main():
    '''Main function to run the application.'''
    # Create an instance of QApplication
    app = QApplication(sys.argv)

    # Create an instance of UserSyncGUI and show it
    ex = UserSyncGUI(csv_file_path, client_id, client_secret, tenant_id)
    ex.show()

    # Start the event loop
    sys.exit(app.exec_())


if __name__ == '__main__':
    # If this script is run as the main script, call the main function
    main()
