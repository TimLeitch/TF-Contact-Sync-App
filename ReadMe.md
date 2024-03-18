# Contact Synchronization Application

This application synchronizes contact information between a CSV file and Microsoft Graph, allowing for automated management of contacts within an organization's Microsoft 365 environment. It supports adding, deleting, and updating contacts based on comparisons between the CSV file and existing contacts in Microsoft 365.

## Prerequisites

Before using this application, ensure you have the following:

- Python 3.6 or later installed on your system.
- A Microsoft Azure subscription and permission to register applications in Azure Active Directory (Azure AD).
- Registered an application in Azure AD to obtain the necessary credentials (client ID, client secret) for API access.

## Installation

To install the application, clone the repository to your local machine and navigate to the application directory in your terminal. Then, install the required Python packages using pip:

```bash
pip install -r requirements.txt
``` 

This will install the required packages for the application to run.

## Configuration

The application requires specific environment variables to be set for authentication with Microsoft Graph and operation. These variables are to be provided in a .env file located at the root of the application directory.

### Creating the .env File

Create a .env file in the root directory of the application with the following content, replacing the placeholder values with your actual Azure AD application registration details and desired CSV file path:

```makefile
CLIENT_ID=your-azure-ad-application-client-id
CLIENT_SECRET=your-azure-ad-application-client-secret
TENANT_ID=your-azure-ad-directory-tenant-id
CSV_FILE_PATH=path-to-your-csv-file.csv
```

- `CLIENT_ID:` The Application (client) ID of your Azure AD application.
- `CLIENT_SECRET:` A client secret generated for your Azure AD application.
- `TENANT_ID:` The Directory (tenant) ID of your Azure AD tenant.
- `CSV_FILE_PATH:` The absolute path to the CSV file containing the contacts you wish to synchronize.

Create a contact CSV file with the required columns and place it in the directory specified by the CSV_FILE_PATH environment variable.
CSV File Format
The CSV file should contain columns for contact information. The application expects the following columns:

```makefile
Given Name,Surname,Email,Business Phone,Mobile,Department,Job Title,Office Location
```

Ensure that the CSV file follows this format for the application to correctly process the contact information.
  
### Usage

To run the application, navigate to the application directory in your terminal and execute the main script:

This initiates a graph api call to populate the users list based on the filter function

    def filter_users(self):
        '''Get users from the Graph API and format them.'''
        users = self.gcm.get_users()
        users = [u for u in users if u.get('mail')]
        # remove users that do not have a givenName and surname
        users = [u for u in users if u.get('givenName') and u.get('surname')]
        # remove users that do not have a department and job title and office location
        users = [u for u in users if u.get('department') and u.get(
            'jobTitle') and u.get('officeLocation')]
        # print(f"Total users: {len(users)}")
        return users

If you need to filter the users based on a different criteria, you can modify the filter_users function to suit your needs.


During the operation of the application, the following files will be created:

- `checked_users.json`: This file is created when the application is closed, saving the last checked state. It helps the application remember which users were checked during the last run.
- `error_log.txt`: This file is created if there are any errors during the process. It logs all the errors that occurred for troubleshooting purposes.
- `sync_results.txt`: This file is generated to store and log the results of each sync. It provides a history of all synchronization operations performed by the application.

Ensure that the application has write permissions to the directory where these files are to be created.

You can update or create a new contact CSV from inside the application by clicking on the `Update Contact List` button. This will open a new GUI window where you can update the contact list and save the changes to the CSV file.


### Troubleshooting

Authentication Errors: Verify that the .env file contains the correct CLIENT_ID, CLIENT_SECRET, and TENANT_ID.
CSV File Path: Ensure the CSV_FILE_PATH in the .env file is correct and points to a valid CSV file.
For further assistance, refer to the Microsoft Graph documentation and ensure your Azure AD application has the necessary permissions for contact management.