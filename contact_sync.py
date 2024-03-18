import csv
import datetime
import traceback

from concurrent.futures import ThreadPoolExecutor, as_completed

from graph import GraphAPIContactManager


class ContactSync:
    '''Class to handle contacts synchronization between a csv file and the Microsoft Graph API.'''

    def __init__(self, client_id, client_secret, tenant_id):
        self.gcm = GraphAPIContactManager(client_id, client_secret, tenant_id)

    def read_csv_file(self, file_path):
        ''' Read a csv file and convert it to a list of dictionaries.'''
        contacts = []
        with open(file_path, 'r') as file:
            reader = csv.DictReader(file)
            for row in reader:
                contacts.append(row)
        return contacts

    def write_to_csv(self, contacts, file_path):
        """This method writes the contact list to a CSV file"""

        # If contacts is empty, write an empty CSV file
        if not contacts:
            with open(file_path, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow([])  # write an empty row
            return

        # If contacts is not empty, proceed as before
        fieldnames = contacts[0].keys()
        with open(file_path, 'w', newline='') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for contact in contacts:
                writer.writerow(contact)

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

    def get_user_folder_id(self, user_id):
        '''Get the folder id for a user. If the folder does not exist, create it. '''
        folders = self.gcm.get_users_folder_id(user_id)
        # if folder name "Work Contacts" exists, return the folder id otherwise create it
        for folder in folders:
            if folder['displayName'] == 'Work Contacts':
                return folder['id']
        # create the folder
        created_folder = self.gcm.create_contact_folder(
            user_id, 'Work Contacts')
        return created_folder['id']

    def get_user_contacts(self, user_id, folder_id):
        '''Get the contacts for a user.'''
        user_contacts = self.gcm.get_user_contacts(user_id, folder_id)
        return user_contacts

    def delete_user_contacts(self, user_id, contacts, folder_id):
        '''Delete the contacts for a user.'''
        delete_contacts = []
        for contact in contacts:
            delete_contacts.append(self.gcm.prepare_delete_contact_request(
                user_id, contact['id'], folder_id))
        response = self.gcm.execute_batch_requests(delete_contacts)
        return response

    def add_user_contacts(self, user_id, contacts, folder_id):
        '''Add the contacts for a user.'''
        add_contacts = []
        for contact in contacts:
            add_contacts.append(self.gcm.prepare_create_contact_request(
                user_id, contact, folder_id))
        response = self.gcm.execute_batch_requests(add_contacts)
        return response

    def update_user_contacts(self, user_id, contacts, folder_id):
        '''Update the contacts for a user.'''
        update_contacts = []
        for contact, contact_id, differences in contacts:
            update_contacts.append(self.gcm.prepare_update_contact_request(
                user_id, differences, folder_id, contact_id))
        response = self.gcm.execute_batch_requests(update_contacts)
        return response

    def format_contact_list(self, contacts):
        '''Format the contacts to match the Graph API schema.'''
        formatted_contacts = []
        for contact in contacts:
            formatted_contact = {
                "givenName": contact.get('givenName', ""),
                "surname": contact.get('surname', ""),
                "emailAddresses": [
                    {
                        "address": contact.get('mail', ""),
                        "name": (contact.get('givenName', "") + ' ' + contact.get('surname', "")).strip()
                    }
                ],
                "mobilePhone": contact.get('mobilePhone', ""),
                "businessPhones": [contact.get('businessPhones', "")],
                "jobTitle": contact.get('jobTitle', ""),
                "department": contact.get('department', ""),
                "officeLocation": contact.get('officeLocation', "")
            }
            formatted_contacts.append(formatted_contact)
            # print(formatted_contacts)
        return formatted_contacts

    def format_user_contacts(self, contacts):
        '''Format the users contacts to match the contacts list schema.'''
        formatted_contacts = []
        for contact in contacts:
            formatted_contact = {
                "id": contact.get('id'),
                "Given Name": contact.get('givenName', ""),
                "Surname": contact.get('surname', ""),
                "Email": contact.get('emailAddresses', [{}])[0].get('address', ""),
                "Mobile": contact.get('mobilePhone', ""),
                "Business Phone": contact.get('businessPhones', [""])[0],
                "Job Title": contact.get('jobTitle', ""),
                "Department": contact.get('department', ""),
                "Office Location": contact.get('officeLocation', "")
            }
            formatted_contacts.append(formatted_contact)
        return formatted_contacts

    def get_field_value(self, contact, field, csv=False):
        '''Get the value of a field from a contact.'''
        # Handle special cases like 'emailAddresses' and 'businessPhones' differently
        if field == 'emailAddresses' and not csv:
            return next((item.get('address', '').strip() for item in contact.get(field, []) if 'address' in item), '')
        elif field == 'businessPhones':
            return next((str(phone).strip() for phone in contact.get(field, []) if phone), '')
        else:
            # Default case for fields that are directly accessible
            return str(contact.get(self.csv_to_graph_field_map(field) if csv else field, '')).strip()

    def compare_contacts(self, user_contacts, csv_contacts):
        '''Compare the contacts from the Graph API and the CSV file.'''
        user_emails = {}
        duplicates = []

        for contact in user_contacts:
            email = contact['emailAddresses'][0]['address']
            if email in user_emails:
                duplicates.append(contact)
            else:
                user_emails[email] = contact

        csv_emails = {c['emailAddresses'][0]
                      ['address']: c for c in csv_contacts}

        to_add = [c for c in csv_contacts if c['emailAddresses']
                  [0]['address'] not in user_emails]

        to_delete = [user_emails[email]
                     for email in user_emails if email not in csv_emails] + duplicates

        to_update = []
        for email, contact in csv_emails.items():
            if email in user_emails:
                differences = self.get_contact_differences(
                    user_emails[email], contact)
                if differences:
                    to_update.append(
                        (contact, user_emails[email]['id'], differences))

        return to_add, to_delete, to_update

    def get_contact_differences(self, graph_contact, csv_contact):
        '''Get the differences between the graph contact and the csv contact.'''
        fields = ['givenName', 'surname', 'mobilePhone',
                  'businessPhones', 'jobTitle', 'department']
        differences = {}

        for field in fields:
            # Get the field value from the graph contact
            graph_value = self.get_field_value(graph_contact, field)
            # Get the field value from the CSV contact
            csv_value = self.get_field_value(csv_contact, field, csv=True)

            if graph_value != csv_value:
                # Add the differences to the differences dictionary
                differences[field] = csv_value

        return differences   # Return the differences

    def csv_to_graph_field_map(self, csv_field):
        '''Map CSV fields to Graph fields if they differ.'''
        mapping = {
            'Given Name': 'givenName',
            'Surname': 'surname',
            'Mobile': 'mobilePhone',
            'Business Phone': 'businessPhones',
            'Job Title': 'jobTitle',
            'Department': 'department'

        }
        return mapping.get(csv_field, csv_field)

    def process_users_concurrently(self, users, contacts):
        '''Process each user in their own thread, potentially in batches.'''
        # print(users)

        def process_user(user):
            '''Process a single user.'''
            # get the folder id for the user or create if it does not exist
            # returns the folder id
            folder_id = self.get_user_folder_id(user['id'])
            user_displayName = user.get('displayName')
            user_id = (user.get('id'))
            print(f"Processing {user_displayName} with ID {user_id}")
            print(f"Folder ID: {folder_id}")

            # get the contacts for the user
            user_contacts = self.gcm.get_user_contacts(user_id, folder_id)

            to_add, to_delete, to_update = self.compare_contacts(
                user_contacts, contacts)

            add_response = self.add_user_contacts(user_id, to_add, folder_id)
            # print(f"Add response: {add_response}")

            delete_response = self.delete_user_contacts(
                user_id, to_delete, folder_id)
            # print(f"Delete response: {delete_response}")

            update_response = self.update_user_contacts(
                user_id, to_update, folder_id)
            # print(f"Update response: {update_response.count}")

            # write to_add, to_delete, and differences to a file
            with open('sync_results.txt', 'a') as file:
                file.write(f"\n\nUser: {user_displayName}\n")
                file.write(f"Timestamp: {datetime.datetime.now()}\n")
                file.write("Contacts to add:\n")
                for contact, response in zip(to_add, add_response):
                    file.write(f"\tContact: {contact}\n")
                    file.write(f"\tResponse: {response}\n")
                file.write("Contacts to delete:\n")
                for contact, response in zip(to_delete, delete_response):
                    file.write(f"\tContact: {contact}\n")
                    file.write(f"\tResponse: {response}\n")
                file.write("Contacts to update:\n")
                for contact, response in zip(to_update, update_response):
                    file.write(f"\tContact: {contact}\n")
                    file.write(f"\tResponse: {response}\n")

            # uncomment the following lines to delete all contacts
            # response = self.delete_user_contacts(
            #     user_id, user_contacts, folder_id)

        with ThreadPoolExecutor(max_workers=20) as executor:
            future_to_user = {executor.submit(
                process_user, user): user for user in users}
            for future in as_completed(future_to_user):
                user = future_to_user[future]
                try:
                    future.result()
                except Exception as exc:
                    user_principal_name = user.get(
                        'userPrincipalName', 'Unknown User')
                    print(f"{user_principal_name} generated an exception: {exc}")
                    # Log the exception details for robust error handling
                    with open('error_log.txt', 'a') as file:
                        file.write(f"\n\nUser: {user_principal_name}\n")
                        file.write(f"Timestamp: {datetime.datetime.now()}\n")
                        file.write(f"Exception: {exc}\n")
                        traceback.print_exc(file=file)
                else:
                    # TODO: update the GUI that the user has been processed so GUI can reduce count
                    pass
