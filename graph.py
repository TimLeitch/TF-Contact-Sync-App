'''graph.py'''
import requests
import json
import uuid
import math
from datetime import datetime, timedelta


class GraphAPIContactManager:
    """A class to manage contacts using Graph API with token management."""

    def __init__(self, client_id, client_secret, tenant_id):
        """Initialize the manager with client credentials."""
        # Store the client ID, client secret, and tenant ID
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id

        # Construct the token URL using the tenant ID
        self.token_url = f"https://login.microsoftonline.com/{
            tenant_id}/oauth2/v2.0/token"

        # Set the API URL to the Microsoft Graph API endpoint
        self.api_url = "https://graph.microsoft.com/v1.0/"

        # Initialize the token and token expiry to None
        self.token = None
        self.token_expiry = datetime.utcnow()

    def get_token(self):
        """Generate a new token using client credentials."""
        if self.token is None or datetime.utcnow() >= self.token_expiry:
            payload = {
                'grant_type': 'client_credentials',
                'client_id': self.client_id,
                'client_secret': self.client_secret,
                'scope': 'https://graph.microsoft.com/.default'
            }
            response = requests.post(self.token_url, data=payload)
            response.raise_for_status()  # Raise an exception for HTTP errors
            token_response = response.json()  # Parse the JSON response
            # Store the access token
            self.token = token_response['access_token']
            # Get the token expiry time
            expires_in = token_response['expires_in']
            self.token_expiry = datetime.utcnow(
            ) + timedelta(seconds=expires_in - 300)  # Buffer time

        return self.token

    def batch_request(self, requests_list):
        """Send a batch request to Graph API."""
        token = self.get_token()
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        # Ensure request bodies are properly serialized for POST requests
        for request in requests_list:
            if 'body' in request and request['method'] == 'POST':
                request['body'] = json.dumps(request['body'])

        batch_payload = requests_list
        # print(batch_payload)  # Debugging
        response = requests.post(
            f"{self.api_url}$batch",
            headers=headers,
            json=batch_payload
        )

        try:
            response.raise_for_status()
        except requests.exceptions.HTTPError as e:
            error_message = f"HTTP Error: {
                e.response.status_code} - {e.response.reason}"
            if 'error' in e.response.json():
                error_message += f"\nError Message: {
                    e.response.json()['error']['message']}"
            raise Exception(error_message)

        return response.json()

    def execute_batch_requests(self, prepared_requests, chunk_size=20):
        """Execute prepared requests in batches."""
        def chunk_requests(reqs, size):
            """Yield successive size chunks from reqs."""
            for i in range(0, len(reqs), size):
                yield reqs[i:i + size]

        all_responses = []

        for chunk in chunk_requests(prepared_requests, chunk_size):
            batch_payload = {"requests": chunk}
            # Debug print to verify payload
            # print(f"Sending batch payload: {batch_payload}\n\n")

            token = self.get_token()
            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            }
            try:
                batch_response = requests.post(
                    f"{self.api_url}$batch",
                    headers=headers,
                    json=batch_payload
                )
                # print(batch_response)
                batch_response.raise_for_status()
                batch_responses = batch_response.json()

                if 'responses' in batch_responses:
                    all_responses.extend(batch_responses['responses'])
                else:
                    print(f"Error or unexpected response format in batch: {
                          batch_responses}")
            except requests.exceptions.RequestException as e:
                print(f"Error occurred during batch request: {e}")
                raise

        # for response in all_responses:
            # print(response)

        return all_responses

    def prepare_get_contacts_request(self, user_id, folder_id):
        """Prepare a request payload for getting contacts."""
        return {
            "id": user_id,
            "url": f"/users/{user_id}/contactFolders/{folder_id}/contacts",
            "method": "GET"
        }

    def prepare_create_contact_folder_request(self, user_id, folder_name):
        """Prepare a request payload for creating a contact folder."""
        return {
            "id": user_id,
            "url": f"/users/{user_id}/contactFolders",
            "method": "POST",
            "body": {"displayName": folder_name},
            "headers": {"Content-Type": "application/json"}
        }

    def prepare_get_contact_folders_request(self, user_id):
        """Prepare a request payload for getting contact folders."""
        return {
            "id": user_id,
            "url": f"/users/{user_id}/contactFolders",
            "method": "GET"
        }

    def prepare_create_contact_request(self, user_id, contact, folder_id):
        """Prepare a request payload for creating a contact."""

        def is_valid_value(value):
            # Checks if the value is not None, not an empty string, and not nan
            if value in [None, "", "nan"]:
                return False
            if isinstance(value, float) and math.isnan(value):
                return False
            return True

        # Exclude 'businessPhones' field if it is blank, null, or NaN
        body = {k: v for k, v in contact.items() if k != "id" and k !=
                "businessPhones" and is_valid_value(v)}

        return {
            "id": str(uuid.uuid4()),  # Convert UUID to string
            "url": f"/users/{user_id}/contactFolders/{folder_id}/contacts",
            "method": "POST",
            "body": body,
            "headers": {"Content-Type": "application/json"}
        }

    def prepare_update_contact_request(self, user_id, contact_differences, folder_id, contact_id):
        """Prepare a request payload for updating a contact."""

        def is_valid_value(value):
            # Checks if the value is not None, not an empty string, and not nan
            if value in [None, "", "nan"]:
                return False
            if isinstance(value, float) and math.isnan(value):
                return False
            return True

        # Create a new body dictionary handling 'businessPhones' and 'mobilePhone' specifically.
        body = {
            k: v for k, v in contact_differences.items() if is_valid_value(v)
        }

        # Handle 'businessPhones' explicitly. If it's not valid or not provided, set it as an empty list to clear it.
        if 'businessPhones' not in body or not is_valid_value(body.get('businessPhones')):
            body['businessPhones'] = []
        else:
            # Ensure businessPhones is properly formatted as a list
            body['businessPhones'] = [body['businessPhones']] if isinstance(
                body['businessPhones'], str) else body['businessPhones']

        # Handle 'mobilePhone' similarly. If it's not valid or not provided, set it to an empty string to clear it.
        if 'mobilePhone' not in body or not is_valid_value(body.get('mobilePhone')):
            body['mobilePhone'] = ""
        else:
            body['mobilePhone'] = body['mobilePhone']

        return {
            "id": str(uuid.uuid4()),  # Request ID, correctly generated here
            "url": f"/users/{user_id}/contactFolders/{folder_id}/contacts/{contact_id}",
            "method": "PATCH",
            "body": body,
            "headers": {"Content-Type": "application/json"}
        }

    def prepare_delete_contact_request(self, user_id, contact_id, folder_id):
        """Prepare a request payload for deleting a contact."""
        # Ensure contact_id is a valid, existing contact ID from Graph API.
        return {
            "id": str(uuid.uuid4()),  # Request ID, correctly generated here
            "url": f"/users/{user_id}/contactFolders/{folder_id}/contacts/{contact_id}",
            "method": "DELETE"
        }

    def get_users(self):
        """Retrieve all users in the organization using the Graph API with pagination."""
        users = []
        token = self.get_token()
        headers = {'Authorization': f'Bearer {token}'}
        api_endpoint = f"{
            self.api_url}users?$select=id,givenName,surname,department,jobTitle,officeLocation,mobilePhone,businessPhones,mail"

        while api_endpoint:
            response = requests.get(api_endpoint, headers=headers)
            response.raise_for_status()
            data = response.json()
            # Assuming 'value' contains the users list
            users.extend(data['value'])

            # Check if there is a next page link
            api_endpoint = data.get('@odata.nextLink')

        return users

    def get_user(self, upn):
        """Retrieve a user by userPrincipalName"""
        token = self.get_token()
        headers = {'Authorization': f'Bearer {token}'}
        api_endpoint = f"{
            self.api_url}users/{upn}"

        response = requests.get(api_endpoint, headers=headers)
        response.raise_for_status()
        data = response.json()

        return data

    def get_user_contacts(self, user_id, folder_id):
        """Retrieve all contacts in a user's contact folder."""
        token = self.get_token()
        headers = {'Authorization': f'Bearer {token}'}
        api_endpoint = f"{
            self.api_url}users/{user_id}/contactFolders/{folder_id}/contacts"

        contacts = []
        while api_endpoint:
            response = requests.get(api_endpoint, headers=headers)
            response.raise_for_status()
            data = response.json()
            contacts.extend(data['value'])
            api_endpoint = data.get('@odata.nextLink')

        return contacts

    def get_users_folder_id(self, user_id):
        """Retrieve the ID of a user's contact folder."""
        token = self.get_token()
        headers = {'Authorization': f'Bearer {token}'}
        api_endpoint = f"{
            self.api_url}users/{user_id}/contactFolders"

        response = requests.get(api_endpoint, headers=headers)
        response.raise_for_status()
        data = response.json()
        return data['value']

    def create_contact_folder(self, user_id, folder_name):
        """Create a new contact folder for a user."""
        token = self.get_token()
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        api_endpoint = f"{
            self.api_url}users/{user_id}/contactFolders"
        payload = json.dumps({"displayName": folder_name})
        response = requests.post(api_endpoint, headers=headers, data=payload)
        response.raise_for_status()
        return response.json()
