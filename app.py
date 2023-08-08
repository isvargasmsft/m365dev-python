"""M365 Developer | Python """
# pylint: disable=no-member

import asyncio

from azure.identity.aio import ClientSecretCredential
from kiota_authentication_azure.azure_identity_authentication_provider import AzureIdentityAuthenticationProvider

from msgraph import GraphRequestAdapter
from msgraph import GraphServiceClient

asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

# Create auth proviver object. Used to authenticate request

credential = ClientSecretCredential("tenantID",
                                    "clientID",
                                    "clientSecret")
scopes = ['https://graph.microsoft.com/.default']
auth_provider = AzureIdentityAuthenticationProvider(credential, scopes=scopes)

# Initialize a request adapter. Handles the HTTP concerns
request_adapter = GraphRequestAdapter(auth_provider)

# Get a service client
client = GraphServiceClient(request_adapter)

# GET emails from user
async def get_user_messages():
     """Getting the messages of a user"""
     try:
         messages = await client.users_by_id("AlexW@M365x86781558.OnMicrosoft.com").messages.get()
         for msg in messages.value:
             print(
                 msg.subject,
                 msg.id,
             )
     except Exception as e_rr:
         print(f'Error: {e_rr.error.message}')

asyncio.run(get_user_messages())

