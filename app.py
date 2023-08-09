"""M365 Developer | Python """
# pylint: disable=no-member

import asyncio

import semantic_kernel as sk

from azure.identity.aio import ClientSecretCredential
from kiota_authentication_azure.azure_identity_authentication_provider import AzureIdentityAuthenticationProvider

from msgraph import GraphRequestAdapter
from msgraph import GraphServiceClient

asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

kernel = sk.Kernel()

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
         messages = await client.users_by_id("AlexW@contoso.com").messages.get()
         for msg in messages.value:
             print(
                 msg.subject,
                 msg.id,
             )
     except Exception as e_rr:
         print(f'Error: {e_rr.error.message}')

asyncio.run(get_user_messages())

# Create ToDo list
from msgraph.generated.models.todo_task_list import TodoTaskList

# variable to store the ID of the new list
my_list_id = "_"

async def create_todo_list():
    """Create a ToDo list"""
    try:
        request_body = TodoTaskList()
        request_body.display_name = 'Action items from emails'
        result = await client.users_by_id("AlexW@contoso.com").todo.lists.post(request_body)
        my_list_id = result.id
        print(my_list_id)
    except Exception as e_rr:
        print(f'Error: {e_rr.error.message}')

asyncio.run(create_todo_list())

# Create new tasks
from msgraph.generated.models.todo_task import TodoTask
async def create_new_task(item):
    """Create a new task"""
    try:
        request_body = TodoTask()
        request_body.title = item
        result = await client.users_by_id("AlexW@contoso.com").todo.lists_by_id(my_list_id).tasks.post(request_body)
    except Exception as e_rr:
        print(f'Error: {e_rr.error.message}')

# Placeholder of actions items. Real action items will be extracted from emails using semantic kernel.
my_action_items = [
    "Create email", 
    "Review presentation",
    "Schedule meeting"
]

for item in my_action_items:
    asyncio.run(create_new_task(item))
