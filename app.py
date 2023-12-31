"""M365 Developer | Microsoft Graph Python SDK and Semantic Kernel"""

import asyncio

import semantic_kernel as sk

from azure.identity.aio import EnvironmentCredential
from kiota_authentication_azure.azure_identity_authentication_provider import AzureIdentityAuthenticationProvider
from kiota_abstractions.api_error import APIError

from msgraph import GraphRequestAdapter
from msgraph import GraphServiceClient

from msgraph.generated.models.todo_task_list import TodoTaskList
from msgraph.generated.models.todo_task import TodoTask

asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

kernel = sk.Kernel()

# Create auth proviver object. Used to authenticate request

credential = EnvironmentCredential()
scopes = ['https://graph.microsoft.com/.default']

# Get a service client
client = GraphServiceClient(credential, scopes)

# GET emails from user
async def get_user_messages():
     """Getting the messages of a user"""
     try:
         messages = await client.users.by_user_id("AlexW@M365x86781558.OnMicrosoft.com").messages.get()
        #  for msg in messages.value:
        #      print(
        #          msg.subject,
        #          msg.id,
        #      )
     except APIError as e_rr:
         print(f'Error: {e_rr.error.message}')


# Create ToDo list
my_list_id = "_"

my_action_items = [
    "Create email", 
    "Review presentation",
    "Schedule meeting"
]

async def create_todo_list_and_tasks():
    """Create a ToDo list"""
    try:
        request_body = TodoTaskList()
        request_body.display_name = 'Action items from emails'
        result = await client.users.by_user_id("AlexW@M365x86781558.OnMicrosoft.com").todo.lists.post(request_body)
        my_list_id = result.id
        # print(my_list_id)

        # Adding tasks
        for item in my_action_items:
            request_body = TodoTask()
            request_body.title = item
            result = await client.users.by_user_id("AlexW@M365x86781558.OnMicrosoft.com").todo.lists.by_todo_task_list_id(my_list_id).tasks.post(request_body)

    except APIError as e_rr:
        print(f'Error: {e_rr.error.message}')


async def main():
    await get_user_messages()
    await create_todo_list_and_tasks()
asyncio.run(main())