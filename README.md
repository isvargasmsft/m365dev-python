# M365 Developer: Using the Microsoft Graph Python SDK and Semantic Kernel

This application reads the emails from a user, extracts actions items and creates a To-Do list.

## Preparation
1. Sign up for a sandbox tenant in the [M365 Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program).
2. Log into the [Azure portal](portal.azure.com) with your M365 sandbox tenant.
3. Go to Azure Active Directory and [register your application in the portal](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app).

## Install the Microsoft Graph Python SDK
```py
pip install msgraph-sdk
```
> Note: Enable long paths in your environment if you receive a `Could not install packages due to an OSError`. For details, see [Enable Long Paths in Windows 10, Version 1607, and Later](https://learn.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation?tabs=powershell#enable-long-paths-in-windows-10-version-1607-and-later).

## Install the Semantic Kernel 

Placeholder. Explore the [Semantic Kernel](https://github.com/microsoft/semantic-kernel).

## Set up authentication and initialize the Microsoft Graph Python SDK
The first piece of this application is initializing the Microsoft Graph service. 
1. In your project, create an app.py file. 
2. Import the ClientSecretCredential library and the AzureIdentityAuthenticationProvider. We'll use these libraries to manage authentication against the Microsoft Graph API. To explore other ways to authenticate, visit the [Microsoft Graph Python SDK repo](https://github.com/microsoftgraph/msgraph-sdk-python).
3. Import the GraphRequestAdapter and the GraphServiceClient. We'll use this to manage the communications with the Microsoft Graph API. 
4. Set up your authentication by initializing an instance of the ClientSecretCredential library and an instance of the GraphServiceClient.

```py
import asyncio

from azure.identity.aio import ClientSecretCredential
from kiota_authentication_azure.azure_identity_authentication_provider import AzureIdentityAuthenticationProvider

from msgraph import GraphRequestAdapter
from msgraph import GraphServiceClient

asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

# Create auth proviver object. Used to authenticate request

credential = ClientSecretCredential("tenantID",
                                    "clientID",
                                    "cientSecret")
scopes = ['https://graph.microsoft.com/.default']
auth_provider = AzureIdentityAuthenticationProvider(credential, scopes=scopes)

# Initialize a request adapter. Handles the HTTP concerns
request_adapter = GraphRequestAdapter(auth_provider)

# Get a service client
client = GraphServiceClient(request_adapter)

```

## Get the emails of a user
Now we are ready to read the emails of the selected user. 
```py
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
         messages = await client.users_by_id("AlexW@contoso.com").messages.get()
         for msg in messages.value:
             print(
                 msg.subject,
                 msg.id,
             )
     except Exception as e_rr:
         print(f'Error: {e_rr.error.message}')
 asyncio.run(get_user_messages())
```

## Extract action items from emails
Placeholder.

## Create a new ToDo list
With the action items ready, we can create a new ToDo list to add our tasks.
```py
# Create ToDo list
from msgraph.generated.models.todo_task_list import TodoTaskList

async def create_todo_list():
    """Create a ToDo list"""
    try:
        request_body = TodoTaskList()
        request_body.display_name = 'Action items from emails'
        result = await client.users_by_id("AlexW@M365x86781558.OnMicrosoft.com").todo.lists.post(request_body)
    except Exception as e_rr:
        print(f'Error: {e_rr.error.message}')

asyncio.run(create_todo_list())
```
