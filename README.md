# Office365-Javascript
A simple javascript library to handle Office365 API Authentification. Example includes getting the oauth2 token and routing between login/authetification pages. 


# App registration
Azure AD App registration: https://dev.outlook.com/AppRegistration​

(You can login with your office365 account)

Make sure to set the app permissions according to needs (Can be changed later).

Save the automatically generated client id. 


# Manage Apps
Go to https://apps.dev.microsoft.com/#/appList to manage your App and set  the following 

"oauth2AllowImplicitFlow": true,

"homepage": "http://localhost:8080/",

"replyUrls": [
        "http://localhost:8080"
    ],


# Configure 
set clientId officeAuth.js same as your Azure AD App

# Test
Host files locally and navigate to localhost:8080/index.html
Click the Start Office Auth Link and the Auth Flow should start.
You will be routed automatically through the office logon and then routed to auth.html
