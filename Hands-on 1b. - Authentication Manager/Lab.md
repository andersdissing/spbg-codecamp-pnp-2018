# Authentication manager

## Setup
1. Create a new Console application
2. Install NuGet SharePointPnPCoreOnline (See lab 1a.)
3. Open `Program.cs`
    1. Add following usings at the top:
        ```cs
        using Microsoft.SharePoint.Client;
        using OfficeDevPnP.Core;
        ```
    2. Add the following scaffolding in the Main method:
        ```cs
        ClientContext clientContext=null;
        AuthenticationManager authenticationManager = new AuthenticationManager();
        string siteUrl = "https://yourTenant.sharepoint.com/sites/Developer";
        string tenant = "yourTenant.onmicrosoft.com";

        // Authenticate

        // Use clientContext
        Web web = clientContext.Web;
        clientContext.Load(web, w => w.Title);
        clientContext.ExecuteQueryRetry();
        Console.WriteLine(web.Title);
        ```
        (Changing siteUrl and tenant)

## Online credentials hardcoded
1. Add the following authentication code:
    ```cs
    // Online credentials hardcoded
    clientContext = authenticationManager.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, $"admin@{tenant}", "Passw0rd");
    ```
    (Changing username and password)
2. Run CTRL+F5

## Online credentials saved
0. Comment out previous authentication code
1. Save credentials
    1. Open windows `Credentials manager`
    2. Click `Windows credentials`
    3. Click `Add a generic credential`
    4. Enter `PnP` as `Internet or network address`
    5. Enter Username and Password
2. Add the following authentication code: 
    ```cs
    // Online credentials saved
    var cred = OfficeDevPnP.Core.Utilities.CredentialManager.GetCredential("PnP");
    clientContext= authenticationManager.GetSharePointOnlineAuthenticatedContextTenant(siteUrl, cred.UserName, cred.SecurePassword);
    ```
3. Run CTRL+F5

## Online credentials requested
0. Comment out previous authentication code
1. Add the following authentication code:
    ```cs
    // Online credentials requested
    clientContext = authenticationManager.GetWebLoginClientContext(siteUrl);
    ```
2. Add reference to System.Drawing
    1. In solution explorer right click `project\references`
    2. Select `Add reference`
    3. Check `System.Drawing`
    4. Click `OK`
3. Run CTRL+F5
4. Login

## AppOnly credentials ACS
0. Comment out previous authentication code
1. Add the following authentication code:
    ```cs
    // AppOnly credentials ACS
    clientContext = authenticationManager.GetAppOnlyAuthenticatedContext(siteUrl, "appId", "appSecret");
    ```
2. Register application
    1. Browse to `https://siteUrl/_layouts/15/appregnew.aspx`
    2. Click `Generate` next to `Client Secret`
    3. Copy the value and paste it into the code instead of `appSecret`
    4. Click `Generate` next to `Client Id`
    5. Copy the value and paste it into the code instead of `appId`
    6. Enter `PnP` as `Title`
    7. Enter `localhost` as `App Domain`
    8. Enter `https://localhost` as `Redirect URI`
    9. Click create
3. Request permissions
    1. Browse to `https://siteUrl/_layouts/15/appinv.aspx`
    2. Paste value of `Client Id` into `App Id`
    3. Click `Lookup`
    4. Paste into `Permission Request XML`
        ```xml
        <AppPermissionRequests AllowAppOnlyPolicy="true">
            <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web" Right="FullControl" />
        </AppPermissionRequests>
        ```
    5. Click `Create`
    6. Trust It
4. Run CTRL+F5

## AppOnly credentials Azure
0. Comment out previous authentication code
1. Add the following authentication code:
    ```cs
    // AppOnly credentials Azure
    clientContext = authenticationManager.GetAzureADAppOnlyAuthenticatedContext(siteUrl, "appId", tenant, @"C:\Pnp\Azure.pfx", (System.Security.SecureString)null);
    ```
2. Create application registration
    1. Visit `https://portal.azure.com`
    2. Click `Azure Active Directory`
    3. Click `App registration`
    4. Click `New application registration`
    5. Enter `PnP` as `Name`
    6. Leave `Web app / API` as `Application type`
    7. Enter `https://localhost` as `Sing-on URL`
    8. Click create
    9. Copy `Application ID` and paste it into the code instead of `appId`
3. Request permissions
    1. Click `Settings`
    2. Click `Required permissions`
    3. Click `Add`
    4. Click `Select an API`
    5. Click `Office 365 SharePoint Online`
    6. Click `Select`
    7. In `Application Permissions` click `Have full control of all site collections`
    8. Click `Select`
    9. Click `Done`
    10. Click `Grant permissions`
    11. Click `Yes`
    12. Close `Settings`
4. Get Certificate to use for authentication
    1. Open `Powershell`
    2. If not already installed install `PnP PowerShell`
        1. Run
            ```powershell
            Install-Module SharePointPnPPowerShellOnline -Scope CurrentUser
            ```
        2. Answer Y a lot of times
    3. Run
        ```powershell
        New-Item c:\PnP -Type Directory
        ```
    4. Run 
        ```powershell
        $cert = New-PnPAzureCertificate -Out "c:\PnP\Azure.pfx"
        ```
5. Register certificate with app registration
    1. In PowerShell run
        ```powershell
        $cert.KeyCredentials | Set-Clipboard
        ```
    2. In App registration in Azure Portal click `Manifest`
    3. Paste inside `[]` in `keyCredentials`
    4. Click `Save`
6. Run CTRL+F5