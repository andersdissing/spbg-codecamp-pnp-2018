# Lab: SharePoint Starter Kit

In this lab, you will walk through 

1. Provisioning SharePoint Starter Kit
2. Update spfx webpart and redeploy.
3. Creating site designs and site scripts.

------

[TOC]

## Links

- [PnP SP Starter Kit](https://github.com/SharePoint/sp-starter-kit)
- [Office UI](https://developer.microsoft.com/en-us/fabric#/components/checkbox)
- [PnP Provisioning](https://github.com/SharePoint/PnP-Provisioning-Schema/blob/master/ProvisioningSchema-2018-07.md)
- [Site script available actions JSON schema](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-json-schema)

# Install and setup Starter Kit

## Get the code

- Clone this repository:

  ```bash
  cd c:\
  md pnpstarterkit
  cd pnpstarterkit
  git clone https://github.com/SharePoint/sp-starter-kit.git
  ```

- in the command line run: 

  ```bash
  npm install
  gulp trust-dev-cert
  gulp serve
  ```

## Preparing your tenant for the PnP SharePoint Starter Kit

- https://github.com/SharePoint/sp-starter-kit/blob/master/documentation/tenant-settings.md#preparing-your-tenant-for-the-pnp-sharepoint-starter-kit

## Connect and Provisioning template

- In PowerShell cmd or ISE

```bash
Connect-PnPOnline https://ZXY.sharepoint.com
Apply-PnPProvisioningHierarchy -Path starterkit.pnp
```

## Test

- Open https://zxy.sharepoint.com/sites/Contosoportal/

# Add filter to PersonalEmail WebPart

- Open solution folder (in VS Code).

- Navigate to .\solution\src\webparts\personalEmail\components.

- Open IPersonalEmailState.ts.
  - Add to line 7

     ````typescript
     onlyUnReadMessage: boolean;
     ````

- Open PersonalEmail.tsx.
  - Add to line 9
     ````typescript
     import { Checkbox } from 'office-ui-fabric-react';
     import { GraphRequest } from '@microsoft/sp-http';
     ````

  - Add to line 16
     ````typescript
     this._onCheckboxChange = this._onCheckboxChange.bind(this);
     ````

  - Add to line 20/21
     ````typescript
     ,
     onlyUnReadMessage: false
     ````

  - Add to line 24
     ````typescript
     private _onCheckboxChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean): void {

       this.setState({ onlyUnReadMessage: isChecked }, () => {
         this._loadMessages();
       });
     }
     ````

  - Add/update from line 47 to 58
     ````typescript
     var request: GraphRequest = this.props.graphClient
      .api("me/messages")
      .version("v1.0")
      .select("bodyPreview,receivedDateTime,from,subject,webLink")
      .top(this.props.nrOfMessages || 5)
      .orderby("receivedDateTime desc");
      
      if (this.state.onlyUnReadMessage) {
        request.filter("isRead eq false");
      }

      request.get((err: any, res: IMessages): void => {
     ````

  - Add to line 125
     ````react
     <Checkbox
       label="Filter: Show only unread messages"
       defaultChecked={this.state.onlyUnReadMessage}
       onChange={this._onCheckboxChange}
     />
     ````
## Deploy (debug) package
- Run bundle and package:

    ````javascript
    gulp bundle package-solution serve
    ````

# Site designs and site scripts

## [Simple site designs and site scripts](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/get-started-create-site-design)

### Walk through cheat sheet

1. Open PowerShell.

2. Navigate to hands-on 3 lab files e.g. 
   C:\github\spbg-codecamp-pnp-2018\Hands-on 3\Files

3. Connect to SharePoint online

   ````powershell
   $adminUPN="<the full email address of a SharePoint administrator account, example: jdoe@contosotoycompany.onmicrosoft.com>"
   $orgName="<name of your Office 365 organization, example: contosotoycompany>"`
   $userCredential = Get-Credential -UserName $adminUPN -Message "Type the password."
   Connect-SPOService -Url https://$orgName-admin.sharepoint.com -Credential $userCredential
   ````

4. Add SiteScript
   ````powershell
   $site_script = Get-Content -Path .\Getting-started.json -raw
   $id = Add-SPOSiteScript -Title "Create customer tracking list" -Content $site_script -Description "Creates list for tracking customer contact information" | select -ExpandProperty ID
   Add-SPOSiteDesign -Title "Contoso customer tracking" -WebTemplate "64" -SiteScripts $id -Description "Tracks key customer data in a list"
   ````

   ##### [Test - Use the new site design](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/get-started-create-site-design#use-the-new-site-design)

   1. Go to the home page of the SharePoint site that you are using for development. 
   2. Choose **Create site**. 
   3. Choose **Team site**. 
   4. In the **Choose a design** drop-down, select your site design **customer orders**. 
   5. In **Site name**, enter a name for the new site **Customer order tracking**. 
   6. Choose **Next**. 
   7. Choose **Finish**. 
   8. A pane indicates that your script is being applied. When it is done, choose **View updated site**. 
   9. You will see the custom list on the page. 

## Site designs and site scripts with Azure functions](https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-pnp-provisioning)