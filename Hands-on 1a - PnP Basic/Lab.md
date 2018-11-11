# PnP Basic Operations

## Use plain CSOM
1. Open PnPBasic.sln
2. Select PnPBasic Project
3. In properties (F4) change Site URL to point at your developer Site Collection
4. Login
5. Run using F5
    1. Login
    2. Trust
    3. Click the CSOM button
    4. When done go back to SharePoint site
    5. Verify that CSOM list is created with 
        * "New View" View
        * "CSOM Item" Content Type
6. Stop using Shift+F5

## Use SharePoint PnP

### Add NuGet Package
1. Right click PnPBasicWeb
2. Select `Manage NuGet Packages...`
3. Select Browse Tab
4. Search for SharePointPnPCoreOnline
5. Select and click Install
6. When prompted Approve (twice)

### Add PnP Button
1. Open `PnPBasicWeb\Pages\Default.aspx`
2. Below the `CSOMButtom` add
    ```html
    <asp:Button ID="PNPButton" runat="server" OnClick="PNPButton_Click" Text="PnP" />    
    ```
3. Right click and select `View Code`
4. Copy/Paste `CSOMButton_Click` rename to `PNPButton_Click`

### Change code to use PnP
1. Add the top add:
    ```cs
    using SharePointPnP;
    ```
2. Change code to create content type to:
    ```cs
   var myCT = web.CreateContentType("PnP Item", "0x010078874F9C61114245806D6F09BC0362F9", "A Lab");
    ```
    Note changed GUID
3. Change code to add field to:
    ```cs
    myCT.AddFieldByName("Categories");
    ```
4. Change code to create list to:
    ```cs
    var list = web.CreateList(ListTemplateType.GenericList, "PnP", false);
    ```
5. Change code to add content type to:
    ```cs
    list.AddContentTypeToList(myCT);
    ```
6. Change code to add view to:
    ```cs
    list.CreateView("New View", ViewType.Html, new[] { "Title", "Categories" }, 10, true);
    ```
7. Change code to set property bag value to:
    ```cs
    web.SetPropertyBagValue("ourwebkey", Guid.NewGuid().ToString());
    ```
8. Change "code" to make property bag value searchable to:
    ```cs
    web.AddIndexedPropertyBagKey("ourwebkey");
    ```

### Fix SharePointContext code
1. Build 
2. Double click error
3. Change catch statement to just be
    ```
    cacch
    ```

### Run and Verify
