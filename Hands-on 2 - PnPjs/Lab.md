# @pnp/PnPjs

## Create SPFx WebPart project
1. Create new folder 'pnpjs'
    ```powershell
    New-Item c:\pnpjs -Type Directory
    ```
2. Enter the folder
    ```powershell
    Set-Location c:\pnpjs
    ```
3. Run the SPFx yeoman generator
    ```powershell
    yo @microsoft/sharepoint
    ```
    Accepting all the defaults
4. Trust dev certificate (only needed once)
    ```powershell
    gulp trust-dev-cert
    ```
    Accept if prompted

## Add @pnp/PnPjs
1. Add npm packages
    ```powershell
    npm install @pnp/sp @pnp/common @pnp/logging @pnp/odata --save
    ```
2. Setup pnp-context to match SPFx context
    1. Open `src\webparts\helloWorld\HelloWorldWebPart.ts`
    2. Add import at top
        ```ts
        import { setup } from '@pnp/common';
        ```
    3. Add onInit method to `HelloWorldWebPart`
        ```ts
        public async onInit(): Promise<void> {
          await super.onInit();
          setup({
            spfxContext: this.context
          });
        }
        ```
3. Change webpart to just output JSON
    1. Add a results property:
        ```ts
        private results={};
        ```
    2. Add a dumpJson method:
        ```ts
        private dumpJson(label, object) {
          this.results[`${Object.keys(this.results).length+1}. ${label}`]=object;
          this.domElement.innerHTML = `
            <pre>${JSON.stringify(this.results,undefined,2)}<pre>`;
        }
        ```
    3. Replace render method:
        ```ts
        public render(): void {
          this.dumpJson('Render', "Done!");
        }
        ```
4. Change `gulp serve` to use hosted workbench
    1. Change `config\serve.json` to
        ```json
        {
          "$schema": "https://developer.microsoft.com/json-schemas/core-build/serve.schema.json",
          "port": 4321,
          "https": true,
          "initialPage": "https://yourTenant.sharepoint.com/sites/developer/_layouts/15/workbench.aspx"
        }
        ```
5. Test
    1. Run `gulp serve`
    2. Login
    3. Add `Hello World` Web Part

## Try Basic User
1. Add import at top
    ```ts
    import { sp } from '@pnp/sp';
    ```
2. Change `render` method to
    ```ts
    public render(): void {
      // Get Web
      sp.web.select("Title").get().then(w => this.dumpJson("getWeb", w));

      // Get 5 most resent items from list PnP
      sp.web.lists.getByTitle("PnP").items
        .top(5).orderBy("Created", false)
        .select("Title", "Categories").get().then(
        li => this.dumpJson("ListItems", li));

      // Update item 1 in list PnP
      sp.web.lists.getByTitle("PnP").items.getById(1).update({
        Title: "My new Title",
        Categories: "Another value"
      }).then(
        iru => this.dumpJson("Update Item", "Done"));

      this.dumpJson('Render', "Done!");
    }
    ```
3. Refresh page in browser

## Try batching
1. Change `render` method to
    ```ts
    public render(): void {
      const batch = sp.web.createBatch();
      let list = sp.web.lists.getByTitle("PnP");
      list.getListItemEntityTypeFullName().then(entityType => {
        list.items.inBatch(batch).add({
          Title: "1st item added in batch",
          Categories: "Batch Item",
          ContentTypeId: "0x010078874F9C61114245806D6F09BC0362F9"
        }, entityType).then(iar => this.dumpJson("Item 1 Added", "Done"));
        list.items.inBatch(batch).add({
          Title: "2nd item added in batch",
          ContentTypeId: "0x01"
        }, entityType).then(iar => this.dumpJson("Item 2 Added", "Done"));
        list.items.inBatch(batch).add({
          Title: "3rd item added in batch",
          Categories: "Batch Item",
          ContentTypeId: "0x010078874F9C61114245806D6F09BC0362F9"
        }, entityType).then(iar => this.dumpJson("Item 3 Added", "Done"));
        list.items.inBatch(batch).add({
          Title: "4th item added in batch",
          ContentTypeId: "0x01"
        }, entityType).then(iar => this.dumpJson("Item 4 Added", "Done"));
        batch.execute().then(_ => {
          sp.web.lists.getByTitle("PnP").items
            .top(5).orderBy("Created", false)
            .select("Title", "Categories").get().then(
              li => this.dumpJson("ListItems", li));
          this.dumpJson("Batch", "Done!");
        });
      });
    }
    ```
2. Refresh page in browser

## Look at documentation

Go to https://pnp.github.io/pnpjs/

