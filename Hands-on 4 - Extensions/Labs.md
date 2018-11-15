# Extensions

## Application Customizer
1. Go to project from PnPjs lab
    ```PowerShell
    Set-Location c:\pnpjs
    ```
2. Add Application Customizer
    1. Run yeoman generator
        ```PowerShell
        yo @microsoft/sharepoint
        ```
    2. Choose `Extension`
    3. Choose `Application Customizer`
    4. Name it `TopBar`
    5. Accept description
3. Make it add top bar
    1. Open Visual studio code
        ```PowerShell
        code .
        ```
    2. Open `src\extensions\topBar\TopBarApplicationCustomizer.ts`
    3. Remove `onInit` method (including `@override`)
    4. Add property to save top placeholder
        ```ts
        private _topPlaceHolder: PlaceholderContent;
        ```
        (use Ctrl+. to add to import)
    5. Add `onPlaceholdersChanged` method
        ```ts
        public onPlaceholdersChanged(placeholderProvider: PlaceholderProvider) {
          if (!this._topPlaceHolder) {
            this._topPlaceHolder = placeholderProvider.tryCreateContent(PlaceholderName.Top);
            if (this._topPlaceHolder && this._topPlaceHolder.domElement) {
              this._topPlaceHolder.domElement.innerHTML = `
                <div>
                  Hello world from ${this.properties.testMessage}
                <div>`;
            }
          }
        }
        ```
4. Test it
    1. Open `config\serve.json`
    2. Copy url in `initialPage` into `serveConfigurations.default.pageUrl`
    3. Remove `initialPage`
    4. Remove `serveConfigurations.default.customActions`
    5. This enables test of web part using `gulp serve`
    6. Change `serveConfigurations.topBar.pageUrl` to be a modern page in your site a good choice is `https://yourTenant.sharepoint.com/sites/Developer/Lists/PnP/New%20View1.aspx`
    7. Change `serveConfigurations.topBar.customActions.<guid>.properties.testMessage` to `SPBG 16/11`
    8. Run
        ```powershell
        gulp serve --config=topBar
        ```
    9. Click `Load debug scripts`
5. Log "debugging"
    1. Open `src\extensions\topBar\TopBarApplicationCustomizer.ts`
    2. Add logging to the beginning of `onPlaceholdersChanged`
        ```ts
        Log.info("TopBar", "onPlaceholdersChanged called");
        ```
    3. Refresh page in browser
    4. Press `Ctrl+F12`
    5. Expand pane
    6. Filter source to `TopBar`
6. Real debugging
    1. Open `.vscode\launch.json`
    2. Copy `Hosted workbench` configuration
    3. Name copy `TopBar`
    4. Copy url from browser
    5. Paste in as Url
    6. Change `sourceMapPathOverrides` to
        ```json
        "sourceMapPathOverrides": {
          "webpack:///../src/*": "${webRoot}/src/*",
          "webpack:///.././src/*": "${webRoot}/src/*",
          "webpack:///../../src/*": "${webRoot}/src/*",
          "webpack:///../../../src/*": "${webRoot}/src/*",
          "webpack:///../../../../src/*": "${webRoot}/src/*",
          "webpack:///../../../../../src/*": "${webRoot}/src/*"
        },
        ```
    7. Open debugger pane
    8. Select `TopBar`
    9. Set a break point in `onPlaceholdersChanged`
    10. Close browser
    11. Press `F5`
    12. Login, accept, ...

## Field customizer
1. Add Field Customizer
    1. Run yeoman generator
        ```PowerShell
        yo @microsoft/sharepoint
        ```
    2. Choose `Extension`
    3. Choose `Field Customizer`
    4. Name it `jsLink`
    5. Accept description
    6. Select `No JavaScript framework`
    7. Add `@microsoft/sp-listview-extensibility`
        ```powershell
        npm install @microsoft/sp-listview-extensibility
        ```
2. Look at code
    1. Open `src\extensions\jsLink\JsLinkFieldCustomizer.ts`
3. Test it
    1. Open `config\serve.json`
    2. Copy `pageUrl` from `topBar` to `jsLink`
    3. Change `InternalFieldName` to `Categories`
    4. Change `sampleText` to `SPBG`
    5. Run
        ```powershell
        gulp serve --config=jsLink
        ```
    6. Click `Load debug scripts`
## Command Set
1. Add Command set
    1. **Take a copy of config\serve.json**
    2. Run yeoman generator
        ```PowerShell
        yo @microsoft/sharepoint
        ```
    3. Choose `Extension`
    4. Choose `ListView Command Set`
    5. Name it `ECB`
    6. Accept description
2. Look at code
    1. Open `src\extensions\ecb\EcbCommandSet.ts`
    2. In the `onExecute` method change COMMAND_1s `Dialog.alert` to
        ```ts
        Dialog.alert(`${this.properties.sampleTextOne} for ${event.selectedRows[0].getValueByName("ID")}`);
        ```
    3. Open `src\extensions\ecb\EcbCommandSet.manifest.json`
    4. Change `items` to
        ```json
          "items": {
            "COMMAND_1": {
              "title": { "default": "SPBG One" },
              "iconImageUrl": "https://static.thenounproject.com/png/2007250-42.png",
              "type": "command"
            },
            "COMMAND_2": {
              "title": { "default": "SPBG Two" },
              "type": "command"
            }
          }
        ```
3. Test it
    1. Open `config\serve.json`
    2. Copy `pageUrl` from `jsLink` to `ecb`
    3. Run
        ```powershell
        gulp serve --config=ecb
        ```
    4. Click `Load debug scripts`
    5. Click `SPBG Two`
    6. Select a single item
    7. Click `SPBG One`
