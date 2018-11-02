cd C:\github\pnp-startkit\sp-starter-kit\provisioning
$user = "zxy@zxy.onmicrosoft.com"
$password = "Zxy" | ConvertTo-SecureString -AsPlainText -Force
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $password
Set-PnPTraceLog -On -Level Debug
$connect = Connect-PnPOnline -Url https://zxy.sharepoint.com -ReturnConnection -Credentials $credential
Apply-PnPProvisioningHierarchy -Path starterkit.pnp -Connection $connect -Verbose -Debug 