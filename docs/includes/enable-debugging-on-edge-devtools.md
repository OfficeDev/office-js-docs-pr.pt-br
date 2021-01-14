Quando o suplemento estiver sendo executado no Microsoft Edge, o código sem interface de usuário não será capaz de anexá-lo a um depurador por padrão.
O código sem interface de usuário é qualquer código executado enquanto o painel de tarefas não está visível, como comandos de suplemento. Para habilitar a depuração, execute os seguintes comandos do [Windows PowerShell](/powershell/scripting/getting-started/getting-started-with-windows-powershell).

1. Execute o comando a seguir para obter informações para o pacote de aplicativos do **Microsoft.Win32WebViewHost**.
    
    ```powershell
    Get-AppxPackage Microsoft.Win32WebViewHost
    ```
    
    O comando lista informações de pacote de aplicativos similares à seguinte saída.
    
    ```powershell
    Name              : Microsoft.Win32WebViewHost
    Publisher         : CN=Microsoft Windows, O=Microsoft Corporation, L=Redmond, S=Washington, C=US
    Architecture      : Neutral
    ResourceId        : neutral
    Version           : 10.0.18362.449
    PackageFullName   : Microsoft.Win32WebViewHost_10.0.18362.449_neutral_neutral_cw5n1h2txyewy
    InstallLocation   : C:\Windows\SystemApps\Microsoft.Win32WebViewHost_cw5n1h2txyewy
    IsFramework       : False
    PackageFamilyName : Microsoft.Win32WebViewHost_cw5n1h2txyewy
    PublisherId       : cw5n1h2txyewy
    IsResourcePackage : False
    IsBundle          : False
    IsDevelopmentMode : False
    NonRemovable      : True
    IsPartiallyStaged : False
    SignatureKind     : System
    Status            : Ok
    ```
    
2. Execute o seguinte comando para habilitar a depuração. Use o valor do **PackageFullName** listado no comando anterior.
    
    ```powershell
    setx JS_DEBUG <PackageFullName>
    ```
    
3. Se o Office já estava em execução, feche-o e reinicie-o para aplicar a alteração de depuração.