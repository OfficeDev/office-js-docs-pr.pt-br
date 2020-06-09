<span data-ttu-id="40f7b-101">Quando o suplemento estiver sendo executado no Microsoft Edge, o código sem interface de usuário não será capaz de anexá-lo a um depurador por padrão.</span><span class="sxs-lookup"><span data-stu-id="40f7b-101">When the add-in is running in Microsoft Edge, UI-less code will not be able to attach to a debugger by default.</span></span>
<span data-ttu-id="40f7b-102">O código sem interface de usuário é qualquer código executado enquanto o painel de tarefas não está visível, como comandos de suplemento.</span><span class="sxs-lookup"><span data-stu-id="40f7b-102">UI-less code is any code running while the task pane is not visible, such as add-in commands.</span></span> <span data-ttu-id="40f7b-103">Para habilitar a depuração, execute os seguintes comandos do [Windows PowerShell](https://docs.microsoft.com/powershell/scripting/getting-started/getting-started-with-windows-powershell).</span><span class="sxs-lookup"><span data-stu-id="40f7b-103">To enable debugging, you need to run the following [Windows PowerShell](https://docs.microsoft.com/powershell/scripting/getting-started/getting-started-with-windows-powershell) commands.</span></span>

1. <span data-ttu-id="40f7b-104">Execute o comando a seguir para obter informações para o pacote de aplicativos do **Microsoft.Win32WebViewHost**.</span><span class="sxs-lookup"><span data-stu-id="40f7b-104">Run the following command to get information for the **Microsoft.Win32WebViewHost** app package.</span></span>
    
    ```powershell
    Get-AppxPackage Microsoft.Win32WebViewHost
    ```
    
    <span data-ttu-id="40f7b-105">O comando lista informações de pacote de aplicativos similares à seguinte saída.</span><span class="sxs-lookup"><span data-stu-id="40f7b-105">The command lists app package information similar to the following output.</span></span>
    
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
    
2. <span data-ttu-id="40f7b-106">Execute o seguinte comando para habilitar a depuração.</span><span class="sxs-lookup"><span data-stu-id="40f7b-106">Run the following command to enable debugging.</span></span> <span data-ttu-id="40f7b-107">Use o valor do **PackageFullName** listado no comando anterior.</span><span class="sxs-lookup"><span data-stu-id="40f7b-107">Use the value for the **PackageFullName** listed from the previous command.</span></span>
    
    ```powershell
    setx JS_DEBUG <PackageFullName>
    ```
    
3. <span data-ttu-id="40f7b-108">Se o Office já estava em execução, feche-o e reinicie-o para aplicar a alteração de depuração.</span><span class="sxs-lookup"><span data-stu-id="40f7b-108">If Office was already running, close and restart Office so that it picks up the debugging change.</span></span>