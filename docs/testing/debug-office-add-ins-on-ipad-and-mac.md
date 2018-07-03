---
title: Depurar suplementos do Office no iPad e no Mac
description: ''
ms.date: 03/21/2018
ms.openlocfilehash: e9efae76aa3341eacfd73d6afcc3a3274536aa9d
ms.sourcegitcommit: 6fbf42723f9c1b72095700c20458fd0e8c572794
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2018
ms.locfileid: "19722328"
---
# <a name="debug-office-add-ins-on-ipad-and-mac"></a><span data-ttu-id="8a468-102">Depurar suplementos do Office no iPad e no Mac</span><span class="sxs-lookup"><span data-stu-id="8a468-102">Debug Office Add-ins on iPad and Mac</span></span>

<span data-ttu-id="8a468-p101">Você pode usar o Visual Studio para desenvolver e depurar suplementos no Windows, mas não pode usá-lo para depurar suplementos no iPad ou no Mac. Como os suplementos são desenvolvidos usando HTML e Javascript, são projetados para funcionar em várias plataformas, mas pode haver diferenças sutis em como cada navegador processa o HTML. Este artigo descreve como depurar suplementos em execução em um iPad ou em um Mac.</span><span class="sxs-lookup"><span data-stu-id="8a468-p101">You can use Visual Studio to develop and debug add-ins on Windows, but you can't use it to debug add-ins on the iPad or Mac. Because add-ins are developed using HTML and Javascript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on an iPad or Mac.</span></span> 

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="8a468-106">Depuração com o Safari Web Inspector em um Mac</span><span class="sxs-lookup"><span data-stu-id="8a468-106">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="8a468-107">Se você tiver um suplemento que mostre a interface do usuário em um painel de tarefas ou em um suplemento de conteúdo, você pode depurar um suplemento do Office usando o Safari Web Inspector.</span><span class="sxs-lookup"><span data-stu-id="8a468-107">If you have add-in that shows UI in a taskpane or in a content add-in, you can debug an Office add-in using Safari Web Inspector.</span></span> 

<span data-ttu-id="8a468-108">Para poder depurar suplementos do Office no Mac, você deve ter o Mac OS High Sierra e Mac Office Versão: 16.9.1 (compilação 18012504) ou posterior.</span><span class="sxs-lookup"><span data-stu-id="8a468-108">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later.</span></span> <span data-ttu-id="8a468-109">Se você não tiver uma compilação do Office Mac, poderá obter uma ao adquirir o [programa Office 365 Developer](https://aka.ms/o365devprogram).</span><span class="sxs-lookup"><span data-stu-id="8a468-109">If you don't have an Office Mac build, you can get one by joining the [Office 365 Developer program](https://aka.ms/o365devprogram).</span></span>

<span data-ttu-id="8a468-110">Para começar, abra um terminal e defina a propriedade `OfficeWebAddinDeveloperExtras` para o aplicativo relevante do Office da seguinte maneira:</span><span class="sxs-lookup"><span data-stu-id="8a468-110">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

<span data-ttu-id="8a468-111">Em seguida, abra o aplicativo do Office e insira seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="8a468-111">Then, open the Office application and insert your add-in.</span></span> <span data-ttu-id="8a468-112">Clique com o botão direito no suplemento e você verá a opção **Inspecionar elemento** no menu de contexto.</span><span class="sxs-lookup"><span data-stu-id="8a468-112">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span>  <span data-ttu-id="8a468-113">Selecione essa opção e ela abrirá o Inspetor, onde você pode definir pontos de interrupção e depurar seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="8a468-113">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="8a468-114">Observe que esse é um recurso experimental e não há garantias de que preservaremos essa funcionalidade em versões futuras de aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="8a468-114">Please note that this is an experimental feature and there are no guarantees that we will preserve this functionality in future versions of Office applications.</span></span>

## <a name="debugging-with-vorlonjs-on-a-ipad-or-mac"></a><span data-ttu-id="8a468-115">Depuração com o Vorlon.JS em um iPad ou Mac</span><span class="sxs-lookup"><span data-stu-id="8a468-115">Debugging with Vorlon.JS on a iPad or Mac</span></span>

<span data-ttu-id="8a468-116">Para depurar um suplemento no iPad ou Mac, você pode usar o Vorlon.JS, um depurador para páginas da Web que é semelhante às ferramentas F12.</span><span class="sxs-lookup"><span data-stu-id="8a468-116">To debug an add-in on iPad or Mac, you can use Vorlon.JS, a debugger for web pages that is similar to the F12 tools.</span></span> <span data-ttu-id="8a468-117">Ele é projetado para funcionar remotamente e permite depurar páginas da Web em dispositivos diferentes.</span><span class="sxs-lookup"><span data-stu-id="8a468-117">It is designed to work remotely and it enables you to debug web pages across different devices.</span></span> <span data-ttu-id="8a468-118">Para saber mais, veja o [site do Vorlon](http://www.vorlonjs.com).</span><span class="sxs-lookup"><span data-stu-id="8a468-118">For more information, see the [Vorlon website](http://www.vorlonjs.com).</span></span>  


### <a name="install-and-set-up-vorlonjs"></a><span data-ttu-id="8a468-119">Instalar e configurar o Vorlon.JS</span><span class="sxs-lookup"><span data-stu-id="8a468-119">Install and set up up Vorlon.JS on a Mac or iPad</span></span>  

1.  <span data-ttu-id="8a468-120">Faça logon no dispositivo como administrador.</span><span class="sxs-lookup"><span data-stu-id="8a468-120">Log on to the device as an administrator.</span></span>

2.  <span data-ttu-id="8a468-121">Instale o [Node.js](https://nodejs.org) se ele ainda não estiver instalado.</span><span class="sxs-lookup"><span data-stu-id="8a468-121">Install [Node.js](https://nodejs.org) if it isn't already installed.</span></span> 

3.  <span data-ttu-id="8a468-p105">Abra uma janela do **Terminal** e digite o comando `npm i -g vorlon`. A ferramenta está instalada em `/usr/local/lib/node_modules/vorlon`.</span><span class="sxs-lookup"><span data-stu-id="8a468-p105">Open a **Terminal** window and enter the command `npm i -g vorlon`. The tool is installed to `/usr/local/lib/node_modules/vorlon`.</span></span>


### <a name="configure-vorlonjs-to-use-https"></a><span data-ttu-id="8a468-124">Configurar o Vorlon.JS para usar HTTPS</span><span class="sxs-lookup"><span data-stu-id="8a468-124">Configure Vorlon.JS to use HTTPS</span></span>

<span data-ttu-id="8a468-p106">Para depurar um aplicativo usando o Vorlon.JS, adicione uma marca `<script>` à página de abertura do aplicativo que carrega um script Vorlon.JS de um local conhecido (veja os detalhes no procedimento a seguir). Se um suplementos for protegido por SSL (HTTPS), todos os scripts usados deverão estar hospedados em um servidor HTTPS, inclusive o script Vorlon.JS. Portanto, você precisará configurar o Vorlon.JS para usar SSL se quiser usar esse script com suplementos.</span><span class="sxs-lookup"><span data-stu-id="8a468-p106">To debug an application using Vorlon.JS, you add a `<script>` tag to the opening page of the application that loads a Vorlon.JS script from a well-known location (for details, see the following procedure). If an add-in is SSL-secured (HTTPS), any scripts that it uses must be hosted from an HTTPS server, including the Vorlon.JS script. Therefore, you must configure Vorlon.JS to use SSL in order to use Vorlon.JS with add-ins.</span></span> 

> [!IMPORTANT]
> [!include[HTTPS guidance](../includes/https-guidance.md)]

1.  <span data-ttu-id="8a468-128">No **Localizador**, acesse `/usr/local/lib/node_modules/vorlon`, abra o menu de contexto (clique com o botão direito do mouse) da pasta `/Server` e escolha **Obter Informações**.</span><span class="sxs-lookup"><span data-stu-id="8a468-128">In **Finder**, go to `/usr/local/lib/node_modules/vorlon`, open the context menu for (right-click) the `/Server` folder, and then select **Get Info**.</span></span>

2.  <span data-ttu-id="8a468-129">Escolha o ícone de cadeado no canto inferior direito da janela **Informações do servidor** para desbloquear a pasta.</span><span class="sxs-lookup"><span data-stu-id="8a468-129">Choose the padlock icon in the lower right corner of the **Server info** window to unlock the folder.</span></span>

3. <span data-ttu-id="8a468-130">Na seção **Compartilhamento e Permissões** da janela, defina o **Privilégio** para o grupo **funcionários** como **Leitura/Gravação**.</span><span class="sxs-lookup"><span data-stu-id="8a468-130">In the **Sharing and Permissions** section of the window, set the **Privilege** for the **staff** group to **Read & Write**.</span></span>

4. <span data-ttu-id="8a468-131">Escolha o ícone de cadeado novamente para ***voltar a bloquear*** a pasta.</span><span class="sxs-lookup"><span data-stu-id="8a468-131">Choose the padlock icon again to ***relock*** the folder.</span></span>

5. <span data-ttu-id="8a468-132">No **Localizador**, expanda a subpasta `/Server`, clique com botão direito no arquivo `config.json` e selecione **Obter Informações**.</span><span class="sxs-lookup"><span data-stu-id="8a468-132">Back in **Finder**, expand the `/Server` subfolder, right-click the file `config.json`, and then select **Get Info**.</span></span>

6. <span data-ttu-id="8a468-p107">Na janela **informações de config.json**, altere os privilégios do arquivo da mesma forma que você fez para sua pasta `/Server` pai. Não se esqueça de bloquear novamente e de fechar a janela.</span><span class="sxs-lookup"><span data-stu-id="8a468-p107">In the **config.json info** window, change the privileges of the file exactly the way you did for its parent `/Server` folder. Be sure to relock and close the window.</span></span>

7. <span data-ttu-id="8a468-p108">No **Localizador**, clique com botão direito do mouse no arquivo `config.json`, selecione **Abrir com**e selecione **TextEdit**. O arquivo é aberto em um editor de texto.</span><span class="sxs-lookup"><span data-stu-id="8a468-p108">Back in **Finder**, right-click the file `config.json`, select **Open with**, and then select **TextEdit**. The file opens in a text editor.</span></span>

8. <span data-ttu-id="8a468-137">Altere a propriedade **useSSL** para `true`.</span><span class="sxs-lookup"><span data-stu-id="8a468-137">Change the value of the **useSSL** property to `true`.</span></span>

9. <span data-ttu-id="8a468-p109">Na seção **plug-ins**, localize o plug-in com a **id** de `OFFICE` e o **nome** de `Office Addin`. Se a propriedade **enabled** do plug-in ainda não estiver como `true`, defina-a como `true`.</span><span class="sxs-lookup"><span data-stu-id="8a468-p109">In the **plugins** section, find the plugin with the **id** of `OFFICE` and the **name** of `Office Addin`. If the **enabled** property for the plug-in is not already `true`, set it to `true`.</span></span>

10. <span data-ttu-id="8a468-140">Salve o arquivo e feche o editor.</span><span class="sxs-lookup"><span data-stu-id="8a468-140">Save the file and close the editor.</span></span>

11. <span data-ttu-id="8a468-141">No **Localizador**, navegue até `/usr/local/lib/node_modules/vorlon`, clique com botão direito do mouse na subpasta `Server` e selecione **Novo terminal na pasta**.</span><span class="sxs-lookup"><span data-stu-id="8a468-141">In **Finder**, navigate to `/usr/local/lib/node_modules/vorlon`, right-click the `Server` subfolder, and select **New terminal at folder**.</span></span> 
    
12. <span data-ttu-id="8a468-p110">Na janela do **Terminal**, digite `sudo vorlon`. Será solicitado que você digite sua senha de administrador. O servidor Vorlon é iniciado. Deixe aberta a janela do **Terminal**.</span><span class="sxs-lookup"><span data-stu-id="8a468-p110">In the **Terminal** window, enter `sudo vorlon`. You will be prompted to enter your administrator password. The Vorlon server starts. Leave the **Terminal** window open.</span></span>

13. <span data-ttu-id="8a468-p111">Abra uma janela do navegador e vá para `https://localhost:1337`, que é a interface do Vorlon.JS. Quando solicitado, escolha **Sempre** para confiar no certificado de segurança.</span><span class="sxs-lookup"><span data-stu-id="8a468-p111">Open a browser window and go to `https://localhost:1337`, which is the Vorlon.JS interface. When prompted, choose **Always** to trust the security certificate.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="8a468-p112">Se não for solicitado, talvez seja necessário confiar no certificado manualmente. O arquivo de certificado é `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt`. Experimente as etapas a seguir. Se você tiver problemas, veja a ajuda do Macintosh ou do iPad.</span><span class="sxs-lookup"><span data-stu-id="8a468-p112">If you are not prompted, you might need to trust the certificate manually. The certificate file is `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt`. Try the following steps. If you have trouble, consult Macintosh or iPad help.</span></span> 
    >
    > 1. <span data-ttu-id="8a468-152">Feche a janela do navegador e na janela do **Terminal** que está executando o servidor Vorlon, use Control-C para parar o servidor.</span><span class="sxs-lookup"><span data-stu-id="8a468-152">Close the browser window and in the **Terminal** window that is running the Vorlon server, use Control-C to stop the server.</span></span>
    > 2. <span data-ttu-id="8a468-p113">No **Localizador**, clique com botão direito do mouse no arquivo `server.crt` e escolha **Acesso ao Conjunto de Chaves**. A janela **Acesso ao Conjunto de Chaves** é exibida.</span><span class="sxs-lookup"><span data-stu-id="8a468-p113">In **Finder**, right-click the `server.crt` file and select **Keychain Access**. The **Keychain Access** window opens.</span></span>
    > 3. <span data-ttu-id="8a468-p114">Na lista **Conjuntos de Chaves** à esquerda, escolha **logon**, caso ainda não estiver marcado, e, em seguida, escolha **Certificados** na seção **Categoria**. Verifique se o **localhost** do certificado está na lista.</span><span class="sxs-lookup"><span data-stu-id="8a468-p114">In the **Keychains** list on the left, select **login** if it is not already selected, and then select **Certificates** in the **Category** section. The certificate **localhost** is listed.</span></span>
    > 4. <span data-ttu-id="8a468-p115">Clique com botão direito do mouse no **localhost** do certificado e escolha **Obter Informações**. Uma janela do **localhost** é exibida.</span><span class="sxs-lookup"><span data-stu-id="8a468-p115">Right-click the certificate **localhost** and select **Get Info**. A **localhost** window opens.</span></span>
    > 5. <span data-ttu-id="8a468-159">Na seção **Confiar**, abra o seletor rotulado como **Ao usar este certificado** e escolha **Sempre Confiar**.</span><span class="sxs-lookup"><span data-stu-id="8a468-159">In the **Trust** section, open the selector labeled **When using this certificate** and select **Always Trust**.</span></span> 
    > 6. <span data-ttu-id="8a468-p116">Feche a janela do **localhost**. Se a ação for bem-sucedida, o certificado do **localhost** na janela **Acesso ao Conjunto de Chaves** exibirá uma cruz branca em um círculo azul no ícone.</span><span class="sxs-lookup"><span data-stu-id="8a468-p116">Close the **localhost** window. If the action was successful, the **localhost** certificate in the **Keychain Access** window has a white cross in a blue circle on its icon.</span></span>


### <a name="configure-the-add-in-for-vorlonjs-debugging"></a><span data-ttu-id="8a468-162">Configurar o suplemento para depuração do Vorlon.JS</span><span class="sxs-lookup"><span data-stu-id="8a468-162">Configure the add-in for Vorlon.JS debugging</span></span>

1. <span data-ttu-id="8a468-163">Adicione a seguinte marca de script à seção `<head>` do arquivo home.html (ou arquivo HTML principal) do seu suplemento:</span><span class="sxs-lookup"><span data-stu-id="8a468-163">Add the following script tag to the `<head>` section of the home.html file (or main HTML file) of your add-in:</span></span>

    ```html
    <script src="https://localhost:1337/vorlon.js"></script>    
    ```  

2. <span data-ttu-id="8a468-164">Implante o aplicativo da Web do suplemento em um servidor Web que pode ser acessado do Mac ou iPad, como um site do Azure.</span><span class="sxs-lookup"><span data-stu-id="8a468-164">Deploy the add-in web application to a web server that is accessible from the Mac or iPad, such as an Azure website.</span></span> 

3. <span data-ttu-id="8a468-165">Atualize a URL do suplemento em todos os locais onde a URL aparece no manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="8a468-165">Update the URL of the add-in in all the places where the URL appears in the add-in manifest.</span></span>

4. <span data-ttu-id="8a468-166">No Mac ou iPad, copie o manifesto do suplemento na seguinte pasta: `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`, onde *{nome_do_host}* é Word, Excel, PowerPoint ou Outlook.</span><span class="sxs-lookup"><span data-stu-id="8a468-166">Copy the add-in manifest to the following folder on the Mac or iPad: `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`, where *{host_name}* is Word, Excel, PowerPoint, or Outlook.</span></span>


### <a name="inspect-an-add-in-in-vorlonjs"></a><span data-ttu-id="8a468-167">Inspecionar um suplemento no Vorlon.JS</span><span class="sxs-lookup"><span data-stu-id="8a468-167">Inspect an add-in in Vorlon.JS</span></span>

1. <span data-ttu-id="8a468-168">Se o servidor Vorlon não estiver sendo executado, no **Localizador**, navegue até `/usr/local/lib/node_modules/vorlon`, clique com botão direito na subpasta `Server` e selecione **Novo terminal na pasta**.</span><span class="sxs-lookup"><span data-stu-id="8a468-168">If the Vorlon server is not running, in **Finder**, navigate to `/usr/local/lib/node_modules/vorlon`, right-click the `Server` subfolder, and select **New terminal at folder**.</span></span> 
    
2.  <span data-ttu-id="8a468-p117">Na janela do **Terminal**, digite `sudo vorlon`. Será solicitado que você digite sua senha de administrador. O servidor Vorlon é iniciado. Deixe aberta a janela do **Terminal**.</span><span class="sxs-lookup"><span data-stu-id="8a468-p117">In the **Terminal** window, enter `sudo vorlon`. You will be prompted to enter your administrator password. The Vorlon server starts. Leave the **Terminal** window open.</span></span>

3.  <span data-ttu-id="8a468-173">Abra uma janela do navegador e vá para `https://localhost:1337`, que é a interface do Vorlon.JS.</span><span class="sxs-lookup"><span data-stu-id="8a468-173">Open a browser window and go to `https://localhost:1337`, which is the Vorlon.JS interface.</span></span>

4. <span data-ttu-id="8a468-p118">Realize o sideload do suplemento. Para o Excel, PowerPoint ou Word, realize o sideload conforme descrito em [Realizar sideload de um suplemento do Office no iPad e no Mac](sideload-an-office-add-in-on-ipad-and-mac.md). Se for um suplemento do Outlook, realize o sideload conforme descrito em [Realizar sideload de suplementos do Outlook para teste](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing). Se o suplemento não usar comandos de suplemento, ele será imediatamente aberto. Caso contrário, escolha o botão para abrir o suplemento. Dependendo da compilação do aplicativo host do Office, o botão será exibido em ambas guias **Página Inicial** ou em uma guia **Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="8a468-p118">Sideload the add-in. If it is for Excel, PowerPoint, or Word, sideload it as described in [Sideload an Office Add-in on iPad and Mac](sideload-an-office-add-in-on-ipad-and-mac.md). If it is an Outlook add-in, sideload it as described in [Sideload Outlook Add-ins for testing](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing). If the add-in does not use add-in commands, it will open immediately. Otherwise, choose the button to open the add-in. Depending on the build of the Office host application, the button will be on either the **Home** tab or an **Add-in** tab.</span></span>

<span data-ttu-id="8a468-180">O suplemento aparecerá na lista de Clientes no Vorlon.JS (no lado esquerdo da interface do Vorlon.JS) como **{OS} - n**, para um determinado número *n* e onde *{OS}* é o tipo de dispositivo, como "Macintosh".</span><span class="sxs-lookup"><span data-stu-id="8a468-180">The add-in will show up in the list of Clients in Vorlon.JS (on the left side of the Vorlon.JS interface) as **{OS} - n**, for some number *n*, and where *{OS}* is the device type, such as "Macintosh".</span></span> 

![Captura de tela que mostra a interface do Vorlon.js](../images/vorlon-interface.png)

<span data-ttu-id="8a468-p119">A ferramenta Vorlon tem uma variedade de plug-ins. Os que estiverem habilitados no momento serão exibidos como guias na parte superior da ferramenta. (É possível habilitar mais plug-ins escolhendo o ícone de engrenagem no canto esquerdo). Esses plug-ins são semelhantes às funções nas ferramentas F12. Por exemplo, você pode realçar elementos DOM, executar comandos e muito mais. Para obter mais detalhes, veja [Principais plug-ins da documentação do Vorlon](http://vorlonjs.com/documentation/#console)</span><span class="sxs-lookup"><span data-stu-id="8a468-p119">The Vorlon tool has a variety of plug-ins. The ones that are currently enabled appear as tabs at the top of the tool. (You can enable more plug-ins by choosing the gears icon on the left.) These plug-ins are  similar to the functions in F12 tools. For example, you can highlight DOM elements, execute commands, and more. For more details, see [Vorlon Documentation Core Plugins](http://vorlonjs.com/documentation/#console)</span></span> 

<span data-ttu-id="8a468-p120">Um plug-in do **Suplemento do Office** adiciona recursos extras ao Office.js, como explorar o modelo de objeto e executar chamadas de Office.js e ler os valores das propriedades de objetos. Para obter instruções, veja [Plug-in do VorlonJS para depuração de suplementos do Office](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/).</span><span class="sxs-lookup"><span data-stu-id="8a468-p120">An **Office Addin** plug-in adds extra capabilities for Office.js, such as exploring the object model, executing Office.js calls, and reading the values of object properties. For instructions, see [VorlonJS plugin for debugging Office Add-in](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/).</span></span>

> [!NOTE]
> <span data-ttu-id="8a468-188">Não é possível definir pontos de interrupção no Vorlon.JS.</span><span class="sxs-lookup"><span data-stu-id="8a468-188">There is no way to set break points in Vorlon.JS.</span></span>


## <a name="clearing-the-office-applications-cache-on-a-mac-or-ipad"></a><span data-ttu-id="8a468-189">Limpar cache do aplicativo do Office em um Mac ou iPad</span><span class="sxs-lookup"><span data-stu-id="8a468-189">Clearing the Office application's cache on a Mac or iPad</span></span>

<span data-ttu-id="8a468-p121">Os Suplementos muitas vezes são armazenados em cache no Office para Mac por questão de desempenho. Normalmente, o cache será limpo quando o suplemento for recarregado. Se houver mais de um suplemento no mesmo documento, é provável que o processo de limpeza automática do cache ao recarregar não seja confiável.</span><span class="sxs-lookup"><span data-stu-id="8a468-p121">Add-ins are cached often in Office for Mac, for performance reasons. Normally, the cache is cleared by reloading the add-in. If  more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.</span></span> 

<span data-ttu-id="8a468-193">No Mac, o cache pode ser limpo manualmente ao excluir tudo na pasta `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="8a468-193">On a Mac, you can clear the cache manually by deleting everything in the `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.</span></span> 

<span data-ttu-id="8a468-p122">No iPad, você pode chamar `window.location.reload(true)` a partir do JavaScript no suplemento para forçar uma recarrega. Uma outra alternativa é reinstalar o Office.</span><span class="sxs-lookup"><span data-stu-id="8a468-p122">On an iPad, you can call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>
