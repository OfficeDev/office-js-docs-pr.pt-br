---
title: Depurar suplementos do Office no iPad e no Mac
description: ''
ms.date: 03/21/2018
ms.openlocfilehash: 5d68fa000e19d81ebbcd1b383a790958f2bbac72
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="debug-office-add-ins-on-ipad-and-mac"></a><span data-ttu-id="17fe2-102">Depurar suplementos do Office no iPad e no Mac</span><span class="sxs-lookup"><span data-stu-id="17fe2-102">Debug Office Add-ins on iPad and Mac</span></span>

<span data-ttu-id="17fe2-p101">Voc? pode usar o Visual Studio para desenvolver e depurar suplementos no Windows, mas n?o pode us?-lo para depurar suplementos no iPad ou no Mac. Como os suplementos s?o desenvolvidos usando HTML e Javascript, s?o projetados para funcionar em v?rias plataformas, mas pode haver diferen?as sutis em como cada navegador processa o HTML. Este artigo descreve como depurar suplementos em execu??o em um iPad ou em um Mac.</span><span class="sxs-lookup"><span data-stu-id="17fe2-p101">You can use Visual Studio to develop and debug add-ins on Windows, but you can't use it to debug add-ins on the iPad or Mac. Because add-ins are developed using HTML and Javascript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on an iPad or Mac.</span></span> 

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="17fe2-106">Depura??o com o Safari Web Inspector em um Mac</span><span class="sxs-lookup"><span data-stu-id="17fe2-106">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="17fe2-107">Voc? pode depurar um suplemento do Office usando o Safari Web Inspector.</span><span class="sxs-lookup"><span data-stu-id="17fe2-107">You can debug an Office add-in using Safari Web Inspector.</span></span> 

<span data-ttu-id="17fe2-108">Para poder depurar suplementos do Office no Mac, voc? deve ter o Mac OS High Sierra e Mac Office Vers?o: 16.9.1 (compila??o 18012504) ou posterior.</span><span class="sxs-lookup"><span data-stu-id="17fe2-108">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later.</span></span> <span data-ttu-id="17fe2-109">Se voc? n?o tiver uma compila??o do Office Mac, poder? obter uma ao adquirir o [programa Office 365 Developer](https://aka.ms/o365devprogram).</span><span class="sxs-lookup"><span data-stu-id="17fe2-109">If you don't have an Office Mac build, you can get one by joining the [Office 365 Developer program](https://aka.ms/o365devprogram).</span></span>

<span data-ttu-id="17fe2-110">Para come?ar, abra um terminal e defina a propriedade `OfficeWebAddinDeveloperExtras` para o aplicativo relevante do Office da seguinte maneira:</span><span class="sxs-lookup"><span data-stu-id="17fe2-110">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

<span data-ttu-id="17fe2-111">Em seguida, abra o aplicativo do Office e insira seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="17fe2-111">Then, open the Office application and insert your add-in.</span></span> <span data-ttu-id="17fe2-112">Clique com o bot?o direito no suplemento e voc? ver? a op??o **Inspecionar elemento** no menu de contexto.</span><span class="sxs-lookup"><span data-stu-id="17fe2-112">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span>  <span data-ttu-id="17fe2-113">Selecione essa op??o e ela abrir? o Inspetor, onde voc? pode definir pontos de interrup??o e depurar seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="17fe2-113">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="17fe2-114">Observe que esse ? um recurso experimental e n?o h? garantias de que preservaremos essa funcionalidade em vers?es futuras de aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="17fe2-114">Please note that this is an experimental feature and there are no guarantees that we will preserve this functionality in future versions of Office applications.</span></span>

## <a name="debugging-with-vorlonjs-on-a-ipad-or-mac"></a><span data-ttu-id="17fe2-115">Depura??o com o Vorlon.JS em um iPad ou Mac</span><span class="sxs-lookup"><span data-stu-id="17fe2-115">Debugging with Vorlon.JS on a iPad or Mac</span></span>

<span data-ttu-id="17fe2-116">Para depurar um suplemento no iPad ou Mac, voc? pode usar o Vorlon.JS, um depurador para p?ginas da Web que ? semelhante ?s ferramentas F12.</span><span class="sxs-lookup"><span data-stu-id="17fe2-116">To debug an add-in on iPad or Mac, you can use Vorlon.JS, a debugger for web pages that is similar to the F12 tools.</span></span> <span data-ttu-id="17fe2-117">Ele ? projetado para funcionar remotamente e permite depurar p?ginas da Web em dispositivos diferentes.</span><span class="sxs-lookup"><span data-stu-id="17fe2-117">It is designed to work remotely and it enables you to debug web pages across different devices.</span></span> <span data-ttu-id="17fe2-118">Para saber mais, veja o [site do Vorlon](http://www.vorlonjs.com).</span><span class="sxs-lookup"><span data-stu-id="17fe2-118">For more information, see the [Vorlon website](http://www.vorlonjs.com).</span></span>  


### <a name="install-and-set-up-vorlonjs"></a><span data-ttu-id="17fe2-119">Instalar e configurar o Vorlon.JS</span><span class="sxs-lookup"><span data-stu-id="17fe2-119">Install and set up up Vorlon.JS on a Mac or iPad</span></span>  

1.  <span data-ttu-id="17fe2-120">Fa?a logon no dispositivo como um administrador.</span><span class="sxs-lookup"><span data-stu-id="17fe2-120">Log on to the device as an administrator.</span></span>

2.  <span data-ttu-id="17fe2-121">Instale o [Node.js](https://nodejs.org) se ele ainda n?o estiver instalado.</span><span class="sxs-lookup"><span data-stu-id="17fe2-121">Install [Node.js](https://nodejs.org) if it isn't already installed.</span></span> 

3.  <span data-ttu-id="17fe2-p105">Abra uma janela do **Terminal** e digite o comando `npm i -g vorlon`. A ferramenta est? instalada em `/usr/local/lib/node_modules/vorlon`.</span><span class="sxs-lookup"><span data-stu-id="17fe2-p105">Open a **Terminal** window and enter the command `npm i -g vorlon`. The tool is installed to `/usr/local/lib/node_modules/vorlon`.</span></span>


### <a name="configure-vorlonjs-to-use-https"></a><span data-ttu-id="17fe2-124">Configurar o Vorlon.JS para usar HTTPS</span><span class="sxs-lookup"><span data-stu-id="17fe2-124">Configure Vorlon.JS to use HTTPS</span></span>

<span data-ttu-id="17fe2-p106">Para depurar um aplicativo usando o Vorlon.JS, adicione uma marca `<script>` ? p?gina de abertura do aplicativo que carrega um script Vorlon.JS de um local conhecido (veja os detalhes no procedimento a seguir). Se um suplementos for protegido por SSL (HTTPS), todos os scripts usados dever?o estar hospedados em um servidor HTTPS, inclusive o script Vorlon.JS. Portanto, voc? precisar? configurar o Vorlon.JS para usar SSL se quiser usar esse script com suplementos.</span><span class="sxs-lookup"><span data-stu-id="17fe2-p106">To debug an application using Vorlon.JS, you add a `<script>` tag to the opening page of the application that loads a Vorlon.JS script from a well-known location (for details, see the following procedure). If an add-in is SSL-secured (HTTPS), any scripts that it uses must be hosted from an HTTPS server, including the Vorlon.JS script. Therefore, you must configure Vorlon.JS to use SSL in order to use Vorlon.JS with add-ins.</span></span> 

> [!IMPORTANT]
> [!include[HTTPS guidance](../includes/https-guidance.md)]

1.  <span data-ttu-id="17fe2-128">No **Localizador**, acesse `/usr/local/lib/node_modules/vorlon`, abra o menu de contexto (clique com o bot?o direito do mouse) da pasta `/Server` e escolha **Obter Informa??es**.</span><span class="sxs-lookup"><span data-stu-id="17fe2-128">In **Finder**, go to `/usr/local/lib/node_modules/vorlon`, open the context menu for (right-click) the `/Server` folder, and then select **Get Info**.</span></span>

2.  <span data-ttu-id="17fe2-129">Escolha o ?cone de cadeado no canto inferior direito da janela **Informa??es do servidor** para desbloquear a pasta.</span><span class="sxs-lookup"><span data-stu-id="17fe2-129">Choose the padlock icon in the lower right corner of the **Server info** window to unlock the folder.</span></span>

3. <span data-ttu-id="17fe2-130">Na se??o **Compartilhamento e Permiss?es** da janela, defina o **Privil?gio** para o grupo **funcion?rios** como **Leitura/Grava??o**.</span><span class="sxs-lookup"><span data-stu-id="17fe2-130">In the **Sharing and Permissions** section of the window, set the **Privilege** for the **staff** group to **Read & Write**.</span></span>

4. <span data-ttu-id="17fe2-131">Escolha o ?cone de cadeado novamente para ***voltar a bloquear*** a pasta.</span><span class="sxs-lookup"><span data-stu-id="17fe2-131">Choose the padlock icon again to ***relock*** the folder.</span></span>

5. <span data-ttu-id="17fe2-132">No **Localizador**, expanda a subpasta `/Server`, clique com bot?o direito no arquivo `config.json` e selecione **Obter Informa??es**.</span><span class="sxs-lookup"><span data-stu-id="17fe2-132">Back in **Finder**, expand the `/Server` subfolder, right-click the file `config.json`, and then select **Get Info**.</span></span>

6. <span data-ttu-id="17fe2-p107">Na janela **informa??es de config.json**, altere os privil?gios do arquivo da mesma forma que voc? fez para sua pasta `/Server` pai. N?o se esque?a de bloquear novamente e de fechar a janela.</span><span class="sxs-lookup"><span data-stu-id="17fe2-p107">In the **config.json info** window, change the privileges of the file exactly the way you did for its parent `/Server` folder. Be sure to relock and close the window.</span></span>

7. <span data-ttu-id="17fe2-p108">No **Localizador**, clique com bot?o direito do mouse no arquivo `config.json`, selecione **Abrir com**e selecione **TextEdit**. O arquivo ? aberto em um editor de texto.</span><span class="sxs-lookup"><span data-stu-id="17fe2-p108">Back in **Finder**, right-click the file `config.json`, select **Open with**, and then select **TextEdit**. The file opens in a text editor.</span></span>

8. <span data-ttu-id="17fe2-137">Altere a propriedade **useSSL** para `true`.</span><span class="sxs-lookup"><span data-stu-id="17fe2-137">Change the value of the **useSSL** property to `true`.</span></span>

9. <span data-ttu-id="17fe2-p109">Na se??o **plug-ins**, localize o plug-in com a **id** de `OFFICE` e o **nome** de `Office Addin`. Se a propriedade **enabled** do plug-in ainda n?o estiver como `true`, defina-a como `true`.</span><span class="sxs-lookup"><span data-stu-id="17fe2-p109">In the **plugins** section, find the plugin with the **id** of `OFFICE` and the **name** of `Office Addin`. If the **enabled** property for the plug-in is not already `true`, set it to `true`.</span></span>

10. <span data-ttu-id="17fe2-140">Salve o arquivo e feche o editor.</span><span class="sxs-lookup"><span data-stu-id="17fe2-140">Save the file and close the editor.</span></span>

11. <span data-ttu-id="17fe2-141">No **Localizador**, navegue at? `/usr/local/lib/node_modules/vorlon`, clique com bot?o direito do mouse na subpasta `Server` e selecione **Novo terminal na pasta**.</span><span class="sxs-lookup"><span data-stu-id="17fe2-141">In **Finder**, navigate to `/usr/local/lib/node_modules/vorlon`, right-click the `Server` subfolder, and select **New terminal at folder**.</span></span> 
    
12. <span data-ttu-id="17fe2-p110">Na janela do **Terminal**, digite `sudo vorlon`. Ser? solicitado que voc? digite sua senha de administrador. O servidor Vorlon ? iniciado. Deixe aberta a janela do **Terminal**.</span><span class="sxs-lookup"><span data-stu-id="17fe2-p110">In the **Terminal** window, enter `sudo vorlon`. You will be prompted to enter your administrator password. The Vorlon server starts. Leave the **Terminal** window open.</span></span>

13. <span data-ttu-id="17fe2-p111">Abra uma janela do navegador e v? para `https://localhost:1337`, que ? a interface do Vorlon.JS. Quando solicitado, escolha **Sempre** para confiar no certificado de seguran?a.</span><span class="sxs-lookup"><span data-stu-id="17fe2-p111">Open a browser window and go to `https://localhost:1337`, which is the Vorlon.JS interface. When prompted, choose **Always** to trust the security certificate.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="17fe2-p112">Se n?o for solicitado, talvez seja necess?rio confiar no certificado manualmente. O arquivo de certificado ? `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt`. Experimente as etapas a seguir. Se voc? tiver problemas, veja a ajuda do Macintosh ou do iPad.</span><span class="sxs-lookup"><span data-stu-id="17fe2-p112">If you are not prompted, you might need to trust the certificate manually. The certificate file is `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt`. Try the following steps. If you have trouble, consult Macintosh or iPad help.</span></span> 
    >
    > 1. <span data-ttu-id="17fe2-152">Feche a janela do navegador e na janela do **Terminal** que est? executando o servidor Vorlon, use Control-C para parar o servidor.</span><span class="sxs-lookup"><span data-stu-id="17fe2-152">Close the browser window and in the **Terminal** window that is running the Vorlon server, use Control-C to stop the server.</span></span>
    > 2. <span data-ttu-id="17fe2-p113">No **Localizador**, clique com bot?o direito do mouse no arquivo `server.crt` e escolha **Acesso ao Conjunto de Chaves**. A janela **Acesso ao Conjunto de Chaves** ? exibida.</span><span class="sxs-lookup"><span data-stu-id="17fe2-p113">In **Finder**, right-click the `server.crt` file and select **Keychain Access**. The **Keychain Access** window opens.</span></span>
    > 3. <span data-ttu-id="17fe2-p114">Na lista **Conjuntos de Chaves** ? esquerda, escolha **logon**, caso ainda n?o estiver marcado, e, em seguida, escolha **Certificados** na se??o **Categoria**. Verifique se o **localhost** do certificado est? na lista.</span><span class="sxs-lookup"><span data-stu-id="17fe2-p114">In the **Keychains** list on the left, select **login** if it is not already selected, and then select **Certificates** in the **Category** section. The certificate **localhost** is listed.</span></span>
    > 4. <span data-ttu-id="17fe2-p115">Clique com bot?o direito do mouse no **localhost** do certificado e escolha **Obter Informa??es**. Uma janela do **localhost** ? exibida.</span><span class="sxs-lookup"><span data-stu-id="17fe2-p115">Right-click the certificate **localhost** and select **Get Info**. A **localhost** window opens.</span></span>
    > 5. <span data-ttu-id="17fe2-159">Na se??o **Confiar**, abra o seletor rotulado como **Ao usar este certificado** e escolha **Sempre Confiar**.</span><span class="sxs-lookup"><span data-stu-id="17fe2-159">In the **Trust** section, open the selector labeled **When using this certificate** and select **Always Trust**.</span></span> 
    > 6. <span data-ttu-id="17fe2-p116">Feche a janela do **localhost**. Se a a??o for bem-sucedida, o certificado do **localhost** na janela **Acesso ao Conjunto de Chaves** exibir? uma cruz branca em um c?rculo azul no ?cone.</span><span class="sxs-lookup"><span data-stu-id="17fe2-p116">Close the **localhost** window. If the action was successful, the **localhost** certificate in the **Keychain Access** window has a white cross in a blue circle on its icon.</span></span>


### <a name="configure-the-add-in-for-vorlonjs-debugging"></a><span data-ttu-id="17fe2-162">Configurar o suplemento para depura??o do Vorlon.JS</span><span class="sxs-lookup"><span data-stu-id="17fe2-162">Configure the add-in for Vorlon.JS debugging</span></span>

1. <span data-ttu-id="17fe2-163">Adicione a seguinte marca de script ? se??o `<head>` do arquivo home.html (ou arquivo HTML principal) do seu suplemento:</span><span class="sxs-lookup"><span data-stu-id="17fe2-163">Add the following script tag to the `<head>` section of the home.html file (or main HTML file) of your add-in:</span></span>

    ```html
    <script src="https://localhost:1337/vorlon.js"></script>    
    ```  

2. <span data-ttu-id="17fe2-164">Implante o aplicativo da Web do suplemento em um servidor Web que pode ser acessado do Mac ou iPad, como um site do Azure.</span><span class="sxs-lookup"><span data-stu-id="17fe2-164">Deploy the add-in web application to a web server that is accessible from the Mac or iPad, such as an Azure website.</span></span> 

3. <span data-ttu-id="17fe2-165">Atualize a URL do suplemento em todos os locais onde a URL aparece no manifesto do suplemento.</span><span class="sxs-lookup"><span data-stu-id="17fe2-165">Update the URL of the add-in in all the places where the URL appears in the add-in manifest.</span></span>

4. <span data-ttu-id="17fe2-166">No Mac ou iPad, copie o manifesto do suplemento na seguinte pasta: `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`, onde *{nome_do_host}* ? Word, Excel, PowerPoint ou Outlook.</span><span class="sxs-lookup"><span data-stu-id="17fe2-166">Copy the add-in manifest to the following folder on the Mac or iPad: `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`, where *{host_name}* is Word, Excel, PowerPoint, or Outlook.</span></span>


### <a name="inspect-an-add-in-in-vorlonjs"></a><span data-ttu-id="17fe2-167">Inspecionar um suplemento no Vorlon.JS</span><span class="sxs-lookup"><span data-stu-id="17fe2-167">Inspect an add-in in Vorlon.JS</span></span>

1. <span data-ttu-id="17fe2-168">Se o servidor Vorlon n?o estiver sendo executado, no **Localizador**, navegue at? `/usr/local/lib/node_modules/vorlon`, clique com bot?o direito na subpasta `Server` e selecione **Novo terminal na pasta**.</span><span class="sxs-lookup"><span data-stu-id="17fe2-168">If the Vorlon server is not running, in **Finder**, navigate to `/usr/local/lib/node_modules/vorlon`, right-click the `Server` subfolder, and select **New terminal at folder**.</span></span> 
    
2.  <span data-ttu-id="17fe2-p117">Na janela do **Terminal**, digite `sudo vorlon`. Ser? solicitado que voc? digite sua senha de administrador. O servidor Vorlon ? iniciado. Deixe aberta a janela do **Terminal**.</span><span class="sxs-lookup"><span data-stu-id="17fe2-p117">In the **Terminal** window, enter `sudo vorlon`. You will be prompted to enter your administrator password. The Vorlon server starts. Leave the **Terminal** window open.</span></span>

3.  <span data-ttu-id="17fe2-173">Abra uma janela do navegador e v? para `https://localhost:1337`, que ? a interface do Vorlon.JS.</span><span class="sxs-lookup"><span data-stu-id="17fe2-173">Open a browser window and go to `https://localhost:1337`, which is the Vorlon.JS interface.</span></span>

4. <span data-ttu-id="17fe2-p118">Realize o sideload do suplemento. Para o Excel, PowerPoint ou Word, realize o sideload conforme descrito em [Realizar sideload de um suplemento do Office no iPad e no Mac](sideload-an-office-add-in-on-ipad-and-mac.md). Se for um suplemento do Outlook, realize o sideload conforme descrito em [Realizar sideload de suplementos do Outlook para teste](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing). Se o suplemento n?o usar comandos de suplemento, ele ser? imediatamente aberto. Caso contr?rio, escolha o bot?o para abrir o suplemento. Dependendo da compila??o do aplicativo host do Office, o bot?o ser? exibido em ambas guias **P?gina Inicial** ou em uma guia **Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="17fe2-p118">Sideload the add-in. If it is for Excel, PowerPoint, or Word, sideload it as described in [Sideload an Office Add-in on iPad and Mac](sideload-an-office-add-in-on-ipad-and-mac.md). If it is an Outlook add-in, sideload it as described in [Sideload Outlook Add-ins for testing](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing). If the add-in does not use add-in commands, it will open immediately. Otherwise, choose the button to open the add-in. Depending on the build of the Office host application, the button will be on either the **Home** tab or an **Add-in** tab.</span></span>

<span data-ttu-id="17fe2-180">O suplemento aparecer? na lista de Clientes no Vorlon.JS (no lado esquerdo da interface do Vorlon.JS) como **{OS} - n**, para um determinado n?mero *n* e onde *{OS}* ? o tipo de dispositivo, como "Macintosh".</span><span class="sxs-lookup"><span data-stu-id="17fe2-180">The add-in will show up in the list of Clients in Vorlon.JS (on the left side of the Vorlon.JS interface) as **{OS} - n**, for some number *n*, and where *{OS}* is the device type, such as "Macintosh".</span></span> 

![Captura de tela que mostra a interface do Vorlon.js](../images/vorlon-interface.png)

<span data-ttu-id="17fe2-p119">A ferramenta Vorlon tem uma variedade de plug-ins. Os que estiverem habilitados no momento ser?o exibidos como guias na parte superior da ferramenta. (? poss?vel habilitar mais plug-ins escolhendo o ?cone de engrenagem no canto esquerdo). Esses plug-ins s?o semelhantes ?s fun??es nas ferramentas F12. Por exemplo, voc? pode real?ar elementos DOM, executar comandos e muito mais. Para obter mais detalhes, veja [Principais plug-ins da documenta??o do Vorlon](http://vorlonjs.com/documentation/#console)</span><span class="sxs-lookup"><span data-stu-id="17fe2-p119">The Vorlon tool has a variety of plug-ins. The ones that are currently enabled appear as tabs at the top of the tool. (You can enable more plug-ins by choosing the gears icon on the left.) These plug-ins are  similar to the functions in F12 tools. For example, you can highlight DOM elements, execute commands, and more. For more details, see [Vorlon Documentation Core Plugins](http://vorlonjs.com/documentation/#console)</span></span> 

<span data-ttu-id="17fe2-p120">Um plug-in do **Suplemento do Office** adiciona recursos extras ao Office.js, como explorar o modelo de objeto e executar chamadas de Office.js e ler os valores das propriedades de objetos. Para obter instru??es, veja [Plug-in do VorlonJS para depura??o de suplementos do Office](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/).</span><span class="sxs-lookup"><span data-stu-id="17fe2-p120">An **Office Addin** plug-in adds extra capabilities for Office.js, such as exploring the object model, executing Office.js calls, and reading the values of object properties. For instructions, see [VorlonJS plugin for debugging Office Add-in](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/).</span></span>

> [!NOTE]
> <span data-ttu-id="17fe2-188">N?o ? poss?vel definir pontos de interrup??o no Vorlon.JS.</span><span class="sxs-lookup"><span data-stu-id="17fe2-188">There is no way to set break points in Vorlon.JS.</span></span>


## <a name="clearing-the-office-applications-cache-on-a-mac-or-ipad"></a><span data-ttu-id="17fe2-189">Limpar cache do aplicativo do Office em um Mac ou iPad</span><span class="sxs-lookup"><span data-stu-id="17fe2-189">Clearing the Office application's cache on a Mac or iPad</span></span>

<span data-ttu-id="17fe2-p121">Os Suplementos muitas vezes s?o armazenados em cache no Office para Mac por quest?o de desempenho. Normalmente, o cache ser? limpo quando o suplemento for recarregado. Se houver mais de um suplemento no mesmo documento, ? prov?vel que o processo de limpeza autom?tica do cache ao recarregar n?o seja confi?vel.</span><span class="sxs-lookup"><span data-stu-id="17fe2-p121">Add-ins are cached often in Office for Mac, for performance reasons. Normally, the cache is cleared by reloading the add-in. If  more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.</span></span> 

<span data-ttu-id="17fe2-193">No Mac, o cache pode ser limpo manualmente ao excluir tudo na pasta `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="17fe2-193">On a Mac, you can clear the cache manually by deleting everything in the `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.</span></span> 

<span data-ttu-id="17fe2-p122">No iPad, voc? pode chamar `window.location.reload(true)` a partir do JavaScript no suplemento para for?ar uma recarrega. Uma outra alternativa ? reinstalar o Office.</span><span class="sxs-lookup"><span data-stu-id="17fe2-p122">On an iPad, you can call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>
