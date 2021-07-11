---
title: Abra Excel página da Web e insiro seu Office Dep.
description: Abra Excel página da Web e insiro seu Office Add-in.
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 18f40b0030f4132a413a879e8b3419af49984b45
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349375"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a><span data-ttu-id="96729-103">Abra Excel página da Web e insiro seu Office Dep.</span><span class="sxs-lookup"><span data-stu-id="96729-103">Open Excel from your web page and embed your Office Add-in</span></span>

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="Imagem do Excel na página da Web abrindo um novo documento Excel com o seu add-in incorporado e abrindo automaticamente.":::

<span data-ttu-id="96729-105">Estenda seu aplicativo Web SaaS para que seus clientes possam abrir seus dados de uma página da Web diretamente para Microsoft Excel.</span><span class="sxs-lookup"><span data-stu-id="96729-105">Extend your SaaS web application so that your customers can open their data from a web page directly to Microsoft Excel.</span></span> <span data-ttu-id="96729-106">Um cenário comum é que os clientes trabalharão com dados em seu aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="96729-106">A common scenario is that customers will be working with data in your web application.</span></span> <span data-ttu-id="96729-107">Em seguida, eles vão querer copiar os dados em um Excel documento.</span><span class="sxs-lookup"><span data-stu-id="96729-107">Then they’ll want to copy the data into an Excel document.</span></span> <span data-ttu-id="96729-108">Por exemplo, eles podem querer executar análises adicionais usando Excel.</span><span class="sxs-lookup"><span data-stu-id="96729-108">For example, they may want to perform additional analysis using Excel.</span></span> <span data-ttu-id="96729-109">Normalmente, o cliente é obrigado a exportar os dados para um arquivo, como um arquivo .csv, e depois importar esses dados para Excel.</span><span class="sxs-lookup"><span data-stu-id="96729-109">Typically, the customer is required to export the data to a file, such as a .csv file, and then import that data into Excel.</span></span> <span data-ttu-id="96729-110">Eles também precisam adicionar manualmente o seu Office Add-in ao documento.</span><span class="sxs-lookup"><span data-stu-id="96729-110">They also have to manually add your Office Add-in to the document.</span></span>

<span data-ttu-id="96729-111">Reduza o número de etapas para um único botão clique em sua página da Web que gera e abre o Excel documento.</span><span class="sxs-lookup"><span data-stu-id="96729-111">Reduce the number of steps to a single button click on your web page that generates and opens the Excel document.</span></span> <span data-ttu-id="96729-112">Você também pode inserir seu Office de usuário dentro do documento e exibi-lo quando o documento for aberto.</span><span class="sxs-lookup"><span data-stu-id="96729-112">You can also embed your Office Add-in inside the document and display it when the document opens.</span></span> <span data-ttu-id="96729-113">Isso garante que o cliente ainda tenha acesso aos recursos do aplicativo.</span><span class="sxs-lookup"><span data-stu-id="96729-113">This ensures the customer still has access to your application features.</span></span> <span data-ttu-id="96729-114">Quando o documento é aberto, os dados selecionados pelo cliente e o seu Office Dedados já estão disponíveis para que eles continuem trabalhando.</span><span class="sxs-lookup"><span data-stu-id="96729-114">When the document opens, the data the customer selected, and your Office Add-in is already available for them to continue working.</span></span>

<span data-ttu-id="96729-115">Este artigo mostra o código e as técnicas para implementar esse cenário em seu próprio aplicativo Web SaaS.</span><span class="sxs-lookup"><span data-stu-id="96729-115">This article shows you code and techniques for implementing this scenario in your own SaaS web application.</span></span>

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a><span data-ttu-id="96729-116">Criar um novo Excel e inserir um Office Dep.</span><span class="sxs-lookup"><span data-stu-id="96729-116">Create a new Excel document and embed an Office Add-in</span></span>

<span data-ttu-id="96729-117">Primeiro, vamos aprender a criar um documento Excel de uma página da Web e inserir um complemento no documento.</span><span class="sxs-lookup"><span data-stu-id="96729-117">First, let’s learn how to create an Excel document from a web page, and embed an add-in into the document.</span></span> <span data-ttu-id="96729-118">O Office de código de entrada [OOXML](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) mostra como inserir o Script Lab [do](https://appsource.microsoft.com/product/office/wa104380862) Script Lab em um novo documento Office.</span><span class="sxs-lookup"><span data-stu-id="96729-118">The [Office OOXML Embed Add-in code sample](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) shows how to embed the [Script Lab add-in](https://appsource.microsoft.com/product/office/wa104380862) into a new Office document.</span></span> <span data-ttu-id="96729-119">Embora o exemplo funcione com qualquer documento Office, vamos nos concentrar Excel planilhas neste artigo.</span><span class="sxs-lookup"><span data-stu-id="96729-119">Although the sample works with any Office document, we’ll just focus on Excel spreadsheets in this article.</span></span> <span data-ttu-id="96729-120">Use as etapas a seguir para criar e executar o exemplo.</span><span class="sxs-lookup"><span data-stu-id="96729-120">Use the following steps to build and run the sample.</span></span>

1. <span data-ttu-id="96729-121">Extraia o código de exemplo  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip em uma pasta no computador.</span><span class="sxs-lookup"><span data-stu-id="96729-121">Extract the sample code from  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip into a folder on your computer.</span></span>
2. <span data-ttu-id="96729-122">Para criar e executar o exemplo, siga as etapas na seção **Para usar o** projeto do readme.</span><span class="sxs-lookup"><span data-stu-id="96729-122">To build and run the sample, follow the steps in the **To use the project** section of the readme.</span></span>
3. <span data-ttu-id="96729-123">Quando você executar o exemplo, ela exibirá uma página da Web semelhante à captura de tela a seguir.</span><span class="sxs-lookup"><span data-stu-id="96729-123">When you run the sample it will display a web page similar to the following screenshot.</span></span> <span data-ttu-id="96729-124">Use a página da Web para criar um novo documento Excel que contém Script Lab quando ele é aberto.</span><span class="sxs-lookup"><span data-stu-id="96729-124">Use the web page to create a new Excel document that contains Script Lab when it opens.</span></span>
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="Captura de tela da página da Web que o exemplo de laboratório de scripts de incorporação exibe para selecionar um arquivo Excel e incorporar o complemento do laboratório de script nele.":::

### <a name="how-the-sample-works"></a><span data-ttu-id="96729-126">Como o exemplo funciona</span><span class="sxs-lookup"><span data-stu-id="96729-126">How the sample works</span></span>

<span data-ttu-id="96729-127">O código de exemplo usa o SDK OOXML para incorporar o Script Lab do Script Lab ao documento de Excel que você escolher.</span><span class="sxs-lookup"><span data-stu-id="96729-127">The sample code uses the OOXML SDK to embed the Script Lab add-in to the Excel document that you choose.</span></span> <span data-ttu-id="96729-128">As informações a seguir são retiradas da [ **seção Sobre o código**](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) no arquivo readme.</span><span class="sxs-lookup"><span data-stu-id="96729-128">The following information is taken from the [**About the code** section](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) in the readme file.</span></span>

<span data-ttu-id="96729-129">O arquivo **Home.aspx.cs**:</span><span class="sxs-lookup"><span data-stu-id="96729-129">The file **Home.aspx.cs**:</span></span>

- <span data-ttu-id="96729-130">Fornece manipuladores de eventos de botão e manipulação básica da interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="96729-130">Provides the button event handlers and basic UI manipulation.</span></span>
- <span data-ttu-id="96729-131">Usa técnicas ASP.NET padrão para carregar e baixar o arquivo.</span><span class="sxs-lookup"><span data-stu-id="96729-131">Uses standard ASP.NET techniques to upload and download the file.</span></span>
- <span data-ttu-id="96729-132">Usa a extensão de nome de arquivo do arquivo carregado (xlsx, docx ou pptx) para determinar o tipo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="96729-132">Uses the file name extension of the uploaded file (xlsx, docx, or pptx) to determine the type of file.</span></span> <span data-ttu-id="96729-133">Isso precisa ser feito no início porque o SDK Open XML geralmente tem APIs distintas para cada tipo de arquivo.</span><span class="sxs-lookup"><span data-stu-id="96729-133">This needs to be done at the outset because the Open XML SDK generally has distinct APIs for each type of file.</span></span>
- <span data-ttu-id="96729-134">Chama o **OOXMLHelper** para validar o arquivo e chama o **AddInEmbedder** para inserir Script Lab no arquivo e definir para abrir automaticamente.</span><span class="sxs-lookup"><span data-stu-id="96729-134">Calls into the **OOXMLHelper** to validate the file and calls into the **AddInEmbedder** to embed Script Lab in the file and set to automatically open.</span></span>

<span data-ttu-id="96729-135">O arquivo **AddInEmbedder.cs**:</span><span class="sxs-lookup"><span data-stu-id="96729-135">The file **AddInEmbedder.cs**:</span></span>

- <span data-ttu-id="96729-136">Fornece a principal lógica de negócios, que neste exemplo é um método que incorpora Script Lab.</span><span class="sxs-lookup"><span data-stu-id="96729-136">Provides the main business logic, which in this sample is a method that embeds Script Lab.</span></span>
- <span data-ttu-id="96729-137">Faz chamadas para o auxiliar OOXML com base no tipo do arquivo.</span><span class="sxs-lookup"><span data-stu-id="96729-137">Makes calls into the OOXML helper based on the type of the file.</span></span>

<span data-ttu-id="96729-138">O arquivo **OOXMLHelper.cs**:</span><span class="sxs-lookup"><span data-stu-id="96729-138">The file **OOXMLHelper.cs**:</span></span>

- <span data-ttu-id="96729-139">Fornece toda a manipulação OOXML detalhada.</span><span class="sxs-lookup"><span data-stu-id="96729-139">Provides all the detailed OOXML manipulation.</span></span>
- <span data-ttu-id="96729-140">Usa uma técnica padrão para validar o arquivo Office, que é simplesmente chamar o **método Document.Open** nele.</span><span class="sxs-lookup"><span data-stu-id="96729-140">Uses a standard technique for validating the Office file, which is simply to call the **Document.Open** method on it.</span></span> <span data-ttu-id="96729-141">Se o arquivo for inválido, o método lançará uma exceção.</span><span class="sxs-lookup"><span data-stu-id="96729-141">If the file is invalid, the method throws an exception.</span></span>
- <span data-ttu-id="96729-142">Contém principalmente o código gerado pelas Ferramentas de Produtividade do SDK open XML 2.5 que estão disponíveis no link para o [SDK Open XML 2.5](/office/open-xml/open-xml-sdk).</span><span class="sxs-lookup"><span data-stu-id="96729-142">Contains mainly code that was generated by the Open XML 2.5 SDK Productivity Tools which are available at the link for the [Open XML 2.5 SDK](/office/open-xml/open-xml-sdk).</span></span>

<span data-ttu-id="96729-143">O **método GenerateWebExtensionPart1Content** no arquivo **OOXMLHelper.cs** define a referência à ID do Script Lab no Microsoft AppSource:</span><span class="sxs-lookup"><span data-stu-id="96729-143">The **GenerateWebExtensionPart1Content** method in the **OOXMLHelper.cs** file sets the reference to the ID of Script Lab in Microsoft AppSource:</span></span>

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- <span data-ttu-id="96729-144">O **valor StoreType** é "OMEX", um alias do Microsoft AppSource.</span><span class="sxs-lookup"><span data-stu-id="96729-144">The **StoreType** value is "OMEX", an alias for Microsoft AppSource.</span></span>
- <span data-ttu-id="96729-145">O **valor** da Loja é "en-US" encontrado na seção Cultura do Microsoft AppSource para Script Lab.</span><span class="sxs-lookup"><span data-stu-id="96729-145">The **Store** value is "en-US" found in the Microsoft AppSource culture section for Script Lab.</span></span>
- <span data-ttu-id="96729-146">O **valor da Id** é a ID de ativo do Microsoft AppSource para Script Lab.</span><span class="sxs-lookup"><span data-stu-id="96729-146">The **Id** value is the Microsoft AppSource asset ID for Script Lab.</span></span>

<span data-ttu-id="96729-147">Se você estiver configurando um complemento de um catálogo de compartilhamento de arquivos para abertura automática, usará valores diferentes:</span><span class="sxs-lookup"><span data-stu-id="96729-147">If you are setting up an add-in from a file share catalog for auto-open, you will use different values:</span></span>

<span data-ttu-id="96729-148">O **valor StoreType** é "FileSystem".</span><span class="sxs-lookup"><span data-stu-id="96729-148">The **StoreType** value is "FileSystem".</span></span>

- <span data-ttu-id="96729-149">O **valor** da Loja é a URL do compartilhamento de rede; por exemplo, " \\ \\ MyComputer \\ MySharedFolder".</span><span class="sxs-lookup"><span data-stu-id="96729-149">The **Store** value is the URL of the network share; for example, "\\\\MyComputer\\MySharedFolder".</span></span> <span data-ttu-id="96729-150">Essa deve ser a URL exata que aparece como o Endereço de Catálogo Confiável do compartilhamento na central de Office Trust Center.</span><span class="sxs-lookup"><span data-stu-id="96729-150">This should be the exact URL that appears as the share's Trusted Catalog Address in the Office Trust Center.</span></span>
- <span data-ttu-id="96729-151">O **valor de Id** é a ID do aplicativo no manifesto dos complementos.</span><span class="sxs-lookup"><span data-stu-id="96729-151">The **Id** value is the app ID in the add-ins manifest.</span></span>
> [!NOTE]
> <span data-ttu-id="96729-152">Para obter mais informações sobre valores alternativos para esses atributos, consulte [Abrir automaticamente](../develop/automatically-open-a-task-pane-with-a-document.md)um painel de tarefas com um documento .</span><span class="sxs-lookup"><span data-stu-id="96729-152">For more information about alternative values for these attributes, see [Automatically open a task pane with a document](../develop/automatically-open-a-task-pane-with-a-document.md).</span></span>

## <a name="use-the-fluent-ui"></a><span data-ttu-id="96729-153">Usar a interface Fluent interface do usuário</span><span class="sxs-lookup"><span data-stu-id="96729-153">Use the Fluent UI</span></span>

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Fluent Ícones de interface do usuário para Word, Excel e PowerPoint.":::

<span data-ttu-id="96729-155">Uma prática prática é usar a interface do usuário Fluent para ajudar os usuários a fazer a transição entre os produtos Microsoft.</span><span class="sxs-lookup"><span data-stu-id="96729-155">A best practice is to use the Fluent UI to help your users transition between Microsoft products.</span></span> <span data-ttu-id="96729-156">Você sempre deve usar um ícone Office para indicar qual aplicativo Office será lançado em sua página da Web.</span><span class="sxs-lookup"><span data-stu-id="96729-156">You should always use an Office icon to indicate which Office application will be launched from your web page.</span></span> <span data-ttu-id="96729-157">Vamos modificar o código de exemplo para usar o ícone Excel para indicar que ele inicia o Excel aplicativo.</span><span class="sxs-lookup"><span data-stu-id="96729-157">Let’s modify the sample code to use the Excel icon to indicate that it launches the Excel application.</span></span>

1. <span data-ttu-id="96729-158">Abra o exemplo em Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="96729-158">Open the sample in Visual Studio.</span></span>
1. <span data-ttu-id="96729-159">Abra a **página Home.aspx.**</span><span class="sxs-lookup"><span data-stu-id="96729-159">Open the **Home.aspx** page.</span></span>
1. <span data-ttu-id="96729-160">Encontre o código a seguir que é o botão de download no formulário.</span><span class="sxs-lookup"><span data-stu-id="96729-160">Find following code that is the download button on the form.</span></span>

    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```

1. <span data-ttu-id="96729-161">Substitua o código do botão pela seguinte marca de imagem.</span><span class="sxs-lookup"><span data-stu-id="96729-161">Replace the button code with the following image tag.</span></span>

    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```

1. <span data-ttu-id="96729-162">Pressione **F5** (ou **Depurar > Iniciar Depuração**).</span><span class="sxs-lookup"><span data-stu-id="96729-162">Press **F5** (or **Debug > Start Debugging**).</span></span> <span data-ttu-id="96729-163">Você verá o ícone aparecer quando a home page for carregada.</span><span class="sxs-lookup"><span data-stu-id="96729-163">You'll see the icon appear when the home page loads.</span></span>

<span data-ttu-id="96729-164">Para obter mais informações, [consulte Office Ícones](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) de Marca no portal Fluent de desenvolvedores da interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="96729-164">For more information, see [Office Brand Icons](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) on the Fluent UI developer portal.</span></span>  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a><span data-ttu-id="96729-165">Upload o documento Excel para Microsoft OneDrive</span><span class="sxs-lookup"><span data-stu-id="96729-165">Upload the Excel document to Microsoft OneDrive</span></span>

<span data-ttu-id="96729-166">Recomendamos carregar novos documentos para OneDrive se seu cliente usa OneDrive.</span><span class="sxs-lookup"><span data-stu-id="96729-166">We recommend uploading new documents to OneDrive if your customer uses OneDrive.</span></span> <span data-ttu-id="96729-167">Isso torna mais fácil para eles encontrar e trabalhar com os documentos.</span><span class="sxs-lookup"><span data-stu-id="96729-167">This makes it easier for them to find and work with the documents.</span></span> <span data-ttu-id="96729-168">Vamos criar um novo exemplo de código e ver como você pode usar o SDK do Microsoft Graph para carregar um novo documento Excel para OneDrive.</span><span class="sxs-lookup"><span data-stu-id="96729-168">Let’s create a new code sample and see how you can use the Microsoft Graph SDK to upload a new Excel document to OneDrive.</span></span>

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a><span data-ttu-id="96729-169">Usar um início rápido para criar um novo aplicativo Web Graph Microsoft</span><span class="sxs-lookup"><span data-stu-id="96729-169">Use a quick-start to build a new Microsoft Graph web application</span></span>

1. <span data-ttu-id="96729-170">Vá para e siga as etapas para criar e abrir um exemplo de código de início rápido que [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) interage com Office serviços.</span><span class="sxs-lookup"><span data-stu-id="96729-170">Go to [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) and follow the steps to create and open a quick start code sample that interacts with Office services.</span></span>
1. <span data-ttu-id="96729-171">Na **etapa 1: Escolher idioma ou plataforma,** escolha **ASP.NET MVC**.</span><span class="sxs-lookup"><span data-stu-id="96729-171">In **step 1: Pick you language or platform**, choose **ASP.NET MVC**.</span></span> <span data-ttu-id="96729-172">Embora as etapas deste procedimento usem a opção ASP.NET MVC, as etapas seguem um padrão que se aplica a qualquer idioma ou plataforma.</span><span class="sxs-lookup"><span data-stu-id="96729-172">Although the steps in this procedure use the ASP.NET MVC option, the steps follow a pattern that apply to any language or platform.</span></span>
1. <span data-ttu-id="96729-173">Na **etapa 2: Obter uma ID do aplicativo e** um segredo, escolha Obter uma **ID do aplicativo e segredo**.</span><span class="sxs-lookup"><span data-stu-id="96729-173">In **step 2: Get an app ID and secret**, choose **Get an app ID and secret**.</span></span>
1. <span data-ttu-id="96729-174">Entre na sua conta Microsoft 365 de usuário.</span><span class="sxs-lookup"><span data-stu-id="96729-174">Sign in to your Microsoft 365 account.</span></span>  
1. <span data-ttu-id="96729-175">Na página Da Web Secreta do **aplicativo,** salve o segredo do aplicativo em um local de arquivo onde você pode recuperá-lo e usá-lo mais tarde.</span><span class="sxs-lookup"><span data-stu-id="96729-175">On the **Please save your app secret** web page, save the app secret to a file location where you can retrieve and use it later.</span></span>
1. <span data-ttu-id="96729-176">Escolha **Got it, take me back to the quick start**.</span><span class="sxs-lookup"><span data-stu-id="96729-176">Choose **Got it, take me back to the quick start**.</span></span>
1. <span data-ttu-id="96729-177">Na **etapa 2: Registro bem-sucedido!**</span><span class="sxs-lookup"><span data-stu-id="96729-177">In **step 2: Registration Successful!**</span></span> <span data-ttu-id="96729-178">Insira o segredo do aplicativo gerado.</span><span class="sxs-lookup"><span data-stu-id="96729-178">Enter the generated app secret.</span></span>
1. <span data-ttu-id="96729-179">Na **etapa 3: Iniciar a codificação,** escolha Baixar o exemplo de código baseado em **SDK.**</span><span class="sxs-lookup"><span data-stu-id="96729-179">In **step 3: Start coding**, choose **Download the SDK-based code sample**.</span></span>
1. <span data-ttu-id="96729-180">Extraia a pasta zip de download em uma pasta local.</span><span class="sxs-lookup"><span data-stu-id="96729-180">Extract the download zip folder into a local folder.</span></span>  
1. <span data-ttu-id="96729-181">Abra o arquivo graph-tutorial.sln no Visual Studio 2019.</span><span class="sxs-lookup"><span data-stu-id="96729-181">Open the graph-tutorial.sln file in Visual Studio 2019.</span></span>
1. <span data-ttu-id="96729-182">Crie e execute a solução e confirme se ela está funcionando corretamente.</span><span class="sxs-lookup"><span data-stu-id="96729-182">Build and run the solution and confirm it is working correctly.</span></span> <span data-ttu-id="96729-183">Você deve poder usar a página da Web do calendário para exibir seu calendário Microsoft 365 calendário.</span><span class="sxs-lookup"><span data-stu-id="96729-183">You should be able to use the calendar web page to view your Microsoft 365 calendar.</span></span>

### <a name="upload-a-file-to-onedrive"></a><span data-ttu-id="96729-184">Upload um arquivo para OneDrive</span><span class="sxs-lookup"><span data-stu-id="96729-184">Upload a file to OneDrive</span></span>

1. <span data-ttu-id="96729-185">Abra a **solução graph-tutorial.sln** no Visual Studio 2019 e abra o arquivo **PrivateSettings.config.**</span><span class="sxs-lookup"><span data-stu-id="96729-185">Open the **graph-tutorial.sln** solution in Visual Studio 2019, and open the **PrivateSettings.config** file.</span></span>
1. <span data-ttu-id="96729-186">Adicione um novo escopo **Files.ReadWrite** à chave   **ida:AppScopes** para que ela se pareça com o código a seguir.</span><span class="sxs-lookup"><span data-stu-id="96729-186">Add a new scope **Files.ReadWrite** to the **ida:AppScopes** key so that it looks like the following code.</span></span>

    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```

1. <span data-ttu-id="96729-187">Abra o **arquivo Index.cshtml.**</span><span class="sxs-lookup"><span data-stu-id="96729-187">Open the **Index.cshtml** file.</span></span>
1. <span data-ttu-id="96729-188">Insira o código ActionLink a seguir para criar um botão para carregar um arquivo no OneDrive.</span><span class="sxs-lookup"><span data-stu-id="96729-188">Insert the following ActionLink code to create a button to upload a file to OneDrive.</span></span>

    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```

1. <span data-ttu-id="96729-189">Abra o **arquivo HomeController.cs.**</span><span class="sxs-lookup"><span data-stu-id="96729-189">Open the **HomeController.cs** file.</span></span>
1. <span data-ttu-id="96729-190">Insira o código a seguir para manipular a solicitação do link de ação.</span><span class="sxs-lookup"><span data-stu-id="96729-190">Insert the following code to handle the request from the action link.</span></span>

    ```csharp
    public void CreateOneDriveFile()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("The contents of the file goes here.")))
            {
                var client = graph_tutorial.Helpers.GraphHelper.UploadFile("/test.txt", stream);
            }
        }
    ```

1. <span data-ttu-id="96729-191">Abra o **arquivo GraphHelper.cs.**</span><span class="sxs-lookup"><span data-stu-id="96729-191">Open the **GraphHelper.cs** file.</span></span>
1. <span data-ttu-id="96729-192">Insira o código a seguir para chamar a API do Microsoft Graph para criar um novo arquivo no OneDrive.</span><span class="sxs-lookup"><span data-stu-id="96729-192">Insert the following code to call the Microsoft Graph API to create a new file on OneDrive.</span></span>

    ```csharp
    public static async Task UploadFile(string fileName, System.IO.MemoryStream stream)
        {
           var graphClient = GetAuthenticatedClient();
            await graphClient.Me
                .Drive
                .Root
                .ItemWithPath(fileName)
                .Content
                .Request()
                .PutAsync<DriveItem>(stream);
            return;
        }
    ```

1. <span data-ttu-id="96729-193">Pressione **F5** (ou **Depurar > Iniciar Depuração**).</span><span class="sxs-lookup"><span data-stu-id="96729-193">Press **F5** (or **Debug > Start Debugging**).</span></span> <span data-ttu-id="96729-194">O aplicativo Web será iniciar.</span><span class="sxs-lookup"><span data-stu-id="96729-194">The web application will start.</span></span>
1. <span data-ttu-id="96729-195">Escolha **Clique aqui para entrar** e entrar.</span><span class="sxs-lookup"><span data-stu-id="96729-195">Choose **Click here to sign in**, and sign in.</span></span>
1. <span data-ttu-id="96729-196">Escolha **Clique aqui para criar um novo arquivo em OneDrive**.</span><span class="sxs-lookup"><span data-stu-id="96729-196">Choose **Click here to create a new file on OneDrive**.</span></span>
1. <span data-ttu-id="96729-197">Abra uma nova guia do navegador e entre na sua OneDrive de usuário.</span><span class="sxs-lookup"><span data-stu-id="96729-197">Open a new browser tab and sign in to your OneDrive account.</span></span> <span data-ttu-id="96729-198">Você verá o arquivo test.txt na pasta raiz.</span><span class="sxs-lookup"><span data-stu-id="96729-198">You'll see the test.txt file in the root folder.</span></span>

<span data-ttu-id="96729-199">Agora que você aprendeu a carregar um arquivo no OneDrive, você pode reutilizar esse código para carregar qualquer documento Excel que você criar.</span><span class="sxs-lookup"><span data-stu-id="96729-199">Now that you've learned how to upload a file to OneDrive, you can reuse this code to upload any Excel document that you create.</span></span>

## <a name="additional-considerations-for-your-solution"></a><span data-ttu-id="96729-200">Considerações adicionais para sua solução</span><span class="sxs-lookup"><span data-stu-id="96729-200">Additional considerations for your solution</span></span>

<span data-ttu-id="96729-201">A solução de todos é diferente em termos de tecnologias e abordagens.</span><span class="sxs-lookup"><span data-stu-id="96729-201">Everyone’s solution is different in terms of technologies and approaches.</span></span> <span data-ttu-id="96729-202">As considerações a seguir ajudarão você a planejar como modificar sua solução para abrir documentos e incorporar seu Office Add-in.</span><span class="sxs-lookup"><span data-stu-id="96729-202">The following considerations will help you plan how to modify your solution to open documents and embed your Office Add-in.</span></span>

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a><span data-ttu-id="96729-203">Criar uma nova Excel na página da Web</span><span class="sxs-lookup"><span data-stu-id="96729-203">Create a new Excel spreadsheet from the web page</span></span>

<span data-ttu-id="96729-204">O exemplo modifica um documento Excel existente.</span><span class="sxs-lookup"><span data-stu-id="96729-204">The sample modifies an existing Excel document.</span></span> <span data-ttu-id="96729-205">Um cenário mais comum é que você criará uma nova planilha Excel de sua página da Web.</span><span class="sxs-lookup"><span data-stu-id="96729-205">A more common scenario is that you’ll create a new Excel spreadsheet from your web page.</span></span> <span data-ttu-id="96729-206">Você pode encontrar detalhes adicionais sobre como criar uma nova planilha em **Criar um documento de** planilha fornecendo um nome de arquivo.</span><span class="sxs-lookup"><span data-stu-id="96729-206">You can find additional details on how to create a new spreadsheet in **Create a spreadsheet document** by providing a file name.</span></span> <span data-ttu-id="96729-207">Este artigo mostra como criar o arquivo localmente, mas você também pode criar o arquivo em um fluxo usando uma sobrecarga no método SpreadsheetDocument.Create.</span><span class="sxs-lookup"><span data-stu-id="96729-207">This article shows how to create the file locally, but you can also create the file in a stream by using an overload on the SpreadsheetDocument.Create method.</span></span>

### <a name="read-custom-properties-when-your-add-in-starts"></a><span data-ttu-id="96729-208">Ler propriedades personalizadas quando o seu complemento for iniciado</span><span class="sxs-lookup"><span data-stu-id="96729-208">Read custom properties when your add-in starts</span></span>

<span data-ttu-id="96729-209">O exemplo de código armazena uma ID de trecho no novo documento Excel usando o SDK OOXML.</span><span class="sxs-lookup"><span data-stu-id="96729-209">The code sample stores a snippet ID in the new Excel document using the OOXML SDK.</span></span> <span data-ttu-id="96729-210">Script Lab lê a ID de trecho do documento Excel e exibe o código do trecho quando ele é aberto.</span><span class="sxs-lookup"><span data-stu-id="96729-210">Script Lab reads the snippet ID from the Excel document and then displays that snippet code when it opens.</span></span> <span data-ttu-id="96729-211">Talvez seja necessário enviar propriedades personalizadas para seu próprio complemento (como uma cadeia de caracteres de consulta ou um token de autenticação temporária).) Consulte **Persistindo o estado** e as configurações do add-in para obter detalhes completos sobre como ler propriedades personalizadas quando o seu complemento for iniciado.</span><span class="sxs-lookup"><span data-stu-id="96729-211">You may need to send custom properties to your own add-in (such as a query string, or temporary authentication token.) See **Persisting add-in state and settings** for complete details on how to read custom properties when your add-in starts.</span></span>

### <a name="initialize-the-excel-document-with-data"></a><span data-ttu-id="96729-212">Inicializar o documento Excel com dados</span><span class="sxs-lookup"><span data-stu-id="96729-212">Initialize the Excel document with data</span></span>

<span data-ttu-id="96729-213">Normalmente, quando o cliente abre um documento Excel de seu site, ele espera que o documento contenha alguns dados do site.</span><span class="sxs-lookup"><span data-stu-id="96729-213">Typically, when the customer opens up an Excel document from your web site, they expect the document to contain some data from the web site.</span></span> <span data-ttu-id="96729-214">Há algumas maneiras de gravar dados no documento.</span><span class="sxs-lookup"><span data-stu-id="96729-214">There are a couple of ways to write data into the document.</span></span>

- <span data-ttu-id="96729-215">**Use o SDK OOXML para gravar os dados**.</span><span class="sxs-lookup"><span data-stu-id="96729-215">**Use the OOXML SDK to write the data**.</span></span> <span data-ttu-id="96729-216">Você pode usar o SDK para gravar diretamente quaisquer dados no documento.</span><span class="sxs-lookup"><span data-stu-id="96729-216">You can use the SDK to directly write any data into the document.</span></span> <span data-ttu-id="96729-217">Essa abordagem será útil se você quiser que os dados sejam disponibilizados no momento em que o documento for aberto.</span><span class="sxs-lookup"><span data-stu-id="96729-217">This approach is useful if you want the data to be available the instant the document is opened.</span></span>
- <span data-ttu-id="96729-218">**Passe uma propriedade de consulta personalizada para seu Office Add-in**.</span><span class="sxs-lookup"><span data-stu-id="96729-218">**Pass a custom query property to your Office Add-in**.</span></span> <span data-ttu-id="96729-219">Ao gerar o documento, você incorpora uma propriedade personalizada para o Office que contém uma cadeia de caracteres de consulta que recupera todos os dados necessários.</span><span class="sxs-lookup"><span data-stu-id="96729-219">When you generate the document, you embed a custom property for the Office Add-in that contains a query string that retrieves all the required data.</span></span> <span data-ttu-id="96729-220">Quando o seu add-in é aberto, ele recupera a consulta, executa a consulta e usa a API JS do Office para inserir o resultado da consulta no documento.</span><span class="sxs-lookup"><span data-stu-id="96729-220">When your add-in opens, it retrieves the query, runs the query, and uses the Office JS API to insert the result of the query into the document.</span></span>

### <a name="working-with-the-ooxml-sdk"></a><span data-ttu-id="96729-221">Trabalhando com o SDK OOXML</span><span class="sxs-lookup"><span data-stu-id="96729-221">Working with the OOXML SDK</span></span>

<span data-ttu-id="96729-222">O SDK OOXML é baseado em .NET.</span><span class="sxs-lookup"><span data-stu-id="96729-222">The OOXML SDK is based on .NET.</span></span> <span data-ttu-id="96729-223">Se o aplicativo Web não for o .NET, você precisará procurar uma maneira alternativa de trabalhar com o OOXML.</span><span class="sxs-lookup"><span data-stu-id="96729-223">If your web application does not .NET, you’ll need to look for an alternative way to work with OOXML.</span></span>

<span data-ttu-id="96729-224">Há uma versão JavaScript do SDK OOXML disponível no [Open XML SDK para JavaScript](https://archive.codeplex.com/?p=openxmlsdkjs).</span><span class="sxs-lookup"><span data-stu-id="96729-224">There is a JavaScript version of the OOXML SDK available at [Open XML SDK for JavaScript](https://archive.codeplex.com/?p=openxmlsdkjs).</span></span>

<span data-ttu-id="96729-225">Você pode colocar o código OOXML em uma função do Azure para separar o código .NET do restante do seu aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="96729-225">You can place the OOXML code in an Azure function to separate the .NET code from the rest of your web application.</span></span> <span data-ttu-id="96729-226">Em seguida, chame a função Azure (para gerar o documento Excel) do aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="96729-226">Then call the Azure function (to generate the Excel document) from your Web application.</span></span> <span data-ttu-id="96729-227">Para obter mais informações sobre as funções do Azure, consulte [Uma introdução às funções do Azure](/azure/azure-functions/functions-overview).</span><span class="sxs-lookup"><span data-stu-id="96729-227">For more information on Azure functions, see [An introduction to Azure Functions](/azure/azure-functions/functions-overview).</span></span>

### <a name="use-single-sign-on"></a><span data-ttu-id="96729-228">Usar o login único</span><span class="sxs-lookup"><span data-stu-id="96729-228">Use single sign-on</span></span>

<span data-ttu-id="96729-229">Para simplificar a autenticação, recomendamos que seu complemento implemente o login único.</span><span class="sxs-lookup"><span data-stu-id="96729-229">To simplify authentication, we recommend your add-in implements single sign-on.</span></span> <span data-ttu-id="96729-230">Para obter mais informações, consulte [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="96729-230">For more information, see [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md)</span></span>

## <a name="see-also"></a><span data-ttu-id="96729-231">Confira também</span><span class="sxs-lookup"><span data-stu-id="96729-231">See also</span></span>

- [<span data-ttu-id="96729-232">Bem-vindo ao SDK Open XML 2.5 para Office</span><span class="sxs-lookup"><span data-stu-id="96729-232">Welcome to the Open XML SDK 2.5 for Office</span></span>](/office/open-xml/open-xml-sdk)
- [<span data-ttu-id="96729-233">Abrir automaticamente um painel de tarefas com um documento</span><span class="sxs-lookup"><span data-stu-id="96729-233">Automatically open a task pane with a document</span></span>](../develop/automatically-open-a-task-pane-with-a-document.md)
- [<span data-ttu-id="96729-234">Persistir o estado e as configurações do suplemento</span><span class="sxs-lookup"><span data-stu-id="96729-234">Persisting add-in state and settings</span></span>](../develop/persisting-add-in-state-and-settings.md)
- [<span data-ttu-id="96729-235">Criar um documento de planilha fornecendo um nome de arquivo</span><span class="sxs-lookup"><span data-stu-id="96729-235">Create a spreadsheet document by providing a file name</span></span>](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)