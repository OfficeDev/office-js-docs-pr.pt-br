---
title: Tutorial de suplemento do PowerPoint
description: Neste tutorial, você criará um suplemento do PowerPoint que insere imagem, texto, obtém metadados do slide e navega entre slides.
ms.date: 05/12/2021
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: 76d40a83155a7a26822b43dc1340e3f9ebda63da
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349403"
---
# <a name="tutorial-create-a-powerpoint-task-pane-add-in"></a><span data-ttu-id="74ba4-103">Tutorial: Criar um Suplemento do Painel de Tarefas</span><span class="sxs-lookup"><span data-stu-id="74ba4-103">Tutorial: Create a PowerPoint task pane add-in</span></span>

<span data-ttu-id="74ba4-104">Neste tutorial, você usará o Visual Studio para criar um Suplementos do Painel de Tarefas do PowerPoint:</span><span class="sxs-lookup"><span data-stu-id="74ba4-104">In this tutorial, you'll use Visual Studio to create an PowerPoint task pane add-in that:</span></span>

> [!div class="checklist"]
>
> - <span data-ttu-id="74ba4-105">Adicionar a foto do dia do [Bing](https://www.bing.com) a um slide</span><span class="sxs-lookup"><span data-stu-id="74ba4-105">Adds the [Bing](https://www.bing.com) photo of the day to a slide</span></span>
> - <span data-ttu-id="74ba4-106">Adicionar texto a um slide</span><span class="sxs-lookup"><span data-stu-id="74ba4-106">Adds text to a slide</span></span>
> - <span data-ttu-id="74ba4-107">Obtém metadados do slide</span><span class="sxs-lookup"><span data-stu-id="74ba4-107">Gets slide metadata</span></span>
> - <span data-ttu-id="74ba4-108">Navega entre slides</span><span class="sxs-lookup"><span data-stu-id="74ba4-108">Navigates between slides</span></span>

## <a name="prerequisites"></a><span data-ttu-id="74ba4-109">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="74ba4-109">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

## <a name="create-your-add-in-project"></a><span data-ttu-id="74ba4-110">Criar seu projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="74ba4-110">Create your add-in project</span></span>

<span data-ttu-id="74ba4-111">Conclua as etapas a seguir para criar um projeto de suplemento do PowerPoint usando o Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="74ba4-111">Complete the following steps to create a PowerPoint add-in project using Visual Studio.</span></span>

1. <span data-ttu-id="74ba4-112">Escolha **Criar um novo projeto**.</span><span class="sxs-lookup"><span data-stu-id="74ba4-112">Choose **Create a new project**.</span></span>

2. <span data-ttu-id="74ba4-113">Usando a caixa de pesquisa, insira **suplemento**.</span><span class="sxs-lookup"><span data-stu-id="74ba4-113">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="74ba4-114">Escolha **Suplemento do PowerPoint Web**, em seguida, selecione **Próximo**.</span><span class="sxs-lookup"><span data-stu-id="74ba4-114">Choose **PowerPoint Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="74ba4-115">Nomeie o projeto como `HelloWorld` e selecione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="74ba4-115">Name the project `HelloWorld`, and select **Create**.</span></span>

4. <span data-ttu-id="74ba4-116">Na caixa de diálogo **Criar suplementos do Office**, escolha **Adicionar novas funcionalidades ao PowerPoint** e depois **Concluir** para criar o projeto.</span><span class="sxs-lookup"><span data-stu-id="74ba4-116">In the **Create Office Add-in** dialog window, choose **Add new functionalities to PowerPoint**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="74ba4-p102">O Visual Studio cria uma solução, e os dois projetos dele são exibidos no **Gerenciador de Soluções**. O arquivo **Home.html** é aberto no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="74ba4-p102">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

     ![A captura de tela da janela do gerenciador de soluções do Microsoft Visual Studio mostrando HelloWorld e HelloWorldWeb, os 2 projetos na solução HelloWorld.](../images/powerpoint-tutorial-solution-explorer.png)

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="74ba4-120">Explorar a solução do Visual Studio</span><span class="sxs-lookup"><span data-stu-id="74ba4-120">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-code"></a><span data-ttu-id="74ba4-121">Código de atualização</span><span class="sxs-lookup"><span data-stu-id="74ba4-121">Update code</span></span>

<span data-ttu-id="74ba4-122">Edite o código do suplemento como mostrado a seguir para criar a estrutura que você usará para implementar a funcionalidade do suplemento nas etapas subsequentes deste tutorial.</span><span class="sxs-lookup"><span data-stu-id="74ba4-122">Edit the add-in code as follows to create the framework that you'll use to implement add-in functionality in subsequent steps of this tutorial.</span></span>

1. <span data-ttu-id="74ba4-p103">**Home.html** especifica o HTML que será renderizado no painel de tarefas do suplemento. Em **Home.html**, encontre a **div** com `id="content-main"`, substitua toda essa **div** pela marcação a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="74ba4-p103">**Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, find the **div** with `id="content-main"`, replace that entire **div** with the following markup, and save the file.</span></span>

    ```html
    <!-- TODO2: Create the content-header div. -->
    <div id="content-main">
        <div class="padding">
            <!-- TODO1: Create the insert-image button. -->
            <!-- TODO3: Create the insert-text button. -->
            <!-- TODO4: Create the get-slide-metadata button. -->
            <!-- TODO5: Create the go-to-slide buttons. -->
        </div>
    </div>
    ```

2. <span data-ttu-id="74ba4-p104">Abra o arquivo **Home.js** na raiz do projeto do aplicativo Web. Este arquivo especifica o script do suplemento. Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="74ba4-p104">Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.</span></span>

    ```js
    (function () {
        "use strict";

        var messageBanner;

        Office.onReady(function () {
            $(document).ready(function () {
                // Initialize the FabricUI notification mechanism and hide it
                var element = document.querySelector('.MessageBanner');
                messageBanner = new components.MessageBanner(element);
                messageBanner.hideBanner();

                // TODO1: Assign event handler for insert-image button.
                // TODO4: Assign event handler for insert-text button.
                // TODO6: Assign event handler for get-slide-metadata button.
                // TODO8: Assign event handlers for the four navigation buttons.
            });
        });

        // TODO2: Define the insertImage function. 

        // TODO3: Define the insertImageFromBase64String function.

        // TODO5: Define the insertText function.

        // TODO7: Define the getSlideMetadata function.

        // TODO9: Define the navigation functions.

        // Helper function for displaying notifications
        function showNotification(header, content) {
            $("#notification-header").text(header);
            $("#notification-body").text(content);
            messageBanner.showBanner();
            messageBanner.toggleExpansion();
        }
    })();
    ```

## <a name="insert-an-image"></a><span data-ttu-id="74ba4-128">Inserir uma imagem</span><span class="sxs-lookup"><span data-stu-id="74ba4-128">Insert an image</span></span>

<span data-ttu-id="74ba4-129">Conclua as seguintes etapas para adicionar o código que recupera a foto do dia do [Bing](https://www.bing.com) e inserir as imagens em um slide.</span><span class="sxs-lookup"><span data-stu-id="74ba4-129">Complete the following steps to add code that retrieves the [Bing](https://www.bing.com) photo of the day and inserts that image into a slide.</span></span>

1. <span data-ttu-id="74ba4-130">Usando o Explorador de soluções, adicione uma nova pasta chamada **Controladores** ao projeto **HelloWorldWeb**.</span><span class="sxs-lookup"><span data-stu-id="74ba4-130">Using Solution Explorer, add a new folder named **Controllers** to the **HelloWorldWeb** project.</span></span>

    ![A captura de tela da janela do Gerenciador de Soluções do Microsoft Visual Studio mostrando a pasta Controladores realçada no projeto HelloWorldWeb.](../images/powerpoint-tutorial-solution-explorer-controllers.png)

2. <span data-ttu-id="74ba4-132">Clique com o botão direito do mouse na pasta **Controladores** e selecione **Adicionar > Novo item com scaffold...**.</span><span class="sxs-lookup"><span data-stu-id="74ba4-132">Right-click the **Controllers** folder and select **Add > New Scaffolded Item...**.</span></span>

3. <span data-ttu-id="74ba4-133">Na janela da caixa de diálogo **Adicionar Scaffold**, selecione **Controlador da Web API 2 – vazio** e escolha o botão **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="74ba4-133">In the **Add Scaffold** dialog window, select **Web API 2 Controller - Empty** and choose the **Add** button.</span></span> 

4. <span data-ttu-id="74ba4-p105">Na janela da caixa de diálogo **Adicionar Controlador**, insira **PhotoController** como nome do controlador e escolha o botão **Adicionar**. O Visual Studio criará e abrirá o arquivo **PhotoController.cs**.</span><span class="sxs-lookup"><span data-stu-id="74ba4-p105">In the **Add Controller** dialog window, enter **PhotoController** as the controller name and choose the **Add** button. Visual Studio creates and opens the **PhotoController.cs** file.</span></span>

5. <span data-ttu-id="74ba4-p106">Substitua todo o conteúdo do arquivo **PhotoController.cs** pelo código a seguir, que chama o serviço do Bing para recuperar a foto do dia como uma cadeia de caracteres com codificação Base64. Quando você usar a API JavaScript do Office para inserir uma imagem em um documento, especifique os dados de imagem como uma cadeia de caracteres com codificação Base64.</span><span class="sxs-lookup"><span data-stu-id="74ba4-p106">Replace the entire contents of the **PhotoController.cs** file with the following code that calls the Bing service to retrieve the photo of the day as a Base64 encoded string. When you use the Office JavaScript API to insert an image into a document, the image data must be specified as a Base64 encoded string.</span></span>

    ```csharp
    using System;
    using System.IO;
    using System.Net;
    using System.Text;
    using System.Web.Http;
    using System.Xml;

    namespace HelloWorldWeb.Controllers
    {
        public class PhotoController : ApiController
        {
            public string Get()
            {
                string url = "http://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1";

                // Create the request.
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                WebResponse response = request.GetResponse();

                using (Stream responseStream = response.GetResponseStream())
                {
                    // Process the result.
                    StreamReader reader = new StreamReader(responseStream, Encoding.UTF8);
                    string result = reader.ReadToEnd();

                    // Parse the xml response and to get the URL.
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(result);
                    string photoURL = "http://bing.com" + doc.SelectSingleNode("/images/image/url").InnerText;

                    // Fetch the photo and return it as a Base64 encoded string.
                    return getPhotoFromURL(photoURL);
                }
            }

            private string getPhotoFromURL(string imageURL)
            {
                var webClient = new WebClient();
                byte[] imageBytes = webClient.DownloadData(imageURL);
                return Convert.ToBase64String(imageBytes);
            }
        }
    }
    ```

6. <span data-ttu-id="74ba4-p107">No arquivo **Home.html**, substitua `TODO1` pela marcação a seguir. Essa marcação define o botão **Inserir Imagem** que aparecerá no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="74ba4-p107">In the **Home.html** file, replace `TODO1` with the following markup. This markup defines the **Insert Image** button that will appear within the add-in's task pane.</span></span>

    ```html
    <button class="Button Button--primary" id="insert-image">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Insert Image</span>
        <span class="Button-description">Gets the photo of the day that shows on the Bing home page and adds it to the slide.</span>
    </button>
    ```

7. <span data-ttu-id="74ba4-140">No arquivo **Home.js**, substitua `TODO1` pelo código a seguir para atribuir o manipulador de eventos ao botão **Inserir Imagem**.</span><span class="sxs-lookup"><span data-stu-id="74ba4-140">In the **Home.js** file, replace `TODO1` with the following code to assign the event handler for the **Insert Image** button.</span></span>

    ```js
    $('#insert-image').click(insertImage);
    ```

8. <span data-ttu-id="74ba4-p108">No arquivo **Home.js**, substitua `TODO2` pelo código a seguir para definir a função `insertImage`. Esta função busca a imagem do serviço Web Bing e chama a função `insertImageFromBase64String` para inserir a imagem no documento.</span><span class="sxs-lookup"><span data-stu-id="74ba4-p108">In the **Home.js** file, replace `TODO2` with the following code to define the `insertImage` function. This function fetches the image from the Bing web service and then calls the `insertImageFromBase64String` function to insert that image into the document.</span></span>

    ```js
    function insertImage() {
        // Get image from from web service (as a Base64 encoded string).
        $.ajax({
            url: "/api/Photo/", success: function (result) {
                insertImageFromBase64String(result);
            }, error: function (xhr, status, error) {
                showNotification("Error", "Oops, something went wrong.");
            }
        });
    }
    ```

9. <span data-ttu-id="74ba4-p109">No arquivo **Home.js**, substitua `TODO3` pelo código a seguir para definir a função `insertImageFromBase64String`. Esta função usa a API JavaScript do Office para inserir a imagem no documento. Observação:</span><span class="sxs-lookup"><span data-stu-id="74ba4-p109">In the **Home.js** file, replace `TODO3` with the following code to define the `insertImageFromBase64String` function. This function uses the Office JavaScript API to insert the image into the document. Note:</span></span>

    - <span data-ttu-id="74ba4-146">A opção `coercionType` especificada como segundo parâmetro da solicitação `setSelectedDataAsync` indica o tipo de dados inserido.</span><span class="sxs-lookup"><span data-stu-id="74ba4-146">The `coercionType` option that's specified as the second parameter of the `setSelectedDataAsync` request indicates the type of data being inserted.</span></span>

    - <span data-ttu-id="74ba4-147">O objeto `asyncResult` encapsula o resultado da solicitação `setSelectedDataAsync`, incluindo informações de status e de erro caso a solicitação tenha falhado.</span><span class="sxs-lookup"><span data-stu-id="74ba4-147">The `asyncResult` object encapsulates the result of the `setSelectedDataAsync` request, including status and error information if the request failed.</span></span>

    ```js
    function insertImageFromBase64String(image) {
        // Call Office.js to insert the image into the document.
        Office.context.document.setSelectedDataAsync(image, {
            coercionType: Office.CoercionType.Image
        },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="74ba4-148">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="74ba4-148">Test the add-in</span></span>

1. <span data-ttu-id="74ba4-p110">Usando o Visual Studio, teste o suplemento do PowerPoint recém-criado, pressionando **F5** ou escolhendo o botão **Iniciar** para abrir o PowerPoint com o botão de suplemento **Mostrar Painel de Tarefas** exibido na faixa de opções. O suplemento será hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="74ba4-p110">Using Visual Studio, test the newly created PowerPoint add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![A captura de tela mostrando o botão Iniciar realçado no Microsoft Visual Studio.](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="74ba4-152">No PowerPoint, selecione o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="74ba4-152">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Captura de tela mostrando o botão Mostrar painel de tarefas realçado na faixa de opções página inicial do PowerPoint.](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="74ba4-154">No painel de tarefas, escolha o botão **Inserir Imagem** para adicionar a foto do dia do Bing ao slide atual.</span><span class="sxs-lookup"><span data-stu-id="74ba4-154">In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide.</span></span>

    ![Captura de tela do suplemento do PowerPoint com o botão Inserir imagem realçado.](../images/powerpoint-tutorial-insert-image-button.png)

4. <span data-ttu-id="74ba4-156">No Visual Studio, interrompa o suplemento pressionando **Shift + F5** ou selecionando o botão **Parar**.</span><span class="sxs-lookup"><span data-stu-id="74ba4-156">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="74ba4-157">O PowerPoint fechará automaticamente quando o suplemento for interrompido.</span><span class="sxs-lookup"><span data-stu-id="74ba4-157">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![Captura de tela mostrando o botão Pare realçado no Microsoft Visual Studio.](../images/powerpoint-tutorial-stop.png)

## <a name="customize-user-interface-ui-elements"></a><span data-ttu-id="74ba4-159">Personalizar os elementos da IU (interface do usuário)</span><span class="sxs-lookup"><span data-stu-id="74ba4-159">Customize User Interface (UI) elements</span></span>

<span data-ttu-id="74ba4-160">Conclua as seguintes etapas para adicionar a marca que personaliza o painel de tarefas da interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="74ba4-160">Complete the following steps to add markup that customizes the task pane UI.</span></span>

1. <span data-ttu-id="74ba4-p112">No arquivo **Home.html**, substitua `TODO2` pela marcação a seguir para adicionar uma seção de cabeçalho e um título ao painel de tarefas. Observação:</span><span class="sxs-lookup"><span data-stu-id="74ba4-p112">In the **Home.html** file, replace `TODO2` with the following markup to add a header section and title to the task pane. Note:</span></span>

    - <span data-ttu-id="74ba4-163">Os estilos que começam com `ms-` são definidos pelo [Fabric Core em Suplementos do Office ](../design/fabric-core.md), uma estrutura de front-end JavaScript para construir experiências do usuário para o Office.</span><span class="sxs-lookup"><span data-stu-id="74ba4-163">The styles that begin with `ms-` are defined by [Fabric Core in Office Add-ins](../design/fabric-core.md), a JavaScript front-end framework for building user experiences for Office.</span></span> <span data-ttu-id="74ba4-164">O arquivo **Home.html** inclui uma referência à folha de estilo do Fabric Core.</span><span class="sxs-lookup"><span data-stu-id="74ba4-164">The **Home.html** file includes a reference to the Fabric Core stylesheet.</span></span>

    ```html
    <div id="content-header">
        <div class="ms-Grid ms-bgColor-neutralPrimary">
            <div class="ms-Grid-row">
                <div class="padding ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"> <div class="ms-font-xl ms-fontColor-white ms-fontWeight-semibold">My PowerPoint add-in</div></div>
            </div>
        </div>
    </div>
    ```

2. <span data-ttu-id="74ba4-165">No arquivo **Home.html**, localize a **div** com `class="footer"` e exclua toda a **div** para remover a seção de rodapé do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="74ba4-165">In the **Home.html** file, find the **div** with `class="footer"` and delete that entire **div** to remove the footer section from the task pane.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="74ba4-166">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="74ba4-166">Test the add-in</span></span>

1. <span data-ttu-id="74ba4-167">Usando o Visual Studio, teste o suplemento do PowerPoint ao pressionar **F5** ou escolha o botão **Iniciar** para abrir o PowerPoint com o botão de suplemento **Mostrar painel de tarefas** exibido na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="74ba4-167">Using Visual Studio, test the PowerPoint add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="74ba4-168">O suplemento será hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="74ba4-168">The add-in will be hosted locally on IIS.</span></span>

    ![Captura de tela mostrando o botão Iniciar realçado no Visual Studio.](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="74ba4-170">No PowerPoint, selecione o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="74ba4-170">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Captura de tela mostrando o botão Mostrar Painel de tarefas realçado na faixa de opções da Página Inicial do PowerPoint.](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="74ba4-172">Observe que agora o painel de tarefas contém uma seção de cabeçalho e um título e não contém mais uma seção de rodapé.</span><span class="sxs-lookup"><span data-stu-id="74ba4-172">Notice that the task pane now contains a header section and title, and no longer contains a footer section.</span></span>

    ![Captura de tela do suplemento do PowerPoint com o botão Inserir Imagem realçado.](../images/powerpoint-tutorial-new-task-pane-ui.png)

4. <span data-ttu-id="74ba4-174">No Visual Studio, interrompa o suplemento pressionando **Shift + F5** ou selecionando o botão **Parar**.</span><span class="sxs-lookup"><span data-stu-id="74ba4-174">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="74ba4-175">O PowerPoint fechará automaticamente quando o suplemento for interrompido.</span><span class="sxs-lookup"><span data-stu-id="74ba4-175">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![Captura de tela exibindo o botão Pare realçado no Microsoft Visual Studio.](../images/powerpoint-tutorial-stop.png)

## <a name="insert-text"></a><span data-ttu-id="74ba4-177">Inserir texto</span><span class="sxs-lookup"><span data-stu-id="74ba4-177">Insert text</span></span>

<span data-ttu-id="74ba4-178">Conclua as seguintes etapas para adicionar o código que insere texto no slide de título que contém as fotos do dia do [Bing](https://www.bing.com).</span><span class="sxs-lookup"><span data-stu-id="74ba4-178">Complete the following steps to add code that inserts text into the title slide which contains the [Bing](https://www.bing.com) photo of the day.</span></span>

1. <span data-ttu-id="74ba4-p116">No arquivo **Home.html**, substitua `TODO3` pela marcação a seguir. Essa marcação define o botão **Inserir Texto** que aparecerá no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="74ba4-p116">In the **Home.html** file, replace `TODO3` with the following markup. This markup defines the **Insert Text** button that will appear within the add-in's task pane.</span></span>

    ```html
        <br /><br />
        <button class="Button Button--primary" id="insert-text">
            <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
            <span class="Button-label">Insert Text</span>
            <span class="Button-description">Inserts text into the slide.</span>
        </button>
    ```

2. <span data-ttu-id="74ba4-181">No arquivo **Home.js**, substitua `TODO4` pelo código a seguir para atribuir o manipulador de eventos ao botão **Inserir Texto**.</span><span class="sxs-lookup"><span data-stu-id="74ba4-181">In the **Home.js** file, replace `TODO4` with the following code to assign the event handler for the **Insert Text** button.</span></span>

    ```js
    $('#insert-text').click(insertText);
    ```

3. <span data-ttu-id="74ba4-p117">No arquivo **Home.js**, substitua `TODO5` pelo código a seguir para definir a função `insertText`. Esta função insere texto no slide atual.</span><span class="sxs-lookup"><span data-stu-id="74ba4-p117">In the **Home.js** file, replace `TODO5` with the following code to define the `insertText` function. This function inserts text into the current slide.</span></span>

    ```js
    function insertText() {
        Office.context.document.setSelectedDataAsync('Hello World!',
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="74ba4-184">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="74ba4-184">Test the add-in</span></span>

1. <span data-ttu-id="74ba4-185">Usando o Visual Studio, teste o suplemento ao pressionar **F5** ou escolha o botão **Iniciar** para abrir o PowerPoint com o botão de suplemento **Mostrar painel de tarefas** exibido na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="74ba4-185">Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="74ba4-186">O suplemento será hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="74ba4-186">The add-in will be hosted locally on IIS.</span></span>

    ![Captura de tela do Microsoft Visual Studio com o botão Iniciar realçado.](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="74ba4-188">No PowerPoint, selecione o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="74ba4-188">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Captura de tela realçando o botão Mostrar Painel de tarefas na faixa de opções da Página inicial no PowerPoint.](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="74ba4-190">No painel de tarefas, escolha o botão **Inserir Imagem** para adicionar a foto do dia do Bing ao slide atual e escolher um design para o slide que contém uma caixa de texto como título.</span><span class="sxs-lookup"><span data-stu-id="74ba4-190">In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide and choose a design for the slide that contains a text box for the title.</span></span>

    ![Captura de tela do PowerPoint com o slide atual realçado e suplemento com o botão Inserir Imagem realçado.](../images/powerpoint-tutorial-insert-image-slide-design.png)

4. <span data-ttu-id="74ba4-192">Coloque o cursor na caixa de texto no slide de título e depois, no painel de tarefas, escolha o botão **Inserir Texto** para adicionar texto ao slide.</span><span class="sxs-lookup"><span data-stu-id="74ba4-192">Put your cursor in the text box on the title slide and then in the task pane, choose the **Insert Text** button to add text to the slide.</span></span>

    ![Captura de tela do suplemento do PowerPoint com o botão Inserir Texto realçado.](../images/powerpoint-tutorial-insert-text.png)

5. <span data-ttu-id="74ba4-194">No Visual Studio, interrompa o suplemento pressionando **Shift + F5** ou selecionando o botão **Parar**.</span><span class="sxs-lookup"><span data-stu-id="74ba4-194">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="74ba4-195">O PowerPoint fechará automaticamente quando o suplemento for interrompido.</span><span class="sxs-lookup"><span data-stu-id="74ba4-195">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![Captura de tela do Microsoft Visual Studio com o botão Pare realçado.](../images/powerpoint-tutorial-stop.png)

## <a name="get-slide-metadata"></a><span data-ttu-id="74ba4-197">Obter metadados do slide</span><span class="sxs-lookup"><span data-stu-id="74ba4-197">Get slide metadata</span></span>

<span data-ttu-id="74ba4-198">Conclua as seguintes etapas para adicionar o código que recupera os metadados para o slide selecionado.</span><span class="sxs-lookup"><span data-stu-id="74ba4-198">Complete the following steps to add code that retrieves metadata for the selected slide.</span></span>

1. <span data-ttu-id="74ba4-p120">No arquivo **Home.html**, substitua `TODO4` pela marcação a seguir. Essa marcação define o botão **Obter metadados do slide** que aparecerá no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="74ba4-p120">In the **Home.html** file, replace `TODO4` with the following markup. This markup defines the **Get Slide Metadata** button that will appear within the add-in's task pane.</span></span>

    ```html
    <br /><br />
    <button class="Button Button--primary" id="get-slide-metadata">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Get Slide Metadata</span>
        <span class="Button-description">Gets metadata for the selected slide(s).</span>
    </button>
    ```

2. <span data-ttu-id="74ba4-201">No arquivo **Home.js**, substitua `TODO6` pelo código a seguir para atribuir o manipulador de eventos para o botão **Obter Metadados do Slide**.</span><span class="sxs-lookup"><span data-stu-id="74ba4-201">In the **Home.js** file, replace `TODO6` with the following code to assign the event handler for the **Get Slide Metadata** button.</span></span>

    ```js
    $('#get-slide-metadata').click(getSlideMetadata);
    ```

3. <span data-ttu-id="74ba4-p121">No arquivo **Home.js**, substitua `TODO7` pelo código a seguir para definir a função `getSlideMetadata`. Esta função recupera metadados dos slides selecionados e os grava em uma janela pop-up da caixa de diálogo no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="74ba4-p121">In the **Home.js** file, replace `TODO7` with the following code to define the `getSlideMetadata` function. This function retrieves metadata for the selected slide(s) and writes it to a popup dialog window within the add-in task pane.</span></span>

    ```js
    function getSlideMetadata() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                } else {
                    showNotification("Metadata for selected slide(s):", JSON.stringify(asyncResult.value), null, 2);
                }
            }
        );
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="74ba4-204">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="74ba4-204">Test the add-in</span></span>

1. <span data-ttu-id="74ba4-205">Usando o Visual Studio, teste o suplemento ao pressionar **F5** ou escolha o botão **Iniciar** para abrir o PowerPoint com o botão de suplemento **Mostrar painel de tarefas** exibido na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="74ba4-205">Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="74ba4-206">O suplemento será hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="74ba4-206">The add-in will be hosted locally on IIS.</span></span>

    ![Captura de tela realçando o botão Iniciar no Microsoft Visual Studio.](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="74ba4-208">No PowerPoint, selecione o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="74ba4-208">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Captura de tela realçando o botão Mostrar painel de tarefas na faixa de opções da Página Inicial do PowerPoint.](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="74ba4-p123">No painel de tarefas, escolha o botão **Obter metadados do slide** para obter os metadados do slide selecionado. Os metadados do slide serão gravados na janela pop-up da caixa de diálogo na parte inferior do painel de tarefas. Nesse caso, a matriz `slides` dos metadados JSON contém um objeto que especifica `id`, `title` e `index` do slide selecionado. Se vários slides tivessem sido selecionados na recuperação de metadados do slide, a matriz `slides` dos metadados JSON conteria um objeto para cada slide selecionado.</span><span class="sxs-lookup"><span data-stu-id="74ba4-p123">In the task pane, choose the **Get Slide Metadata** button to get the metadata for the selected slide. The slide metadata is written to the popup dialog window at the bottom of the task pane. In this case, the `slides` array within the JSON metadata contains one object that specifies the `id`, `title`, and `index` of the selected slide. If multiple slides had been selected when you retrieved slide metadata, the `slides` array within the JSON metadata would contain one object for each selected slide.</span></span>

    ![Captura de tela do suplemento do PowerPoint com o botão Obter metadados do slide realçado.](../images/powerpoint-tutorial-get-slide-metadata.png)

4. <span data-ttu-id="74ba4-215">No Visual Studio, interrompa o suplemento pressionando **Shift + F5** ou selecionando o botão **Parar**.</span><span class="sxs-lookup"><span data-stu-id="74ba4-215">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="74ba4-216">O PowerPoint fechará automaticamente quando o suplemento for interrompido.</span><span class="sxs-lookup"><span data-stu-id="74ba4-216">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![Captura de tela realçando o botão Pare no Microsoft Visual Studio.](../images/powerpoint-tutorial-stop.png)

## <a name="navigate-between-slides"></a><span data-ttu-id="74ba4-218">Navegar entre slides</span><span class="sxs-lookup"><span data-stu-id="74ba4-218">Navigate between slides</span></span>

<span data-ttu-id="74ba4-219">Conclua as seguintes etapas para adicionar o código que navega entre os slides de um documento.</span><span class="sxs-lookup"><span data-stu-id="74ba4-219">Complete the following steps to add code that navigates between the slides of a document.</span></span>

1. <span data-ttu-id="74ba4-p125">No arquivo **Home.html**, substitua `TODO5` pela marcação a seguir. Essa marcação define os quatro botões de navegação que aparecerão no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="74ba4-p125">In the **Home.html** file, replace `TODO5` with the following markup. This markup defines the four navigation buttons that will appear within the add-in's task pane.</span></span>

    ```html
    <br /><br />
    <button class="Button Button--primary" id="go-to-first-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to First Slide</span>
        <span class="Button-description">Go to the first slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-next-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Next Slide</span>
        <span class="Button-description">Go to the next slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-previous-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Previous Slide</span>
        <span class="Button-description">Go to the previous slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-last-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Last Slide</span>
        <span class="Button-description">Go to the last slide.</span>
    </button>
    ```

2. <span data-ttu-id="74ba4-222">No arquivo **Home.js**, substitua `TODO8` pelo código a seguir para atribuir o manipulador de eventos aos quatro botões de navegação.</span><span class="sxs-lookup"><span data-stu-id="74ba4-222">In the **Home.js** file, replace `TODO8` with the following code to assign the event handlers for the four navigation buttons.</span></span>

    ```js
    $('#go-to-first-slide').click(goToFirstSlide);
    $('#go-to-next-slide').click(goToNextSlide);
    $('#go-to-previous-slide').click(goToPreviousSlide);
    $('#go-to-last-slide').click(goToLastSlide);
    ```

3. <span data-ttu-id="74ba4-223">No arquivo **Home.js**, substitua `TODO9` pelo código a seguir para definir as funções de navegação.</span><span class="sxs-lookup"><span data-stu-id="74ba4-223">In the **Home.js** file, replace `TODO9` with the following code to define the navigation functions.</span></span> <span data-ttu-id="74ba4-224">Cada uma dessas funções usa a função `goToByIdAsync` para selecionar um slide com base em sua posição no documento (primeiro, último, anterior e próximo).</span><span class="sxs-lookup"><span data-stu-id="74ba4-224">Each of these functions uses the `goToByIdAsync` function to select a slide based upon its position in the document (first, last, previous, and next).</span></span>

    ```js
    function goToFirstSlide() {
        Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToLastSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToPreviousSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToNextSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="74ba4-225">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="74ba4-225">Test the add-in</span></span>

1. <span data-ttu-id="74ba4-226">Usando o Visual Studio, teste o suplemento ao pressionar **F5** ou escolha o botão **Iniciar** para abrir o PowerPoint com o botão de suplemento **Mostrar painel de tarefas** exibido na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="74ba4-226">Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="74ba4-227">O suplemento será hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="74ba4-227">The add-in will be hosted locally on IIS.</span></span>

    ![Captura de tela mostrando o botão Iniciar realçado na barra de ferramentas do Microsoft Visual Studio.](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="74ba4-229">No PowerPoint, selecione o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="74ba4-229">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Captura de tela mostrando o botão Mostrar Painel de tarefas realçado na faixa de opções da Página inicial no PowerPoint.](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="74ba4-231">Use o botão **Novo Slide** na faixa de opções da guia **Página Inicial** para adicionar dois novos slides ao documento.</span><span class="sxs-lookup"><span data-stu-id="74ba4-231">Use the **New Slide** button in the ribbon of the **Home** tab to add two new slides to the document.</span></span>

4. <span data-ttu-id="74ba4-p128">No painel de tarefas, escolha o botão **Ir para o primeiro Slide**. O primeiro slide no documento é selecionado e exibido.</span><span class="sxs-lookup"><span data-stu-id="74ba4-p128">In the task pane, choose the **Go to First Slide** button. The first slide in the document is selected and displayed.</span></span>

    ![Captura de tela do suplemento do PowerPoint com o botão Ir Para o Primeiro Slide realçado.](../images/powerpoint-tutorial-go-to-first-slide.png)

5. <span data-ttu-id="74ba4-p129">No painel de tarefas, escolha o botão **Ir para o próximo Slide**. O próximo slide no documento é selecionado e exibido.</span><span class="sxs-lookup"><span data-stu-id="74ba4-p129">In the task pane, choose the **Go to Next Slide** button. The next slide in the document is selected and displayed.</span></span>

    ![Captura de tela do suplemento do PowerPoint com o botão Ir Para o Próximo Slide realçado.](../images/powerpoint-tutorial-go-to-next-slide.png)

6. <span data-ttu-id="74ba4-p130">No painel de tarefas, escolha o botão **Ir Para o Slide Anterior**. O slide anterior no documento é selecionado e exibido.</span><span class="sxs-lookup"><span data-stu-id="74ba4-p130">In the task pane, choose the **Go to Previous Slide** button. The previous slide in the document is selected and displayed.</span></span>

    ![Captura de tela do suplemento do PowerPoint com o botão Ir Para o Slide Anterior realçado.](../images/powerpoint-tutorial-go-to-previous-slide.png)

7. <span data-ttu-id="74ba4-p131">No painel de tarefas, escolha o botão **Ir Para o Último Slide**. O último slide no documento é selecionado e exibido.</span><span class="sxs-lookup"><span data-stu-id="74ba4-p131">In the task pane, choose the **Go to Last Slide** button. The last slide in the document is selected and displayed.</span></span>

    ![Captura de tela do suplemento do PowerPoint com o botão Ir Para o Último Slide realçado.](../images/powerpoint-tutorial-go-to-last-slide.png)

8. <span data-ttu-id="74ba4-244">No Visual Studio, interrompa o suplemento pressionando **Shift + F5** ou selecionando o botão **Parar**.</span><span class="sxs-lookup"><span data-stu-id="74ba4-244">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="74ba4-245">O PowerPoint fechará automaticamente quando o suplemento for interrompido.</span><span class="sxs-lookup"><span data-stu-id="74ba4-245">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![Captura de tela mostrando o botão Pare realçado na barra de ferramentas do Microsoft Visual Studio.](../images/powerpoint-tutorial-stop.png)

## <a name="next-steps"></a><span data-ttu-id="74ba4-247">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="74ba4-247">Next steps</span></span>

<span data-ttu-id="74ba4-248">Neste tutorial, você criou um suplemento do PowerPoint que insere imagem, texto, obtém metadados do slide e navega entre slides.</span><span class="sxs-lookup"><span data-stu-id="74ba4-248">In this tutorial, you've created a PowerPoint add-in that inserts an image, inserts text, gets slide metadata, and navigates between slides.</span></span> <span data-ttu-id="74ba4-249">Para saber mais sobre como criar suplementos do PowerPoint, prossiga para o artigo a seguir.</span><span class="sxs-lookup"><span data-stu-id="74ba4-249">To learn more about building PowerPoint add-ins, continue to the following article.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="74ba4-250">Visão geral dos Suplementos do SharePoint</span><span class="sxs-lookup"><span data-stu-id="74ba4-250">PowerPoint add-ins overview</span></span>](../powerpoint/powerpoint-add-ins.md)

## <a name="see-also"></a><span data-ttu-id="74ba4-251">Confira também</span><span class="sxs-lookup"><span data-stu-id="74ba4-251">See also</span></span>

- [<span data-ttu-id="74ba4-252">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="74ba4-252">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="74ba4-253">Desenvolver Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="74ba4-253">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
