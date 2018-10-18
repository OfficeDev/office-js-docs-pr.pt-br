<span data-ttu-id="5a0ce-101">Neste tutorial, comece configurando seu projeto de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="5a0ce-101">You'll begin this tutorial by setting up your development project.</span></span> 

> [!NOTE]
> <span data-ttu-id="5a0ce-102">Esta página descreve uma etapa individual do tutorial de suplemento do PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="5a0ce-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="5a0ce-103">Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou via outro link direto, acesse a página de Introdução do [tutorial de suplemento do PowerPoint](../tutorials/powerpoint-tutorial.yml) para começá-lo do início.</span><span class="sxs-lookup"><span data-stu-id="5a0ce-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="5a0ce-104">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="5a0ce-104">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

## <a name="setup"></a><span data-ttu-id="5a0ce-105">Configurar</span><span class="sxs-lookup"><span data-stu-id="5a0ce-105">Setup</span></span>

<span data-ttu-id="5a0ce-106">Neste tutorial, você criará um suplemento com o Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="5a0ce-106">In this tutorial, you'll create an add-in using Visual Studio.</span></span>

### <a name="create-the-add-in-project"></a><span data-ttu-id="5a0ce-107">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="5a0ce-107">Create the add-in project</span></span>

1. <span data-ttu-id="5a0ce-108">Na barra de menus do Visual Studio, selecione **Arquivo**  >  **Novo**  >  **Projeto**.</span><span class="sxs-lookup"><span data-stu-id="5a0ce-108">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="5a0ce-109">Na lista de tipos de projeto em **Visual C#** ou no **Visual Basic**, expanda a opção **Office/SharePoint**, escolha **Suplementos** e depois **Suplemento da Web do PowerPoint** como o tipo de projeto.</span><span class="sxs-lookup"><span data-stu-id="5a0ce-109">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **PowerPoint Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="5a0ce-110">Nomeie o projeto como **HelloWorld** e depois selecione o botão **OK**.</span><span class="sxs-lookup"><span data-stu-id="5a0ce-110">Name the project **HelloWorld**, and then choose the **OK** button.</span></span>

4. <span data-ttu-id="5a0ce-111">Na caixa de diálogo **Criar suplementos do Office**, escolha **Adicionar novas funcionalidades ao PowerPoint**e depois **Concluir** para criar o projeto.</span><span class="sxs-lookup"><span data-stu-id="5a0ce-111">In the **Create Office Add-in** dialog window, choose **Add new functionalities to PowerPoint**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="5a0ce-p102">O Visual Studio cria uma solução, e os dois projetos dele são exibidos no **Gerenciador de Soluções**. O arquivo **Home.html** é aberto no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="5a0ce-p102">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

     ![Tutorial do PowerPoint: janela do Explorador de soluções do Visual Studio mostrando os dois projetos na solução HelloWorld](../images/powerpoint-tutorial-solution-explorer.png)

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="5a0ce-115">Explorar a solução do Visual Studio</span><span class="sxs-lookup"><span data-stu-id="5a0ce-115">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-code"></a><span data-ttu-id="5a0ce-116">Código de atualização</span><span class="sxs-lookup"><span data-stu-id="5a0ce-116">Update code</span></span> 

<span data-ttu-id="5a0ce-117">Edite o código do suplemento como mostrado a seguir para criar a estrutura que você usará para implementar a funcionalidade do suplemento nas etapas subsequentes deste tutorial.</span><span class="sxs-lookup"><span data-stu-id="5a0ce-117">Edit the add-in code as follows, to create the framework that you'll use to implement add-in functionality in subsequent steps of this tutorial.</span></span>

1. <span data-ttu-id="5a0ce-118">**Home.html** especifica o HTML que será renderizado no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5a0ce-118">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="5a0ce-119">Em **Home.html**, encontre a **div** com `id="content-main"`, substitua toda essa **div** pela marcação a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="5a0ce-119">In **Home.html**, find the **div** with `id="content-main"`, replace that entire **div** with the following markup, and save the file.</span></span>

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

2. <span data-ttu-id="5a0ce-120">Abra o arquivo **Home.js** na raiz do projeto do aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="5a0ce-120">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="5a0ce-121">Este arquivo especifica o script do suplemento.</span><span class="sxs-lookup"><span data-stu-id="5a0ce-121">This file specifies the script for the add-in.</span></span> <span data-ttu-id="5a0ce-122">Substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="5a0ce-122">Replace the entire contents with the following code and save the file.</span></span>

    ```javascript
    (function () {
        "use strict";

        var messageBanner;

        Office.initialize = function (reason) {
            $(document).ready(function () {
                // Initialize the FabricUI notification mechanism and hide it
                var element = document.querySelector('.ms-MessageBanner');
                messageBanner = new fabric.MessageBanner(element);
                messageBanner.hideBanner();

                // TODO1: Assign event handler for insert-image button.
                // TODO4: Assign event handler for insert-text button.
                // TODO6: Assign event handler for get-slide-metadata button.
                // TODO8: Assign event handlers for the four navigation buttons.
            });
        };

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
