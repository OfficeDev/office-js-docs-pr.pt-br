---
title: Tutorial de suplemento do PowerPoint
description: Neste tutorial, você criará um suplemento do PowerPoint que insere imagem, texto, obtém metadados do slide e navega entre slides.
ms.date: 02/09/2021
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: 35359f58831ebd4b8874247378a09e9da97e4d69
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50238075"
---
# <a name="tutorial-create-a-powerpoint-task-pane-add-in"></a>Tutorial: Criar um Suplemento do Painel de Tarefas

Neste tutorial, você usará o Visual Studio para criar um Suplementos do Painel de Tarefas do PowerPoint:

> [!div class="checklist"]
>
> - Adicionar a foto do dia do [Bing](https://www.bing.com) a um slide
> - Adicionar texto a um slide
> - Obtém metadados do slide
> - Navega entre slides

## <a name="prerequisites"></a>Pré-requisitos

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

## <a name="create-your-add-in-project"></a>Criar seu projeto do suplemento

Conclua as etapas a seguir para criar um projeto de suplemento do PowerPoint usando o Visual Studio.

1. Escolha **Criar um novo projeto**.

2. Usando a caixa de pesquisa, insira **suplemento**. Escolha **Suplemento do PowerPoint Web**, em seguida, selecione **Próximo**.

3. Nomeie o projeto como `HelloWorld` e selecione **Criar**.

4. Na caixa de diálogo **Criar suplementos do Office**, escolha **Adicionar novas funcionalidades ao PowerPoint** e depois **Concluir** para criar o projeto.

5. O Visual Studio cria uma solução, e os dois projetos dele são exibidos no **Gerenciador de Soluções**. O arquivo **Home.html** é aberto no Visual Studio.

     ![A captura de tela da janela do gerenciador de soluções do Microsoft Visual Studio mostrando HelloWorld e HelloWorldWeb, os 2 projetos na solução HelloWorld](../images/powerpoint-tutorial-solution-explorer.png)

### <a name="explore-the-visual-studio-solution"></a>Explorar a solução do Visual Studio

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-code"></a>Código de atualização

Edite o código do suplemento como mostrado a seguir para criar a estrutura que você usará para implementar a funcionalidade do suplemento nas etapas subsequentes deste tutorial.

1. **Home.html** especifica o HTML que será renderizado no painel de tarefas do suplemento. Em **Home.html**, encontre a **div** com `id="content-main"`, substitua toda essa **div** pela marcação a seguir e salve o arquivo.

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

2. Abra o arquivo **Home.js** na raiz do projeto do aplicativo Web. Este arquivo especifica o script do suplemento. Substitua todo o conteúdo pelo código a seguir e salve o arquivo.

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

## <a name="insert-an-image"></a>Inserir uma imagem

Conclua as seguintes etapas para adicionar o código que recupera a foto do dia do [Bing](https://www.bing.com) e inserir as imagens em um slide.

1. Usando o Explorador de soluções, adicione uma nova pasta chamada **Controladores** ao projeto **HelloWorldWeb**.

    ![A captura de tela da janela do Gerenciador de Soluções do Microsoft Visual Studio mostrando a pasta Controladores realçada no projeto HelloWorldWeb](../images/powerpoint-tutorial-solution-explorer-controllers.png)

2. Clique com o botão direito do mouse na pasta **Controladores** e selecione **Adicionar > Novo item com scaffold...**.

3. Na janela da caixa de diálogo **Adicionar Scaffold**, selecione **Controlador da Web API 2 – vazio** e escolha o botão **Adicionar**. 

4. Na janela da caixa de diálogo **Adicionar Controlador**, insira **PhotoController** como nome do controlador e escolha o botão **Adicionar**. O Visual Studio criará e abrirá o arquivo **PhotoController.cs**.

5. Substitua todo o conteúdo do arquivo **PhotoController.cs** pelo código a seguir, que chama o serviço do Bing para recuperar a foto do dia como uma cadeia de caracteres com codificação Base64. Quando você usar a API JavaScript do Office para inserir uma imagem em um documento, especifique os dados de imagem como uma cadeia de caracteres com codificação Base64.

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

6. No arquivo **Home.html**, substitua `TODO1` pela marcação a seguir. Essa marcação define o botão **Inserir Imagem** que aparecerá no painel de tarefas do suplemento.

    ```html
    <button class="Button Button--primary" id="insert-image">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Insert Image</span>
        <span class="Button-description">Gets the photo of the day that shows on the Bing home page and adds it to the slide.</span>
    </button>
    ```

7. No arquivo **Home.js**, substitua `TODO1` pelo código a seguir para atribuir o manipulador de eventos ao botão **Inserir Imagem**.

    ```js
    $('#insert-image').click(insertImage);
    ```

8. No arquivo **Home.js**, substitua `TODO2` pelo código a seguir para definir a função `insertImage`. Esta função busca a imagem do serviço Web Bing e chama a função `insertImageFromBase64String` para inserir a imagem no documento.

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

9. No arquivo **Home.js**, substitua `TODO3` pelo código a seguir para definir a função `insertImageFromBase64String`. Esta função usa a API JavaScript do Office para inserir a imagem no documento. Observação:

    - A opção `coercionType` especificada como segundo parâmetro da solicitação `setSelectedDataAsync` indica o tipo de dados inserido.

    - O objeto `asyncResult` encapsula o resultado da solicitação `setSelectedDataAsync`, incluindo informações de status e de erro caso a solicitação tenha falhado.

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

### <a name="test-the-add-in"></a>Testar o suplemento

1. Usando o Visual Studio, teste o suplemento do PowerPoint recém-criado, pressionando **F5** ou escolhendo o botão **Iniciar** para abrir o PowerPoint com o botão de suplemento **Mostrar Painel de Tarefas** exibido na faixa de opções. O suplemento será hospedado localmente no IIS.

    ![A captura de tela mostrando o botão Iniciar realçado no Microsoft Visual Studio](../images/powerpoint-tutorial-start.png)

2. No PowerPoint, selecione o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.

    ![Captura de tela mostrando o botão Mostrar painel de tarefas realçado na faixa de opções página inicial do PowerPoint](../images/powerpoint-tutorial-show-taskpane-button.png)

3. No painel de tarefas, escolha o botão **Inserir Imagem** para adicionar a foto do dia do Bing ao slide atual.

    ![Captura de tela do suplemento do PowerPoint com o botão Inserir imagem realçado](../images/powerpoint-tutorial-insert-image-button.png)

4. No Visual Studio, interrompa o suplemento pressionando **Shift + F5** ou selecionando o botão **Parar**. O PowerPoint fechará automaticamente quando o suplemento for interrompido.

    ![Captura de tela mostrando o botão Pare realçado no Microsoft Visual Studio](../images/powerpoint-tutorial-stop.png)

## <a name="customize-user-interface-ui-elements"></a>Personalizar os elementos da IU (interface do usuário)

Conclua as seguintes etapas para adicionar a marca que personaliza o painel de tarefas da interface do usuário.

1. No arquivo **Home.html**, substitua `TODO2` pela marcação a seguir para adicionar uma seção de cabeçalho e um título ao painel de tarefas. Observação:

    - Os estilos que começam com `ms-` são definidos pelo [Office UI Fabric](../design/office-ui-fabric.md), uma estrutura de front-end JavaScript para criar experiências do usuário do Office. O arquivo **Home.html** inclui uma referência à folha de estilos do Fabric.

    ```html
    <div id="content-header">
        <div class="ms-Grid ms-bgColor-neutralPrimary">
            <div class="ms-Grid-row">
                <div class="padding ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"> <div class="ms-font-xl ms-fontColor-white ms-fontWeight-semibold">My PowerPoint add-in</div></div>
            </div>
        </div>
    </div>
    ```

2. No arquivo **Home.html**, localize a **div** com `class="footer"` e exclua toda a **div** para remover a seção de rodapé do painel de tarefas.

### <a name="test-the-add-in"></a>Testar o suplemento

1. Usando o Visual Studio, teste o suplemento do PowerPoint ao pressionar **F5** ou escolha o botão **Iniciar** para abrir o PowerPoint com o botão de suplemento **Mostrar painel de tarefas** exibido na faixa de opções. O suplemento será hospedado localmente no IIS.

    ![Captura de tela mostrando o botão Iniciar realçado no Visual Studio](../images/powerpoint-tutorial-start.png)

2. No PowerPoint, selecione o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.

    ![Captura de tela mostrando o botão Mostrar Painel de tarefas realçado na faixa de opções da Página Inicial do PowerPoint](../images/powerpoint-tutorial-show-taskpane-button.png)

3. Observe que agora o painel de tarefas contém uma seção de cabeçalho e um título e não contém mais uma seção de rodapé.

    ![Captura de tela do suplemento do PowerPoint com o botão Inserir Imagem realçado](../images/powerpoint-tutorial-new-task-pane-ui.png)

4. No Visual Studio, interrompa o suplemento pressionando **Shift + F5** ou selecionando o botão **Parar**. O PowerPoint fechará automaticamente quando o suplemento for interrompido.

    ![Captura de tela exibindo o botão Pare realçado no Microsoft Visual Studio](../images/powerpoint-tutorial-stop.png)

## <a name="insert-text"></a>Inserir texto

Conclua as seguintes etapas para adicionar o código que insere texto no slide de título que contém as fotos do dia do [Bing](https://www.bing.com).

1. No arquivo **Home.html**, substitua `TODO3` pela marcação a seguir. Essa marcação define o botão **Inserir Texto** que aparecerá no painel de tarefas do suplemento.

    ```html
        <br /><br />
        <button class="Button Button--primary" id="insert-text">
            <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
            <span class="Button-label">Insert Text</span>
            <span class="Button-description">Inserts text into the slide.</span>
        </button>
    ```

2. No arquivo **Home.js**, substitua `TODO4` pelo código a seguir para atribuir o manipulador de eventos ao botão **Inserir Texto**.

    ```js
    $('#insert-text').click(insertText);
    ```

3. No arquivo **Home.js**, substitua `TODO5` pelo código a seguir para definir a função `insertText`. Esta função insere texto no slide atual.

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

### <a name="test-the-add-in"></a>Testar o suplemento

1. Usando o Visual Studio, teste o suplemento ao pressionar **F5** ou escolha o botão **Iniciar** para abrir o PowerPoint com o botão de suplemento **Mostrar painel de tarefas** exibido na faixa de opções. O suplemento será hospedado localmente no IIS.

    ![Captura de tela do Microsoft Visual Studio com o botão Iniciar realçado](../images/powerpoint-tutorial-start.png)

2. No PowerPoint, selecione o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.

    ![Captura de tela realçando o botão Mostrar Painel de tarefas na faixa de opções da Página inicial no PowerPoint](../images/powerpoint-tutorial-show-taskpane-button.png)

3. No painel de tarefas, escolha o botão **Inserir Imagem** para adicionar a foto do dia do Bing ao slide atual e escolher um design para o slide que contém uma caixa de texto como título.

    ![Captura de tela do PowerPoint com o slide atual realçado e suplemento com o botão Inserir Imagem realçado](../images/powerpoint-tutorial-insert-image-slide-design.png)

4. Coloque o cursor na caixa de texto no slide de título e depois, no painel de tarefas, escolha o botão **Inserir Texto** para adicionar texto ao slide.

    ![Captura de tela do suplemento do PowerPoint com o botão Inserir Texto realçado](../images/powerpoint-tutorial-insert-text.png)

5. No Visual Studio, interrompa o suplemento pressionando **Shift + F5** ou selecionando o botão **Parar**. O PowerPoint fechará automaticamente quando o suplemento for interrompido.

    ![Captura de tela do Microsoft Visual Studio com o botão Pare realçado](../images/powerpoint-tutorial-stop.png)

## <a name="get-slide-metadata"></a>Obter metadados do slide

Conclua as seguintes etapas para adicionar o código que recupera os metadados para o slide selecionado.

1. No arquivo **Home.html**, substitua `TODO4` pela marcação a seguir. Essa marcação define o botão **Obter metadados do slide** que aparecerá no painel de tarefas do suplemento.

    ```html
    <br /><br />
    <button class="Button Button--primary" id="get-slide-metadata">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Get Slide Metadata</span>
        <span class="Button-description">Gets metadata for the selected slide(s).</span>
    </button>
    ```

2. No arquivo **Home.js**, substitua `TODO6` pelo código a seguir para atribuir o manipulador de eventos para o botão **Obter Metadados do Slide**.

    ```js
    $('#get-slide-metadata').click(getSlideMetadata);
    ```

3. No arquivo **Home.js**, substitua `TODO7` pelo código a seguir para definir a função `getSlideMetadata`. Esta função recupera metadados dos slides selecionados e os grava em uma janela pop-up da caixa de diálogo no painel de tarefas do suplemento.

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

### <a name="test-the-add-in"></a>Testar o suplemento

1. Usando o Visual Studio, teste o suplemento ao pressionar **F5** ou escolha o botão **Iniciar** para abrir o PowerPoint com o botão de suplemento **Mostrar painel de tarefas** exibido na faixa de opções. O suplemento será hospedado localmente no IIS.

    ![Captura de tela realçando o botão Iniciar no Microsoft Visual Studio](../images/powerpoint-tutorial-start.png)

2. No PowerPoint, selecione o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.

    ![Captura de tela realçando o botão Mostrar painel de tarefas na faixa de opções da Página Inicial do PowerPoint](../images/powerpoint-tutorial-show-taskpane-button.png)

3. No painel de tarefas, escolha o botão **Obter metadados do slide** para obter os metadados do slide selecionado. Os metadados do slide serão gravados na janela pop-up da caixa de diálogo na parte inferior do painel de tarefas. Nesse caso, a matriz `slides` dos metadados JSON contém um objeto que especifica `id`, `title` e `index` do slide selecionado. Se vários slides tivessem sido selecionados na recuperação de metadados do slide, a matriz `slides` dos metadados JSON conteria um objeto para cada slide selecionado.

    ![Captura de tela do suplemento do PowerPoint com o botão Obter metadados do slide realçado](../images/powerpoint-tutorial-get-slide-metadata.png)

4. No Visual Studio, interrompa o suplemento pressionando **Shift + F5** ou selecionando o botão **Parar**. O PowerPoint fechará automaticamente quando o suplemento for interrompido.

    ![Captura de tela realçando o botão Pare no Microsoft Visual Studio](../images/powerpoint-tutorial-stop.png)

## <a name="navigate-between-slides"></a>Navegar entre slides

Conclua as seguintes etapas para adicionar o código que navega entre os slides de um documento.

1. No arquivo **Home.html**, substitua `TODO5` pela marcação a seguir. Essa marcação define os quatro botões de navegação que aparecerão no painel de tarefas do suplemento.

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

2. No arquivo **Home.js**, substitua `TODO8` pelo código a seguir para atribuir o manipulador de eventos aos quatro botões de navegação.

    ```js
    $('#go-to-first-slide').click(goToFirstSlide);
    $('#go-to-next-slide').click(goToNextSlide);
    $('#go-to-previous-slide').click(goToPreviousSlide);
    $('#go-to-last-slide').click(goToLastSlide);
    ```

3. No arquivo **Home.js**, substitua `TODO9` pelo código a seguir para definir as funções de navegação. Cada uma dessas funções usa a função `goToByIdAsync` para selecionar um slide com base em sua posição no documento (primeiro, último, anterior e próximo).

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

### <a name="test-the-add-in"></a>Testar o suplemento

1. Usando o Visual Studio, teste o suplemento ao pressionar **F5** ou escolha o botão **Iniciar** para abrir o PowerPoint com o botão de suplemento **Mostrar painel de tarefas** exibido na faixa de opções. O suplemento será hospedado localmente no IIS.

    ![Captura de tela mostrando o botão Iniciar realçado na barra de ferramentas do Microsoft Visual Studio](../images/powerpoint-tutorial-start.png)

2. No PowerPoint, selecione o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.

    ![Captura de tela mostrando o botão Mostrar Painel de tarefas realçado na faixa de opções da Página inicial no PowerPoint](../images/powerpoint-tutorial-show-taskpane-button.png)

3. Use o botão **Novo Slide** na faixa de opções da guia **Página Inicial** para adicionar dois novos slides ao documento.

4. No painel de tarefas, escolha o botão **Ir para o primeiro Slide**. O primeiro slide no documento é selecionado e exibido.

    ![Captura de tela do suplemento do PowerPoint com o botão Ir Para o Primeiro Slide realçado](../images/powerpoint-tutorial-go-to-first-slide.png)

5. No painel de tarefas, escolha o botão **Ir para o próximo Slide**. O próximo slide no documento é selecionado e exibido.

    ![Captura de tela do suplemento do PowerPoint com o botão Ir Para o Próximo Slide realçado](../images/powerpoint-tutorial-go-to-next-slide.png)

6. No painel de tarefas, escolha o botão **Ir Para o Slide Anterior**. O slide anterior no documento é selecionado e exibido.

    ![Captura de tela do suplemento do PowerPoint com o botão Ir Para o Slide Anterior realçado](../images/powerpoint-tutorial-go-to-previous-slide.png)

7. No painel de tarefas, escolha o botão **Ir Para o Último Slide**. O último slide no documento é selecionado e exibido.

    ![Captura de tela do suplemento do PowerPoint com o botão Ir Para o Último Slide realçado](../images/powerpoint-tutorial-go-to-last-slide.png)

8. No Visual Studio, interrompa o suplemento pressionando **Shift + F5** ou selecionando o botão **Parar**. O PowerPoint fechará automaticamente quando o suplemento for interrompido.

    ![Captura de tela mostrando o botão Pare realçado na barra de ferramentas do Microsoft Visual Studio](../images/powerpoint-tutorial-stop.png)

## <a name="next-steps"></a>Próximas etapas

Neste tutorial, você criou um suplemento do PowerPoint que insere imagem, texto, obtém metadados do slide e navega entre slides. Para saber mais sobre a criação de suplementos do PowerPoint, continue no seguinte artigo:

> [!div class="nextstepaction"]
> [Visão geral dos Suplementos do SharePoint](../powerpoint/powerpoint-add-ins.md)

## <a name="see-also"></a>Confira também

- [Visão geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
- [Desenvolver Suplementos do Office ](../develop/develop-overview.md)
