<span data-ttu-id="4ff7d-101">Nesta etapa do tutorial, você vai recuperar a foto do dia do [Bing](https://www.bing.com) e inseri-la em um slide.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-101">In this step of the tutorial, you'll retrieve the [Bing](https://www.bing.com) photo of the day and insert that image into a slide.</span></span>

> [!NOTE]
> <span data-ttu-id="4ff7d-102">Esta página descreve uma etapa individual do tutorial de suplemento do PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="4ff7d-103">Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou via outro link direto, acesse a página de Introdução do [tutorial de suplemento do PowerPoint](../tutorials/powerpoint-tutorial.yml) para começá-lo do início.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="add-the-bing-photo-of-the-day-to-a-slide"></a><span data-ttu-id="4ff7d-104">Adicionar a foto do dia do Bing a um slide</span><span class="sxs-lookup"><span data-stu-id="4ff7d-104">Add the Bing photo of the day to a slide</span></span>

1. <span data-ttu-id="4ff7d-105">Usando o Explorador de soluções, adicione uma nova pasta chamada **Controladores** ao projeto **HelloWorldWeb**.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-105">Using Solution Explorer, add a new folder named **Controllers** to the **HelloWorldWeb** project.</span></span>

    ![Tutorial do PowerPoint: janela do Explorador de soluções do Visual Studio que realça a pasta Controladores no projeto HelloWorldWeb](../images/powerpoint-tutorial-solution-explorer-controllers.png)

2. <span data-ttu-id="4ff7d-107">Clique com o botão direito do mouse na pasta **Controladores** e selecione **Adicionar > Novo item com scaffold...**.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-107">Right-click the **Controllers** folder and select **Add > New Scaffolded Item...**.</span></span>

3. <span data-ttu-id="4ff7d-108">Na janela da caixa de diálogo **Adicionar Scaffold**, selecione **Controlador da Web API 2 – vazio** e escolha o botão **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-108">In the **Add Scaffold** dialog window, select **Web API 2 Controller - Empty** and choose the **Add** button.</span></span> 

4. <span data-ttu-id="4ff7d-109">Na janela da caixa de diálogo **Adicionar Controlador**, insira **PhotoController** como nome do controlador e escolha o botão **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-109">In the **Add Controller** dialog window, enter **PhotoController** as the controller name and choose the **Add** button.</span></span> <span data-ttu-id="4ff7d-110">O Visual Studio criará e abrirá o arquivo **PhotoController.cs**.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-110">Visual Studio creates and opens the **PhotoController.cs** file.</span></span>

5. <span data-ttu-id="4ff7d-111">Substitua todo o conteúdo do arquivo **PhotoController.cs** pelo código a seguir, que chama o serviço do Bing para recuperar a foto do dia como uma cadeia de caracteres com codificação Base64.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-111">Replace the entire contents of the **PhotoController.cs** file with the following code that calls the Bing service to retrieve the photo of the day as a Base64 encoded string.</span></span> <span data-ttu-id="4ff7d-112">Quando você usar a API JavaScript do Office para inserir uma imagem em um documento, especifique os dados de imagem como uma cadeia de caracteres com codificação Base64.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-112">When you use the Office JavaScript API to insert an image into a document, the image data must be specified as a Base64 encoded string.</span></span>

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

6. <span data-ttu-id="4ff7d-113">No arquivo **Home.html**, substitua `TODO1` pela marcação a seguir.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-113">In the **Home.html** file, replace `TODO1` with the following markup.</span></span> <span data-ttu-id="4ff7d-114">Essa marcação define o botão **Inserir Imagem** que aparecerá no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-114">This markup defines the **Insert Image** button that will appear within the add-in's task pane.</span></span>

    ```html
    <button class="ms-Button ms-Button--primary" id="insert-image">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Insert Image</span>
        <span class="ms-Button-description">Gets the photo of the day that shows on the Bing home page and adds it to the slide.</span>
    </button>
    ```

7. <span data-ttu-id="4ff7d-115">No arquivo **Home.js**, substitua `TODO1` pelo código a seguir para atribuir o manipulador de eventos ao botão **Inserir Imagem**.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-115">In the **Home.js** file, replace `TODO1` with the following code to assign the event handler for the **Insert Image** button.</span></span>

    ```js
    $('#insert-image').click(insertImage);
    ```

8. <span data-ttu-id="4ff7d-116">No arquivo **Home.js**, substitua `TODO2` pelo código a seguir para definir a função **insertImage**.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-116">In the **Home.js** file, replace `TODO2` with the following code to define the **insertImage** function.</span></span> <span data-ttu-id="4ff7d-117">Esta função busca a imagem do serviço Web Bing e chama a função `insertImageFromBase64String` para inserir a imagem no documento.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-117">This function fetches the image from the Bing web service and then calls the `insertImageFromBase64String` function to insert that image into the document.</span></span>

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

9. <span data-ttu-id="4ff7d-118">No arquivo **Home.js**, substitua `TODO3` pelo código a seguir para definir a função `insertImageFromBase64String`.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-118">In the **Home.js** file, replace `TODO3` with the following code to define the `insertImageFromBase64String` function.</span></span> <span data-ttu-id="4ff7d-119">Esta função usa a API JavaScript do Office para inserir a imagem no documento.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-119">This function uses the Office JavaScript API to insert the image into the document.</span></span> <span data-ttu-id="4ff7d-120">Observação:</span><span class="sxs-lookup"><span data-stu-id="4ff7d-120">Note:</span></span> 

    - <span data-ttu-id="4ff7d-121">A opção `coercionType` especificada como segundo parâmetro da solicitação `setSelectedDataAsyc` indica o tipo de dados inserido.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-121">The `coercionType` option that's specified as the second parameter of the `setSelectedDataAsyc` request indicates the type of data being inserted.</span></span> 

    - <span data-ttu-id="4ff7d-122">O objeto `asyncResult` encapsula o resultado da solicitação `setSelectedDataAsync`, incluindo informações de status e de erro caso a solicitação tenha falhado.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-122">The `asyncResult` object encapsulates the result of the `setSelectedDataAsync` request, including status and error information if the request failed.</span></span>

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

## <a name="test-the-add-in"></a><span data-ttu-id="4ff7d-123">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="4ff7d-123">Test the add-in</span></span>

1. <span data-ttu-id="4ff7d-p107">Usando o Visual Studio, teste o suplemento do PowerPoint recém-criado pressionando `F5` ou escolhendo o botão **Iniciar** para abrir o PowerPoint com o botão de suplemento **Mostrar painel de tarefas** exibido na faixa de opções. O suplemento será hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-p107">Using Visual Studio, test the newly created PowerPoint add-in by pressing `F5` or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![Captura de tela do Visual Studio com o botão Iniciar realçado](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="4ff7d-127">No PowerPoint, selecione o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-127">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Captura de tela do Visual Studio com o botão Mostrar Painel de Tarefas realçado na faixa de opções Página Inicial](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="4ff7d-129">No painel de tarefas, escolha o botão **Inserir Imagem** para adicionar a foto do dia do Bing ao slide atual.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-129">In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide.</span></span>

    ![Captura de tela do suplemento do PowerPoint com o botão Inserir imagem realçado](../images/powerpoint-tutorial-insert-image-button.png)

4. <span data-ttu-id="4ff7d-131">No Visual Studio, interrompa o suplemento pressionando `Shift + F5` ou selecionando o botão **Parar**.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-131">In Visual Studio, stop the add-in by pressing `Shift + F5` or choosing the **Stop** button.</span></span> <span data-ttu-id="4ff7d-132">O PowerPoint fechará automaticamente quando o suplemento for interrompido.</span><span class="sxs-lookup"><span data-stu-id="4ff7d-132">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![Captura de tela do Visual Studio com o botão Parar realçado](../images/powerpoint-tutorial-stop.png)