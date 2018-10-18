<span data-ttu-id="572c2-101">Nesta etapa do tutorial, você vai adicionar texto ao Slide de Título que contém a foto do dia do [Bing](https://www.bing.com).</span><span class="sxs-lookup"><span data-stu-id="572c2-101">In this step of the tutorial, you'll add text to the title slide that contains the [Bing](https://www.bing.com) photo of the day.</span></span>

> [!NOTE]
> <span data-ttu-id="572c2-102">Esta página descreve uma etapa individual do tutorial de suplemento do PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="572c2-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="572c2-103">Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou via outro link direto, acesse a página de Introdução do [tutorial de suplemento do PowerPoint](../tutorials/powerpoint-tutorial.yml) para começá-lo do início.</span><span class="sxs-lookup"><span data-stu-id="572c2-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="add-text-to-a-slide"></a><span data-ttu-id="572c2-104">Adicionar texto a um slide</span><span class="sxs-lookup"><span data-stu-id="572c2-104">Add text to a slide</span></span> 

1. <span data-ttu-id="572c2-105">No arquivo **Home.html**, substitua `TODO3` pela marcação a seguir.</span><span class="sxs-lookup"><span data-stu-id="572c2-105">In the **Home.html** file, replace `TODO3` with the following markup.</span></span> <span data-ttu-id="572c2-106">Essa marcação define o botão **Inserir Texto** que aparecerá no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="572c2-106">This markup defines the **Insert Text** button that will appear within the add-in's task pane.</span></span>

    ```html
        <br /><br />
        <button class="ms-Button ms-Button--primary" id="insert-text">
            <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
            <span class="ms-Button-label">Insert Text</span>
            <span class="ms-Button-description">Inserts text into the slide.</span>
        </button>
    ```

2. <span data-ttu-id="572c2-107">No arquivo **Home.js**, substitua `TODO4` pelo código a seguir para atribuir o manipulador de eventos ao botão **Inserir Texto**.</span><span class="sxs-lookup"><span data-stu-id="572c2-107">In the **Home.js** file, replace `TODO4` with the following code to assign the event handler for the **Insert Text** button.</span></span>

    ```js
    $('#insert-text').click(insertText);
    ```

3. <span data-ttu-id="572c2-108">No arquivo **Home.js**, substitua `TODO5` pelo código a seguir para definir a função **insertText**.</span><span class="sxs-lookup"><span data-stu-id="572c2-108">In the **Home.js** file, replace `TODO5` with the following code to define the **insertText** function.</span></span> <span data-ttu-id="572c2-109">Esta função insere texto no slide atual.</span><span class="sxs-lookup"><span data-stu-id="572c2-109">This function inserts text into the current slide.</span></span>

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

## <a name="test-the-add-in"></a><span data-ttu-id="572c2-110">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="572c2-110">Test the add-in</span></span>

1. <span data-ttu-id="572c2-p104">Usando o Visual Studio, teste o suplemento pressionando `F5` ou escolhendo o botão **Iniciar** para abrir o PowerPoint com o botão de suplemento **Mostrar Painel de Tarefas** exibido na faixa de opções. O suplemento será hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="572c2-p104">Using Visual Studio, test the add-in by pressing `F5` or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![Captura de tela do Visual Studio com o botão Iniciar realçado](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="572c2-114">No PowerPoint, selecione o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="572c2-114">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Captura de tela do Visual Studio com o botão Mostrar Painel de Tarefas realçado na faixa de opções Página Inicial](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="572c2-116">No painel de tarefas, escolha o botão **Inserir Imagem** para adicionar a foto do dia do Bing ao slide atual e escolher um design para o slide que contém uma caixa de texto como título.</span><span class="sxs-lookup"><span data-stu-id="572c2-116">In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide and choose a design for the slide that contains a text box for the title.</span></span>

    ![Captura de tela do suplemento do PowerPoint com o botão Inserir imagem realçado](../images/powerpoint-tutorial-insert-image-slide-design.png)

4. <span data-ttu-id="572c2-118">Coloque o cursor na caixa de texto no slide de título e depois, no painel de tarefas, escolha o botão **Inserir Texto** para adicionar texto ao slide.</span><span class="sxs-lookup"><span data-stu-id="572c2-118">Put your cursor in the text box on the title slide and then in the task pane, choose the **Insert Text** button to add text to the slide.</span></span>

    ![Captura de tela do suplemento do PowerPoint com o botão Inserir Texto realçado](../images/powerpoint-tutorial-insert-text.png)


5. <span data-ttu-id="572c2-120">No Visual Studio, interrompa o suplemento pressionando `Shift + F5` ou selecionando o botão **Parar**.</span><span class="sxs-lookup"><span data-stu-id="572c2-120">In Visual Studio, stop the add-in by pressing `Shift + F5` or choosing the **Stop** button.</span></span> <span data-ttu-id="572c2-121">O PowerPoint fechará automaticamente quando o suplemento for interrompido.</span><span class="sxs-lookup"><span data-stu-id="572c2-121">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![Captura de tela do Visual Studio com o botão Parar realçado](../images/powerpoint-tutorial-stop.png)