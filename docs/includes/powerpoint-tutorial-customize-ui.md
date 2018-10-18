<span data-ttu-id="b7bf9-101">Nesta etapa do tutorial, você vai personalizar a IU (interface do usuário) do painel tarefas.</span><span class="sxs-lookup"><span data-stu-id="b7bf9-101">In this step of the tutorial, you'll customize the task pane user interface (UI).</span></span>

> [!NOTE]
> <span data-ttu-id="b7bf9-102">Esta página descreve uma etapa individual do tutorial de suplemento do PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="b7bf9-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="b7bf9-103">Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou via outro link direto, acesse a página de Introdução do [tutorial de suplemento do PowerPoint](../tutorials/powerpoint-tutorial.yml) para começá-lo do início.</span><span class="sxs-lookup"><span data-stu-id="b7bf9-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="customize-the-task-pane-ui"></a><span data-ttu-id="b7bf9-104">Personalizar a interface do usuário do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b7bf9-104">Customize the task pane UI</span></span> 

1. <span data-ttu-id="b7bf9-105">No arquivo **Home.html**, substitua `TODO2` pela marcação a seguir para adicionar uma seção de cabeçalho e um título ao painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="b7bf9-105">In the **Home.html** file, replace `TODO2` with the following markup to add a header section and title to the task pane.</span></span> <span data-ttu-id="b7bf9-106">Observação:</span><span class="sxs-lookup"><span data-stu-id="b7bf9-106">Note:</span></span>

    - <span data-ttu-id="b7bf9-107">Os estilos que começam com `ms-` são definidos pelo [Office UI Fabric](../design/office-ui-fabric.md), uma estrutura de front-end JavaScript para criar experiências do usuário do Office e Office 365.</span><span class="sxs-lookup"><span data-stu-id="b7bf9-107">The styles that begin with `ms-` are defined by [Office UI Fabric](../design/office-ui-fabric.md), a JavaScript front-end framework for building user experiences for Office and Office 365.</span></span> <span data-ttu-id="b7bf9-108">O arquivo **Home.html** inclui uma referência à folha de estilos do Fabric.</span><span class="sxs-lookup"><span data-stu-id="b7bf9-108">The **Home.html** file includes a reference to the Fabric stylesheet.</span></span>

    ```html
    <div id="content-header">
        <div class="ms-Grid ms-bgColor-neutralPrimary">
            <div class="ms-Grid-row">
                <div class="padding ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"> <div class="ms-font-xl ms-fontColor-white ms-fontWeight-semibold">My PowerPoint add-in</div></div>
            </div>
        </div>
    </div>
    ```

2. <span data-ttu-id="b7bf9-109">No arquivo **Home.html**, localize a **div** com `class="footer"` e exclua toda a **div** para remover a seção de rodapé do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="b7bf9-109">In the **Home.html** file, find the **div** with `class="footer"` and delete that entire **div** to remove the footer section from the task pane.</span></span>

## <a name="test-the-add-in"></a><span data-ttu-id="b7bf9-110">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="b7bf9-110">Test the add-in</span></span>

1. <span data-ttu-id="b7bf9-p104">Usando o Visual Studio, teste o suplemento do PowerPoint pressionando `F5` ou escolhendo o botão **Iniciar** para abrir o PowerPoint com o botão de suplemento **Mostrar painel de tarefas** exibido na faixa de opções. O suplemento será hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="b7bf9-p104">Using Visual Studio, test the PowerPoint add-in by pressing `F5` or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![Captura de tela do Visual Studio com o botão Iniciar realçado](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="b7bf9-114">No PowerPoint, selecione o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="b7bf9-114">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Captura de tela do Visual Studio com o botão Mostrar Painel de Tarefas realçado na faixa de opções Página Inicial](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="b7bf9-116">Observe que agora o painel de tarefas contém uma seção de cabeçalho e um título e não contém mais uma seção de rodapé.</span><span class="sxs-lookup"><span data-stu-id="b7bf9-116">Notice that the task pane now contains a header section and title, and no longer contains a footer section.</span></span>

    ![Captura de tela do suplemento do PowerPoint com o botão Inserir imagem realçado](../images/powerpoint-tutorial-new-task-pane-ui.png)

4. <span data-ttu-id="b7bf9-118">No Visual Studio, interrompa o suplemento pressionando `Shift + F5` ou selecionando o botão **Parar**.</span><span class="sxs-lookup"><span data-stu-id="b7bf9-118">In Visual Studio, stop the add-in by pressing `Shift + F5` or choosing the **Stop** button.</span></span> <span data-ttu-id="b7bf9-119">O PowerPoint fechará automaticamente quando o suplemento for interrompido.</span><span class="sxs-lookup"><span data-stu-id="b7bf9-119">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![Captura de tela do Visual Studio com o botão Parar realçado](../images/powerpoint-tutorial-stop.png)

