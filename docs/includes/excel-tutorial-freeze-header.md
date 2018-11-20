<span data-ttu-id="8a7ab-101">Quando uma tabela for longa o suficiente para que um usuário precise rolar para ver algumas linhas, a linha de cabeçalho poderá ficar fora da vista.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-101">When a table is long enough that a user must scroll to see some rows, the header row can scroll out of sight.</span></span> <span data-ttu-id="8a7ab-102">Nesta etapa do tutorial, você precisará congelar a linha do cabeçalho da tabela que criou anteriormente para que ela permaneça visível, mesmo que o usuário role ao longo da planilha.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-102">In this step of the tutorial, you'll freeze the header row of the table that you created previously, so that it remains visible even as the user scrolls down the worksheet.</span></span> 

> [!NOTE]
> <span data-ttu-id="8a7ab-103">Esta página descreve uma etapa individual do tutorial de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-103">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="8a7ab-104">Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou via outro link direto, acesse a página de Introdução do [tutorial de suplemento do Excel](../tutorials/excel-tutorial.yml) para começá-lo do início.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-104">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="freeze-the-tables-header-row"></a><span data-ttu-id="8a7ab-105">Congelar a linha de cabeçalho da tabela</span><span class="sxs-lookup"><span data-stu-id="8a7ab-105">Freeze the table's header row</span></span>

1. <span data-ttu-id="8a7ab-106">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-106">Open the project in your code editor.</span></span>
2. <span data-ttu-id="8a7ab-107">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-107">Open the file index.html.</span></span>
3. <span data-ttu-id="8a7ab-108">Abaixo do `div` que contém o botão `create-chart`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="8a7ab-108">Below the `div` that contains the `create-chart` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="freeze-header">Freeze Header</button>
    </div>
    ```

4. <span data-ttu-id="8a7ab-109">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-109">Open the app.js file.</span></span>

5. <span data-ttu-id="8a7ab-110">Abaixo da linha que atribui um identificador de clique ao botão `create-chart`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="8a7ab-110">Below the line that assigns a click handler to the `create-chart` button, add the following code:</span></span>

    ```js
    $('#freeze-header').click(freezeHeader);
    ```

6. <span data-ttu-id="8a7ab-111">Abaixo da função `createChart`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="8a7ab-111">Below the `createChart` function add the following function:</span></span>

    ```js
    function freezeHeader() {
        Excel.run(function (context) {

            // TODO1: Queue commands to keep the header visible when the user scrolls.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

7. <span data-ttu-id="8a7ab-p103">Substitua `TODO1` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="8a7ab-p103">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="8a7ab-114">A coleção `Worksheet.freezePanes` é um conjunto de painéis da planilha que fica congelado ou fixado no mesmo lugar quando rolamos a planilha.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-114">The `Worksheet.freezePanes` collection is a set of panes in the worksheet that are pinned, or frozen, in place when the worksheet is scrolled.</span></span>
   - <span data-ttu-id="8a7ab-p104">O método `freezeRows` considera como parâmetro o número de linhas, começando da parte superior, que devem ser fixadas no local. Passamos `1` para fixar a primeira linha no local.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-p104">The `freezeRows` method takes as a parameter the number of rows, from the top that are to be pinned in place. We pass `1` to pin the first row in place.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ```

## <a name="test-the-add-in"></a><span data-ttu-id="8a7ab-117">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="8a7ab-117">Test the add-in</span></span>

1. <span data-ttu-id="8a7ab-118">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite Ctrl + C duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-118">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="8a7ab-119">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-119">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="8a7ab-120">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-120">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="8a7ab-121">Para fazer isso, interrompa o processo do servidor para obter uma solicitação para inserir o comando de compilação.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-121">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="8a7ab-122">Após a compilação, reinicie o servidor.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-122">After the build, you restart the server.</span></span> <span data-ttu-id="8a7ab-123">As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-123">The next few steps carry out this process.</span></span>

1. <span data-ttu-id="8a7ab-124">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte no Internet Explorer (que é usada em segundo plano pelo Excel para executar os suplementos do Excel).</span><span class="sxs-lookup"><span data-stu-id="8a7ab-124">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
2. <span data-ttu-id="8a7ab-125">Execute o comando `npm start` para iniciar um servidor Web em um host local.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-125">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="8a7ab-126">Feche o painel de tarefas para recarregá-lo e, no menu **Início**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-126">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
6. <span data-ttu-id="8a7ab-127">Se a tabela estiver na planilha, exclua-a.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-127">If the table is in the worksheet, delete it.</span></span>
7. <span data-ttu-id="8a7ab-128">No painel de tarefas, escolha **Criar Tabela**.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-128">In the taskpane, choose **Create Table**.</span></span>
8. <span data-ttu-id="8a7ab-129">Escolha o botão **Congelar Cabeçalho**.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-129">Choose the **Freeze Header** button.</span></span>
9. <span data-ttu-id="8a7ab-130">Role a planilha para baixo, o suficiente para ver que o cabeçalho da tabela permanece visível na parte superior mesmo ao rolar até que as primeiras linhas fiquem fora da vista.</span><span class="sxs-lookup"><span data-stu-id="8a7ab-130">Scroll down the worksheet enough to to see that the table header remains visible at the top even when the higher rows scroll out of sight.</span></span>

    ![Tutorial do Excel: congelar cabeçalho](../images/excel-tutorial-freeze-header.png)
