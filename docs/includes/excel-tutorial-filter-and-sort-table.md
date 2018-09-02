<span data-ttu-id="0a8aa-101">Nesta etapa do tutorial, você vai filtrar e classificar a tabela que criou anteriormente.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-101">In this step of the tutorial, you'll filter and sort the table that you created previously.</span></span>

> [!NOTE]
> <span data-ttu-id="0a8aa-102">Esta página descreve uma etapa individual do tutorial de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="0a8aa-103">Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou via outro link direto, acesse a página de Introdução do [tutorial de suplemento do Excel](../tutorials/excel-tutorial.yml) para começá-lo do início.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="filter-the-table"></a><span data-ttu-id="0a8aa-104">Filtrar a tabela</span><span class="sxs-lookup"><span data-stu-id="0a8aa-104">Filter the table</span></span>

1. <span data-ttu-id="0a8aa-105">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="0a8aa-106">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-106">Open the file index.html.</span></span>
3. <span data-ttu-id="0a8aa-107">Abaixo do `div`, que contém o botão `create-table`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="0a8aa-107">Just below the `div` that contains the `create-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="filter-table">Filter Table</button>            
    </div>
    ```

4. <span data-ttu-id="0a8aa-108">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-108">Open the app.js file.</span></span>

5. <span data-ttu-id="0a8aa-109">Logo abaixo da linha que atribui um identificador de clique ao botão `create-table`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="0a8aa-109">Just below the line that assigns a click handler to the `create-table` button, add the following code:</span></span>

    ```js
    $('#filter-table').click(filterTable);
    ```

6. <span data-ttu-id="0a8aa-110">Logo abaixo da função `createTable`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="0a8aa-110">Just below the `createTable` function, add the following function:</span></span>

    ```js
    function filterTable() {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to filter out all expense categories except 
            //        Groceries and Education.

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

7. <span data-ttu-id="0a8aa-p102">Substitua `TODO1` pelo código a seguir. Nota:</span><span class="sxs-lookup"><span data-stu-id="0a8aa-p102">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="0a8aa-113">O código primeiro faz referência à coluna que precisa de filtragem ao passar o nome da coluna para o método `getItem`, em vez de passar o índice para o método `getItemAt` como o método `createTable` faz.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-113">The code first gets a reference to the column that needs filtering by passing the column name to the `getItem` method, instead of passing its index to the `getItemAt` method as the `createTable` method does.</span></span> <span data-ttu-id="0a8aa-114">Como os usuários podem mover as colunas da tabela, a coluna de um determinado índice pode mudar depois da criação da tabela.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-114">Since users can move table columns, the column at a given index might change after the table is created.</span></span> <span data-ttu-id="0a8aa-115">Portanto, é mais seguro usar o nome da coluna como referência dela.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-115">Hence, it is safer to use the column name to get a reference to the column.</span></span> <span data-ttu-id="0a8aa-116">Usamos de forma segura `getItemAt` em um tutorial anterior porque usamos o mesmo método que cria a tabela. Assim não existe a chance de um usuário mover a coluna.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-116">We used `getItemAt` safely in the preceding tutorial, because we used it in the very same method that creates the table, so there is no chance that a user has moved the column.</span></span>
   - <span data-ttu-id="0a8aa-117">O método `applyValuesFilter` é um dos vários métodos de filtragem do objeto `Filter`.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-117">The `applyValuesFilter` method is one of several filtering methods on the `Filter` object.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);
    ``` 

## <a name="sort-the-table"></a><span data-ttu-id="0a8aa-118">Classificar a tabela</span><span class="sxs-lookup"><span data-stu-id="0a8aa-118">Sort the table</span></span>

1. <span data-ttu-id="0a8aa-119">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-119">Open the file index.html.</span></span>
2. <span data-ttu-id="0a8aa-120">Abaixo do `div` que contém o botão `filter-table`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="0a8aa-120">Below the `div` that contains the `filter-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="sort-table">Sort Table</button>            
    </div>
    ```

3. <span data-ttu-id="0a8aa-121">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-121">Open the app.js file.</span></span>

4. <span data-ttu-id="0a8aa-122">Abaixo da linha que atribui um identificador de clique ao botão `filter-table`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="0a8aa-122">Below the line that assigns a click handler to the `filter-table` button, add the following code:</span></span>

    ```js
    $('#sort-table').click(sortTable);
    ```

5. <span data-ttu-id="0a8aa-123">Abaixo da função `filterTable`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-123">Below the `filterTable` function add the following function.</span></span>

    ```js
    function sortTable() {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to sort the table by Merchant name.

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

7. <span data-ttu-id="0a8aa-p104">Substitua `TODO1` pelo código a seguir. Nota:</span><span class="sxs-lookup"><span data-stu-id="0a8aa-p104">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="0a8aa-126">O código cria uma matriz de objetos `SortField` que tem apenas um membro, já que o suplemento só classifica a coluna Comerciante.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-126">The code creates an array of `SortField` objects which has just one member since the add-in only sorts on the Merchant column.</span></span>
   - <span data-ttu-id="0a8aa-127">A propriedade `key` de um objeto `SortField` é o índice com base em zero da coluna a classificar.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-127">The `key` property of a `SortField` object is the zero-based index of the column to sort-on.</span></span>
   - <span data-ttu-id="0a8aa-128">O membro `sort` de uma `Table` é um objeto `TableSort`, não um método.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-128">The `sort` member of a `Table` is a `TableSort` object, not a method.</span></span> <span data-ttu-id="0a8aa-129">Os `SortField`s são passados para o método `TableSort` do objeto `apply`.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-129">The `SortField`s are passed the `TableSort` object's `apply` method.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const sortFields = [
        { 
            key: 1,            // Merchant column
            ascending: false,
        }
    ];

    expensesTable.sort.apply(sortFields);
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="0a8aa-130">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="0a8aa-130">Test the add-in</span></span>

1. <span data-ttu-id="0a8aa-131">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite Ctrl + C duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-131">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="0a8aa-132">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-132">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="0a8aa-133">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-133">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="0a8aa-134">Para fazer isso, interrompa o processo do servidor para obter uma solicitação para inserir o comando de compilação.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-134">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="0a8aa-135">Após a compilação, reinicie o servidor.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-135">After the build, you restart the server.</span></span> <span data-ttu-id="0a8aa-136">As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-136">The next few steps carry out this process.</span></span>

1. <span data-ttu-id="0a8aa-137">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte no Internet Explorer (que é usada em segundo plano pelo Excel para executar os suplementos do Excel).</span><span class="sxs-lookup"><span data-stu-id="0a8aa-137">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
2. <span data-ttu-id="0a8aa-138">Execute o comando `npm start` para iniciar um servidor Web em um host local.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-138">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="0a8aa-139">Feche o painel de tarefas para recarregá-lo e, no menu **Início**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-139">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="0a8aa-140">Se, por algum motivo, a tabela não estiver na planilha aberta, no painel de tarefas, escolha **Criar tabela**.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-140">If for any reason the table is not in the open worksheet, in the taskpane, choose **Create Table**.</span></span> 
6. <span data-ttu-id="0a8aa-141">Escolha os botões **Filtrar tabela** e **Classificar tabela** em qualquer ordem.</span><span class="sxs-lookup"><span data-stu-id="0a8aa-141">Choose the **Filter Table** and **Sort Table** buttons, in either order.</span></span>

    ![Tutorial do Excel: filtrar e classificar tabela](../images/excel-tutorial-filter-and-sort-table.png)
