<span data-ttu-id="8cbb3-101">Nesta etapa do tutorial, você vai criar um gráfico com dados da tabela que você criou anteriormente e depois vai formatar o gráfico.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-101">In this step of the tutorial, you'll create a chart using data from the table that you created previously, and then format the chart.</span></span>

> [!NOTE]
> <span data-ttu-id="8cbb3-102">Esta página descreve uma etapa individual do tutorial de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="8cbb3-103">Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou via outro link direto, acesse a página de Introdução do [tutorial de suplemento do Excel](../tutorials/excel-tutorial.yml) para começá-lo do início.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="chart-table-data"></a><span data-ttu-id="8cbb3-104">Dados de tabela do gráfico</span><span class="sxs-lookup"><span data-stu-id="8cbb3-104">Chart table data</span></span>

1. <span data-ttu-id="8cbb3-105">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="8cbb3-106">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-106">Open the file index.html.</span></span>
3. <span data-ttu-id="8cbb3-107">Abaixo do `div` que contém o botão `sort-table`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="8cbb3-107">Below the `div` that contains the `sort-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="create-chart">Create Chart</button>            
    </div>
    ```

4. <span data-ttu-id="8cbb3-108">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-108">Open the app.js file.</span></span>

5. <span data-ttu-id="8cbb3-109">Abaixo da linha que atribui um identificador de clique ao botão `sort-chart`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="8cbb3-109">Below the line that assigns a click handler to the `sort-chart` button, add the following code:</span></span>

    ```js
    $('#create-chart').click(createChart);
    ```

6. <span data-ttu-id="8cbb3-110">Abaixo da função `sortTable`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-110">Below the `sortTable` function add the following function.</span></span>

    ```js
    function createChart() {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to get the range of data to be charted.

            // TODO2: Queue command to create the chart and define its type.

            // TODO3: Queue commands to position and format the chart.

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

7. <span data-ttu-id="8cbb3-p102">Substitua `TODO1` pelo código a seguir. Para excluir a linha de cabeçalho, o código usa o método `Table.getDataBodyRange` para acessar o intervalo de dados que você deseja representar graficamente em vez do método `getRange`.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-p102">Replace `TODO1` with the following code. Note that in order to exclude the header row, the code uses the `Table.getDataBodyRange` method to get the range of data you want to chart instead of the `getRange` method.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const dataRange = expensesTable.getDataBodyRange();
    ``` 

8. <span data-ttu-id="8cbb3-113">Substitua `TODO2` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-113">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="8cbb3-114">Observe os seguintes parâmetros:</span><span class="sxs-lookup"><span data-stu-id="8cbb3-114">Note the following parameters:</span></span>
   - <span data-ttu-id="8cbb3-p104">O primeiro parâmetro para o método `add` especifica o tipo de gráfico. Há diversos tipos.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-p104">The first parameter to the `add` method specifies the type of chart. There are several dozen types.</span></span> 
   - <span data-ttu-id="8cbb3-117">O segundo parâmetro especifica um intervalo de dados a incluir no gráfico.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-117">The second parameter specifies the range of data to include in the chart.</span></span> 
   - <span data-ttu-id="8cbb3-118">O terceiro parâmetro determina se uma série de pontos de dados da tabela deve estar representada por linha ou por coluna.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-118">The third parameter determines whether a series of data points from the table should be charted rowwise or columnwise.</span></span> <span data-ttu-id="8cbb3-119">A opção `auto` informa ao Excel para decidir o melhor método.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-119">The option `auto` tells Excel to decide the best method.</span></span>

    ```js
    let chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ``` 

9. <span data-ttu-id="8cbb3-120">Substitua `TODO3` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-120">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="8cbb3-121">A maior parte do código é autoexplicativa.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-121">Most of this code is self-explanatory.</span></span> <span data-ttu-id="8cbb3-122">Observação:</span><span class="sxs-lookup"><span data-stu-id="8cbb3-122">Note:</span></span>
   - <span data-ttu-id="8cbb3-123">Os parâmetros do método `setPosition` especificam as células da esquerda superior e da direita inferior da área da planilha que deve conter o gráfico.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-123">The parameters to the `setPosition` method specify the upper left and lower right cells of the worksheet area that should contain the chart.</span></span> <span data-ttu-id="8cbb3-124">O Excel ajusta detalhes como a largura da linha para criar uma boa aparência para o gráfico no espaço fornecido.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-124">Excel can adjust things like line width to make the chart look good in the space it has been given.</span></span>
   - <span data-ttu-id="8cbb3-125">"Série" é um conjunto de pontos de dados de uma coluna da tabela.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-125">A "series" is a set of data points from a column of the table.</span></span> <span data-ttu-id="8cbb3-126">Como há apenas uma coluna sem cadeia de caracteres na tabela, o Excel deduz que essa é a única coluna de pontos de dados no gráfico.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-126">Since there is only one non-string column in the table, Excel infers that the column is the only column of data points to chart.</span></span> <span data-ttu-id="8cbb3-127">Ele interpreta outras colunas como rótulos do gráfico.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-127">It interprets the other columns as chart labels.</span></span> <span data-ttu-id="8cbb3-128">Portanto, haverá apenas uma série no gráfico e será necessário o índice 0.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-128">So there will be just one series in the chart and it will have index 0.</span></span> <span data-ttu-id="8cbb3-129">Ele será rotulado como "Valor em €".</span><span class="sxs-lookup"><span data-stu-id="8cbb3-129">This is the one to label with "Value in €".</span></span> 

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in €';
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="8cbb3-130">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="8cbb3-130">Test the add-in</span></span>


1. <span data-ttu-id="8cbb3-131">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite Ctrl + C duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-131">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="8cbb3-132">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-132">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="8cbb3-133">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-133">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="8cbb3-134">Para fazer isso, interrompa o processo do servidor para obter uma solicitação para inserir o comando de compilação.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-134">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="8cbb3-135">Após a compilação, reinicie o servidor.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-135">After the build, you restart the server.</span></span> <span data-ttu-id="8cbb3-136">As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-136">The next few steps carry out this process.</span></span>

1. <span data-ttu-id="8cbb3-137">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte no Internet Explorer (que é usada em segundo plano pelo Excel para executar os suplementos do Excel).</span><span class="sxs-lookup"><span data-stu-id="8cbb3-137">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
2. <span data-ttu-id="8cbb3-138">Execute o comando `npm start` para iniciar um servidor Web em um host local.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-138">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="8cbb3-139">Feche o painel de tarefas para recarregá-lo e, no menu **Início**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-139">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="8cbb3-140">Se, por algum motivo, a tabela estiver não na planilha aberta, no painel de tarefas, escolha **Criar tabela** e depois os botões **Filtrar tabela** e \*\*Classificar tabela \*\* em qualquer ordem.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-140">If for any reason the table is not in the open worksheet, in the taskpane, choose **Create Table** and then **Filter Table** and **Sort Table** buttons, in either order.</span></span>
6. <span data-ttu-id="8cbb3-141">Clique no botão **Criar gráfico**.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-141">Choose the **Create Chart** button.</span></span> <span data-ttu-id="8cbb3-142">Um gráfico é criado e incluirá somente os dados das linhas que foram filtradas.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-142">A chart is created and only the data from the rows that have been filtered are included.</span></span> <span data-ttu-id="8cbb3-143">Os rótulos dos pontos de dados na parte inferior estão na ordem de classificação do gráfico, ou seja, nomes de comerciantes em ordem alfabética inversa.</span><span class="sxs-lookup"><span data-stu-id="8cbb3-143">The labels on the data points across the bottom are in the sort order of the chart; that is, merchant names in reverse alphabetical order.</span></span>

    ![Tutorial do Excel - Criar gráfico](../images/excel-tutorial-create-chart.png)
