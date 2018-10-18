<span data-ttu-id="4d03e-101">Nesta etapa do tutorial, você testará no programa se o suplemento é compatível com a versão atual do Excel do usuário, adicionará uma tabela a uma planilha, depois preencherá e formatará a tabela com os dados.</span><span class="sxs-lookup"><span data-stu-id="4d03e-101">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Excel, add a table to a worksheet, populate the table with data, and format it.</span></span>

> [!NOTE]
> <span data-ttu-id="4d03e-102">Esta página descreve uma etapa individual do tutorial de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="4d03e-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="4d03e-103">Se você chegou aqui por meio dos resultados de mecanismos de pesquisa ou via outro link direto, acesse a página de Introdução do [tutorial de suplemento do Excel](../tutorials/excel-tutorial.yml) para começá-lo do início.</span><span class="sxs-lookup"><span data-stu-id="4d03e-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="code-the-add-in"></a><span data-ttu-id="4d03e-104">Codificação do suplemento</span><span class="sxs-lookup"><span data-stu-id="4d03e-104">Code the add-in</span></span>

1. <span data-ttu-id="4d03e-105">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="4d03e-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="4d03e-106">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="4d03e-106">Open the file index.html.</span></span>
3. <span data-ttu-id="4d03e-107">Substitua `TODO1` pela marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="4d03e-107">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="create-table">Create Table</button>
    ```

4. <span data-ttu-id="4d03e-108">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="4d03e-108">Open the app.js file.</span></span>
5. <span data-ttu-id="4d03e-109">Substitua o `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="4d03e-109">Replace the `TODO1` with the following code.</span></span> <span data-ttu-id="4d03e-110">O código determina se a versão do Excel do usuário proporciona suporte a uma versão do Excel.js que inclua as APIs com esta série de tutoriais.</span><span class="sxs-lookup"><span data-stu-id="4d03e-110">This code determines whether the user's version of Excel supports a version of Excel.js that includes all the APIs that this series of tutorials will use.</span></span> <span data-ttu-id="4d03e-111">Em um suplemento de produção, use o corpo do bloco condicional para ocultar ou desabilitar a interface do usuário que chame a APIs sem suporte.</span><span class="sxs-lookup"><span data-stu-id="4d03e-111">In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs.</span></span> <span data-ttu-id="4d03e-112">Dessa forma, permitirá que o usuário ainda use as partes do suplemento às quais a versão do Excel dá suporte.</span><span class="sxs-lookup"><span data-stu-id="4d03e-112">This will enable the user to still make use of the parts of the add-in that are supported by their version of Excel.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    } 
    ```

6. <span data-ttu-id="4d03e-113">Substitua o `TODO2` pelo código a seguir:</span><span class="sxs-lookup"><span data-stu-id="4d03e-113">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#create-table').click(createTable);
    ```

7. <span data-ttu-id="4d03e-114">Substitua o `TODO3` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="4d03e-114">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="4d03e-115">Observe o seguinte:</span><span class="sxs-lookup"><span data-stu-id="4d03e-115">Note the following:</span></span>
   - <span data-ttu-id="4d03e-116">A lógica de negócios de Excel.js será adicionada à função que passar por `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="4d03e-116">Your Excel.js business logic will be added to the function that is passed to `Excel.run`.</span></span> <span data-ttu-id="4d03e-117">Essa lógica não é executada imediatamente.</span><span class="sxs-lookup"><span data-stu-id="4d03e-117">This logic does not execute immediately.</span></span> <span data-ttu-id="4d03e-118">Em vez disso, ela é adicionada à fila de comandos pendentes.</span><span class="sxs-lookup"><span data-stu-id="4d03e-118">Instead, it is added to a queue of pending commands.</span></span>
   - <span data-ttu-id="4d03e-119">O método `context.sync` envia todos os comandos da fila para execução no Excel.</span><span class="sxs-lookup"><span data-stu-id="4d03e-119">The `context.sync` method sends all queued commands to Excel for execution.</span></span>
   - <span data-ttu-id="4d03e-120">é seguido por um bloco `catch`.`Excel.run`</span><span class="sxs-lookup"><span data-stu-id="4d03e-120">The `Excel.run` is followed by a `catch` block.</span></span> <span data-ttu-id="4d03e-121">Essa é uma prática recomendada que você sempre deve seguir.</span><span class="sxs-lookup"><span data-stu-id="4d03e-121">This is a best practice that you should always follow.</span></span> 

    ```js
    function createTable() {
        Excel.run(function (context) {
            
            // TODO4: Queue table creation logic here.

            // TODO5: Queue commands to populate the table with data.

            // TODO6: Queue commands to format the table.

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

8. <span data-ttu-id="4d03e-p106">Substitua `TODO4` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="4d03e-p106">Replace `TODO4` with the following code. Note:</span></span>
   - <span data-ttu-id="4d03e-124">O código cria uma tabela usando o método `add` de conjunto de tabela da planilha, que sempre existe mesmo que ela esteja vazia.</span><span class="sxs-lookup"><span data-stu-id="4d03e-124">The code creates a table by using `add` method of a worksheet's table collection, which always exists even if it is empty.</span></span> <span data-ttu-id="4d03e-125">Essa é a maneira padrão de criar objetos no Excel.js.</span><span class="sxs-lookup"><span data-stu-id="4d03e-125">This is the standard way that Excel.js objects are created.</span></span> <span data-ttu-id="4d03e-126">Não há nenhuma API do construtor de classe e você nunca usará um operador `new` para criar um objeto do Excel.</span><span class="sxs-lookup"><span data-stu-id="4d03e-126">There are no class constructor APIs, and you never use a `new` operator to create an Excel object.</span></span> <span data-ttu-id="4d03e-127">Em vez disso, adicione a um objeto de conjunto pai.</span><span class="sxs-lookup"><span data-stu-id="4d03e-127">Instead, you add to a parent collection object.</span></span> 
   - <span data-ttu-id="4d03e-128">O primeiro parâmetro do método `add`é o intervalo apenas da linha superior da tabela, não o intervalo inteiro que a tabela por fim usará.</span><span class="sxs-lookup"><span data-stu-id="4d03e-128">The first parameter of the `add` method is the range of only the top row of the table, not the entire range the table will ultimately use.</span></span> <span data-ttu-id="4d03e-129">Isso ocorre porque, quando o suplemento preenche as linhas de dados (na próxima etapa), ele adicionará novas linhas à tabela, em vez de gravar os valores nas células das linhas existentes.</span><span class="sxs-lookup"><span data-stu-id="4d03e-129">This is because when the add-in populates the data rows (in the next step), it will add new rows to the table instead of writing values to the cells of existing rows.</span></span> <span data-ttu-id="4d03e-130">Esse é um padrão mais comum, porque o número de linhas em uma tabela geralmente não é conhecido quando a tabela é criada.</span><span class="sxs-lookup"><span data-stu-id="4d03e-130">This is a more common pattern because the number of rows that a table will have is often not known when the table is created.</span></span> 
   - <span data-ttu-id="4d03e-131">Os nomes de tabelas devem ser exclusivos pela pasta de trabalho inteira, não só na planilha.</span><span class="sxs-lookup"><span data-stu-id="4d03e-131">Table names must be unique across the entire workbook, not just the worksheet.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ``` 

9. <span data-ttu-id="4d03e-p109">Substitua `TODO5` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="4d03e-p109">Replace `TODO5` with the following code. Note:</span></span>
   - <span data-ttu-id="4d03e-134">Os valores das células de um intervalo são definidos em uma matriz de matrizes.</span><span class="sxs-lookup"><span data-stu-id="4d03e-134">The cell values of a range are set with an array of arrays.</span></span>
   - <span data-ttu-id="4d03e-135">Novas linhas são criadas em uma tabela ao chamar o método `add` do conjunto de linhas da tabela.</span><span class="sxs-lookup"><span data-stu-id="4d03e-135">New rows are created in a table by calling the `add` method of the table's row collection.</span></span> <span data-ttu-id="4d03e-136">Você pode adicionar várias linhas em uma única chamada de `add` ao incluir várias matrizes de valores de células na matriz pai que é passada como segundo parâmetro.</span><span class="sxs-lookup"><span data-stu-id="4d03e-136">You can add multiple rows in a single call of `add` by including multiple cell value arrays in the parent array that is passed as the second parameter.</span></span>

    ```js
    expensesTable.getHeaderRowRange().values = 
        [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add at the end*/, [
        ["1/1/2017", "The Phone Company", "Communications", "120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
        ["1/11/2017", "Bellows College", "Education", "350.1"],
        ["1/15/2017", "Trey Research", "Other", "135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
    ]);
    ``` 

10. <span data-ttu-id="4d03e-p111">Substitua `TODO6` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="4d03e-p111">Replace `TODO6` with the following code. Note:</span></span>
   - <span data-ttu-id="4d03e-139">O código recebe uma referência para a coluna **quantidade** ao passar o índice com base em zero para o método `getItemAt` do conjunto de colunas da tabela.</span><span class="sxs-lookup"><span data-stu-id="4d03e-139">The code gets a reference to the **Amount** column by passing its zero-based index to the `getItemAt` method of the table's column collection.</span></span> 

     > [!NOTE]
     > <span data-ttu-id="4d03e-140">Os objetos do conjunto Excel.js, como `TableCollection`, `WorksheetCollection`, e `TableColumnCollection`, têm a propriedade `items` que é como uma matriz dos tipos de objetos filhos, como `Table` ou `Worksheet` ou `TableColumn`; mas um objeto `*Collection` não é uma matriz.</span><span class="sxs-lookup"><span data-stu-id="4d03e-140">Excel.js collection objects, such as `TableCollection`, `WorksheetCollection`, and `TableColumnCollection` have an `items` property that is an array of the child object types, such as `Table` or `Worksheet` or `TableColumn`; but a `*Collection` object is not itself an array.</span></span>

   - <span data-ttu-id="4d03e-141">O código formata o intervalo da coluna **quantidade** como Euros com um segundo decimal.</span><span class="sxs-lookup"><span data-stu-id="4d03e-141">The code then formats the range of the **Amount** column as Euros to the second decimal.</span></span> 
   - <span data-ttu-id="4d03e-142">Por fim, isso garante que a largura das colunas e a altura das linhas sejam grandes o suficiente para o maior (ou o mais alto) item de dados.</span><span class="sxs-lookup"><span data-stu-id="4d03e-142">Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item.</span></span> <span data-ttu-id="4d03e-143">Observe que o código deve receber os objetos `Range` a formatar.</span><span class="sxs-lookup"><span data-stu-id="4d03e-143">Notice that the code must get `Range` objects to format.</span></span> <span data-ttu-id="4d03e-144">`TableColumn` Os objetos `TableColumn` e `TableRow` não têm propriedades de formato.</span><span class="sxs-lookup"><span data-stu-id="4d03e-144">`TableColumn` and `TableRow` objects do not have format properties.</span></span>

        ```js
        expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
        expensesTable.getRange().format.autofitColumns();
        expensesTable.getRange().format.autofitRows();
        ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="4d03e-145">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="4d03e-145">Test the add-in</span></span>

1. <span data-ttu-id="4d03e-146">Abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue para a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="4d03e-146">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
2. <span data-ttu-id="4d03e-147">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte no Internet Explorer (que é usada em segundo plano pelo Excel para executar os suplementos do Excel).</span><span class="sxs-lookup"><span data-stu-id="4d03e-147">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
3. <span data-ttu-id="4d03e-148">Execute o comando `npm start` para iniciar um servidor Web em um host local.</span><span class="sxs-lookup"><span data-stu-id="4d03e-148">Run the command `npm start` to start a web server running on localhost.</span></span>   
4. <span data-ttu-id="4d03e-149">Realize o sideload do suplemento usando um dos métodos a seguir:</span><span class="sxs-lookup"><span data-stu-id="4d03e-149">Sideload the add-in by using one of the following methods:</span></span>
    - <span data-ttu-id="4d03e-150">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="4d03e-150">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="4d03e-151">Excel Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="4d03e-151">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="4d03e-152">iPad e Mac: [Realizar o sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="4d03e-152">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>
5. <span data-ttu-id="4d03e-153">Sobre o menu da **Página Inicial**, escolha **Mostrar painel de tarefas**.</span><span class="sxs-lookup"><span data-stu-id="4d03e-153">On the **Home** menu, choose **Show Taskpane**.</span></span>
6. <span data-ttu-id="4d03e-154">No painel de tarefas, escolha **Criar tabela**.</span><span class="sxs-lookup"><span data-stu-id="4d03e-154">In the taskpane, choose **Create Table**.</span></span>

    ![Tutorial do Excel - Criar tabela](../images/excel-tutorial-create-table.png)
