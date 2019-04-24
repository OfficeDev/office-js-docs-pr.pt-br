---
title: Tutorial de suplemento do Excel
description: Neste tutorial, você criará um suplemento do Excel que cria, preenche, filtra e classifica uma tabela, cria um gráfico, congela um cabeçalho de tabela, protege uma planilha e abre uma caixa de diálogo
ms.date: 01/28/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 410b2391d207f7c83f9accb349448dbc0c92a0e2
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451279"
---
# <a name="tutorial-create-an-excel-task-pane-add-in"></a><span data-ttu-id="55761-103">Tutorial: criar um suplemento do painel de tarefas no Excel</span><span class="sxs-lookup"><span data-stu-id="55761-103">Tutorial: Create an Excel task pane add-in</span></span>

<span data-ttu-id="55761-104">Neste tutorial: você criará um suplemento do painel de tarefas no Excel</span><span class="sxs-lookup"><span data-stu-id="55761-104">In this tutorial, you'll create an Excel task pane add-in that:</span></span>

> [!div class="checklist"]
> * <span data-ttu-id="55761-105">Cria uma tabela</span><span class="sxs-lookup"><span data-stu-id="55761-105">Creates a table</span></span>
> * <span data-ttu-id="55761-106">Filtra e classifica uma tabela</span><span class="sxs-lookup"><span data-stu-id="55761-106">Filters and sorts a table</span></span>
> * <span data-ttu-id="55761-107">Cria um gráfico</span><span class="sxs-lookup"><span data-stu-id="55761-107">Creates a chart</span></span>
> * <span data-ttu-id="55761-108">Congela um cabeçalho de tabela</span><span class="sxs-lookup"><span data-stu-id="55761-108">Freezes a table header</span></span>
> * <span data-ttu-id="55761-109">Protege uma planilha</span><span class="sxs-lookup"><span data-stu-id="55761-109">Protects a worksheet</span></span>
> * <span data-ttu-id="55761-110">Abre uma caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="55761-110">Opens a dialog</span></span>

## <a name="prerequisites"></a><span data-ttu-id="55761-111">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="55761-111">Prerequisites</span></span>

<span data-ttu-id="55761-112">Para usar este tutorial, você precisa instalar o seguinte.</span><span class="sxs-lookup"><span data-stu-id="55761-112">To use this tutorial, you need to have the following installed.</span></span> 

- <span data-ttu-id="55761-p101">Excel 2016, versão 1711 (build 8730.1000 do Clique para Executar) ou posterior. Talvez você precise ser um participante do programa Office Insider para ter essa versão. Para saber mais, confira a página [Seja um Office Insider](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="55761-p101">Excel 2016, version 1711 (Build 8730.1000 Click-to-Run) or later. You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span>

- [<span data-ttu-id="55761-116">Nó</span><span class="sxs-lookup"><span data-stu-id="55761-116">Node</span></span>](https://nodejs.org/en/) 

- <span data-ttu-id="55761-117">[Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)</span><span class="sxs-lookup"><span data-stu-id="55761-117">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

- <span data-ttu-id="55761-118">Você precisa ter uma conexão com a Internet para testar o suplemento neste tutorial.</span><span class="sxs-lookup"><span data-stu-id="55761-118">You need to have an Internet connection to test the add-in in this tutorial.</span></span>

## <a name="create-your-add-in-project"></a><span data-ttu-id="55761-119">Criar seu projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="55761-119">Create your add-in project</span></span>

<span data-ttu-id="55761-120">Conclua as etapas a seguir para criar o projeto de suplemento do Excel que você vai usar como base para este tutorial.</span><span class="sxs-lookup"><span data-stu-id="55761-120">Complete the following steps to create the Excel add-in project that you'll use as the basis for this tutorial.</span></span>

1. <span data-ttu-id="55761-121">Clone o repositório do GitHub com o [Tutorial de suplemento do Excel](https://github.com/OfficeDev/Excel-Add-in-Tutorial).</span><span class="sxs-lookup"><span data-stu-id="55761-121">Clone the GitHub repository [Excel add-in tutorial](https://github.com/OfficeDev/Excel-Add-in-Tutorial).</span></span>

2. <span data-ttu-id="55761-122">Abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue para a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="55761-122">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

3. <span data-ttu-id="55761-123">Execute o comando `npm install` para instalar as ferramentas e bibliotecas listadas no arquivo package.json.</span><span class="sxs-lookup"><span data-stu-id="55761-123">Run the command `npm install` to install the tools and libraries listed in the package.json file.</span></span> 

4. <span data-ttu-id="55761-124">Execute as etapas em [Adicionar certificados autoassinados como certificado raiz confiáveis](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) para confiar no certificado do sistema operacional do seu computador de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="55761-124">Carry out the steps in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to trust the certificate for your development computer's operating system.</span></span>

## <a name="create-a-table"></a><span data-ttu-id="55761-125">Criar uma tabela</span><span class="sxs-lookup"><span data-stu-id="55761-125">Create a table</span></span>

<span data-ttu-id="55761-126">Nesta etapa do tutorial, você testará no programa se o suplemento é compatível com a versão atual do Excel do usuário, adicionará uma tabela a uma planilha, depois preencherá e formatará a tabela com os dados.</span><span class="sxs-lookup"><span data-stu-id="55761-126">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Excel, add a table to a worksheet, populate the table with data, and format it.</span></span>

### <a name="code-the-add-in"></a><span data-ttu-id="55761-127">Codificação do suplemento</span><span class="sxs-lookup"><span data-stu-id="55761-127">Code the add-in</span></span>

1. <span data-ttu-id="55761-128">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="55761-128">Open the project in your code editor.</span></span>

2. <span data-ttu-id="55761-129">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="55761-129">Open the file index.html.</span></span>

3. <span data-ttu-id="55761-130">Substitua `TODO1` pela marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="55761-130">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="create-table">Create Table</button>
    ```

4. <span data-ttu-id="55761-131">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="55761-131">Open the app.js file.</span></span>

5. <span data-ttu-id="55761-p102">Substitua o `TODO1` pelo código a seguir. O código determina se a versão do Excel do usuário proporciona suporte a uma versão do Excel.js que inclua as APIs com esta série de tutoriais. Em um suplemento de produção, use o corpo do bloco condicional para ocultar ou desabilitar a interface do usuário que chame a APIs sem suporte. Dessa forma, permitirá que o usuário ainda use as partes do suplemento às quais a versão do Excel dá suporte.</span><span class="sxs-lookup"><span data-stu-id="55761-p102">Replace the `TODO1` with the following code. This code determines whether the user's version of Excel supports a version of Excel.js that includes all the APIs that this series of tutorials will use. In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs. This will enable the user to still make use of the parts of the add-in that are supported by their version of Excel.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }
    ```

6. <span data-ttu-id="55761-136">Substitua o `TODO2` pelo código a seguir:</span><span class="sxs-lookup"><span data-stu-id="55761-136">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#create-table').click(createTable);
    ```

7. <span data-ttu-id="55761-137">Substitua o `TODO3` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="55761-137">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="55761-138">Observação:</span><span class="sxs-lookup"><span data-stu-id="55761-138">Note:</span></span>

   - <span data-ttu-id="55761-p104">A lógica de negócios de Excel.js será adicionada à função que passar por `Excel.run`. Essa lógica não é executada imediatamente. Em vez disso, ela é adicionada à fila de comandos pendentes.</span><span class="sxs-lookup"><span data-stu-id="55761-p104">Your Excel.js business logic will be added to the function that is passed to `Excel.run`. This logic does not execute immediately. Instead, it is added to a queue of pending commands.</span></span>

   - <span data-ttu-id="55761-142">O método `context.sync` envia todos os comandos da fila para execução no Excel.</span><span class="sxs-lookup"><span data-stu-id="55761-142">The `context.sync` method sends all queued commands to Excel for execution.</span></span>

   - <span data-ttu-id="55761-p105">`Excel.run` é seguido por um bloco `catch`. Essa é uma prática recomendada que você sempre deve seguir.</span><span class="sxs-lookup"><span data-stu-id="55761-p105">The `Excel.run` is followed by a `catch` block. This is a best practice that you should always follow.</span></span> 

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

8. <span data-ttu-id="55761-145">Substitua `TODO4` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="55761-145">Replace `TODO4` with the following code.</span></span> <span data-ttu-id="55761-146">Observação:</span><span class="sxs-lookup"><span data-stu-id="55761-146">Note:</span></span>

   - <span data-ttu-id="55761-p107">O código cria uma tabela usando o método `add` de conjunto de tabela da planilha, que sempre existe mesmo que ela esteja vazia. Essa é a maneira padrão de criar objetos no Excel.js. Não há nenhuma API do construtor de classe e você nunca usará um operador `new` para criar um objeto do Excel. Em vez disso, adicione a um objeto de conjunto pai.</span><span class="sxs-lookup"><span data-stu-id="55761-p107">The code creates a table by using `add` method of a worksheet's table collection, which always exists even if it is empty. This is the standard way that Excel.js objects are created. There are no class constructor APIs, and you never use a `new` operator to create an Excel object. Instead, you add to a parent collection object.</span></span>

   - <span data-ttu-id="55761-p108">O primeiro parâmetro do método `add`é o intervalo apenas da linha superior da tabela, não o intervalo inteiro que a tabela por fim usará. Isso ocorre porque, quando o suplemento preenche as linhas de dados (na próxima etapa), ele adicionará novas linhas à tabela, em vez de gravar os valores nas células das linhas existentes. Esse é um padrão mais comum, porque o número de linhas em uma tabela geralmente não é conhecido quando a tabela é criada.</span><span class="sxs-lookup"><span data-stu-id="55761-p108">The first parameter of the `add` method is the range of only the top row of the table, not the entire range the table will ultimately use. This is because when the add-in populates the data rows (in the next step), it will add new rows to the table instead of writing values to the cells of existing rows. This is a more common pattern because the number of rows that a table will have is often not known when the table is created.</span></span>

   - <span data-ttu-id="55761-154">Os nomes de tabelas devem ser exclusivos pela pasta de trabalho inteira, não só na planilha.</span><span class="sxs-lookup"><span data-stu-id="55761-154">Table names must be unique across the entire workbook, not just the worksheet.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ```

9. <span data-ttu-id="55761-155">Substitua `TODO5` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="55761-155">Replace `TODO5` with the following code.</span></span> <span data-ttu-id="55761-156">Observação:</span><span class="sxs-lookup"><span data-stu-id="55761-156">Note:</span></span>

   - <span data-ttu-id="55761-157">Os valores das células de um intervalo são definidos em uma matriz de matrizes.</span><span class="sxs-lookup"><span data-stu-id="55761-157">The cell values of a range are set with an array of arrays.</span></span>

   - <span data-ttu-id="55761-p110">Novas linhas são criadas em uma tabela ao chamar o método `add` do conjunto de linhas da tabela. Você pode adicionar várias linhas em uma única chamada de `add` ao incluir várias matrizes de valores de células na matriz pai que é passada como segundo parâmetro.</span><span class="sxs-lookup"><span data-stu-id="55761-p110">New rows are created in a table by calling the `add` method of the table's row collection. You can add multiple rows in a single call of `add` by including multiple cell value arrays in the parent array that is passed as the second parameter.</span></span>

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

10. <span data-ttu-id="55761-160">Substitua `TODO6` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="55761-160">Replace `TODO6` with the following code.</span></span> <span data-ttu-id="55761-161">Observação:</span><span class="sxs-lookup"><span data-stu-id="55761-161">Note:</span></span>

   - <span data-ttu-id="55761-162">O código recebe uma referência para a coluna **quantidade** ao passar o índice com base em zero para o método `getItemAt` do conjunto de colunas da tabela.</span><span class="sxs-lookup"><span data-stu-id="55761-162">The code gets a reference to the **Amount** column by passing its zero-based index to the `getItemAt` method of the table's column collection.</span></span>

     > [!NOTE]
     > <span data-ttu-id="55761-163">Os objetos do conjunto Excel.js, como `TableCollection`, `WorksheetCollection`, e `TableColumnCollection`, têm a propriedade `items` que é como uma matriz dos tipos de objetos filhos, como `Table` ou `Worksheet` ou `TableColumn`; mas um objeto `*Collection` não é uma matriz.</span><span class="sxs-lookup"><span data-stu-id="55761-163">Excel.js collection objects, such as `TableCollection`, `WorksheetCollection`, and `TableColumnCollection` have an `items` property that is an array of the child object types, such as `Table` or `Worksheet` or `TableColumn`; but a `*Collection` object is not itself an array.</span></span>

   - <span data-ttu-id="55761-164">O código formata o intervalo da coluna **quantidade** como Euros com um segundo decimal.</span><span class="sxs-lookup"><span data-stu-id="55761-164">The code then formats the range of the **Amount** column as Euros to the second decimal.</span></span> 

   - <span data-ttu-id="55761-p112">Por fim, isso garante que a largura das colunas e a altura das linhas sejam grandes o suficiente para o maior (ou o mais alto) item de dados. Observe que o código deve receber os objetos `Range` a formatar. Os objetos `TableColumn` e `TableRow` não têm propriedades de formato.</span><span class="sxs-lookup"><span data-stu-id="55761-p112">Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item. Notice that the code must get `Range` objects to format. `TableColumn` and `TableRow` objects do not have format properties.</span></span>

        ```js
        expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
        expensesTable.getRange().format.autofitColumns();
        expensesTable.getRange().format.autofitRows();
        ```

### <a name="test-the-add-in"></a><span data-ttu-id="55761-168">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="55761-168">Test the add-in</span></span>

1. <span data-ttu-id="55761-169">Abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue para a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="55761-169">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

2. <span data-ttu-id="55761-170">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte no Internet Explorer (que é usada em segundo plano pelo Excel para executar os suplementos do Excel).</span><span class="sxs-lookup"><span data-stu-id="55761-170">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="55761-171">Execute o comando `npm start` para iniciar um servidor Web em um host local.</span><span class="sxs-lookup"><span data-stu-id="55761-171">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="55761-172">Realize o sideload do suplemento usando um dos métodos a seguir:</span><span class="sxs-lookup"><span data-stu-id="55761-172">Sideload the add-in by using one of the following methods:</span></span>

    - <span data-ttu-id="55761-173">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="55761-173">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>

    - <span data-ttu-id="55761-174">Excel Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="55761-174">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>

    - <span data-ttu-id="55761-175">iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="55761-175">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

5. <span data-ttu-id="55761-176">No menu **Página Inicial**, escolha **Mostrar Painel de Tarefas**.</span><span class="sxs-lookup"><span data-stu-id="55761-176">On the **Home** menu, choose **Show Taskpane**.</span></span>

6. <span data-ttu-id="55761-177">No painel de tarefas, escolha **Criar Tabela**.</span><span class="sxs-lookup"><span data-stu-id="55761-177">In the task pane, choose **Create Table**.</span></span>

    ![Tutorial do Excel: Criar tabela](../images/excel-tutorial-create-table.png)

## <a name="filter-and-sort-a-table"></a><span data-ttu-id="55761-179">Filtrar e classificar uma tabela</span><span class="sxs-lookup"><span data-stu-id="55761-179">Filter and sort a table</span></span>

<span data-ttu-id="55761-180">Nesta etapa do tutorial, você vai filtrar e classificar a tabela que criou anteriormente.</span><span class="sxs-lookup"><span data-stu-id="55761-180">In this step of the tutorial, you'll filter and sort the table that you created previously.</span></span>

### <a name="filter-the-table"></a><span data-ttu-id="55761-181">Filtrar a tabela</span><span class="sxs-lookup"><span data-stu-id="55761-181">Filter the table</span></span>

1. <span data-ttu-id="55761-182">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="55761-182">Open the project in your code editor.</span></span>

2. <span data-ttu-id="55761-183">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="55761-183">Open the file index.html.</span></span>

3. <span data-ttu-id="55761-184">Abaixo do `div`, que contém o botão `create-table`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="55761-184">Just below the `div` that contains the `create-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="filter-table">Filter Table</button>
    </div>
    ```

4. <span data-ttu-id="55761-185">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="55761-185">Open the app.js file.</span></span>

5. <span data-ttu-id="55761-186">Logo abaixo da linha que atribui um identificador de clique ao botão `create-table`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="55761-186">Just below the line that assigns a click handler to the `create-table` button, add the following code:</span></span>

    ```js
    $('#filter-table').click(filterTable);
    ```

6. <span data-ttu-id="55761-187">Logo abaixo da função `createTable`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="55761-187">Just below the `createTable` function, add the following function:</span></span>

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

7. <span data-ttu-id="55761-188">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="55761-188">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="55761-189">Observação:</span><span class="sxs-lookup"><span data-stu-id="55761-189">Note:</span></span>

   - <span data-ttu-id="55761-p114">O código primeiro faz referência à coluna que precisa de filtragem ao passar o nome da coluna para o método `getItem`, em vez de passar o índice para o método `getItemAt` como o método `createTable` faz. Como os usuários podem mover as colunas da tabela, a coluna de um determinado índice pode mudar depois da criação da tabela. Portanto, é mais seguro usar o nome da coluna como referência dela. Usamos de forma segura `getItemAt` em um tutorial anterior porque usamos o mesmo método que cria a tabela. Assim não existe a chance de um usuário mover a coluna.</span><span class="sxs-lookup"><span data-stu-id="55761-p114">The code first gets a reference to the column that needs filtering by passing the column name to the `getItem` method, instead of passing its index to the `getItemAt` method as the `createTable` method does. Since users can move table columns, the column at a given index might change after the table is created. Hence, it is safer to use the column name to get a reference to the column. We used `getItemAt` safely in the preceding tutorial, because we used it in the very same method that creates the table, so there is no chance that a user has moved the column.</span></span>

   - <span data-ttu-id="55761-194">O método `applyValuesFilter` é um dos vários métodos de filtragem do objeto `Filter`.</span><span class="sxs-lookup"><span data-stu-id="55761-194">The `applyValuesFilter` method is one of several filtering methods on the `Filter` object.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);
    ``` 

### <a name="sort-the-table"></a><span data-ttu-id="55761-195">Classificar a tabela</span><span class="sxs-lookup"><span data-stu-id="55761-195">Sort the table</span></span>

1. <span data-ttu-id="55761-196">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="55761-196">Open the file index.html.</span></span>

2. <span data-ttu-id="55761-197">Abaixo do `div` que contém o botão `filter-table`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="55761-197">Below the `div` that contains the `filter-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="sort-table">Sort Table</button>
    </div>
    ```

3. <span data-ttu-id="55761-198">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="55761-198">Open the app.js file.</span></span>

4. <span data-ttu-id="55761-199">Abaixo da linha que atribui um identificador de clique ao botão `filter-table`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="55761-199">Below the line that assigns a click handler to the `filter-table` button, add the following code:</span></span>

    ```js
    $('#sort-table').click(sortTable);
    ```

5. <span data-ttu-id="55761-200">Abaixo da função `filterTable`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="55761-200">Below the `filterTable` function add the following function.</span></span>

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

6. <span data-ttu-id="55761-201">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="55761-201">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="55761-202">Observação:</span><span class="sxs-lookup"><span data-stu-id="55761-202">Note:</span></span>

   - <span data-ttu-id="55761-203">O código cria uma matriz de objetos `SortField` que tem apenas um membro, já que o suplemento só classifica a coluna Comerciante.</span><span class="sxs-lookup"><span data-stu-id="55761-203">The code creates an array of `SortField` objects which has just one member since the add-in only sorts on the Merchant column.</span></span>

   - <span data-ttu-id="55761-204">A propriedade `key` de um objeto `SortField` é o índice com base em zero da coluna a classificar.</span><span class="sxs-lookup"><span data-stu-id="55761-204">The `key` property of a `SortField` object is the zero-based index of the column to sort-on.</span></span>

   - <span data-ttu-id="55761-205">O membro `sort` de uma `Table` é um objeto `TableSort`, não um método.</span><span class="sxs-lookup"><span data-stu-id="55761-205">The `sort` member of a `Table` is a `TableSort` object, not a method.</span></span> <span data-ttu-id="55761-206">Os `SortField`s são passados para o método `apply` do objeto `TableSort`.</span><span class="sxs-lookup"><span data-stu-id="55761-206">The `SortField`s are passed to the `TableSort` object's `apply` method.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var sortFields = [
        {
            key: 1,            // Merchant column
            ascending: false,
        }
    ];

    expensesTable.sort.apply(sortFields);
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="55761-207">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="55761-207">Test the add-in</span></span>

1. <span data-ttu-id="55761-208">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite **Ctrl + C** duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="55761-208">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter **Ctrl+C** twice to stop the running web server.</span></span> <span data-ttu-id="55761-209">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="55761-209">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="55761-210">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor.</span><span class="sxs-lookup"><span data-stu-id="55761-210">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="55761-211">Para fazer isso, interrompa o processo do servidor para obter uma solicitação para inserir o comando de compilação.</span><span class="sxs-lookup"><span data-stu-id="55761-211">In order to do this, you need to kill the server process so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="55761-212">Após a compilação, reinicie o servidor.</span><span class="sxs-lookup"><span data-stu-id="55761-212">After the build, you restart the server.</span></span> <span data-ttu-id="55761-213">As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="55761-213">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="55761-214">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte no Internet Explorer (que é usada em segundo plano pelo Excel para executar os suplementos do Excel).</span><span class="sxs-lookup"><span data-stu-id="55761-214">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="55761-215">Execute o comando `npm start` para iniciar um servidor Web em um host local.</span><span class="sxs-lookup"><span data-stu-id="55761-215">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="55761-216">Feche o painel de tarefas para recarregá-lo e, no menu **Página Inicial**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="55761-216">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="55761-217">Se por qualquer motivo a tabela não estiver na planilha aberta, no painel de tarefas, escolha **Criar Tabela**.</span><span class="sxs-lookup"><span data-stu-id="55761-217">If for any reason the table is not in the open worksheet, in the task pane, choose **Create Table**.</span></span>

6. <span data-ttu-id="55761-218">Escolha os botões **Filtrar Tabela** e **Classificar Tabela** em qualquer ordem.</span><span class="sxs-lookup"><span data-stu-id="55761-218">Choose the **Filter Table** and **Sort Table** buttons, in either order.</span></span>

    ![Tutorial do Excel: filtrar e classificar tabela](../images/excel-tutorial-filter-and-sort-table.png)

## <a name="create-a-chart"></a><span data-ttu-id="55761-220">Criar um gráfico</span><span class="sxs-lookup"><span data-stu-id="55761-220">Create a chart</span></span>

<span data-ttu-id="55761-221">Nesta etapa do tutorial, você vai criar um gráfico com dados da tabela que você criou anteriormente e depois vai formatar o gráfico.</span><span class="sxs-lookup"><span data-stu-id="55761-221">In this step of the tutorial, you'll create a chart using data from the table that you created previously, and then format the chart.</span></span>

### <a name="chart-a-chart-using-table-data"></a><span data-ttu-id="55761-222">Gráfico de um gráfico com dados de tabela</span><span class="sxs-lookup"><span data-stu-id="55761-222">Chart a chart using table data</span></span>

1. <span data-ttu-id="55761-223">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="55761-223">Open the project in your code editor.</span></span>

2. <span data-ttu-id="55761-224">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="55761-224">Open the file index.html.</span></span>

3. <span data-ttu-id="55761-225">Abaixo do `div` que contém o botão `sort-table`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="55761-225">Below the `div` that contains the `sort-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="create-chart">Create Chart</button>
    </div>
    ```

4. <span data-ttu-id="55761-226">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="55761-226">Open the app.js file.</span></span>

5. <span data-ttu-id="55761-227">Abaixo da linha que atribui um identificador de clique ao botão `sort-chart`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="55761-227">Below the line that assigns a click handler to the `sort-chart` button, add the following code:</span></span>

    ```js
    $('#create-chart').click(createChart);
    ```

6. <span data-ttu-id="55761-228">Abaixo da função `sortTable`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="55761-228">Below the `sortTable` function add the following function.</span></span>

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

7. <span data-ttu-id="55761-p119">Substitua `TODO1` pelo código a seguir. Para excluir a linha de cabeçalho, o código usa o método `Table.getDataBodyRange` para acessar o intervalo de dados que você deseja representar graficamente em vez do método `getRange`.</span><span class="sxs-lookup"><span data-stu-id="55761-p119">Replace `TODO1` with the following code. Note that in order to exclude the header row, the code uses the `Table.getDataBodyRange` method to get the range of data you want to chart instead of the `getRange` method.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var dataRange = expensesTable.getDataBodyRange();
    ```

8. <span data-ttu-id="55761-p120">Substitua `TODO2` pelo código a seguir. Observe os seguintes parâmetros:</span><span class="sxs-lookup"><span data-stu-id="55761-p120">Replace `TODO2` with the following code. Note the following parameters:</span></span>

   - <span data-ttu-id="55761-p121">O primeiro parâmetro para o método `add` especifica o tipo de gráfico. Há diversos tipos.</span><span class="sxs-lookup"><span data-stu-id="55761-p121">The first parameter to the `add` method specifies the type of chart. There are several dozen types.</span></span>

   - <span data-ttu-id="55761-235">O segundo parâmetro especifica um intervalo de dados a incluir no gráfico.</span><span class="sxs-lookup"><span data-stu-id="55761-235">The second parameter specifies the range of data to include in the chart.</span></span>

   - <span data-ttu-id="55761-236">O terceiro parâmetro determina se uma série de pontos de dados da tabela deve estar representada por linha ou por coluna.</span><span class="sxs-lookup"><span data-stu-id="55761-236">The third parameter determines whether a series of data points from the table should be charted row-wise or column-wise.</span></span> <span data-ttu-id="55761-237">A opção `auto` informa ao Excel para decidir o melhor método.</span><span class="sxs-lookup"><span data-stu-id="55761-237">The option `auto` tells Excel to decide the best method.</span></span>

    ```js
    var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ```

9. <span data-ttu-id="55761-p123">Substitua `TODO3` pelo código a seguir. A maior parte do código é autoexplicativa. Observação:</span><span class="sxs-lookup"><span data-stu-id="55761-p123">Replace `TODO3` with the following code. Most of this code is self-explanatory. Note:</span></span>
   
   - <span data-ttu-id="55761-p124">Os parâmetros do método `setPosition` especificam as células da esquerda superior e da direita inferior da área da planilha que deve conter o gráfico. O Excel ajusta detalhes como a largura da linha para criar uma boa aparência para o gráfico no espaço fornecido.</span><span class="sxs-lookup"><span data-stu-id="55761-p124">The parameters to the `setPosition` method specify the upper left and lower right cells of the worksheet area that should contain the chart. Excel can adjust things like line width to make the chart look good in the space it has been given.</span></span>
   
   - <span data-ttu-id="55761-p125">"Série" é um conjunto de pontos de dados de uma coluna da tabela. Como há apenas uma coluna sem cadeia de caracteres na tabela, o Excel deduz que essa é a única coluna de pontos de dados no gráfico. Ele interpreta outras colunas como rótulos do gráfico. Portanto, haverá apenas uma série no gráfico e será necessário o índice 0. Ele será rotulado como "Valor em €".</span><span class="sxs-lookup"><span data-stu-id="55761-p125">A "series" is a set of data points from a column of the table. Since there is only one non-string column in the table, Excel infers that the column is the only column of data points to chart. It interprets the other columns as chart labels. So there will be just one series in the chart and it will have index 0. This is the one to label with "Value in €".</span></span>

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in €';
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="55761-248">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="55761-248">Test the add-in</span></span>

1. <span data-ttu-id="55761-249">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite **Ctrl + C** duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="55761-249">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter **Ctrl+C** twice to stop the running web server.</span></span> <span data-ttu-id="55761-250">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="55761-250">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="55761-p127">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor. Para fazer isso, interrompa o processo do servidor para obter uma solicitação para inserir o comando de compilação. Após a compilação, reinicie o servidor. As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="55761-p127">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command. After the build, you restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="55761-255">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte no Internet Explorer (que é usada em segundo plano pelo Excel para executar os suplementos do Excel).</span><span class="sxs-lookup"><span data-stu-id="55761-255">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="55761-256">Execute o comando `npm start` para iniciar um servidor Web em um host local.</span><span class="sxs-lookup"><span data-stu-id="55761-256">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="55761-257">Feche o painel de tarefas para recarregá-lo e, no menu **Página Inicial**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="55761-257">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="55761-258">Se, por algum motivo, a tabela não estiver na planilha aberta, no painel de tarefas, escolha **Criar Tabela** e depois os botões **Filtrar Tabela** e \*\*Classificar Tabela \*\* em qualquer ordem.</span><span class="sxs-lookup"><span data-stu-id="55761-258">If for any reason the table is not in the open worksheet, in the task pane, choose **Create Table** and then **Filter Table** and **Sort Table** buttons, in either order.</span></span>

6. <span data-ttu-id="55761-p128">Clique no botão **Criar gráfico**. Um gráfico é criado e incluirá somente os dados das linhas que foram filtradas. Os rótulos dos pontos de dados na parte inferior estão na ordem de classificação do gráfico, ou seja, nomes de comerciantes em ordem alfabética inversa.</span><span class="sxs-lookup"><span data-stu-id="55761-p128">Choose the **Create Chart** button. A chart is created and only the data from the rows that have been filtered are included. The labels on the data points across the bottom are in the sort order of the chart; that is, merchant names in reverse alphabetical order.</span></span>

    ![Tutorial do Excel - Criar gráfico ](../images/excel-tutorial-create-chart.png)

## <a name="freeze-a-table-header"></a><span data-ttu-id="55761-263">Congelar um cabeçalho de tabela</span><span class="sxs-lookup"><span data-stu-id="55761-263">Freeze a table header</span></span>

<span data-ttu-id="55761-p129">Quando uma tabela for longa o suficiente para que um usuário precise rolar para ver algumas linhas, a linha de cabeçalho poderá ficar fora da vista. Nesta etapa do tutorial, você precisará congelar a linha do cabeçalho da tabela que criou anteriormente para que ela permaneça visível, mesmo que o usuário role ao longo da planilha.</span><span class="sxs-lookup"><span data-stu-id="55761-p129">When a table is long enough that a user must scroll to see some rows, the header row can scroll out of sight. In this step of the tutorial, you'll freeze the header row of the table that you created previously, so that it remains visible even as the user scrolls down the worksheet.</span></span>

### <a name="freeze-the-tables-header-row"></a><span data-ttu-id="55761-266">Congelar a linha de cabeçalho da tabela</span><span class="sxs-lookup"><span data-stu-id="55761-266">Freeze the table's header row</span></span>

1. <span data-ttu-id="55761-267">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="55761-267">Open the project in your code editor.</span></span>

2. <span data-ttu-id="55761-268">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="55761-268">Open the file index.html.</span></span>

3. <span data-ttu-id="55761-269">Abaixo do `div` que contém o botão `create-chart`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="55761-269">Below the `div` that contains the `create-chart` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="freeze-header">Freeze Header</button>
    </div>
    ```

4. <span data-ttu-id="55761-270">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="55761-270">Open the app.js file.</span></span>

5. <span data-ttu-id="55761-271">Abaixo da linha que atribui um identificador de clique ao botão `create-chart`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="55761-271">Below the line that assigns a click handler to the `create-chart` button, add the following code:</span></span>

    ```js
    $('#freeze-header').click(freezeHeader);
    ```

6. <span data-ttu-id="55761-272">Abaixo da função `createChart`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="55761-272">Below the `createChart` function add the following function:</span></span>

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

7. <span data-ttu-id="55761-273">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="55761-273">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="55761-274">Observação:</span><span class="sxs-lookup"><span data-stu-id="55761-274">Note:</span></span>

   - <span data-ttu-id="55761-275">A coleção `Worksheet.freezePanes` é um conjunto de painéis da planilha que fica congelado ou fixado no mesmo lugar quando rolamos a planilha.</span><span class="sxs-lookup"><span data-stu-id="55761-275">The `Worksheet.freezePanes` collection is a set of panes in the worksheet that are pinned, or frozen, in place when the worksheet is scrolled.</span></span>

   - <span data-ttu-id="55761-p131">O método `freezeRows` considera como parâmetro o número de linhas, começando da parte superior, que devem ser fixadas no local. Passamos `1` para fixar a primeira linha no local.</span><span class="sxs-lookup"><span data-stu-id="55761-p131">The `freezeRows` method takes as a parameter the number of rows, from the top that are to be pinned in place. We pass `1` to pin the first row in place.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="55761-278">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="55761-278">Test the add-in</span></span>

1. <span data-ttu-id="55761-279">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite **Ctrl + C** duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="55761-279">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter **Ctrl+C** twice to stop the running web server.</span></span> <span data-ttu-id="55761-280">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="55761-280">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="55761-p133">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor. Para fazer isso, interrompa o processo do servidor para obter uma solicitação para inserir o comando de compilação. Após a compilação, reinicie o servidor. As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="55761-p133">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command. After the build, you restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="55761-285">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte no Internet Explorer (que é usada em segundo plano pelo Excel para executar os suplementos do Excel).</span><span class="sxs-lookup"><span data-stu-id="55761-285">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="55761-286">Execute o comando `npm start` para iniciar um servidor Web em um host local.</span><span class="sxs-lookup"><span data-stu-id="55761-286">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="55761-287">Feche o painel de tarefas para recarregá-lo e, no menu **Início**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="55761-287">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="55761-288">Se a tabela estiver na planilha, exclua-a.</span><span class="sxs-lookup"><span data-stu-id="55761-288">If the table is in the worksheet, delete it.</span></span>

6. <span data-ttu-id="55761-289">No painel de tarefas, escolha **Criar Tabela**.</span><span class="sxs-lookup"><span data-stu-id="55761-289">In the task pane, choose **Create Table**.</span></span>

7. <span data-ttu-id="55761-290">Escolha o botão **Congelar Cabeçalho**.</span><span class="sxs-lookup"><span data-stu-id="55761-290">Choose the **Freeze Header** button.</span></span>

8. <span data-ttu-id="55761-291">Role a planilha para baixo, o suficiente para ver que o cabeçalho da tabela permanece visível na parte superior mesmo ao rolar até que as primeiras linhas fiquem fora da vista.</span><span class="sxs-lookup"><span data-stu-id="55761-291">Scroll down the worksheet enough to to see that the table header remains visible at the top even when the higher rows scroll out of sight.</span></span>

    ![Tutorial do Excel: congelar cabeçalho](../images/excel-tutorial-freeze-header.png)

## <a name="protect-a-worksheet"></a><span data-ttu-id="55761-293">Proteger uma planilha</span><span class="sxs-lookup"><span data-stu-id="55761-293">Protect a worksheet</span></span>

<span data-ttu-id="55761-294">Nesta etapa do tutorial, você adicionará outro botão à faixa de opções que, quando selecionado, executa uma função que você precisará definir para ativar e desativar a proteção da planilha.</span><span class="sxs-lookup"><span data-stu-id="55761-294">In this step of the tutorial, you'll add another button to the ribbon that, when chosen, executes a function that you'll define to toggle worksheet protection on and off.</span></span>

### <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a><span data-ttu-id="55761-295">Configure o manifesto para adicionar um segundo botão à faixa de opções</span><span class="sxs-lookup"><span data-stu-id="55761-295">Configure the manifest to add a second ribbon button</span></span>

1. <span data-ttu-id="55761-296">Abra o arquivo de manifesto my-office-add-in-manifest.xml.</span><span class="sxs-lookup"><span data-stu-id="55761-296">Open the manifest file my-office-add-in-manifest.xml.</span></span>

2. <span data-ttu-id="55761-p134">Encontre o elemento `<Control>`. Esse elemento define o botão **Mostrar Painel de Tarefas** na faixa de opções **Início** que você usa para iniciar o suplemento. Vamos adicionar um segundo botão ao mesmo grupo na faixa de opções **Início**. Entre a marca de Controle final (`</Control>`) e a marca de Grupo final (`</Group>`), adicione a marcação a seguir.</span><span class="sxs-lookup"><span data-stu-id="55761-p134">Find the `<Control>` element. This element defines the **Show Taskpane** button on the **Home** ribbon you have been using to launch the add-in. We're going to add a second button to the same group on the **Home** ribbon. In between the end Control tag (`</Control>`) and the end Group tag (`</Group>`), add the following markup.</span></span>

    ```xml
    <Control xsi:type="Button" id="<!--TODO1: Unique (in manifest) name for button -->">
        <Label resid="<!--TODO2: Button label -->" />
        <Supertip>            
            <Title resid="<!-- TODO3: Button tool tip title -->" />
            <Description resid="<!-- TODO4: Button tool tip description -->" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="<!-- TODO5: Specify the type of action-->">
            <!-- TODO6: Identify the function.-->
        </Action>
    </Control>
    ```

3. <span data-ttu-id="55761-301">Substitua `TODO1` por uma cadeia de caracteres que fornece ao botão uma ID exclusiva no arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="55761-301">Replace `TODO1` with a string that gives the button an ID that is unique within this manifest file.</span></span> <span data-ttu-id="55761-302">Como nosso botão ativará ou desativará a proteção da planilha, use "ToggleProtection".</span><span class="sxs-lookup"><span data-stu-id="55761-302">Since our button is going to toggle protection of the worksheet on and off, use "ToggleProtection".</span></span> <span data-ttu-id="55761-303">Quando terminar, a marca de Controle de início inteira deve se parecer com o seguinte:</span><span class="sxs-lookup"><span data-stu-id="55761-303">When you are done, the entire start Control tag should look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. <span data-ttu-id="55761-p136">Os próximos três `TODO`s definem “resid”, que significa ID de recurso. Um recurso é uma cadeia de caracteres e você criará essas três cadeias de caracteres em uma etapa posterior. Por enquanto, você precisa fornecer IDs aos recursos. O rótulo do botão deve ser "Toggle Protection", mas a *ID* dessa cadeia de caracteres será "ProtectionButtonLabel", de forma que o elemento `Label` completo deve se parecer com o código a seguir:</span><span class="sxs-lookup"><span data-stu-id="55761-p136">The next three `TODO`s set "resid"s, which is short for resource ID. A resource is a string, and you'll create these three strings in a later step. For now, you need to give IDs to the resources. The button label should read "Toggle Protection", but the *ID* of this string should be "ProtectionButtonLabel", so the completed `Label` element should look like the following code:</span></span>

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. <span data-ttu-id="55761-p137">O elemento `SuperTip` define a dica de ferramenta do botão. O título da dica de ferramenta deve ser o mesmo que o rótulo do botão, por isso, usamos a mesma ID de recurso: "ProtectionButtonLabel". A descrição da dica de ferramenta será "Click to turn protection of the worksheet on and off". Mas o `ID` será "ProtectionButtonToolTip". Portanto, quando terminar, a marcação `SuperTip` inteira deve se parecer com o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="55761-p137">The `SuperTip` element defines the tool tip for the button. The tool tip title should be the same as the button label, so we use the very same resource ID: "ProtectionButtonLabel". The tool tip description will be "Click to turn protection of the worksheet on and off". But the `ID` should be "ProtectionButtonToolTip". So, when you are done, the whole `SuperTip` markup should look like the following code:</span></span> 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > <span data-ttu-id="55761-p138">Em um suplemento de produção,não é recomendável usar o mesmo ícone para dois botões diferentes; mas, para simplificar este tutorial, faremos isso. Portanto, a marcação `Icon` em nosso novo `Control` é apenas uma cópia do elemento `Icon` do `Control` existente.</span><span class="sxs-lookup"><span data-stu-id="55761-p138">In a production add-in, you would not want to use the same icon for two different buttons; but to simplify this tutorial, we'll do that. So the `Icon` markup in our new `Control` is just a copy of the `Icon` element from the existing `Control`.</span></span> 

6. <span data-ttu-id="55761-p139">O elemento `Action` dentro do elemento `Control` original já está presente no manifesto, tem seu tipo definido como `ShowTaskpane`, mas nosso novo botão não abrirá um painel de tarefas, mas sim executará uma função personalizada criada em uma etapa posterior. Portanto, substitua `TODO5` por `ExecuteFunction`, que é o tipo de ação para botões que acionam funções personalizadas. A marca `Action` de início deve ser similar ao código abaixo:</span><span class="sxs-lookup"><span data-stu-id="55761-p139">The `Action` element inside the original `Control` element that was already present in the manifest, has its type set to `ShowTaskpane`, but our new button isn't going to open a task pane; it's going to run a custom function that you create in a later step. So replace `TODO5` with `ExecuteFunction` which is the action type for buttons that trigger custom functions. The start `Action` tag should look like the following code:</span></span>
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. <span data-ttu-id="55761-p140">O elemento `Action` original tem elementos filhos que especificam uma ID do painel de tarefas e uma URL da página que deve ser aberta no painel de tarefas. No entanto, um elemento `Action` do tipo `ExecuteFunction` tem um único elemento filho que nomeia a função executada pelo controle. Você criará essa função em uma etapa posterior e ela será chamada de `toggleProtection`. Então, substitua `TODO6` pela marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="55761-p140">The original `Action` element has child elements that specify a task pane ID and a URL of the page that should be opened in the task pane. But an `Action` element of the `ExecuteFunction` type has a single child element that names the function that the control executes. You'll create that function in a later step, and it will be called `toggleProtection`. So, replace `TODO6` with the following markup:</span></span>
 
    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    <span data-ttu-id="55761-322">A marcação `Control` inteira deve ter a aparência a seguir:</span><span class="sxs-lookup"><span data-stu-id="55761-322">The entire `Control` markup should now look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
        <Label resid="ProtectionButtonLabel" />
        <Supertip>            
            <Title resid="ProtectionButtonLabel" />
            <Description resid="ProtectionButtonToolTip" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Contoso.tpicon_16x16" />
            <bt:Image size="32" resid="Contoso.tpicon_32x32" />
            <bt:Image size="80" resid="Contoso.tpicon_80x80" />
        </Icon>
        <Action xsi:type="ExecuteFunction">
           <FunctionName>toggleProtection</FunctionName>
        </Action>
    </Control>
    ```

8. <span data-ttu-id="55761-323">Role para baixo até a seção `Resources` do manifesto.</span><span class="sxs-lookup"><span data-stu-id="55761-323">Scroll down to the `Resources` section of the manifest.</span></span>

9. <span data-ttu-id="55761-324">Adicione a seguinte marcação como filho do elemento `bt:ShortStrings`.</span><span class="sxs-lookup"><span data-stu-id="55761-324">Add the following markup as a child of the `bt:ShortStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. <span data-ttu-id="55761-325">Adicione a seguinte marcação como filho do elemento `bt:LongStrings`.</span><span class="sxs-lookup"><span data-stu-id="55761-325">Add the following markup as a child of the `bt:LongStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. <span data-ttu-id="55761-326">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="55761-326">Save the file.</span></span>

### <a name="create-the-function-that-protects-the-sheet"></a><span data-ttu-id="55761-327">Criar a função que protege a planilha</span><span class="sxs-lookup"><span data-stu-id="55761-327">Create the function that protects the sheet</span></span>

1. <span data-ttu-id="55761-328">Abra o arquivo \function-file\function-file.js.</span><span class="sxs-lookup"><span data-stu-id="55761-328">Open the file \function-file\function-file.js.</span></span>

2. <span data-ttu-id="55761-329">O arquivo já tem uma Expressão de Função Invocada Imediatamente (IFFE).</span><span class="sxs-lookup"><span data-stu-id="55761-329">The file already has an Immediately Invoked Function Expression (IFFE).</span></span> <span data-ttu-id="55761-330">*Fora do IIFE*, adicione o código a seguir.</span><span class="sxs-lookup"><span data-stu-id="55761-330">*Outside of the IIFE*, add the following code.</span></span> <span data-ttu-id="55761-331">Observe que é possível especificar um parâmetro `args` para o método e a última linha do método chamará `args.completed`.</span><span class="sxs-lookup"><span data-stu-id="55761-331">Note that we specify an `args` parameter to the method and the very last line of the method calls `args.completed`.</span></span> <span data-ttu-id="55761-332">Esse é um requisito para todos os comandos de suplemento do tipo **ExecuteFunction**.</span><span class="sxs-lookup"><span data-stu-id="55761-332">This is a requirement for all add-in commands of type **ExecuteFunction**.</span></span> <span data-ttu-id="55761-333">Ele sinaliza para o aplicativo host do Office que a função terminou e que a interface do usuário podem ficar responsiva novamente.</span><span class="sxs-lookup"><span data-stu-id="55761-333">It signals the Office host application that the function has finished and the UI can become responsive again.</span></span>

    ```js
    function toggleProtection(args) {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to reverse the protection status of the current worksheet.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```

3. <span data-ttu-id="55761-p142">Substitua `TODO1` pelo código a seguir. O código usa propriedade de proteção do objeto de planilha em um padrão de botão de alternância padrão. O `TODO2` será explicado na próxima seção.</span><span class="sxs-lookup"><span data-stu-id="55761-p142">Replace `TODO1` with the following code. This code uses the worksheet object's protection property in a standard toggle pattern. The `TODO2` will be explained in the next section.</span></span>

    ```js
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // TODO2: Queue command to load the sheet's "protection.protected" property from
    //        the document and re-synchronize the document and task pane.

     if (sheet.protection.protected) {
        sheet.protection.unprotect();
    } else {
        sheet.protection.protect();
    }
    ``` 

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="55761-337">Adicione código para buscar propriedades do documento em objetos de script do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="55761-337">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="55761-p143">Em todas as funções anteriores desta série de tutoriais, você colocou em fila comandos para *gravar* no documento do Office. Cada função terminou com uma chamada para o método `context.sync()`, que envia os comandos em fila para o documento a ser executado. Entretanto, o código adicionado na última etapa chama a propriedade `sheet.protection.protected` e essa é uma grande diferença das funções anteriores que você escreveu, pois o objeto `sheet` é apenas um objeto de proxy que existe no script do seu painel de tarefas. Ele não sabe qual é o estado real de proteção do documento, portanto, sua propriedade `protection.protected` não pode ter um valor real. É necessário primeiro buscar o status de proteção do documento e definir o valor de `sheet.protection.protected`. Somente então será possível chamar `sheet.protection.protected` sem causar uma exceção. Esse processo de busca tem três etapas:</span><span class="sxs-lookup"><span data-stu-id="55761-p143">In all the earlier functions in this series of tutorials, you queued commands to *write* to the Office document. Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed. But the code you added in the last step calls the `sheet.protection.protected` property, and this is a significant difference from the earlier functions you wrote, because the `sheet` object is only a proxy object that exists in your task pane's script. It doesn't know what the actual protection state of the document is, so its `protection.protected` property can't have a real value. It is necessary to first fetch the protection status from the document and use it set the value of `sheet.protection.protected`. Only then can `sheet.protection.protected` be called without causing an exception to be thrown. This fetching process has three steps:</span></span>

   1. <span data-ttu-id="55761-345">Coloque em fila um comando para carregar (ou seja, fetch) as propriedades que seu código precisa ler.</span><span class="sxs-lookup"><span data-stu-id="55761-345">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>

   2. <span data-ttu-id="55761-346">Chame o método `sync` do objeto de contexto para enviar o comando em fila para o documento para execução e retornar as informações solicitadas.</span><span class="sxs-lookup"><span data-stu-id="55761-346">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>

   3. <span data-ttu-id="55761-347">Como o método `sync` é assíncrono, certifique-se de que ele tenha sido concluído antes que o código chame as propriedades que foram buscadas.</span><span class="sxs-lookup"><span data-stu-id="55761-347">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="55761-348">Essas etapas devem ser concluídas sempre que seu código precisar *ler* informações do documento do Office.</span><span class="sxs-lookup"><span data-stu-id="55761-348">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="55761-p144">Na função `toggleProtection`, substitua `TODO2` pelo seguinte código. Observação:</span><span class="sxs-lookup"><span data-stu-id="55761-p144">In the `toggleProtection` function, replace `TODO2` with the following code. Note:</span></span>
   
   - <span data-ttu-id="55761-p145">Todos os objetos do Excel têm um método `load`. Especifique as propriedades do objeto que você deseja ler no parâmetro como uma cadeia de caracteres de nomes delimitados por vírgulas. Nesse caso, a propriedade que você precisa ler é uma subpropriedade de `protection`. Referencie a subpropriedade quase exatamente como você faria em qualquer lugar do seu código, mas usando uma barra (“/”) em vez de um ponto (".").</span><span class="sxs-lookup"><span data-stu-id="55761-p145">Every Excel object has a `load` method. You specify the properties of the object that you want to read in the parameter as a string of comma-delimited names. In this case, the property you need to read is a subproperty of the `protection` property. You reference the subproperty almost exactly as you would anywhere else in your code, with the exception that you use a forward slash ('/') character instead of a "." character.</span></span>

   - <span data-ttu-id="55761-355">Para garantir que a lógica de botão de alternância, `sheet.protection.protected`, não seja executada até após `sync` ser concluído e o `sheet.protection.protected` ser atribuída ao valor correto buscado no documento, ele será movido (na próxima etapa) para uma função `then` que não será executada até `sync` ser concluído.</span><span class="sxs-lookup"><span data-stu-id="55761-355">To ensure that the toggle logic, which reads `sheet.protection.protected`, does not run until after the `sync` is complete and the `sheet.protection.protected` has been assigned the correct value that is fetched from the document, it will be moved (in the next step) into a `then` function that won't run until the `sync` has completed.</span></span> 

    ```js
    sheet.load('protection/protected');
    return context.sync()
        .then(
            function() {
                // TODO3: Move the queued toggle logic here.
            }
        )
        // TODO4: Move the final call of `context.sync` here and ensure that it
        //        does not run until the toggle logic has been queued.
    ``` 

2. <span data-ttu-id="55761-p146">Você não pode ter duas instruções `return` no mesmo caminho de código sem ramificações, portanto, exclua a linha final `return context.sync();` no final de `Excel.run`. Você adicionará um novo `context.sync` final em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="55761-p146">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Excel.run`. You will add a new final `context.sync`, in a later step.</span></span>

3. <span data-ttu-id="55761-358">Recorte a estrutura `if ... else` na função `toggleProtection` e a cole no lugar de `TODO3`.</span><span class="sxs-lookup"><span data-stu-id="55761-358">Cut the `if ... else` structure in the `toggleProtection` function and paste it in place of `TODO3`.</span></span>

4. <span data-ttu-id="55761-p147">Substitua `TODO4` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="55761-p147">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="55761-361">Passar o método `sync` para uma função `then` garante que ele não seja executado até que `sheet.protection.unprotect()` ou `sheet.protection.protect()` seja enfileirado.</span><span class="sxs-lookup"><span data-stu-id="55761-361">Passing the `sync` method to a `then` function ensures that it does not run until either `sheet.protection.unprotect()` or `sheet.protection.protect()` has been queued.</span></span>

   - <span data-ttu-id="55761-362">O método `then` invoca qualquer função que é passada para ele e não é recomendável que `sync` seja chamado duas vezes, portanto, remova os “()” do fim de `context.sync`.</span><span class="sxs-lookup"><span data-stu-id="55761-362">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of `context.sync`.</span></span>

    ```js
    .then(context.sync);
    ```

   <span data-ttu-id="55761-363">Quando terminar, a função inteira deve se parecer com o seguinte:</span><span class="sxs-lookup"><span data-stu-id="55761-363">When you are done, the entire function should look like the following:</span></span>

    ```js
    function toggleProtection(args) {
        Excel.run(function (context) {            
          var sheet = context.workbook.worksheets.getActiveWorksheet();          
          sheet.load('protection/protected');

          return context.sync()
              .then(
                  function() {
                    if (sheet.protection.protected) {
                        sheet.protection.unprotect();
                    } else {
                        sheet.protection.protect();
                    }
                  }
              )
              .then(context.sync);
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```

### <a name="configure-the-script-loading-html-file"></a><span data-ttu-id="55761-364">Configure o arquivo HTML de carregamento de script</span><span class="sxs-lookup"><span data-stu-id="55761-364">Configure the script-loading HTML file</span></span>

<span data-ttu-id="55761-p148">Abra o arquivo /function-file/function-file.html. Esse é um arquivo HTML sem IU que é chamado quando o usuário pressiona o botão **Ativar/Desativar Proteção da Planilha**. O objetivo é carregar o método JavaScript que deve ser executado quando botão é pressionado. Esse arquivo não será alterado. Basta observar que a segunda marca `<script>` carrega o functionfile.js.</span><span class="sxs-lookup"><span data-stu-id="55761-p148">Open the /function-file/function-file.html file. This is a UI-less HTML file that is called when the user presses the **Toggle Worksheet Protection** button. Its purpose is to load the JavaScript method that should run when the button is pushed. You are not going to change this file. Simply note that the second `<script>` tag loads the functionfile.js.</span></span>

   > [!NOTE]
   > <span data-ttu-id="55761-p149">O arquivo function-file.html e o arquivo function-file.js carregado são executados em um processo do IE completamente separado de painel de tarefas do suplemento. Se o function-file.js foi transcompilado no mesmo arquivo bundle.js que o arquivo app.js, o suplemento precisará carregar duas cópias do arquivo bundle.js, o que anule o propósito do agrupamento. Além disso, o arquivo function-file.js não contém qualquer JavaScript incompatível com o Internet Explorer. Por esses dois motivos, esse suplemento não transcompila o function-file.js.</span><span class="sxs-lookup"><span data-stu-id="55761-p149">The function-file.html file and the function-file.js file that it loads run in an entirely separate IE process from the add-in's task pane. If the function-file.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling. In addition, the function-file.js file does not contain any JavaScript that is unsupported by IE. For these two reasons, this add-in does not transpile the function-file.js at all.</span></span> 

### <a name="test-the-add-in"></a><span data-ttu-id="55761-374">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="55761-374">Test the add-in</span></span>

1. <span data-ttu-id="55761-375">Feche todos os aplicativos do Office, incluindo o Excel.</span><span class="sxs-lookup"><span data-stu-id="55761-375">Close all Office applications, including Excel.</span></span> 

2. <span data-ttu-id="55761-p150">Para excluir o cache do Office, exclua o conteúdo da pasta de cache. Isso é necessário para limpar totalmente a versão anterior do suplemento do host.</span><span class="sxs-lookup"><span data-stu-id="55761-p150">Delete the Office cache by deleting the contents of the cache folder. This is necessary to completely clear the old version of the add-in from the host.</span></span> 

    - <span data-ttu-id="55761-378">No Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="55761-378">For Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

    - <span data-ttu-id="55761-379">No Mac: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="55761-379">For Mac: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

3. <span data-ttu-id="55761-p151">Se, por algum motivo, o servidor não estiver executando, em uma janela do Git Bash ou em um prompt do sistema habilitado para Node.JS, acesse a pasta **Iniciar** do projeto e execute o comando `npm start`. Não é necessário recriar o projeto, pois o único arquivo JavaScript que você alterou não faz parte do bundle.js interno.</span><span class="sxs-lookup"><span data-stu-id="55761-p151">If for any reason, your server is not running, then in a Git Bash window, or Node.JS-enabled system prompt, navigate to the **Start** folder of the project and run the command `npm start`. You do not need to rebuild the project because the only JavaScript file you changed is not part of the built bundle.js.</span></span>

4. <span data-ttu-id="55761-p152">Usando a nova versão do arquivo de manifesto alterado, repita o processo de sideloading usando um dos seguintes métodos. *Você deve substituir a cópia anterior do arquivo de manifesto.*</span><span class="sxs-lookup"><span data-stu-id="55761-p152">Using the new version of the changed manifest file, repeat the sideloading process by using one of the following methods. *You should overwrite the previous copy of the manifest file.*</span></span>

    - <span data-ttu-id="55761-384">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="55761-384">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>

    - <span data-ttu-id="55761-385">Excel Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="55761-385">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>

    - <span data-ttu-id="55761-386">iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="55761-386">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

5. <span data-ttu-id="55761-387">Abra qualquer planilha no Excel.</span><span class="sxs-lookup"><span data-stu-id="55761-387">Open any worksheet in Excel.</span></span>

6. <span data-ttu-id="55761-p153">Na Faixa de Opções, em **Página Inicial**, escolha **Ativar Proteger Planilha**. Observe que a maioria dos controles na Faixa de Opções está desabilitada (e visualmente esmaecida) conforme mostrado na captura de tela abaixo.</span><span class="sxs-lookup"><span data-stu-id="55761-p153">On the **Home** ribbon, choose **Toggle Worksheet Protection**. Note that most of the controls on the ribbon are disabled (and visually grayed-out) as seen in screenshot below.</span></span> 

7. <span data-ttu-id="55761-p154">Escolha uma célula como se quisesse alterar o conteúdo. Você receberá um erro informando que a planilha está protegida.</span><span class="sxs-lookup"><span data-stu-id="55761-p154">Choose a cell as you would if you wanted to change its content. You get an error telling you that the worksheet is protected.</span></span>

8. <span data-ttu-id="55761-392">Escolha **Ativar/Desativar Proteção da Planilha** novamente e os controles serão reabilitados e você poderá alterar os valores das células.</span><span class="sxs-lookup"><span data-stu-id="55761-392">Choose **Toggle Worksheet Protection** again, and the controls are reenabled, and you can change cell values again.</span></span>

    ![Tutorial do Excel - Faixa de Opções com a Proteção Ativada](../images/excel-tutorial-ribbon-with-protection-on.png)

## <a name="open-a-dialog"></a><span data-ttu-id="55761-394">Abrir uma caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="55761-394">Open a dialog</span></span>

<span data-ttu-id="55761-p155">Nesta etapa final do tutorial, você abre uma caixa de diálogo no suplemento, passa uma mensagem do processo de caixa de diálogo para o processo de painel de tarefas e fecha a caixa de diálogo. As caixas de diálogo do Suplemento do Office são *não modais*: o usuário pode continuar a interagir com o documento no aplicativo do Office do host e com a página host no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="55761-p155">In this final step of the tutorial, you'll open a dialog in your add-in, pass a message from the dialog process to the task pane process, and close the dialog. Office Add-in dialogs are *nonmodal*: a user can continue to interact with both the document in the host Office application and with the host page in the task pane.</span></span>

### <a name="create-the-dialog-page"></a><span data-ttu-id="55761-397">Crie a página da caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="55761-397">Create the dialog page</span></span>

1. <span data-ttu-id="55761-398">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="55761-398">Open the project in your code editor.</span></span>

2. <span data-ttu-id="55761-399">Crie um arquivo chamado popup.html na raiz do projeto (onde se encontra index.html).</span><span class="sxs-lookup"><span data-stu-id="55761-399">Create a file in the root of the project (where index.html is) called popup.html.</span></span>

3. <span data-ttu-id="55761-p156">Adicione a marcação a seguir em popup.html. Observação:</span><span class="sxs-lookup"><span data-stu-id="55761-p156">Add the following markup to popup.html. Note:</span></span>

   - <span data-ttu-id="55761-402">a página tem um `<input>` em que o usuário insere o nome dele e um botão que envia o nome para a página no painel de tarefas onde ele será exibido.</span><span class="sxs-lookup"><span data-stu-id="55761-402">The page has a `<input>` where the user will enter their name and a button that will send the name to the page in the task pane where it will be displayed.</span></span>

   - <span data-ttu-id="55761-403">A marcação carrega um script chamado popup.js que você criará em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="55761-403">The markup loads a script called popup.js that you will create in a later step.</span></span>

   - <span data-ttu-id="55761-404">Ela também carrega uma biblioteca Office.JS e jQuery porque elas serão usadas em popup.js.</span><span class="sxs-lookup"><span data-stu-id="55761-404">It also loads the Office.JS library and jQuery because they will be used in popup.js.</span></span>

    ```html
    <!DOCTYPE html>
    <html>
        <head lang="en">
            <title>Dialog for My Office Add-in</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">

            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.min.css" />
            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.components.css" />
            <link rel="stylesheet" href="app.css" />

            <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
            <script type="text/javascript" src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.2.1.min.js"></script>
            <script type="text/javascript" src="popup.js"></script>

        </head>
        <body style="display:flex;flex-direction:column;align-items:center;justify-content:center">
            <div class="padding">
                <p class="ms-font-xl">ENTER YOUR NAME</p>
            </div>
            <div class="padding">
                <input id="name-box" type="text"/>
            </div>
            <div class="padding">
                <button id="ok-button" class="ms-Button">OK</button>
            </div>
        </body>
    </html>
    ```

4. <span data-ttu-id="55761-405">Crie um arquivo chamado popup.js na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="55761-405">Create a file in the root of the project called popup.js.</span></span>

5. <span data-ttu-id="55761-406">Adicione o código a seguir a popup.js.</span><span class="sxs-lookup"><span data-stu-id="55761-406">Add the following code to popup.js.</span></span> <span data-ttu-id="55761-407">Observe o seguinte a respeito deste código:</span><span class="sxs-lookup"><span data-stu-id="55761-407">Note the following about this code:</span></span>

   - <span data-ttu-id="55761-408">*Todas as páginas que chamam APIs na biblioteca Office.JS devem primeiro garantir que a biblioteca tenha sido totalmente inicializada.*</span><span class="sxs-lookup"><span data-stu-id="55761-408">*Every page that calls APIs in the Office.JS library must first ensure that the library is fully initialized.*</span></span> <span data-ttu-id="55761-409">A melhor maneira de fazer isso é chamando o método `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="55761-409">The best way to do that is to call the `Office.onReady()` method.</span></span> <span data-ttu-id="55761-410">Se o suplemento possuir as próprias tarefas de inicialização, o código deverá ser colocado em um método `then()` encadeado à chamada de `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="55761-410">If your add-in has its own initialization tasks, the code should go in a `then()` method that is chained to the call of `Office.onReady()`.</span></span> <span data-ttu-id="55761-411">Para um exemplo, veja o arquivo app.js na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="55761-411">For an example, see the app.js file in the project root.</span></span> <span data-ttu-id="55761-412">A chamada de `Office.onReady()` deve ser executada antes de qualquer chamada para Office.JS; por isso, a tarefa se encontra em um arquivo de script que é carregado pela página, como neste caso.</span><span class="sxs-lookup"><span data-stu-id="55761-412">The call of `Office.onReady()` must run before any calls to Office.JS; hence the assignment is in a script file that is loaded by the page, as it is in this case.</span></span>
   - <span data-ttu-id="55761-413">A função `ready` do jQuery é chamada dentro do método `then()`.</span><span class="sxs-lookup"><span data-stu-id="55761-413">The jQuery `ready` function is called inside the `then()` method.</span></span> <span data-ttu-id="55761-414">Na maioria dos casos, o carregamento, a inicialização ou o código de bootstrap de outras bibliotecas JavaScript devem ficar dentro do método `then()` encadeado à chamada de `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="55761-414">In most cases, the loading, initializing, or bootstrapping code of other JavaScript libraries should be inside the `then()` method that is chained to the call of `Office.onReady()`.</span></span>

    ```js
    (function () {
    "use strict";

        Office.onReady()
            .then(function() {
                $(document).ready(function () {  

                    // TODO1: Assign handler to the OK button.

                });
            });

        // TODO2: Create the OK button handler

    }());
    ```

6. <span data-ttu-id="55761-p160">Substitua `TODO1` pelo código a seguir. Você criará a função `sendStringToParentPage` na próxima etapa.</span><span class="sxs-lookup"><span data-stu-id="55761-p160">Replace `TODO1` with the following code. You'll create the `sendStringToParentPage` function in the next step.</span></span>

    ```js
    $('#ok-button').click(sendStringToParentPage);
    ```

7. <span data-ttu-id="55761-p161">Substitua `TODO2` pelo código a seguir. O método `messageParent` passa seu parâmetro para a página pai, neste caso, a página no painel de tarefas. O parâmetro pode ser um booliano ou uma cadeia de caracteres, que inclui tudo o que pode ser serializado como uma cadeia de caracteres, como XML ou JSON.</span><span class="sxs-lookup"><span data-stu-id="55761-p161">Replace `TODO2` with the following code. The `messageParent` method passes its parameter to the parent page, in this case, the page in the task pane. The parameter can be a boolean or a string, which includes anything that can be serialized as a string, such as XML or JSON.</span></span>

    ```js
    function sendStringToParentPage() {
        var userName = $('#name-box').val();
        Office.context.ui.messageParent(userName);
    }
    ```

8. <span data-ttu-id="55761-420">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="55761-420">Save the file.</span></span>

   > [!NOTE]
   > <span data-ttu-id="55761-p162">O arquivo popup.html e o arquivo popup.js carregado são executados em um processo do Internet Explorer completamente separado de painel de tarefas do suplemento. Se o popup.js foi transcompilado no mesmo arquivo bundle.js que o arquivo app.js, o suplemento precisará carregar duas cópias do arquivo bundle.js, o que anule o propósito do agrupamento. Além disso, o arquivo popup.js não contém qualquer JavaScript incompatível com o Internet Explorer. Por esses dois motivos, esse suplemento não transcompila o popup.js.</span><span class="sxs-lookup"><span data-stu-id="55761-p162">The popup.html file, and the popup.js file that it loads, run in an entirely separate Internet Explorer process from the add-in's task pane. If the popup.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling. In addition, the popup.js file does not contain any JavaScript that is unsupported by IE. For these two reasons, this add-in does not transpile the popup.js file at all.</span></span>

### <a name="open-the-dialog-from-the-task-pane"></a><span data-ttu-id="55761-425">Abra a caixa de diálogo do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="55761-425">Open the dialog from the task pane</span></span>

1. <span data-ttu-id="55761-426">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="55761-426">Open the file index.html.</span></span>

2. <span data-ttu-id="55761-427">Abaixo do `div` que contém o botão `freeze-header`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="55761-427">Below the `div` that contains the `freeze-header` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="open-dialog">Open Dialog</button>
    </div>
    ```

3. <span data-ttu-id="55761-p163">A caixa de diálogo solicitará que o usuário insira um nome e passará o nome de usuário para o painel de tarefas. O painel de tarefas o exibirá em um rótulo. Imediatamente abaixo do `div` que você adicionou, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="55761-p163">The dialog will prompt the user to enter a name and pass the user's name to the task pane. The task pane will display it in a label. Immediately below the `div` that you just added, add the following markup:</span></span>

    ```html
    <div class="padding">
        <label id="user-name"></label>
    </div>
    ```

4. <span data-ttu-id="55761-431">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="55761-431">Open the app.js file.</span></span>

5. <span data-ttu-id="55761-p164">Abaixo da linha que atribui um identificador de clique ao botão `freeze-header`, adicione o seguinte código. Você criará o método `openDialog` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="55761-p164">Below the line that assigns a click handler to the `freeze-header` button, add the following code. You'll create the `openDialog` method in a later step.</span></span>

    ```js
    $('#open-dialog').click(openDialog);
    ```

6. <span data-ttu-id="55761-p165">Abaixo da função `freezeHeader`, adicione a declaração seguinte. Essa variável é usada para armazenar um objeto no contexto de execução da página pai que atua como um intermediador no contexto de execução da página da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="55761-p165">Below the `freezeHeader` function add the following declaration. This variable is used to hold an object in the parent page's execution context that acts as an intermediator to the dialog page's execution context.</span></span>

    ```js
    var dialog = null;
    ```

7. <span data-ttu-id="55761-p166">Abaixo da declaração de `dialog`, adicione a função a seguir. É importante observar o que esse código *não* contém: não há nenhuma chamada de `Excel.run`. Isso ocorre porque a API para abrir uma caixa de diálogo é compartilhada com todos os hosts do Office, portanto, ela faz parte da API de Office JavaScript Common, não da API específica do Excel.</span><span class="sxs-lookup"><span data-stu-id="55761-p166">Below the declaration of `dialog`, add the following function. The important thing to notice about this code is what is *not* there: there is no call of `Excel.run`. This is because the API to open a dialog is shared among all Office hosts, so it is part of the Office JavaScript Common API, not the Excel-specific API.</span></span>

    ```js
    function openDialog() {
        // TODO1: Call the Office Common API that opens a dialog
    }
    ```

8. <span data-ttu-id="55761-439">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="55761-439">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="55761-440">Observação:</span><span class="sxs-lookup"><span data-stu-id="55761-440">Note:</span></span>

   - <span data-ttu-id="55761-441">O método`displayDialogAsync` abre uma caixa de diálogo no centro da tela.</span><span class="sxs-lookup"><span data-stu-id="55761-441">The `displayDialogAsync` method opens a dialog in the center of the screen.</span></span>

   - <span data-ttu-id="55761-442">O primeiro parâmetro é a URL da página a ser aberta.</span><span class="sxs-lookup"><span data-stu-id="55761-442">The first parameter is the URL of the page to open.</span></span>

   - <span data-ttu-id="55761-p168">O segundo parâmetro passa opções. `height` e `width` são porcentagens do tamanho da janela do aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="55761-p168">The second parameter passes options. `height` and `width` are percentages of the size of the Office application's window.</span></span>

    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},

        // TODO2: Add callback parameter.
    );
    ```

### <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a><span data-ttu-id="55761-445">Processar a mensagem da caixa de diálogo e depois fechá-la</span><span class="sxs-lookup"><span data-stu-id="55761-445">Process the message from the dialog and close the dialog</span></span>

1. <span data-ttu-id="55761-p169">Continue no arquivo app.js e substitua `TODO2` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="55761-p169">Continue in the app.js file, and replace `TODO2` with the following code. Note:</span></span>

   - <span data-ttu-id="55761-448">O retorno de chamada é executado logo após a caixa de diálogo ser aberta com êxito e antes de o usuário executar qualquer ação nela.</span><span class="sxs-lookup"><span data-stu-id="55761-448">The callback is executed immediately after the dialog successfully opens and before the user has taken any action in the dialog.</span></span>

   - <span data-ttu-id="55761-449">O `result.value` é o objeto que funciona como um tipo de intermediário entre contextos execução das páginas de pai e de caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="55761-449">The `result.value` is the object that acts as a kind of middleman between the execution contexts of the parent and dialog pages.</span></span>

   - <span data-ttu-id="55761-p170">A função `processMessage` será criada em uma etapa posterior. Esse identificador processará os valores que sejam enviados da página da caixa de diálogo com chamadas da função `messageParent`.</span><span class="sxs-lookup"><span data-stu-id="55761-p170">The `processMessage` function will be created in a later step. This handler will process any values that are sent from the dialog page with calls of the `messageParent` function.</span></span>

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. <span data-ttu-id="55761-452">Abaixo da função `openDialog`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="55761-452">Below the `openDialog` function, add the following function.</span></span>

    ```js
    function processMessage(arg) {
        $('#user-name').text(arg.message);
        dialog.close();
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="55761-453">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="55761-453">Test the add-in</span></span>

1. <span data-ttu-id="55761-454">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite **Ctrl + C** duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="55761-454">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter **Ctrl+C** twice to stop the running web server.</span></span> <span data-ttu-id="55761-455">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="55761-455">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="55761-p172">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor. Para fazer isso, interrompa o processo do servidor para obter uma solicitação para inserir o comando de compilação. Após a compilação, reinicie o servidor. As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="55761-p172">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command. After the build, you restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="55761-460">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte no Internet Explorer (que é usada em segundo plano pelo Excel para executar os suplementos do Excel).</span><span class="sxs-lookup"><span data-stu-id="55761-460">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="55761-461">Execute o comando `npm start` para iniciar um servidor Web em um host local.</span><span class="sxs-lookup"><span data-stu-id="55761-461">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="55761-462">Feche o painel de tarefas para recarregá-lo e, no menu **Início**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="55761-462">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="55761-463">Escolha o botão **Abrir Caixa de Diálogo** no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="55761-463">Choose the **Open Dialog** button in the task pane.</span></span>

6. <span data-ttu-id="55761-464">Quando a caixa de diálogo estiver aberta, arraste-a e redimensione-a.</span><span class="sxs-lookup"><span data-stu-id="55761-464">While the dialog is open, drag it and resize it.</span></span> <span data-ttu-id="55761-465">Observe que você pode interagir com a planilha e pressionar outros botões no painel de tarefas. No entanto, não é possível iniciar uma segunda caixa de diálogo na mesma página do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="55761-465">Note that you can interact with the worksheet and press other buttons on the task pane, but you cannot launch a second dialog from the same task pane page.</span></span>

7. <span data-ttu-id="55761-p174">Na caixa de diálogo, digite um nome e escolha **OK**. O nome aparecerá no painel de tarefas e a caixa de diálogo será fechada.</span><span class="sxs-lookup"><span data-stu-id="55761-p174">In the dialog, enter a name and choose **OK**. The name appears on the task pane and the dialog closes.</span></span>

8. <span data-ttu-id="55761-p175">Opcionalmente, comente a linha `dialog.close();` na função `processMessage`. Em seguida, repita as etapas desta seção. A caixa de diálogo permanece aberta e você pode alterar o nome. É possível fechá-la manualmente pressionando o botão **X** no canto superior direito.</span><span class="sxs-lookup"><span data-stu-id="55761-p175">Optionally, comment out the line `dialog.close();` in the `processMessage` function. Then repeat the steps of this section. The dialog stays open and you can change the name. You can close it manually by pressing the **X** button in the upper right corner.</span></span>

    ![Tutorial do Excel - Caixa de diálogo](../images/excel-tutorial-dialog-open.png)

## <a name="next-steps"></a><span data-ttu-id="55761-473">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="55761-473">Next steps</span></span>

<span data-ttu-id="55761-474">Neste tutorial você criou um suplemento do Excel que interage com tabelas, gráficos, planilhas e caixas de diálogo em uma pasta de trabalho do Excel.</span><span class="sxs-lookup"><span data-stu-id="55761-474">In this tutorial, you've created an Excel task pane add-in that interacts with tables, charts, worksheets, and dialogs in an Excel workbook.</span></span> <span data-ttu-id="55761-475">Para saber mais sobre o desenvolvimento de suplementos do Excel, continue no seguinte artigo:</span><span class="sxs-lookup"><span data-stu-id="55761-475">To learn more about building Excel add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="55761-476">Visão geral dos suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="55761-476">Excel add-ins overview</span></span>](../excel/excel-add-ins-overview.md)
