---
title: Tutorial de suplemento do Excel
description: Neste tutorial, você criará um suplemento do Excel que cria, preenche, filtra e classifica uma tabela, cria um gráfico, congela um cabeçalho de tabela, protege uma planilha e abre uma caixa de diálogo
ms.date: 12/31/2018
ms.topic: tutorial
ms.openlocfilehash: fe4350f5f3fdbe34250c1739c7651a1dde1e28ef
ms.sourcegitcommit: 3007bf57515b0811ff98a7e1518ecc6fc9462276
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/04/2019
ms.locfileid: "27724934"
---
# <a name="tutorial-create-an-excel-task-pane-add-in"></a><span data-ttu-id="3746f-103">Tutorial: criar um suplemento do painel de tarefas no Excel</span><span class="sxs-lookup"><span data-stu-id="3746f-103">Tutorial: Create an Excel task pane add-in</span></span>

<span data-ttu-id="3746f-104">Neste tutorial: você criará um suplemento do painel de tarefas no Excel</span><span class="sxs-lookup"><span data-stu-id="3746f-104">In this tutorial, you'll create an Excel task pane add-in that:</span></span>

> [!div class="checklist"]
> * <span data-ttu-id="3746f-105">Cria uma tabela</span><span class="sxs-lookup"><span data-stu-id="3746f-105">Creates a new table.</span></span>
> * <span data-ttu-id="3746f-106">Filtra e classifica uma tabela</span><span class="sxs-lookup"><span data-stu-id="3746f-106">Filters and sorts a table</span></span>
> * <span data-ttu-id="3746f-107">Cria um gráfico</span><span class="sxs-lookup"><span data-stu-id="3746f-107">Creates a new chart.</span></span>
> * <span data-ttu-id="3746f-108">Congela um cabeçalho de tabela</span><span class="sxs-lookup"><span data-stu-id="3746f-108">Freezes a table header</span></span>
> * <span data-ttu-id="3746f-109">Protege uma planilha</span><span class="sxs-lookup"><span data-stu-id="3746f-109">Protects a worksheet</span></span>
> * <span data-ttu-id="3746f-110">Abre uma caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="3746f-110">Opens a dialog</span></span>

## <a name="prerequisites"></a><span data-ttu-id="3746f-111">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="3746f-111">Prerequisites</span></span>

<span data-ttu-id="3746f-112">Para usar este tutorial, você precisa instalar o seguinte.</span><span class="sxs-lookup"><span data-stu-id="3746f-112">To use this tutorial, you need to have the following installed.</span></span> 

- <span data-ttu-id="3746f-113">Excel 2016, versão 1711 (build 8730.1000 do Clique para Executar) ou posterior.</span><span class="sxs-lookup"><span data-stu-id="3746f-113">Excel 2016, version 1711 (Build 8730.1000 Click-to-Run) or later.</span></span> <span data-ttu-id="3746f-114">Talvez você precise ser um participante do programa Office Insider para ter essa versão.</span><span class="sxs-lookup"><span data-stu-id="3746f-114">You might need to be an Office Insider to get this version.</span></span> <span data-ttu-id="3746f-115">Para saber mais, confira a página [Seja um Office Insider](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="3746f-115">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span>

- [<span data-ttu-id="3746f-116">Nó</span><span class="sxs-lookup"><span data-stu-id="3746f-116">Node</span></span>](https://nodejs.org/en/) 

- <span data-ttu-id="3746f-117">[Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)</span><span class="sxs-lookup"><span data-stu-id="3746f-117">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

## <a name="create-your-add-in-project"></a><span data-ttu-id="3746f-118">Criar seu projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="3746f-118">Create your add-in project</span></span>

<span data-ttu-id="3746f-119">Conclua as etapas a seguir para criar o projeto de suplemento do Excel que você vai usar como base para este tutorial.</span><span class="sxs-lookup"><span data-stu-id="3746f-119">Complete the following steps to create the Excel add-in project that you'll use as the basis for this tutorial.</span></span>

1. <span data-ttu-id="3746f-120">Clone o repositório do GitHub com o [Tutorial de suplemento do Excel](https://github.com/OfficeDev/Excel-Add-in-Tutorial).</span><span class="sxs-lookup"><span data-stu-id="3746f-120">Clone the GitHub repository [Excel Add-in Tutorial](https://github.com/OfficeDev/Excel-Add-in-Tutorial).</span></span>

2. <span data-ttu-id="3746f-121">Abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue para a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="3746f-121">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

3. <span data-ttu-id="3746f-122">Execute o comando `npm install` para instalar as ferramentas e bibliotecas listadas no arquivo package.json.</span><span class="sxs-lookup"><span data-stu-id="3746f-122">Run the command `npm install` to install the tools and libraries listed in the package.json file.</span></span> 

4. <span data-ttu-id="3746f-123">Execute as etapas em [Adicionar certificados autoassinados como certificado raiz confiáveis](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) para confiar no certificado do sistema operacional do seu computador de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="3746f-123">Carry out the steps in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to trust the certificate for your development computer's operating system.</span></span>

## <a name="create-a-table"></a><span data-ttu-id="3746f-124">Criar uma tabela</span><span class="sxs-lookup"><span data-stu-id="3746f-124">Create a table</span></span>

<span data-ttu-id="3746f-125">Nesta etapa do tutorial, você testará no programa se o suplemento é compatível com a versão atual do Excel do usuário, adicionará uma tabela a uma planilha, depois preencherá e formatará a tabela com os dados.</span><span class="sxs-lookup"><span data-stu-id="3746f-125">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Excel, add a table to a worksheet, populate the table with data, and format it.</span></span>

### <a name="code-the-add-in"></a><span data-ttu-id="3746f-126">Codificação do suplemento</span><span class="sxs-lookup"><span data-stu-id="3746f-126">Code the add-in</span></span>

1. <span data-ttu-id="3746f-127">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="3746f-127">Open the project in your code editor.</span></span>

2. <span data-ttu-id="3746f-128">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="3746f-128">Open the file index.html.</span></span>

3. <span data-ttu-id="3746f-129">Substitua `TODO1` pela marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="3746f-129">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="create-table">Create Table</button>
    ```

4. <span data-ttu-id="3746f-130">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="3746f-130">Open the app.js file.</span></span>

5. <span data-ttu-id="3746f-131">Substitua o `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="3746f-131">Replace the `TODO1` with the following code.</span></span> <span data-ttu-id="3746f-132">O código determina se a versão do Excel do usuário proporciona suporte a uma versão do Excel.js que inclua as APIs com esta série de tutoriais.</span><span class="sxs-lookup"><span data-stu-id="3746f-132">This code determines whether the user's version of Excel supports a version of Excel.js that includes all the APIs that this series of tutorials will use.</span></span> <span data-ttu-id="3746f-133">Em um suplemento de produção, use o corpo do bloco condicional para ocultar ou desabilitar a interface do usuário que chame a APIs sem suporte.</span><span class="sxs-lookup"><span data-stu-id="3746f-133">In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs.</span></span> <span data-ttu-id="3746f-134">Dessa forma, permitirá que o usuário ainda use as partes do suplemento às quais a versão do Excel dá suporte.</span><span class="sxs-lookup"><span data-stu-id="3746f-134">This will enable the user to still make use of the parts of the add-in that are supported by their version of Excel.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }
    ```

6. <span data-ttu-id="3746f-135">Substitua o `TODO2` pelo código a seguir:</span><span class="sxs-lookup"><span data-stu-id="3746f-135">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#create-table').click(createTable);
    ```

7. <span data-ttu-id="3746f-136">Substitua o `TODO3` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="3746f-136">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="3746f-137">Observação:</span><span class="sxs-lookup"><span data-stu-id="3746f-137">Note:</span></span>

   - <span data-ttu-id="3746f-138">A lógica de negócios de Excel.js será adicionada à função que passar por `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="3746f-138">Your Excel.js business logic will be added to the function that is passed to `Excel.run`.</span></span> <span data-ttu-id="3746f-139">Essa lógica não é executada imediatamente.</span><span class="sxs-lookup"><span data-stu-id="3746f-139">This logic does not execute immediately.</span></span> <span data-ttu-id="3746f-140">Em vez disso, ela é adicionada à fila de comandos pendentes.</span><span class="sxs-lookup"><span data-stu-id="3746f-140">Instead, it is added to a queue of pending commands.</span></span>

   - <span data-ttu-id="3746f-141">O método `context.sync` envia todos os comandos da fila para execução no Excel.</span><span class="sxs-lookup"><span data-stu-id="3746f-141">The `context.sync` method sends all queued commands to Excel for execution.</span></span>

   - <span data-ttu-id="3746f-142">`Excel.run` é seguido por um bloco `catch`.</span><span class="sxs-lookup"><span data-stu-id="3746f-142">The `Excel.run` is followed by a `catch` block.</span></span> <span data-ttu-id="3746f-143">Essa é uma prática recomendada que você sempre deve seguir.</span><span class="sxs-lookup"><span data-stu-id="3746f-143">This is a best practice that you should always follow.</span></span> 

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

8. <span data-ttu-id="3746f-p106">Substitua `TODO4` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="3746f-p106">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="3746f-146">O código cria uma tabela usando o método `add` de conjunto de tabela da planilha, que sempre existe mesmo que ela esteja vazia.</span><span class="sxs-lookup"><span data-stu-id="3746f-146">The code creates a table by using `add` method of a worksheet's table collection, which always exists even if it is empty.</span></span> <span data-ttu-id="3746f-147">Essa é a maneira padrão de criar objetos no Excel.js.</span><span class="sxs-lookup"><span data-stu-id="3746f-147">This is the standard way that Excel.js objects are created.</span></span> <span data-ttu-id="3746f-148">Não há nenhuma API do construtor de classe e você nunca usará um operador `new` para criar um objeto do Excel.</span><span class="sxs-lookup"><span data-stu-id="3746f-148">There are no class constructor APIs, and you never use a `new` operator to create an Excel object.</span></span> <span data-ttu-id="3746f-149">Em vez disso, adicione a um objeto de conjunto pai.</span><span class="sxs-lookup"><span data-stu-id="3746f-149">Instead, you add to a parent collection object.</span></span>

   - <span data-ttu-id="3746f-150">O primeiro parâmetro do método `add`é o intervalo apenas da linha superior da tabela, não o intervalo inteiro que a tabela por fim usará.</span><span class="sxs-lookup"><span data-stu-id="3746f-150">The first parameter of the `add` method is the range of only the top row of the table, not the entire range the table will ultimately use.</span></span> <span data-ttu-id="3746f-151">Isso ocorre porque, quando o suplemento preenche as linhas de dados (na próxima etapa), ele adicionará novas linhas à tabela, em vez de gravar os valores nas células das linhas existentes.</span><span class="sxs-lookup"><span data-stu-id="3746f-151">This is because when the add-in populates the data rows (in the next step), it will add new rows to the table instead of writing values to the cells of existing rows.</span></span> <span data-ttu-id="3746f-152">Esse é um padrão mais comum, porque o número de linhas em uma tabela geralmente não é conhecido quando a tabela é criada.</span><span class="sxs-lookup"><span data-stu-id="3746f-152">This is a more common pattern because the number of rows that a table will have is often not known when the table is created.</span></span>

   - <span data-ttu-id="3746f-153">Os nomes de tabelas devem ser exclusivos pela pasta de trabalho inteira, não só na planilha.</span><span class="sxs-lookup"><span data-stu-id="3746f-153">Table names must be unique across the entire workbook, not just the worksheet.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ```

9. <span data-ttu-id="3746f-p109">Substitua `TODO5` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="3746f-p109">Replace `TODO5` with the following code. Note:</span></span>

   - <span data-ttu-id="3746f-156">Os valores das células de um intervalo são definidos em uma matriz de matrizes.</span><span class="sxs-lookup"><span data-stu-id="3746f-156">The cell values of a range are set with an array of arrays.</span></span>

   - <span data-ttu-id="3746f-157">Novas linhas são criadas em uma tabela ao chamar o método `add` do conjunto de linhas da tabela.</span><span class="sxs-lookup"><span data-stu-id="3746f-157">New rows are created in a table by calling the `add` method of the table's row collection.</span></span> <span data-ttu-id="3746f-158">Você pode adicionar várias linhas em uma única chamada de `add` ao incluir várias matrizes de valores de células na matriz pai que é passada como segundo parâmetro.</span><span class="sxs-lookup"><span data-stu-id="3746f-158">You can add multiple rows in a single call of `add` by including multiple cell value arrays in the parent array that is passed as the second parameter.</span></span>

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

10. <span data-ttu-id="3746f-p111">Substitua `TODO6` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="3746f-p111">Replace `TODO6` with the following code. Note:</span></span>

   - <span data-ttu-id="3746f-161">O código recebe uma referência para a coluna **quantidade** ao passar o índice com base em zero para o método `getItemAt` do conjunto de colunas da tabela.</span><span class="sxs-lookup"><span data-stu-id="3746f-161">The code gets a reference to the **Amount** column by passing its zero-based index to the `getItemAt` method of the table's column collection.</span></span>

     > [!NOTE]
     > <span data-ttu-id="3746f-162">Os objetos do conjunto Excel.js, como `TableCollection`, `WorksheetCollection`, e `TableColumnCollection`, têm a propriedade `items` que é como uma matriz dos tipos de objetos filhos, como `Table` ou `Worksheet` ou `TableColumn`; mas um objeto `*Collection` não é uma matriz.</span><span class="sxs-lookup"><span data-stu-id="3746f-162">Excel.js collection objects, such as `TableCollection`, `WorksheetCollection`, and `TableColumnCollection` have an `items` property that is an array of the child object types, such as `Table` or `Worksheet` or `TableColumn`; but a `*Collection` object is not itself an array.</span></span>

   - <span data-ttu-id="3746f-163">O código formata o intervalo da coluna **quantidade** como Euros com um segundo decimal.</span><span class="sxs-lookup"><span data-stu-id="3746f-163">The code then formats the range of the **Amount** column as Euros to the second decimal.</span></span> 

   - <span data-ttu-id="3746f-164">Por fim, isso garante que a largura das colunas e a altura das linhas sejam grandes o suficiente para o maior (ou o mais alto) item de dados.</span><span class="sxs-lookup"><span data-stu-id="3746f-164">Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item.</span></span> <span data-ttu-id="3746f-165">Observe que o código deve receber os objetos `Range` a formatar.</span><span class="sxs-lookup"><span data-stu-id="3746f-165">Notice that the code must get `Range` objects to format.</span></span> <span data-ttu-id="3746f-166">Os objetos `TableColumn` e `TableRow` não têm propriedades de formato.</span><span class="sxs-lookup"><span data-stu-id="3746f-166">`TableColumn` and `TableRow` objects do not have format properties.</span></span>

        ```js
        expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
        expensesTable.getRange().format.autofitColumns();
        expensesTable.getRange().format.autofitRows();
        ```

### <a name="test-the-add-in"></a><span data-ttu-id="3746f-167">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="3746f-167">Test the add-in</span></span>

1. <span data-ttu-id="3746f-168">Abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue para a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="3746f-168">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

2. <span data-ttu-id="3746f-169">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte no Internet Explorer (que é usada em segundo plano pelo Excel para executar os suplementos do Excel).</span><span class="sxs-lookup"><span data-stu-id="3746f-169">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="3746f-170">Execute o comando `npm start` para iniciar um servidor Web em um host local.</span><span class="sxs-lookup"><span data-stu-id="3746f-170">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="3746f-171">Realize o sideload do suplemento usando um dos métodos a seguir:</span><span class="sxs-lookup"><span data-stu-id="3746f-171">Sideload the add-in by using one of the following methods:</span></span>

    - <span data-ttu-id="3746f-172">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="3746f-172">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>

    - <span data-ttu-id="3746f-173">Excel Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="3746f-173">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>

    - <span data-ttu-id="3746f-174">iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="3746f-174">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

5. <span data-ttu-id="3746f-175">No menu **Página Inicial**, escolha **Mostrar Painel de Tarefas**.</span><span class="sxs-lookup"><span data-stu-id="3746f-175">On the **Home** menu, choose **Show Taskpane**.</span></span>

6. <span data-ttu-id="3746f-176">No painel de tarefas, escolha **Criar Tabela**.</span><span class="sxs-lookup"><span data-stu-id="3746f-176">In the task pane, choose **Create Table**.</span></span>

    ![Tutorial do Excel: Criar tabela](../images/excel-tutorial-create-table.png)

## <a name="filter-and-sort-a-table"></a><span data-ttu-id="3746f-178">Filtrar e classificar uma tabela</span><span class="sxs-lookup"><span data-stu-id="3746f-178">Filter and sort a table</span></span>

<span data-ttu-id="3746f-179">Nesta etapa do tutorial, você vai filtrar e classificar a tabela que criou anteriormente.</span><span class="sxs-lookup"><span data-stu-id="3746f-179">In this step of the tutorial, you'll filter and sort the table that you created previously.</span></span>

### <a name="filter-the-table"></a><span data-ttu-id="3746f-180">Filtrar a tabela</span><span class="sxs-lookup"><span data-stu-id="3746f-180">Filter the table</span></span>

1. <span data-ttu-id="3746f-181">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="3746f-181">Open the project in your code editor.</span></span>

2. <span data-ttu-id="3746f-182">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="3746f-182">Open the file index.html.</span></span>

3. <span data-ttu-id="3746f-183">Abaixo do `div`, que contém o botão `create-table`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="3746f-183">Just below the `div` that contains the `create-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="filter-table">Filter Table</button>
    </div>
    ```

4. <span data-ttu-id="3746f-184">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="3746f-184">Open the app.js file.</span></span>

5. <span data-ttu-id="3746f-185">Logo abaixo da linha que atribui um identificador de clique ao botão `create-table`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="3746f-185">Just below the line that assigns a click handler to the `create-table` button, add the following code:</span></span>

    ```js
    $('#filter-table').click(filterTable);
    ```

6. <span data-ttu-id="3746f-186">Logo abaixo da função `createTable`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="3746f-186">Just below the `createTable` function, add the following function:</span></span>

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

7. <span data-ttu-id="3746f-p113">Substitua `TODO1` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="3746f-p113">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="3746f-189">O código primeiro faz referência à coluna que precisa de filtragem ao passar o nome da coluna para o método `getItem`, em vez de passar o índice para o método `getItemAt` como o método `createTable` faz.</span><span class="sxs-lookup"><span data-stu-id="3746f-189">The code first gets a reference to the column that needs filtering by passing the column name to the `getItem` method, instead of passing its index to the `getItemAt` method as the `createTable` method does.</span></span> <span data-ttu-id="3746f-190">Como os usuários podem mover as colunas da tabela, a coluna de um determinado índice pode mudar depois da criação da tabela.</span><span class="sxs-lookup"><span data-stu-id="3746f-190">Since users can move table columns, the column at a given index might change after the table is created.</span></span> <span data-ttu-id="3746f-191">Portanto, é mais seguro usar o nome da coluna como referência dela.</span><span class="sxs-lookup"><span data-stu-id="3746f-191">Hence, it is safer to use the column name to get a reference to the column.</span></span> <span data-ttu-id="3746f-192">Usamos de forma segura `getItemAt` em um tutorial anterior porque usamos o mesmo método que cria a tabela. Assim não existe a chance de um usuário mover a coluna.</span><span class="sxs-lookup"><span data-stu-id="3746f-192">We used `getItemAt` safely in the preceding tutorial, because we used it in the very same method that creates the table, so there is no chance that a user has moved the column.</span></span>

   - <span data-ttu-id="3746f-193">O método `applyValuesFilter` é um dos vários métodos de filtragem do objeto `Filter`.</span><span class="sxs-lookup"><span data-stu-id="3746f-193">The `applyValuesFilter` method is one of several filtering methods on the `Filter` object.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);
    ``` 

### <a name="sort-the-table"></a><span data-ttu-id="3746f-194">Classificar a tabela</span><span class="sxs-lookup"><span data-stu-id="3746f-194">Sort the table</span></span>

1. <span data-ttu-id="3746f-195">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="3746f-195">Open the file index.html.</span></span>

2. <span data-ttu-id="3746f-196">Abaixo do `div` que contém o botão `filter-table`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="3746f-196">Below the `div` that contains the `filter-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="sort-table">Sort Table</button>
    </div>
    ```

3. <span data-ttu-id="3746f-197">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="3746f-197">Open the app.js file.</span></span>

4. <span data-ttu-id="3746f-198">Abaixo da linha que atribui um identificador de clique ao botão `filter-table`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="3746f-198">Below the line that assigns a click handler to the `filter-table` button, add the following code:</span></span>

    ```js
    $('#sort-table').click(sortTable);
    ```

5. <span data-ttu-id="3746f-199">Abaixo da função `filterTable`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="3746f-199">Below the `filterTable` function add the following function.</span></span>

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

6. <span data-ttu-id="3746f-p115">Substitua `TODO1` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="3746f-p115">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="3746f-202">O código cria uma matriz de objetos `SortField` que tem apenas um membro, já que o suplemento só classifica a coluna Comerciante.</span><span class="sxs-lookup"><span data-stu-id="3746f-202">The code creates an array of `SortField` objects which has just one member since the add-in only sorts on the Merchant column.</span></span>

   - <span data-ttu-id="3746f-203">A propriedade `key` de um objeto `SortField` é o índice com base em zero da coluna a classificar.</span><span class="sxs-lookup"><span data-stu-id="3746f-203">The `key` property of a `SortField` object is the zero-based index of the column to sort-on.</span></span>

   - <span data-ttu-id="3746f-204">O membro `sort` de uma `Table` é um objeto `TableSort`, não um método.</span><span class="sxs-lookup"><span data-stu-id="3746f-204">The `sort` member of a `Table` is a `TableSort` object, not a method.</span></span> <span data-ttu-id="3746f-205">Os `SortField`s são passados para o método `apply` do objeto `TableSort`.</span><span class="sxs-lookup"><span data-stu-id="3746f-205">The `SortField`s are passed to the `TableSort` object's `apply` method.</span></span>

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

### <a name="test-the-add-in"></a><span data-ttu-id="3746f-206">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="3746f-206">Test the add-in</span></span>

1. <span data-ttu-id="3746f-207">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite **Ctrl + C** duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="3746f-207">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="3746f-208">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="3746f-208">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="3746f-209">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor.</span><span class="sxs-lookup"><span data-stu-id="3746f-209">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="3746f-210">Para fazer isso, interrompa o processo do servidor para obter uma solicitação para inserir o comando de compilação.</span><span class="sxs-lookup"><span data-stu-id="3746f-210">In order to do this, you need to kill the server process so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="3746f-211">Após a compilação, reinicie o servidor.</span><span class="sxs-lookup"><span data-stu-id="3746f-211">After the build, you restart the server.</span></span> <span data-ttu-id="3746f-212">As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="3746f-212">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="3746f-213">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte no Internet Explorer (que é usada em segundo plano pelo Excel para executar os suplementos do Excel).</span><span class="sxs-lookup"><span data-stu-id="3746f-213">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="3746f-214">Execute o comando `npm start` para iniciar um servidor Web em um host local.</span><span class="sxs-lookup"><span data-stu-id="3746f-214">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="3746f-215">Feche o painel de tarefas para recarregá-lo e, no menu **Página Inicial**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="3746f-215">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="3746f-216">Se por qualquer motivo a tabela não estiver na planilha aberta, no painel de tarefas, escolha **Criar Tabela**.</span><span class="sxs-lookup"><span data-stu-id="3746f-216">If for any reason the table is not in the open worksheet, in the task pane, choose **Create Table**.</span></span>

6. <span data-ttu-id="3746f-217">Escolha os botões **Filtrar Tabela** e **Classificar Tabela** em qualquer ordem.</span><span class="sxs-lookup"><span data-stu-id="3746f-217">Choose the **Filter Table** and **Sort Table** buttons, in either order.</span></span>

    ![Tutorial do Excel: filtrar e classificar tabela](../images/excel-tutorial-filter-and-sort-table.png)

## <a name="create-a-chart"></a><span data-ttu-id="3746f-219">Criar um gráfico</span><span class="sxs-lookup"><span data-stu-id="3746f-219">Create a chart</span></span>

<span data-ttu-id="3746f-220">Nesta etapa do tutorial, você vai criar um gráfico com dados da tabela que você criou anteriormente e depois vai formatar o gráfico.</span><span class="sxs-lookup"><span data-stu-id="3746f-220">In this step of the tutorial, you'll create a chart using data from the table that you created previously, and then format the chart.</span></span>

### <a name="chart-a-chart-using-table-data"></a><span data-ttu-id="3746f-221">Gráfico de um gráfico com dados de tabela</span><span class="sxs-lookup"><span data-stu-id="3746f-221">Chart a chart using table data</span></span>

1. <span data-ttu-id="3746f-222">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="3746f-222">Open the project in your code editor.</span></span>

2. <span data-ttu-id="3746f-223">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="3746f-223">Open the file index.html.</span></span>

3. <span data-ttu-id="3746f-224">Abaixo do `div` que contém o botão `sort-table`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="3746f-224">Below the `div` that contains the `sort-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="create-chart">Create Chart</button>
    </div>
    ```

4. <span data-ttu-id="3746f-225">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="3746f-225">Open the app.js file.</span></span>

5. <span data-ttu-id="3746f-226">Abaixo da linha que atribui um identificador de clique ao botão `sort-chart`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="3746f-226">Below the line that assigns a click handler to the `sort-chart` button, add the following code:</span></span>

    ```js
    $('#create-chart').click(createChart);
    ```

6. <span data-ttu-id="3746f-227">Abaixo da função `sortTable`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="3746f-227">Below the `sortTable` function add the following function.</span></span>

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

7. <span data-ttu-id="3746f-p119">Substitua `TODO1` pelo código a seguir. Para excluir a linha de cabeçalho, o código usa o método `Table.getDataBodyRange` para acessar o intervalo de dados que você deseja representar graficamente em vez do método `getRange`.</span><span class="sxs-lookup"><span data-stu-id="3746f-p119">Replace `TODO1` with the following code. Note that in order to exclude the header row, the code uses the `Table.getDataBodyRange` method to get the range of data you want to chart instead of the `getRange` method.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var dataRange = expensesTable.getDataBodyRange();
    ```

8. <span data-ttu-id="3746f-230">Substitua `TODO2` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="3746f-230">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="3746f-231">Observe os seguintes parâmetros:</span><span class="sxs-lookup"><span data-stu-id="3746f-231">Note the following parameters:</span></span>

   - <span data-ttu-id="3746f-p121">O primeiro parâmetro para o método `add` especifica o tipo de gráfico. Há diversos tipos.</span><span class="sxs-lookup"><span data-stu-id="3746f-p121">The first parameter to the `add` method specifies the type of chart. There are several dozen types.</span></span>

   - <span data-ttu-id="3746f-234">O segundo parâmetro especifica um intervalo de dados a incluir no gráfico.</span><span class="sxs-lookup"><span data-stu-id="3746f-234">The second parameter specifies the range of data to include in the chart.</span></span>

   - <span data-ttu-id="3746f-235">O terceiro parâmetro determina se uma série de pontos de dados da tabela deve estar representada por linha ou por coluna.</span><span class="sxs-lookup"><span data-stu-id="3746f-235">The third parameter determines whether a series of data points from the table should be charted row-wise or column-wise.</span></span> <span data-ttu-id="3746f-236">A opção `auto` informa ao Excel para decidir o melhor método.</span><span class="sxs-lookup"><span data-stu-id="3746f-236">The option `auto` tells Excel to decide the best method.</span></span>

    ```js
    var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ```

9. <span data-ttu-id="3746f-237">Substitua `TODO3` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="3746f-237">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="3746f-238">A maior parte do código é autoexplicativa.</span><span class="sxs-lookup"><span data-stu-id="3746f-238">Most of this code is self-explanatory.</span></span> <span data-ttu-id="3746f-239">Observação:</span><span class="sxs-lookup"><span data-stu-id="3746f-239">Note:</span></span>
   
   - <span data-ttu-id="3746f-240">Os parâmetros do método `setPosition` especificam as células da esquerda superior e da direita inferior da área da planilha que deve conter o gráfico.</span><span class="sxs-lookup"><span data-stu-id="3746f-240">The parameters to the `setPosition` method specify the upper left and lower right cells of the worksheet area that should contain the chart.</span></span> <span data-ttu-id="3746f-241">O Excel ajusta detalhes como a largura da linha para criar uma boa aparência para o gráfico no espaço fornecido.</span><span class="sxs-lookup"><span data-stu-id="3746f-241">Excel can adjust things like line width to make the chart look good in the space it has been given.</span></span>
   
   - <span data-ttu-id="3746f-242">"Série" é um conjunto de pontos de dados de uma coluna da tabela.</span><span class="sxs-lookup"><span data-stu-id="3746f-242">A "series" is a set of data points from a column of the table.</span></span> <span data-ttu-id="3746f-243">Como há apenas uma coluna sem cadeia de caracteres na tabela, o Excel deduz que essa é a única coluna de pontos de dados no gráfico.</span><span class="sxs-lookup"><span data-stu-id="3746f-243">Since there is only one non-string column in the table, Excel infers that the column is the only column of data points to chart.</span></span> <span data-ttu-id="3746f-244">Ele interpreta outras colunas como rótulos do gráfico.</span><span class="sxs-lookup"><span data-stu-id="3746f-244">It interprets the other columns as chart labels.</span></span> <span data-ttu-id="3746f-245">Portanto, haverá apenas uma série no gráfico e será necessário o índice 0.</span><span class="sxs-lookup"><span data-stu-id="3746f-245">So there will be just one series in the chart and it will have index 0.</span></span> <span data-ttu-id="3746f-246">Ele será rotulado como "Valor em €".</span><span class="sxs-lookup"><span data-stu-id="3746f-246">This is the one to label with "Value in €".</span></span>

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in €';
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="3746f-247">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="3746f-247">Test the add-in</span></span>

1. <span data-ttu-id="3746f-248">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite **Ctrl + C** duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="3746f-248">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="3746f-249">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="3746f-249">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="3746f-250">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor.</span><span class="sxs-lookup"><span data-stu-id="3746f-250">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="3746f-251">Para fazer isso, interrompa o processo do servidor para obter uma solicitação para inserir o comando de compilação.</span><span class="sxs-lookup"><span data-stu-id="3746f-251">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="3746f-252">Após a compilação, reinicie o servidor.</span><span class="sxs-lookup"><span data-stu-id="3746f-252">After the build, you restart the server.</span></span> <span data-ttu-id="3746f-253">As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="3746f-253">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="3746f-254">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte no Internet Explorer (que é usada em segundo plano pelo Excel para executar os suplementos do Excel).</span><span class="sxs-lookup"><span data-stu-id="3746f-254">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="3746f-255">Execute o comando `npm start` para iniciar um servidor Web em um host local.</span><span class="sxs-lookup"><span data-stu-id="3746f-255">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="3746f-256">Feche o painel de tarefas para recarregá-lo e, no menu **Página Inicial**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="3746f-256">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="3746f-257">Se, por algum motivo, a tabela não estiver na planilha aberta, no painel de tarefas, escolha **Criar Tabela** e depois os botões **Filtrar Tabela** e \*\*Classificar Tabela \*\* em qualquer ordem.</span><span class="sxs-lookup"><span data-stu-id="3746f-257">If for any reason the table is not in the open worksheet, in the task pane, choose **Create Table** and then **Filter Table** and **Sort Table** buttons, in either order.</span></span>

6. <span data-ttu-id="3746f-258">Escolha o botão **Criar gráfico**.</span><span class="sxs-lookup"><span data-stu-id="3746f-258">Choose the **Create Chart** button.</span></span> <span data-ttu-id="3746f-259">Um gráfico é criado e incluirá somente os dados das linhas que foram filtradas.</span><span class="sxs-lookup"><span data-stu-id="3746f-259">A chart is created and only the data from the rows that have been filtered are included.</span></span> <span data-ttu-id="3746f-260">Os rótulos dos pontos de dados na parte inferior estão na ordem de classificação do gráfico, ou seja, nomes de comerciantes em ordem alfabética inversa.</span><span class="sxs-lookup"><span data-stu-id="3746f-260">The labels on the data points across the bottom are in the sort order of the chart; that is, merchant names in reverse alphabetical order.</span></span>

    ![Tutorial do Excel - Criar gráfico ](../images/excel-tutorial-create-chart.png)

## <a name="freeze-a-table-header"></a><span data-ttu-id="3746f-262">Congelar um cabeçalho de tabela</span><span class="sxs-lookup"><span data-stu-id="3746f-262">Freeze a table header in place</span></span>

<span data-ttu-id="3746f-263">Quando uma tabela for longa o suficiente para que um usuário precise rolar para ver algumas linhas, a linha de cabeçalho poderá ficar fora da vista.</span><span class="sxs-lookup"><span data-stu-id="3746f-263">When a table is long enough that a user must scroll to see some rows, the header row can scroll out of sight.</span></span> <span data-ttu-id="3746f-264">Nesta etapa do tutorial, você precisará congelar a linha do cabeçalho da tabela que criou anteriormente para que ela permaneça visível, mesmo que o usuário role ao longo da planilha.</span><span class="sxs-lookup"><span data-stu-id="3746f-264">In this step of the tutorial, you'll freeze the header row of the table that you created previously, so that it remains visible even as the user scrolls down the worksheet.</span></span>

### <a name="freeze-the-tables-header-row"></a><span data-ttu-id="3746f-265">Congelar a linha de cabeçalho da tabela</span><span class="sxs-lookup"><span data-stu-id="3746f-265">Freeze the table's header row</span></span>

1. <span data-ttu-id="3746f-266">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="3746f-266">Open the project in your code editor.</span></span>

2. <span data-ttu-id="3746f-267">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="3746f-267">Open the file index.html.</span></span>

3. <span data-ttu-id="3746f-268">Abaixo do `div` que contém o botão `create-chart`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="3746f-268">Below the `div` that contains the `create-chart` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="freeze-header">Freeze Header</button>
    </div>
    ```

4. <span data-ttu-id="3746f-269">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="3746f-269">Open the app.js file.</span></span>

5. <span data-ttu-id="3746f-270">Abaixo da linha que atribui um identificador de clique ao botão `create-chart`, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="3746f-270">Below the line that assigns a click handler to the `create-chart` button, add the following code:</span></span>

    ```js
    $('#freeze-header').click(freezeHeader);
    ```

6. <span data-ttu-id="3746f-271">Abaixo da função `createChart`, adicione a função a seguir:</span><span class="sxs-lookup"><span data-stu-id="3746f-271">Below the `createChart` function add the following function:</span></span>

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

7. <span data-ttu-id="3746f-p130">Substitua `TODO1` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="3746f-p130">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="3746f-274">A coleção `Worksheet.freezePanes` é um conjunto de painéis da planilha que fica congelado ou fixado no mesmo lugar quando rolamos a planilha.</span><span class="sxs-lookup"><span data-stu-id="3746f-274">The `Worksheet.freezePanes` collection is a set of panes in the worksheet that are pinned, or frozen, in place when the worksheet is scrolled.</span></span>

   - <span data-ttu-id="3746f-p131">O método `freezeRows` considera como parâmetro o número de linhas, começando da parte superior, que devem ser fixadas no local. Passamos `1` para fixar a primeira linha no local.</span><span class="sxs-lookup"><span data-stu-id="3746f-p131">The `freezeRows` method takes as a parameter the number of rows, from the top that are to be pinned in place. We pass `1` to pin the first row in place.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="3746f-277">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="3746f-277">Test the add-in</span></span>

1. <span data-ttu-id="3746f-278">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite **Ctrl + C** duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="3746f-278">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="3746f-279">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="3746f-279">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="3746f-280">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor.</span><span class="sxs-lookup"><span data-stu-id="3746f-280">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="3746f-281">Para fazer isso, interrompa o processo do servidor para obter uma solicitação para inserir o comando de compilação.</span><span class="sxs-lookup"><span data-stu-id="3746f-281">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="3746f-282">Após a compilação, reinicie o servidor.</span><span class="sxs-lookup"><span data-stu-id="3746f-282">After the build, you restart the server.</span></span> <span data-ttu-id="3746f-283">As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="3746f-283">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="3746f-284">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte no Internet Explorer (que é usada em segundo plano pelo Excel para executar os suplementos do Excel).</span><span class="sxs-lookup"><span data-stu-id="3746f-284">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="3746f-285">Execute o comando `npm start` para iniciar um servidor Web em um host local.</span><span class="sxs-lookup"><span data-stu-id="3746f-285">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="3746f-286">Feche o painel de tarefas para recarregá-lo e, no menu **Início**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="3746f-286">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="3746f-287">Se a tabela estiver na planilha, exclua-a.</span><span class="sxs-lookup"><span data-stu-id="3746f-287">If the table is in the worksheet, delete it.</span></span>

6. <span data-ttu-id="3746f-288">No painel de tarefas, escolha **Criar Tabela**.</span><span class="sxs-lookup"><span data-stu-id="3746f-288">In the task pane, choose **Create Table**.</span></span>

7. <span data-ttu-id="3746f-289">Escolha o botão **Congelar Cabeçalho**.</span><span class="sxs-lookup"><span data-stu-id="3746f-289">Choose the **Freeze Header** button.</span></span>

8. <span data-ttu-id="3746f-290">Role a planilha para baixo, o suficiente para ver que o cabeçalho da tabela permanece visível na parte superior mesmo ao rolar até que as primeiras linhas fiquem fora da vista.</span><span class="sxs-lookup"><span data-stu-id="3746f-290">Scroll down the worksheet enough to to see that the table header remains visible at the top even when the higher rows scroll out of sight.</span></span>

    ![Tutorial do Excel: congelar cabeçalho](../images/excel-tutorial-freeze-header.png)

## <a name="protect-a-worksheet"></a><span data-ttu-id="3746f-292">Proteger uma planilha</span><span class="sxs-lookup"><span data-stu-id="3746f-292">Protect a worksheet from changes</span></span>

<span data-ttu-id="3746f-293">Nesta etapa do tutorial, você adicionará outro botão à faixa de opções que, quando selecionado, executa uma função que você precisará definir para ativar e desativar a proteção da planilha.</span><span class="sxs-lookup"><span data-stu-id="3746f-293">In this step of the tutorial, you'll add another button to the ribbon that, when chosen, executes a function that you'll define to toggle worksheet protection on and off.</span></span>

### <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a><span data-ttu-id="3746f-294">Configure o manifesto para adicionar um segundo botão à faixa de opções</span><span class="sxs-lookup"><span data-stu-id="3746f-294">Configure the manifest to add a second ribbon button</span></span>

1. <span data-ttu-id="3746f-295">Abra o arquivo de manifesto my-office-add-in-manifest.xml.</span><span class="sxs-lookup"><span data-stu-id="3746f-295">Open the manifest file my-office-add-in-manifest.xml.</span></span>

2. <span data-ttu-id="3746f-296">Encontre o elemento `<Control>`.</span><span class="sxs-lookup"><span data-stu-id="3746f-296">Find the `<Control>` element.</span></span> <span data-ttu-id="3746f-297">Esse elemento define o botão **Mostrar Painel de Tarefas** na faixa de opções **Início** que você usa para iniciar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="3746f-297">This element defines the **Show Taskpane** button on the **Home** ribbon you have been using to launch the add-in.</span></span> <span data-ttu-id="3746f-298">Vamos adicionar um segundo botão ao mesmo grupo na faixa de opções **Início**.</span><span class="sxs-lookup"><span data-stu-id="3746f-298">We're going to add a second button to the same group on the **Home** ribbon.</span></span> <span data-ttu-id="3746f-299">Entre a marca de Controle final (`</Control>`) e a marca de Grupo final (`</Group>`), adicione a marcação a seguir.</span><span class="sxs-lookup"><span data-stu-id="3746f-299">In between the end Control tag (`</Control>`) and the end Group tag (`</Group>`), add the following markup.</span></span>

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

3. <span data-ttu-id="3746f-300">Substitua `TODO1` por uma cadeia de caracteres que fornece ao botão uma ID exclusiva no arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="3746f-300">Replace `TODO1` with a string that gives the button an ID that is unique within this manifest file.</span></span> <span data-ttu-id="3746f-301">Como nosso botão ativará ou desativará a proteção da planilha, use "ToggleProtection".</span><span class="sxs-lookup"><span data-stu-id="3746f-301">Since our button is going to toggle protection of the worksheet on and off, use "ToggleProtection".</span></span> <span data-ttu-id="3746f-302">Quando terminar, a marca de Controle de início inteira deve se parecer com o seguinte:</span><span class="sxs-lookup"><span data-stu-id="3746f-302">When you are done, the entire start Control tag should look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. <span data-ttu-id="3746f-303">Os próximos três `TODO`s definem “resid”, que significa ID de recurso.</span><span class="sxs-lookup"><span data-stu-id="3746f-303">The next three `TODO`s set "resid"s, which is short for resource ID.</span></span> <span data-ttu-id="3746f-304">Um recurso é uma cadeia de caracteres e você criará essas três cadeias de caracteres em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="3746f-304">A resource is a string, and you'll create these three strings in a later step.</span></span> <span data-ttu-id="3746f-305">Por enquanto, você precisa fornecer IDs aos recursos.</span><span class="sxs-lookup"><span data-stu-id="3746f-305">For now, you need to give IDs to the resources.</span></span> <span data-ttu-id="3746f-306">O rótulo do botão deve ser "Toggle Protection", mas a *ID* dessa cadeia de caracteres será "ProtectionButtonLabel", de forma que o elemento `Label` completo deve se parecer com o código a seguir:</span><span class="sxs-lookup"><span data-stu-id="3746f-306">The button label should read "Toggle Protection", but the *ID* of this string should be "ProtectionButtonLabel", so the completed `Label` element should look like the following code:</span></span>

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. <span data-ttu-id="3746f-307">O elemento `SuperTip` define a dica de ferramenta do botão.</span><span class="sxs-lookup"><span data-stu-id="3746f-307">The `SuperTip` element defines the tool tip for the button.</span></span> <span data-ttu-id="3746f-308">O título da dica de ferramenta deve ser o mesmo que o rótulo do botão, por isso, usamos a mesma ID de recurso: "ProtectionButtonLabel".</span><span class="sxs-lookup"><span data-stu-id="3746f-308">The tool tip title should be the same as the button label, so we use the very same resource ID: "ProtectionButtonLabel".</span></span> <span data-ttu-id="3746f-309">A descrição da dica de ferramenta será "Click to turn protection of the worksheet on and off".</span><span class="sxs-lookup"><span data-stu-id="3746f-309">The tool tip description will be "Click to turn protection of the worksheet on and off".</span></span> <span data-ttu-id="3746f-310">Mas o `ID` será "ProtectionButtonToolTip".</span><span class="sxs-lookup"><span data-stu-id="3746f-310">But the `ID` should be "ProtectionButtonToolTip".</span></span> <span data-ttu-id="3746f-311">Portanto, quando terminar, a marcação `SuperTip` inteira deve se parecer com o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="3746f-311">So, when you are done, the whole `SuperTip` markup should look like the following code:</span></span> 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > <span data-ttu-id="3746f-312">Em um suplemento de produção,não é recomendável usar o mesmo ícone para dois botões diferentes; mas, para simplificar este tutorial, faremos isso.</span><span class="sxs-lookup"><span data-stu-id="3746f-312">In a production add-in, you would not want to use the same icon for two different buttons; but to simplify this tutorial, we'll do that.</span></span> <span data-ttu-id="3746f-313">Portanto, a marcação `Icon` em nosso novo `Control` é apenas uma cópia do elemento `Icon` do `Control` existente.</span><span class="sxs-lookup"><span data-stu-id="3746f-313">So the `Icon` markup in our new `Control` is just a copy of the `Icon` element from the existing `Control`.</span></span> 

6. <span data-ttu-id="3746f-314">O elemento `Action` dentro do elemento `Control` original já está presente no manifesto, tem seu tipo definido como `ShowTaskpane`, mas nosso novo botão não abrirá um painel de tarefas, mas sim executará uma função personalizada criada em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="3746f-314">The `Action` element inside the original `Control` element that was already present in the manifest, has its type set to `ShowTaskpane`, but our new button isn't going to open a task pane; it's going to run a custom function that you create in a later step.</span></span> <span data-ttu-id="3746f-315">Portanto, substitua `TODO5` por `ExecuteFunction`, que é o tipo de ação para botões que acionam funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="3746f-315">So replace `TODO5` with `ExecuteFunction` which is the action type for buttons that trigger custom functions.</span></span> <span data-ttu-id="3746f-316">A marca `Action` de início deve ser similar ao código abaixo:</span><span class="sxs-lookup"><span data-stu-id="3746f-316">The start `Action` tag should look like the following code:</span></span>
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. <span data-ttu-id="3746f-317">O elemento `Action` original tem elementos filhos que especificam uma ID do painel de tarefas e uma URL da página que deve ser aberta no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="3746f-317">The original `Action` element has child elements that specify a task pane ID and a URL of the page that should be opened in the task pane.</span></span> <span data-ttu-id="3746f-318">No entanto, um elemento `Action` do tipo `ExecuteFunction` tem um único elemento filho que nomeia a função executada pelo controle.</span><span class="sxs-lookup"><span data-stu-id="3746f-318">But an `Action` element of the `ExecuteFunction` type has a single child element that names the function that the control executes.</span></span> <span data-ttu-id="3746f-319">Você criará essa função em uma etapa posterior e ela será chamada de `toggleProtection`.</span><span class="sxs-lookup"><span data-stu-id="3746f-319">You'll create that function in a later step, and it will be called `toggleProtection`.</span></span> <span data-ttu-id="3746f-320">Então, substitua `TODO6` pela marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="3746f-320">So, replace `TODO6` with the following markup:</span></span>
 
    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    <span data-ttu-id="3746f-321">A marcação `Control` inteira deve ter a aparência a seguir:</span><span class="sxs-lookup"><span data-stu-id="3746f-321">The entire `Control` markup should now look like the following:</span></span>

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

8. <span data-ttu-id="3746f-322">Role para baixo até a seção `Resources` do manifesto.</span><span class="sxs-lookup"><span data-stu-id="3746f-322">Scroll down to the `Resources` section of the manifest.</span></span>

9. <span data-ttu-id="3746f-323">Adicione a seguinte marcação como filho do elemento `bt:ShortStrings`.</span><span class="sxs-lookup"><span data-stu-id="3746f-323">Add the following markup as a child of the `bt:ShortStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. <span data-ttu-id="3746f-324">Adicione a seguinte marcação como filho do elemento `bt:LongStrings`.</span><span class="sxs-lookup"><span data-stu-id="3746f-324">Add the following markup as a child of the `bt:LongStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. <span data-ttu-id="3746f-325">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="3746f-325">Save the file.</span></span>

### <a name="create-the-function-that-protects-the-sheet"></a><span data-ttu-id="3746f-326">Criar a função que protege a planilha</span><span class="sxs-lookup"><span data-stu-id="3746f-326">Create the function that protects the sheet</span></span>

1. <span data-ttu-id="3746f-327">Abra o arquivo \function-file\function-file.js.</span><span class="sxs-lookup"><span data-stu-id="3746f-327">Open the file \function-file\function-file.js.</span></span>

2. <span data-ttu-id="3746f-328">O arquivo já tem uma Expressão de Função Invocada Imediatamente (IFFE).</span><span class="sxs-lookup"><span data-stu-id="3746f-328">The file already has an Immediately Invoked Function Expression (IFFE).</span></span> <span data-ttu-id="3746f-329">Não é necessário ter uma lógica de inicialização personalizada, portanto, deixe a função atribuída a `Office.initialize` com um corpo vazio.</span><span class="sxs-lookup"><span data-stu-id="3746f-329">No custom initialization logic is needed, so leave the function that is assigned to `Office.initialize` with an empty body.</span></span> <span data-ttu-id="3746f-330">(Mas não a exclua.</span><span class="sxs-lookup"><span data-stu-id="3746f-330">(But do not delete it.</span></span> <span data-ttu-id="3746f-331">A propriedade `Office.initialize` não pode ser nula ou indefinida.) *Fora da IIFE*, adicione o seguinte código.</span><span class="sxs-lookup"><span data-stu-id="3746f-331">The `Office.initialize` property cannot be null or undefined.) *Outside of the IIFE*, add the following code.</span></span> <span data-ttu-id="3746f-332">Observe que é possível especificar um parâmetro `args` para o método e a última linha do método chamará `args.completed`.</span><span class="sxs-lookup"><span data-stu-id="3746f-332">Note that we specify an `args` parameter to the method and the very last line of the method calls `args.completed`.</span></span> <span data-ttu-id="3746f-333">Esse é um requisito para todos os comandos de suplemento do tipo **ExecuteFunction**.</span><span class="sxs-lookup"><span data-stu-id="3746f-333">This is a requirement for all add-in commands of type **ExecuteFunction**.</span></span> <span data-ttu-id="3746f-334">Ele sinaliza para o aplicativo host do Office que a função terminou e que a interface do usuário podem ficar responsiva novamente.</span><span class="sxs-lookup"><span data-stu-id="3746f-334">It signals the Office host application that the function has finished and the UI can become responsive again.</span></span>

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

3. <span data-ttu-id="3746f-335">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="3746f-335">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="3746f-336">O código usa propriedade de proteção do objeto de planilha em um padrão de botão de alternância padrão.</span><span class="sxs-lookup"><span data-stu-id="3746f-336">This code uses the worksheet object's protection property in a standard toggle pattern.</span></span> <span data-ttu-id="3746f-337">O `TODO2` será explicado na próxima seção.</span><span class="sxs-lookup"><span data-stu-id="3746f-337">The `TODO2` will be explained in the next section.</span></span>

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

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="3746f-338">Adicione código para buscar propriedades do documento em objetos de script do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="3746f-338">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="3746f-339">Em todas as funções anteriores desta série de tutoriais, você colocou em fila comandos para *gravar* no documento do Office.</span><span class="sxs-lookup"><span data-stu-id="3746f-339">In all the earlier functions in this series of tutorials, you queued commands to *write* to the Office document.</span></span> <span data-ttu-id="3746f-340">Cada função terminou com uma chamada para o método `context.sync()`, que envia os comandos em fila para o documento a ser executado.</span><span class="sxs-lookup"><span data-stu-id="3746f-340">Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed.</span></span> <span data-ttu-id="3746f-341">Entretanto, o código adicionado na última etapa chama a propriedade `sheet.protection.protected` e essa é uma grande diferença das funções anteriores que você escreveu, pois o objeto `sheet` é apenas um objeto de proxy que existe no script do seu painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="3746f-341">But the code you added in the last step calls the `sheet.protection.protected` property, and this is a significant difference from the earlier functions you wrote, because the `sheet` object is only a proxy object that exists in your task pane's script.</span></span> <span data-ttu-id="3746f-342">Ele não sabe qual é o estado real de proteção do documento, portanto, sua propriedade `protection.protected` não pode ter um valor real.</span><span class="sxs-lookup"><span data-stu-id="3746f-342">It doesn't know what the actual protection state of the document is, so its `protection.protected` property can't have a real value.</span></span> <span data-ttu-id="3746f-343">É necessário primeiro buscar o status de proteção do documento e definir o valor de `sheet.protection.protected`.</span><span class="sxs-lookup"><span data-stu-id="3746f-343">It is necessary to first fetch the protection status from the document and use it set the value of `sheet.protection.protected`.</span></span> <span data-ttu-id="3746f-344">Somente então será possível chamar `sheet.protection.protected` sem causar uma exceção.</span><span class="sxs-lookup"><span data-stu-id="3746f-344">Only then can `sheet.protection.protected` be called without causing an exception to be thrown.</span></span> <span data-ttu-id="3746f-345">Esse processo de busca tem três etapas:</span><span class="sxs-lookup"><span data-stu-id="3746f-345">This fetching process has three steps:</span></span>

   1. <span data-ttu-id="3746f-346">Coloque em fila um comando para carregar (ou seja, fetch) as propriedades que seu código precisa ler.</span><span class="sxs-lookup"><span data-stu-id="3746f-346">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>

   2. <span data-ttu-id="3746f-347">Chame o método `sync` do objeto de contexto para enviar o comando em fila para o documento para execução e retornar as informações solicitadas.</span><span class="sxs-lookup"><span data-stu-id="3746f-347">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>

   3. <span data-ttu-id="3746f-348">Como o método `sync` é assíncrono, certifique-se de que ele tenha sido concluído antes que o código chame as propriedades que foram buscadas.</span><span class="sxs-lookup"><span data-stu-id="3746f-348">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="3746f-349">Essas etapas devem ser concluídas sempre que seu código precisar *ler* informações do documento do Office.</span><span class="sxs-lookup"><span data-stu-id="3746f-349">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="3746f-p144">Na função `toggleProtection`, substitua `TODO2` pelo seguinte código. Observação:</span><span class="sxs-lookup"><span data-stu-id="3746f-p144">In the `toggleProtection` function, replace `TODO2` with the following code. Note:</span></span>
   
   - <span data-ttu-id="3746f-352">Todos os objetos do Excel têm um método `load`.</span><span class="sxs-lookup"><span data-stu-id="3746f-352">Every Excel object has a `load` method.</span></span> <span data-ttu-id="3746f-353">Especifique as propriedades do objeto que você deseja ler no parâmetro como uma cadeia de caracteres de nomes delimitados por vírgulas.</span><span class="sxs-lookup"><span data-stu-id="3746f-353">You specify the properties of the object that you want to read in the parameter as a string of comma-delimited names.</span></span> <span data-ttu-id="3746f-354">Nesse caso, a propriedade que você precisa ler é uma subpropriedade de `protection`.</span><span class="sxs-lookup"><span data-stu-id="3746f-354">In this case, the property you need to read is a subproperty of the `protection` property.</span></span> <span data-ttu-id="3746f-355">Referencie a subpropriedade quase exatamente como você faria em qualquer lugar do seu código, mas usando uma barra (“/”) em vez de um ponto (".").</span><span class="sxs-lookup"><span data-stu-id="3746f-355">You reference the subproperty almost exactly as you would anywhere else in your code, with the exception that you use a forward slash ('/') character instead of a "." character.</span></span>

   - <span data-ttu-id="3746f-356">Para garantir que a lógica de botão de alternância, `sheet.protection.protected`, não seja executada até após `sync` ser concluído e o `sheet.protection.protected` ser atribuída ao valor correto buscado no documento, ele será movido (na próxima etapa) para uma função `then` que não será executada até `sync` ser concluído.</span><span class="sxs-lookup"><span data-stu-id="3746f-356">To ensure that the toggle logic, which reads `sheet.protection.protected`, does not run until after the `sync` is complete and the `sheet.protection.protected` has been assigned the correct value that is fetched from the document, it will be moved (in the next step) into a `then` function that won't run until the `sync` has completed.</span></span> 

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

2. <span data-ttu-id="3746f-357">Você não pode ter duas instruções `return` no mesmo caminho de código sem ramificações, portanto, exclua a linha final `return context.sync();` no final de `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="3746f-357">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Excel.run`.</span></span> <span data-ttu-id="3746f-358">Você adicionará um novo `context.sync` final em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="3746f-358">You will add a new final `context.sync`, in a later step.</span></span>

3. <span data-ttu-id="3746f-359">Recorte a estrutura `if ... else` na função `toggleProtection` e a cole no lugar de `TODO3`.</span><span class="sxs-lookup"><span data-stu-id="3746f-359">Cut the `if ... else` structure in the `toggleProtection` function and paste it in place of `TODO3`.</span></span>

4. <span data-ttu-id="3746f-p147">Substitua `TODO4` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="3746f-p147">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="3746f-362">Passar o método `sync` para uma função `then` garante que ele não seja executado até que `sheet.protection.unprotect()` ou `sheet.protection.protect()` seja enfileirado.</span><span class="sxs-lookup"><span data-stu-id="3746f-362">Passing the `sync` method to a `then` function ensures that it does not run until either `sheet.protection.unprotect()` or `sheet.protection.protect()` has been queued.</span></span>

   - <span data-ttu-id="3746f-363">O método `then` invoca qualquer função que é passada para ele e não é recomendável que `sync` seja chamado duas vezes, portanto, remova os “()” do fim de `context.sync`.</span><span class="sxs-lookup"><span data-stu-id="3746f-363">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of `context.sync`.</span></span>

    ```js
    .then(context.sync);
    ```

   <span data-ttu-id="3746f-364">Quando terminar, a função inteira deve se parecer com o seguinte:</span><span class="sxs-lookup"><span data-stu-id="3746f-364">When you are done, the entire function should look like the following:</span></span>

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

### <a name="configure-the-script-loading-html-file"></a><span data-ttu-id="3746f-365">Configure o arquivo HTML de carregamento de script</span><span class="sxs-lookup"><span data-stu-id="3746f-365">Configure the script-loading HTML file</span></span>

<span data-ttu-id="3746f-366">Abra o arquivo /function-file/function-file.html.</span><span class="sxs-lookup"><span data-stu-id="3746f-366">Open the /function-file/function-file.html file.</span></span> <span data-ttu-id="3746f-367">Esse é um arquivo HTML sem IU que é chamado quando o usuário pressiona o botão **Ativar/Desativar Proteção da Planilha**.</span><span class="sxs-lookup"><span data-stu-id="3746f-367">This is a UI-less HTML file that is called when the user presses the **Toggle Worksheet Protection** button.</span></span> <span data-ttu-id="3746f-368">O objetivo é carregar o método JavaScript que deve ser executado quando botão é pressionado.</span><span class="sxs-lookup"><span data-stu-id="3746f-368">Its purpose is to load the JavaScript method that should run when the button is pushed.</span></span> <span data-ttu-id="3746f-369">Esse arquivo não será alterado.</span><span class="sxs-lookup"><span data-stu-id="3746f-369">You are not going to change this file.</span></span> <span data-ttu-id="3746f-370">Basta observar que a segunda marca `<script>` carrega o functionfile.js.</span><span class="sxs-lookup"><span data-stu-id="3746f-370">Simply note that the second `<script>` tag loads the functionfile.js.</span></span>

   > [!NOTE]
   > <span data-ttu-id="3746f-371">O arquivo function-file.html e o arquivo function-file.js carregado são executados em um processo do IE completamente separado de painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="3746f-371">The function-file.html file and the function-file.js file that it loads run in an entirely separate IE process from the add-in's task pane.</span></span> <span data-ttu-id="3746f-372">Se o function-file.js foi transcompilado no mesmo arquivo bundle.js que o arquivo app.js, o suplemento precisará carregar duas cópias do arquivo bundle.js, o que anule o propósito do agrupamento.</span><span class="sxs-lookup"><span data-stu-id="3746f-372">If the function-file.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="3746f-373">Além disso, o arquivo function-file.js não contém qualquer JavaScript incompatível com o Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="3746f-373">In addition, the function-file.js file does not contain any JavaScript that is unsupported by IE.</span></span> <span data-ttu-id="3746f-374">Por esses dois motivos, esse suplemento não transcompila o function-file.js.</span><span class="sxs-lookup"><span data-stu-id="3746f-374">For these two reasons, this add-in does not transpile the function-file.js at all.</span></span> 

### <a name="test-the-add-in"></a><span data-ttu-id="3746f-375">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="3746f-375">Test the add-in</span></span>

1. <span data-ttu-id="3746f-376">Feche todos os aplicativos do Office, incluindo o Excel.</span><span class="sxs-lookup"><span data-stu-id="3746f-376">Close all Office applications, including Excel.</span></span> 

2. <span data-ttu-id="3746f-377">Para excluir o cache do Office, exclua o conteúdo da pasta de cache.</span><span class="sxs-lookup"><span data-stu-id="3746f-377">Delete the Office cache by deleting the contents of the cache folder.</span></span> <span data-ttu-id="3746f-378">Isso é necessário para limpar totalmente a versão anterior do suplemento do host.</span><span class="sxs-lookup"><span data-stu-id="3746f-378">This is necessary to completely clear the old version of the add-in from the host.</span></span> 

    - <span data-ttu-id="3746f-379">No Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="3746f-379">For Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

    - <span data-ttu-id="3746f-380">No Mac: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="3746f-380">For Mac: `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

3. <span data-ttu-id="3746f-381">Se, por algum motivo, o servidor não estiver executando, em uma janela do Git Bash ou em um prompt do sistema habilitado para Node.JS, acesse a pasta **Iniciar** do projeto e execute o comando `npm start`.</span><span class="sxs-lookup"><span data-stu-id="3746f-381">If for any reason, your server is not running, then in a Git Bash window, or Node.JS-enabled system prompt, navigate to the **Start** folder of the project and run the command `npm start`.</span></span> <span data-ttu-id="3746f-382">Não é necessário recriar o projeto, pois o único arquivo JavaScript que você alterou não faz parte do bundle.js interno.</span><span class="sxs-lookup"><span data-stu-id="3746f-382">You do not need to rebuild the project because the only JavaScript file you changed is not part of the built bundle.js.</span></span>

4. <span data-ttu-id="3746f-383">Usando a nova versão do arquivo de manifesto alterado, repita o processo de sideloading usando um dos seguintes métodos.</span><span class="sxs-lookup"><span data-stu-id="3746f-383">Using the new version of the changed manifest file, repeat the sideloading process by using one of the following methods.</span></span> <span data-ttu-id="3746f-384">*Você deve substituir a cópia anterior do arquivo de manifesto.*</span><span class="sxs-lookup"><span data-stu-id="3746f-384">*You should overwrite the previous copy of the manifest file.*</span></span>

    - <span data-ttu-id="3746f-385">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="3746f-385">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>

    - <span data-ttu-id="3746f-386">Excel Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="3746f-386">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>

    - <span data-ttu-id="3746f-387">iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="3746f-387">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

5. <span data-ttu-id="3746f-388">Abra qualquer planilha no Excel.</span><span class="sxs-lookup"><span data-stu-id="3746f-388">Open any worksheet in Excel.</span></span>

6. <span data-ttu-id="3746f-p153">Na Faixa de Opções, em **Página Inicial**, escolha **Ativar Proteger Planilha**. Observe que a maioria dos controles na Faixa de Opções está desabilitada (e visualmente esmaecida) conforme mostrado na captura de tela abaixo.</span><span class="sxs-lookup"><span data-stu-id="3746f-p153">On the **Home** ribbon, choose **Toggle Worksheet Protection**. Note that most of the controls on the ribbon are disabled (and visually grayed-out) as seen in screenshot below.</span></span> 

7. <span data-ttu-id="3746f-391">Escolha uma célula como se quisesse alterar o conteúdo.</span><span class="sxs-lookup"><span data-stu-id="3746f-391">Choose a cell as you would if you wanted to change its content.</span></span> <span data-ttu-id="3746f-392">Você receberá um erro informando que a planilha está protegida.</span><span class="sxs-lookup"><span data-stu-id="3746f-392">You get an error telling you that the worksheet is protected.</span></span>

8. <span data-ttu-id="3746f-393">Escolha **Ativar/Desativar Proteção da Planilha** novamente e os controles serão reabilitados e você poderá alterar os valores das células.</span><span class="sxs-lookup"><span data-stu-id="3746f-393">Choose **Toggle Worksheet Protection** again, and the controls are reenabled, and you can change cell values again.</span></span>

    ![Tutorial do Excel - Faixa de Opções com a Proteção Ativada](../images/excel-tutorial-ribbon-with-protection-on.png)

## <a name="open-a-dialog"></a><span data-ttu-id="3746f-395">Abrir uma caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="3746f-395">Open a dialog box.</span></span>

<span data-ttu-id="3746f-396">Nesta etapa final do tutorial, você abre uma caixa de diálogo no suplemento, passa uma mensagem do processo de caixa de diálogo para o processo de painel de tarefas e fecha a caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="3746f-396">In this final step of the tutorial, you'll open a dialog in your add-in, pass a message from the dialog process to the task pane process, and close the dialog.</span></span> <span data-ttu-id="3746f-397">As caixas de diálogo do Suplemento do Office são *não modais*: o usuário pode continuar a interagir com o documento no aplicativo do Office do host e com a página host no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="3746f-397">Office Add-in dialogs are *nonmodal*: a user can continue to interact with both the document in the host Office application and with the host page in the task pane.</span></span>

### <a name="create-the-dialog-page"></a><span data-ttu-id="3746f-398">Crie a página da caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="3746f-398">Create the dialog page</span></span>

1. <span data-ttu-id="3746f-399">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="3746f-399">Open the project in your code editor.</span></span>

2. <span data-ttu-id="3746f-400">Crie um arquivo chamado popup.html na raiz do projeto (onde se encontra index.html).</span><span class="sxs-lookup"><span data-stu-id="3746f-400">Create a file in the root of the project (where index.html is) called popup.html.</span></span>

3. <span data-ttu-id="3746f-p156">Adicione a marcação a seguir em popup.html. Observação:</span><span class="sxs-lookup"><span data-stu-id="3746f-p156">Add the following markup to popup.html. Note:</span></span>

   - <span data-ttu-id="3746f-403">a página tem um `<input>` em que o usuário insere o nome dele e um botão que envia o nome para a página no painel de tarefas onde ele será exibido.</span><span class="sxs-lookup"><span data-stu-id="3746f-403">The page has a `<input>` where the user will enter their name and a button that will send the name to the page in the task pane where it will be displayed.</span></span>

   - <span data-ttu-id="3746f-404">A marcação carrega um script chamado popup.js que você criará em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="3746f-404">The markup loads a script called popup.js that you will create in a later step.</span></span>

   - <span data-ttu-id="3746f-405">Ela também carrega uma biblioteca Office.JS e jQuery porque elas serão usadas em popup.js.</span><span class="sxs-lookup"><span data-stu-id="3746f-405">It also loads the Office.JS library and jQuery because they will be used in popup.js.</span></span>

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

4. <span data-ttu-id="3746f-406">Crie um arquivo chamado popup.js na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="3746f-406">Create a file in the root of the project called popup.js.</span></span>

5. <span data-ttu-id="3746f-p157">Adicione o código a seguir a popup.js. Observação:</span><span class="sxs-lookup"><span data-stu-id="3746f-p157">Add the following code to popup.js. Note:</span></span>

   - <span data-ttu-id="3746f-409">*Todas as páginas que chamam APIs na biblioteca Office.JS devem atribuir uma função à propriedade `Office.initialize`.*</span><span class="sxs-lookup"><span data-stu-id="3746f-409">*Every page that calls APIs in the Office.JS library must assign a function to the `Office.initialize` property.*</span></span> <span data-ttu-id="3746f-410">Se nenhuma inicialização for necessária, a função poderá ter um corpo vazio, mas a propriedade não deve ser deixada indefinida, atribuída a nulo ou a um valor que não seja uma função.</span><span class="sxs-lookup"><span data-stu-id="3746f-410">If no initialization is needed, then the function can have an empty body, but the property must not be left undefined, assigned to null or to a non-function value.</span></span> <span data-ttu-id="3746f-411">Por exemplo, veja o arquivo app.js na raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="3746f-411">For an example, see the app.js file in the project root.</span></span> <span data-ttu-id="3746f-412">O código que cria a tarefa deve ser executado antes de qualquer chamada para Office.JS; por isso, a tarefa se encontra em um arquivo de script que é carregado pela página, como neste caso.</span><span class="sxs-lookup"><span data-stu-id="3746f-412">The code that makes the assignment must run before any calls to Office.JS; hence the assignment is in a script file that is loaded by the page, as it is in this case.</span></span>
   
   - <span data-ttu-id="3746f-p159">A função jQuery `ready` é chamada dentro do método `initialize`. É uma regra quase universal que o código de carregamento, inicialização ou bootstrapping de outras bibliotecas JavaScript deva estar dentro da função `Office.initialize`.</span><span class="sxs-lookup"><span data-stu-id="3746f-p159">The jQuery `ready` function is called inside the `initialize` method. It is an almost universal rule that the loading, initializing, or bootstrapping code of other JavaScript libraries should be inside the `Office.initialize` function.</span></span>

    ```js
    (function () {
    "use strict";

        Office.initialize = function() {
            $(document).ready(function () {  

                // TODO1: Assign handler to the OK button.

            });
        }

        // TODO2: Create the OK button handler

    }());
    ```

6. <span data-ttu-id="3746f-415">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="3746f-415">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="3746f-416">Você criará a função `sendStringToParentPage` na próxima etapa.</span><span class="sxs-lookup"><span data-stu-id="3746f-416">You'll create the `sendStringToParentPage` function in the next step.</span></span>

    ```js
    $('#ok-button').click(sendStringToParentPage);
    ```

7. <span data-ttu-id="3746f-417">Substitua `TODO2` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="3746f-417">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="3746f-418">O método `messageParent` passa seu parâmetro para a página pai, neste caso, a página no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="3746f-418">The `messageParent` method passes its parameter to the parent page, in this case, the page in the task pane.</span></span> <span data-ttu-id="3746f-419">O parâmetro pode ser um booliano ou uma cadeia de caracteres, que inclui tudo o que pode ser serializado como uma cadeia de caracteres, como XML ou JSON.</span><span class="sxs-lookup"><span data-stu-id="3746f-419">The parameter can be a boolean or a string, which includes anything that can be serialized as a string, such as XML or JSON.</span></span>

    ```js
    function sendStringToParentPage() {
        var userName = $('#name-box').val();
        Office.context.ui.messageParent(userName);
    }
    ```

8. <span data-ttu-id="3746f-420">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="3746f-420">Save the file.</span></span>

   > [!NOTE]
   > <span data-ttu-id="3746f-421">O arquivo popup.html e o arquivo popup.js carregado são executados em um processo do Internet Explorer completamente separado de painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="3746f-421">The popup.html file, and the popup.js file that it loads, run in an entirely separate Internet Explorer process from the add-in's task pane.</span></span> <span data-ttu-id="3746f-422">Se o popup.js foi transcompilado no mesmo arquivo bundle.js que o arquivo app.js, o suplemento precisará carregar duas cópias do arquivo bundle.js, o que anule o propósito do agrupamento.</span><span class="sxs-lookup"><span data-stu-id="3746f-422">If the popup.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="3746f-423">Além disso, o arquivo popup.js não contém qualquer JavaScript incompatível com o Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="3746f-423">In addition, the popup.js file does not contain any JavaScript that is unsupported by IE.</span></span> <span data-ttu-id="3746f-424">Por esses dois motivos, esse suplemento não transcompila o popup.js.</span><span class="sxs-lookup"><span data-stu-id="3746f-424">For these two reasons, this add-in does not transpile the popup.js file at all.</span></span>

### <a name="open-the-dialog-from-the-task-pane"></a><span data-ttu-id="3746f-425">Abra a caixa de diálogo do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="3746f-425">Open the dialog from the task pane</span></span>

1. <span data-ttu-id="3746f-426">Abra o arquivo index.html.</span><span class="sxs-lookup"><span data-stu-id="3746f-426">Open the file index.html.</span></span>

2. <span data-ttu-id="3746f-427">Abaixo do `div` que contém o botão `freeze-header`, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="3746f-427">Below the `div` that contains the `freeze-header` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="open-dialog">Open Dialog</button>
    </div>
    ```

3. <span data-ttu-id="3746f-428">A caixa de diálogo solicitará que o usuário insira um nome e passará o nome de usuário para o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="3746f-428">The dialog will prompt the user to enter a name and pass the user's name to the task pane.</span></span> <span data-ttu-id="3746f-429">O painel de tarefas o exibirá em um rótulo.</span><span class="sxs-lookup"><span data-stu-id="3746f-429">The task pane will display it in a label.</span></span> <span data-ttu-id="3746f-430">Imediatamente abaixo do `div` que você adicionou, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="3746f-430">Immediately below the `div` that you just added, add the following markup:</span></span>

    ```html
    <div class="padding">
        <label id="user-name"></label>
    </div>
    ```

4. <span data-ttu-id="3746f-431">Abra o arquivo app.js.</span><span class="sxs-lookup"><span data-stu-id="3746f-431">Open the app.js file.</span></span>

5. <span data-ttu-id="3746f-432">Abaixo da linha que atribui um identificador de clique ao botão `freeze-header`, adicione o seguinte código.</span><span class="sxs-lookup"><span data-stu-id="3746f-432">Below the line that assigns a click handler to the `freeze-header` button, add the following code.</span></span> <span data-ttu-id="3746f-433">Você criará o método `openDialog` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="3746f-433">You'll create the `openDialog` method in a later step.</span></span>

    ```js
    $('#open-dialog').click(openDialog);
    ```

6. <span data-ttu-id="3746f-p165">Abaixo da função `freezeHeader`, adicione a declaração seguinte. Essa variável é usada para armazenar um objeto no contexto de execução da página pai que atua como um intermediador no contexto de execução da página da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="3746f-p165">Below the `freezeHeader` function add the following declaration. This variable is used to hold an object in the parent page's execution context that acts as an intermediator to the dialog page's execution context.</span></span>

    ```js
    var dialog = null;
    ```

7. <span data-ttu-id="3746f-436">Abaixo da declaração de `dialog`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="3746f-436">Below the declaration of `dialog`, add the following function.</span></span> <span data-ttu-id="3746f-437">É importante observar o que esse código *não* contém: não há nenhuma chamada de `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="3746f-437">The important thing to notice about this code is what is *not* there: there is no call of `Excel.run`.</span></span> <span data-ttu-id="3746f-438">Isso ocorre porque a API para abrir uma caixa de diálogo é compartilhada com todos os hosts do Office, portanto, ela faz parte da API de Office JavaScript Common, não da API específica do Excel.</span><span class="sxs-lookup"><span data-stu-id="3746f-438">This is because the API to open a dialog is shared among all Office hosts, so it is part of the Office JavaScript Common API, not the Excel-specific API.</span></span>

    ```js
    function openDialog() {
        // TODO1: Call the Office Common API that opens a dialog
    }
    ```

8. <span data-ttu-id="3746f-p167">Substitua `TODO1` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="3746f-p167">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="3746f-441">O método`displayDialogAsync` abre uma caixa de diálogo no centro da tela.</span><span class="sxs-lookup"><span data-stu-id="3746f-441">The `displayDialogAsync` method opens a dialog in the center of the screen.</span></span>

   - <span data-ttu-id="3746f-442">O primeiro parâmetro é a URL da página a ser aberta.</span><span class="sxs-lookup"><span data-stu-id="3746f-442">The first parameter is the URL of the page to open.</span></span>

   - <span data-ttu-id="3746f-p168">O segundo parâmetro passa opções. `height` e `width` são porcentagens do tamanho da janela do aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="3746f-p168">The second parameter passes options. `height` and `width` are percentages of the size of the Office application's window.</span></span>

    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},

        // TODO2: Add callback parameter.
    );
    ```

### <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a><span data-ttu-id="3746f-445">Processar a mensagem da caixa de diálogo e depois fechá-la</span><span class="sxs-lookup"><span data-stu-id="3746f-445">Process the message from the dialog and close the dialog</span></span>

1. <span data-ttu-id="3746f-p169">Continue no arquivo app.js e substitua `TODO2` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="3746f-p169">Continue in the app.js file, and replace `TODO2` with the following code. Note:</span></span>

   - <span data-ttu-id="3746f-448">O retorno de chamada é executado logo após a caixa de diálogo ser aberta com êxito e antes de o usuário executar qualquer ação nela.</span><span class="sxs-lookup"><span data-stu-id="3746f-448">The callback is executed immediately after the dialog successfully opens and before the user has taken any action in the dialog.</span></span>

   - <span data-ttu-id="3746f-449">O `result.value` é o objeto que funciona como um tipo de intermediário entre contextos execução das páginas de pai e de caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="3746f-449">The `result.value` is the object that acts as a kind of middleman between the execution contexts of the parent and dialog pages.</span></span>

   - <span data-ttu-id="3746f-450">A função `processMessage` será criada em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="3746f-450">The `processMessage` function will be created in a later step.</span></span> <span data-ttu-id="3746f-451">Esse identificador processará os valores que sejam enviados da página da caixa de diálogo com chamadas da função `messageParent`.</span><span class="sxs-lookup"><span data-stu-id="3746f-451">This handler will process any values that are sent from the dialog page with calls of the `messageParent` function.</span></span>

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. <span data-ttu-id="3746f-452">Abaixo da função `openDialog`, adicione a função a seguir.</span><span class="sxs-lookup"><span data-stu-id="3746f-452">Below the `openDialog` function, add the following function.</span></span>

    ```js
    function processMessage(arg) {
        $('#user-name').text(arg.message);
        dialog.close();
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="3746f-453">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="3746f-453">Test the add-in</span></span>

1. <span data-ttu-id="3746f-454">Se a janela Git bash ou o prompt de sistema habilitado para Node.JS do tutorial anterior ainda estiverem abertos, digite **Ctrl + C** duas vezes para interromper a execução do servidor Web.</span><span class="sxs-lookup"><span data-stu-id="3746f-454">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="3746f-455">Caso contrário, abra uma janela Git bash ou um prompt de sistema habilitado para Node.JS e navegue até a pasta **Iniciar** do projeto.</span><span class="sxs-lookup"><span data-stu-id="3746f-455">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="3746f-456">Embora o servidor de sincronização do navegador recarregue o suplemento no painel de tarefas sempre que você fizer uma alteração em algum arquivo, incluindo o arquivo app.js, ele não transcompila o JavaScript, portanto, é necessário repetir o comando de compilação para que as alterações em app.js as entrem em vigor.</span><span class="sxs-lookup"><span data-stu-id="3746f-456">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="3746f-457">Para fazer isso, interrompa o processo do servidor para obter uma solicitação para inserir o comando de compilação.</span><span class="sxs-lookup"><span data-stu-id="3746f-457">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="3746f-458">Após a compilação, reinicie o servidor.</span><span class="sxs-lookup"><span data-stu-id="3746f-458">After the build, you restart the server.</span></span> <span data-ttu-id="3746f-459">As próximas etapas executam esse processo.</span><span class="sxs-lookup"><span data-stu-id="3746f-459">The next few steps carry out this process.</span></span>

2. <span data-ttu-id="3746f-460">Execute o comando `npm run build` para transcompilar seu código-fonte ES6 para uma versão anterior do JavaScript com suporte no Internet Explorer (que é usada em segundo plano pelo Excel para executar os suplementos do Excel).</span><span class="sxs-lookup"><span data-stu-id="3746f-460">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>

3. <span data-ttu-id="3746f-461">Execute o comando `npm start` para iniciar um servidor Web em um host local.</span><span class="sxs-lookup"><span data-stu-id="3746f-461">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="3746f-462">Feche o painel de tarefas para recarregá-lo e, no menu **Início**, selecione **Mostrar Painel de Tarefas** para reabrir o suplemento.</span><span class="sxs-lookup"><span data-stu-id="3746f-462">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="3746f-463">Escolha o botão **Abrir Caixa de Diálogo** no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="3746f-463">Choose the **Open Dialog** button in the task pane.</span></span>

6. <span data-ttu-id="3746f-464">Quando a caixa de diálogo estiver aberta, arraste-a e redimensione-a.</span><span class="sxs-lookup"><span data-stu-id="3746f-464">While the dialog is open, drag it and resize it.</span></span> <span data-ttu-id="3746f-465">Observe que você pode interagir com a planilha e pressionar outros botões no painel de tarefas. No entanto, não é possível iniciar uma segunda caixa de diálogo na mesma página do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="3746f-465">Note that you can interact with the worksheet and press other buttons on the task pane, but you cannot launch a second dialog from the same task pane page.</span></span>

7. <span data-ttu-id="3746f-466">Na caixa de diálogo, digite um nome e escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="3746f-466">In the dialog, enter a name and choose **OK**.</span></span> <span data-ttu-id="3746f-467">O nome aparecerá no painel de tarefas e a caixa de diálogo será fechada.</span><span class="sxs-lookup"><span data-stu-id="3746f-467">The name appears on the task pane and the dialog closes.</span></span>

8. <span data-ttu-id="3746f-468">Opcionalmente, comente a linha `dialog.close();` na função `processMessage`.</span><span class="sxs-lookup"><span data-stu-id="3746f-468">Optionally, comment out the line `dialog.close();` in the `processMessage` function.</span></span> <span data-ttu-id="3746f-469">Em seguida, repita as etapas desta seção.</span><span class="sxs-lookup"><span data-stu-id="3746f-469">Then repeat the steps of this section.</span></span> <span data-ttu-id="3746f-470">A caixa de diálogo permanece aberta e você pode alterar o nome.</span><span class="sxs-lookup"><span data-stu-id="3746f-470">The dialog stays open and you can change the name.</span></span> <span data-ttu-id="3746f-471">É possível fechá-la manualmente pressionando o botão **X** no canto superior direito.</span><span class="sxs-lookup"><span data-stu-id="3746f-471">You can close it manually by pressing the **X** button in the upper right corner.</span></span>

    ![Tutorial do Excel - Caixa de diálogo](../images/excel-tutorial-dialog-open.png)

## <a name="next-steps"></a><span data-ttu-id="3746f-473">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="3746f-473">Next steps</span></span>

<span data-ttu-id="3746f-474">Neste tutorial você criou um suplemento do Excel que interage com tabelas, gráficos, planilhas e caixas de diálogo em uma pasta de trabalho do Excel.</span><span class="sxs-lookup"><span data-stu-id="3746f-474">In this tutorial, you've created an Excel task pane add-in that interacts with tables, charts, worksheets, and dialogs in an Excel workbook.</span></span> <span data-ttu-id="3746f-475">Para saber mais sobre o desenvolvimento de suplementos do Excel, continue no seguinte artigo:</span><span class="sxs-lookup"><span data-stu-id="3746f-475">To learn more about developing Outlook add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="3746f-476">Visão geral dos suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="3746f-476">Excel add-ins overview</span></span>](../excel/excel-add-ins-overview.md)
