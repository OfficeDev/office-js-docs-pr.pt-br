---
title: Tutorial de suplemento do Excel
description: Neste tutorial, você criará um suplemento do Excel que cria, preenche, filtra e classifica uma tabela, cria um gráfico, congela um cabeçalho de tabela, protege uma planilha e abre uma caixa de diálogo
ms.date: 01/16/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 70df5e7e78abf64bf36d33cade0b40ff8e3c18f4
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950891"
---
# <a name="tutorial-create-an-excel-task-pane-add-in"></a><span data-ttu-id="f62e0-103">Tutorial: criar um suplemento do painel de tarefas no Excel</span><span class="sxs-lookup"><span data-stu-id="f62e0-103">Tutorial: Create an Excel task pane add-in</span></span>

<span data-ttu-id="f62e0-104">Neste tutorial: você criará um suplemento do painel de tarefas no Excel</span><span class="sxs-lookup"><span data-stu-id="f62e0-104">In this tutorial, you'll create an Excel task pane add-in that:</span></span>

> [!div class="checklist"]
> * <span data-ttu-id="f62e0-105">Cria uma tabela</span><span class="sxs-lookup"><span data-stu-id="f62e0-105">Creates a table</span></span>
> * <span data-ttu-id="f62e0-106">Filtra e classifica uma tabela</span><span class="sxs-lookup"><span data-stu-id="f62e0-106">Filters and sorts a table</span></span>
> * <span data-ttu-id="f62e0-107">Cria um gráfico</span><span class="sxs-lookup"><span data-stu-id="f62e0-107">Creates a chart</span></span>
> * <span data-ttu-id="f62e0-108">Congela um cabeçalho de tabela</span><span class="sxs-lookup"><span data-stu-id="f62e0-108">Freezes a table header</span></span>
> * <span data-ttu-id="f62e0-109">Protege uma planilha</span><span class="sxs-lookup"><span data-stu-id="f62e0-109">Protects a worksheet</span></span>
> * <span data-ttu-id="f62e0-110">Abre uma caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="f62e0-110">Opens a dialog</span></span>

> [!TIP]
> <span data-ttu-id="f62e0-111">Se você já concluiu o início rápido [Criar um suplemento do painel de tarefas do Excel](../quickstarts/excel-quickstart-jquery.md) e gostaria de usar esse projeto como ponto de partida para este tutorial, vá diretamente para a seção [Criar uma tabela](#create-a-table).</span><span class="sxs-lookup"><span data-stu-id="f62e0-111">If you've already completed the [Build an Excel task pane add-in](../quickstarts/excel-quickstart-jquery.md) quick start, and want to use that project as a starting point for this tutorial, go directly to the [Create a table](#create-a-table) section to start this tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="f62e0-112">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="f62e0-112">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-your-add-in-project"></a><span data-ttu-id="f62e0-113">Criar seu projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="f62e0-113">Create your add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="f62e0-114">**Escolha o tipo de projeto:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="f62e0-114">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="f62e0-115">**Escolha o tipo de script:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="f62e0-115">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="f62e0-116">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="f62e0-116">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="f62e0-117">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?**</span><span class="sxs-lookup"><span data-stu-id="f62e0-117">**Which Office client application would you like to support?**</span></span> `Excel`

![Gerador do Yeoman](../images/yo-office-excel.png)

<span data-ttu-id="f62e0-119">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="f62e0-119">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="create-a-table"></a><span data-ttu-id="f62e0-120">Criar uma tabela</span><span class="sxs-lookup"><span data-stu-id="f62e0-120">Create a table</span></span>

<span data-ttu-id="f62e0-121">Nesta etapa do tutorial, você testará no programa se o suplemento é compatível com a versão atual do Excel do usuário, adicionará uma tabela a uma planilha, depois preencherá e formatará a tabela com os dados.</span><span class="sxs-lookup"><span data-stu-id="f62e0-121">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Excel, add a table to a worksheet, populate the table with data, and format it.</span></span>

### <a name="code-the-add-in"></a><span data-ttu-id="f62e0-122">Codificação do suplemento</span><span class="sxs-lookup"><span data-stu-id="f62e0-122">Code the add-in</span></span>

1. <span data-ttu-id="f62e0-123">Abra o projeto em seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="f62e0-123">Open the project in your code editor.</span></span>

2. <span data-ttu-id="f62e0-124">Abra o arquivo **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-124">Open the file **./src/taskpane/taskpane.html**.</span></span>  <span data-ttu-id="f62e0-125">Ele contém a marcação HTML para o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="f62e0-125">This file contains the HTML markup for the task pane.</span></span>

3. <span data-ttu-id="f62e0-126">Localize o elemento `<main>` e exclua todas as linhas que aparecem após a marca de abertura `<main>` e antes da marca de fechamento `</main>`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-126">Locate the `<main>` element and delete all lines that appear after the opening `<main>` tag and before the closing `</main>` tag.</span></span>

4. <span data-ttu-id="f62e0-127">Adicione a seguinte marcação imediatamente após a marca de abertura `<main>`:</span><span class="sxs-lookup"><span data-stu-id="f62e0-127">Add the following markup immediately after the opening `<main>` tag:</span></span>

    ```html
    <button class="ms-Button" id="create-table">Create Table</button><br/><br/>
    ```

5. <span data-ttu-id="f62e0-128">Abra o arquivo **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-128">Open the file **./src/taskpane/taskpane.js**.</span></span> <span data-ttu-id="f62e0-129">Ele contém o código da API JavaScript do Office que facilita a interação entre o painel de tarefas e o aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="f62e0-129">This file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

6. <span data-ttu-id="f62e0-130">Remova todas as referências ao botão `run` e à função `run()` da seguinte forma:</span><span class="sxs-lookup"><span data-stu-id="f62e0-130">Remove all references to the `run` button and the `run()` function by doing the following:</span></span>

    - <span data-ttu-id="f62e0-131">Localize e exclua a linha `document.getElementById("run").onclick = run;`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-131">Locate and delete the line `document.getElementById("run").onclick = run;`.</span></span>

    - <span data-ttu-id="f62e0-132">Localize e exclua toda a função `run()`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-132">Locate and delete the entire `run()` function.</span></span>

7. <span data-ttu-id="f62e0-133">Na chamada do método `Office.onReady`, localize a linha `if (info.host === Office.HostType.Excel) {` e adicione o seguinte código imediatamente após ela.</span><span class="sxs-lookup"><span data-stu-id="f62e0-133">Within the `Office.onReady` method call, locate the line `if (info.host === Office.HostType.Excel) {` and add the following code immediately after that line.</span></span> <span data-ttu-id="f62e0-134">Observação:</span><span class="sxs-lookup"><span data-stu-id="f62e0-134">Note:</span></span>

    - <span data-ttu-id="f62e0-135">a primeira parte desse código determina se a versão do usuário do Excel é compatível com uma versão do Excel.js que inclui todas as APIs que esta série de tutoriais usará.</span><span class="sxs-lookup"><span data-stu-id="f62e0-135">The first part of this code determines whether the user's version of Excel supports a version of Excel.js that includes all the APIs that this series of tutorials will use.</span></span> <span data-ttu-id="f62e0-136">Em um suplemento de produção, use o corpo do bloco condicional para ocultar ou desabilitar a interface do usuário que chame a APIs sem suporte.</span><span class="sxs-lookup"><span data-stu-id="f62e0-136">In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs.</span></span> <span data-ttu-id="f62e0-137">Dessa forma, permitirá que o usuário ainda use as partes do suplemento às quais a versão do Excel dá suporte.</span><span class="sxs-lookup"><span data-stu-id="f62e0-137">This will enable the user to still make use of the parts of the add-in that are supported by their version of Excel.</span></span>

    - <span data-ttu-id="f62e0-138">A segunda parte desse código adiciona um manipulador de eventos para o botão `create-table`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-138">The second part of this code adds an event handler for the `create-table` button.</span></span>

    ```js
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("create-table").onclick = createTable;
    ```

8. <span data-ttu-id="f62e0-139">Adicione a seguinte função ao final do arquivo.</span><span class="sxs-lookup"><span data-stu-id="f62e0-139">Add the following function to the end of the file.</span></span> <span data-ttu-id="f62e0-140">Observação:</span><span class="sxs-lookup"><span data-stu-id="f62e0-140">Note:</span></span>

    - <span data-ttu-id="f62e0-p106">A lógica de negócios de Excel.js será adicionada à função que passar por `Excel.run`. Essa lógica não é executada imediatamente. Em vez disso, ela é adicionada à fila de comandos pendentes.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p106">Your Excel.js business logic will be added to the function that is passed to `Excel.run`. This logic does not execute immediately. Instead, it is added to a queue of pending commands.</span></span>

    - <span data-ttu-id="f62e0-144">O método `context.sync` envia todos os comandos da fila para execução no Excel.</span><span class="sxs-lookup"><span data-stu-id="f62e0-144">The `context.sync` method sends all queued commands to Excel for execution.</span></span>

    - <span data-ttu-id="f62e0-p107">`Excel.run` é seguido por um bloco `catch`. Essa é uma prática recomendada que você sempre deve seguir.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p107">The `Excel.run` is followed by a `catch` block. This is a best practice that you should always follow.</span></span> 

    ```js
    function createTable() {
        Excel.run(function (context) {

            // TODO1: Queue table creation logic here.

            // TODO2: Queue commands to populate the table with data.

            // TODO3: Queue commands to format the table.

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

9. <span data-ttu-id="f62e0-147">Na função `createTable()`, substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="f62e0-147">Within the `createTable()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="f62e0-148">Observação:</span><span class="sxs-lookup"><span data-stu-id="f62e0-148">Note:</span></span>

    - <span data-ttu-id="f62e0-p109">O código cria uma tabela usando o método `add` de conjunto de tabela da planilha, que sempre existe mesmo que ela esteja vazia. Essa é a maneira padrão de criar objetos no Excel.js. Não há nenhuma API do construtor de classe e você nunca usará um operador `new` para criar um objeto do Excel. Em vez disso, adicione a um objeto de conjunto pai.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p109">The code creates a table by using `add` method of a worksheet's table collection, which always exists even if it is empty. This is the standard way that Excel.js objects are created. There are no class constructor APIs, and you never use a `new` operator to create an Excel object. Instead, you add to a parent collection object.</span></span>

    - <span data-ttu-id="f62e0-p110">O primeiro parâmetro do método `add`é o intervalo apenas da linha superior da tabela, não o intervalo inteiro que a tabela por fim usará. Isso ocorre porque, quando o suplemento preenche as linhas de dados (na próxima etapa), ele adicionará novas linhas à tabela, em vez de gravar os valores nas células das linhas existentes. Esse é um padrão mais comum, porque o número de linhas em uma tabela geralmente não é conhecido quando a tabela é criada.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p110">The first parameter of the `add` method is the range of only the top row of the table, not the entire range the table will ultimately use. This is because when the add-in populates the data rows (in the next step), it will add new rows to the table instead of writing values to the cells of existing rows. This is a more common pattern because the number of rows that a table will have is often not known when the table is created.</span></span>

    - <span data-ttu-id="f62e0-156">Os nomes de tabelas devem ser exclusivos pela pasta de trabalho inteira, não só na planilha.</span><span class="sxs-lookup"><span data-stu-id="f62e0-156">Table names must be unique across the entire workbook, not just the worksheet.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ```

10. <span data-ttu-id="f62e0-157">Na função `createTable()`, substitua `TODO2` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="f62e0-157">Within the `createTable()` function, replace `TODO2` with the following code.</span></span> <span data-ttu-id="f62e0-158">Observação:</span><span class="sxs-lookup"><span data-stu-id="f62e0-158">Note:</span></span>

    - <span data-ttu-id="f62e0-159">Os valores das células de um intervalo são definidos em uma matriz de matrizes.</span><span class="sxs-lookup"><span data-stu-id="f62e0-159">The cell values of a range are set with an array of arrays.</span></span>

    - <span data-ttu-id="f62e0-p112">Novas linhas são criadas em uma tabela ao chamar o método `add` do conjunto de linhas da tabela. Você pode adicionar várias linhas em uma única chamada de `add` ao incluir várias matrizes de valores de células na matriz pai que é passada como segundo parâmetro.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p112">New rows are created in a table by calling the `add` method of the table's row collection. You can add multiple rows in a single call of `add` by including multiple cell value arrays in the parent array that is passed as the second parameter.</span></span>

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

11. <span data-ttu-id="f62e0-162">Na função `createTable()`, substitua `TODO3` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="f62e0-162">Within the `createTable()` function, replace `TODO3` with the following code.</span></span> <span data-ttu-id="f62e0-163">Observação:</span><span class="sxs-lookup"><span data-stu-id="f62e0-163">Note:</span></span>

    - <span data-ttu-id="f62e0-164">O código recebe uma referência para a coluna **quantidade** ao passar o índice com base em zero para o método `getItemAt` do conjunto de colunas da tabela.</span><span class="sxs-lookup"><span data-stu-id="f62e0-164">The code gets a reference to the **Amount** column by passing its zero-based index to the `getItemAt` method of the table's column collection.</span></span>

        > [!NOTE]
        > <span data-ttu-id="f62e0-165">Os objetos do conjunto Excel.js, como `TableCollection`, `WorksheetCollection`, e `TableColumnCollection`, têm a propriedade `items` que é como uma matriz dos tipos de objetos filhos, como `Table` ou `Worksheet` ou `TableColumn`; mas um objeto `*Collection` não é uma matriz.</span><span class="sxs-lookup"><span data-stu-id="f62e0-165">Excel.js collection objects, such as `TableCollection`, `WorksheetCollection`, and `TableColumnCollection` have an `items` property that is an array of the child object types, such as `Table` or `Worksheet` or `TableColumn`; but a `*Collection` object is not itself an array.</span></span>

    - <span data-ttu-id="f62e0-166">O código formata o intervalo da coluna **quantidade** como Euros com um segundo decimal.</span><span class="sxs-lookup"><span data-stu-id="f62e0-166">The code then formats the range of the **Amount** column as Euros to the second decimal.</span></span> 

    - <span data-ttu-id="f62e0-p114">Por fim, isso garante que a largura das colunas e a altura das linhas sejam grandes o suficiente para o maior (ou o mais alto) item de dados. Observe que o código deve receber os objetos `Range` a formatar. Os objetos `TableColumn` e `TableRow` não têm propriedades de formato.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p114">Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item. Notice that the code must get `Range` objects to format. `TableColumn` and `TableRow` objects do not have format properties.</span></span>

    ```js
    expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();
    ```

12. <span data-ttu-id="f62e0-170">Verifique se você salvou todas as alterações feitas no projeto.</span><span class="sxs-lookup"><span data-stu-id="f62e0-170">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="f62e0-171">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="f62e0-171">Test the add-in</span></span>

1. <span data-ttu-id="f62e0-172">Conclua as etapas a seguir para iniciar o servidor Web local e fazer o sideload do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="f62e0-172">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f62e0-173">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="f62e0-173">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="f62e0-174">Se você for solicitado a instalar um certificado após executar um dos seguintes comandos, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="f62e0-174">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    > [!TIP]
    > <span data-ttu-id="f62e0-175">Se você estiver testando seu suplemento no Mac, execute o seguinte comando no diretório raiz do seu projeto antes de continuar.</span><span class="sxs-lookup"><span data-stu-id="f62e0-175">If you're testing your add-in on Mac, run the following command in the root directory of your project before proceeding.</span></span> <span data-ttu-id="f62e0-176">O servidor Web local é iniciado quando este comando é executado.</span><span class="sxs-lookup"><span data-stu-id="f62e0-176">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="f62e0-177">Para testar o seu suplemento no Excel, execute o seguinte comando no diretório raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="f62e0-177">To test your add-in in Excel, run the following command in the root directory of your project.</span></span> <span data-ttu-id="f62e0-178">Isso inicia o servidor Web local (caso ainda não esteja em execução) e abre o Excel com o suplemento carregado.</span><span class="sxs-lookup"><span data-stu-id="f62e0-178">This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="f62e0-179">Para testar seu suplemento no Excel na Web, execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="f62e0-179">To test your add-in in Excel on the web, run the following command in the root directory of your project.</span></span> <span data-ttu-id="f62e0-180">Quando você executar este comando, o servidor da Web local será iniciado (se ainda não estiver em execução).</span><span class="sxs-lookup"><span data-stu-id="f62e0-180">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="f62e0-181">Para usar o seu suplemento, abra um novo documento no Excel na Web e em seguida realize o sideload no suplemento de acordo com as instruções em [Realizar Sideload nos Suplementos do Office no Office na Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="f62e0-181">To use your add-in, open a new document in Excel on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

2. <span data-ttu-id="f62e0-182">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="f62e0-182">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-3b.png)

3. <span data-ttu-id="f62e0-184">No painel de tarefas, escolha o botão **Criar tabela**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-184">In the task pane, choose the **Create Table** button.</span></span>

    ![Tutorial do Excel: Criar tabela](../images/excel-tutorial-create-table-2.png)

## <a name="filter-and-sort-a-table"></a><span data-ttu-id="f62e0-186">Filtrar e classificar uma tabela</span><span class="sxs-lookup"><span data-stu-id="f62e0-186">Filter and sort a table</span></span>

<span data-ttu-id="f62e0-187">Nesta etapa do tutorial, você vai filtrar e classificar a tabela que criou anteriormente.</span><span class="sxs-lookup"><span data-stu-id="f62e0-187">In this step of the tutorial, you'll filter and sort the table that you created previously.</span></span>

### <a name="filter-the-table"></a><span data-ttu-id="f62e0-188">Filtrar a tabela</span><span class="sxs-lookup"><span data-stu-id="f62e0-188">Filter the table</span></span>

1. <span data-ttu-id="f62e0-189">Abra o arquivo **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-189">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="f62e0-190">Localize o elemento `<button>` para o botão `create-table` e adicione a seguinte marcação após essa linha:</span><span class="sxs-lookup"><span data-stu-id="f62e0-190">Locate the `<button>` element for the `create-table` button, and add the following markup after that line:</span></span>

    ```html
    <button class="ms-Button" id="filter-table">Filter Table</button><br/><br/>
    ```

3. <span data-ttu-id="f62e0-191">Abra o arquivo **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-191">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="f62e0-192">Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `create-table` e adicione o seguinte código após ela:</span><span class="sxs-lookup"><span data-stu-id="f62e0-192">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `create-table` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("filter-table").onclick = filterTable;
    ```

5. <span data-ttu-id="f62e0-193">Adicione a seguinte função ao final do arquivo:</span><span class="sxs-lookup"><span data-stu-id="f62e0-193">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="f62e0-194">Na função `filterTable()`, substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="f62e0-194">Within the `filterTable()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="f62e0-195">Observação:</span><span class="sxs-lookup"><span data-stu-id="f62e0-195">Note:</span></span>

   - <span data-ttu-id="f62e0-p120">O código primeiro faz referência à coluna que precisa de filtragem ao passar o nome da coluna para o método `getItem`, em vez de passar o índice para o método `getItemAt` como o método `createTable` faz. Como os usuários podem mover as colunas da tabela, a coluna de um determinado índice pode mudar depois da criação da tabela. Portanto, é mais seguro usar o nome da coluna como referência dela. Usamos de forma segura `getItemAt` em um tutorial anterior porque usamos o mesmo método que cria a tabela. Assim não existe a chance de um usuário mover a coluna.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p120">The code first gets a reference to the column that needs filtering by passing the column name to the `getItem` method, instead of passing its index to the `getItemAt` method as the `createTable` method does. Since users can move table columns, the column at a given index might change after the table is created. Hence, it is safer to use the column name to get a reference to the column. We used `getItemAt` safely in the preceding tutorial, because we used it in the very same method that creates the table, so there is no chance that a user has moved the column.</span></span>

   - <span data-ttu-id="f62e0-200">O método `applyValuesFilter` é um dos vários métodos de filtragem do objeto `Filter`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-200">The `applyValuesFilter` method is one of several filtering methods on the `Filter` object.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(['Education', 'Groceries']);
    ``` 

### <a name="sort-the-table"></a><span data-ttu-id="f62e0-201">Classificar a tabela</span><span class="sxs-lookup"><span data-stu-id="f62e0-201">Sort the table</span></span>

1. <span data-ttu-id="f62e0-202">Abra o arquivo **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-202">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="f62e0-203">Localize o elemento `<button>` para o botão `filter-table` e adicione a seguinte marcação após essa linha:</span><span class="sxs-lookup"><span data-stu-id="f62e0-203">Locate the `<button>` element for the `filter-table` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="sort-table">Sort Table</button><br/><br/>
    ```

3. <span data-ttu-id="f62e0-204">Abra o arquivo **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-204">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="f62e0-205">Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `filter-table` e adicione o seguinte código após ela:</span><span class="sxs-lookup"><span data-stu-id="f62e0-205">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `filter-table` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("sort-table").onclick = sortTable;
    ```

5. <span data-ttu-id="f62e0-206">Adicione a seguinte função ao final do arquivo:</span><span class="sxs-lookup"><span data-stu-id="f62e0-206">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="f62e0-207">Na função `sortTable()`, substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="f62e0-207">Within the `sortTable()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="f62e0-208">Observação:</span><span class="sxs-lookup"><span data-stu-id="f62e0-208">Note:</span></span>

   - <span data-ttu-id="f62e0-209">O código cria uma matriz de objetos `SortField` que tem apenas um membro, já que o suplemento só classifica a coluna Comerciante.</span><span class="sxs-lookup"><span data-stu-id="f62e0-209">The code creates an array of `SortField` objects which has just one member since the add-in only sorts on the Merchant column.</span></span>

   - <span data-ttu-id="f62e0-210">A propriedade `key` de um objeto `SortField` é o índice com base em zero da coluna a classificar.</span><span class="sxs-lookup"><span data-stu-id="f62e0-210">The `key` property of a `SortField` object is the zero-based index of the column to sort-on.</span></span>

   - <span data-ttu-id="f62e0-211">O membro `sort` de uma `Table` é um objeto `TableSort`, não um método.</span><span class="sxs-lookup"><span data-stu-id="f62e0-211">The `sort` member of a `Table` is a `TableSort` object, not a method.</span></span> <span data-ttu-id="f62e0-212">Os `SortField`s são passados para o método `apply` do objeto `TableSort`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-212">The `SortField`s are passed to the `TableSort` object's `apply` method.</span></span>

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

7. <span data-ttu-id="f62e0-213">Verifique se você salvou todas as alterações feitas no projeto.</span><span class="sxs-lookup"><span data-stu-id="f62e0-213">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="f62e0-214">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="f62e0-214">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="f62e0-215">Se o painel de tarefas do suplemento ainda não estiver aberto no Excel, vá para a guia **Página Inicial** e escolha o botão **Mostrar Painel de Tarefas** na faixa de opções para abri-lo.</span><span class="sxs-lookup"><span data-stu-id="f62e0-215">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="f62e0-216">Se a tabela que você adicionou anteriormente neste tutorial não estiver presente na planilha aberta, escolha o botão **Criar tabela** no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="f62e0-216">If the table you added previously in this tutorial is not present in the open worksheet, choose the **Create Table** button in the task pane.</span></span>

4. <span data-ttu-id="f62e0-217">Escolha os botões **Filtrar Tabela** e **Classificar Tabela**, em qualquer ordem.</span><span class="sxs-lookup"><span data-stu-id="f62e0-217">Choose the **Filter Table** button and the **Sort Table** button, in either order.</span></span>

    ![Tutorial do Excel: filtrar e classificar tabela](../images/excel-tutorial-filter-and-sort-table-2.png)

## <a name="create-a-chart"></a><span data-ttu-id="f62e0-219">Criar um gráfico</span><span class="sxs-lookup"><span data-stu-id="f62e0-219">Create a chart</span></span>

<span data-ttu-id="f62e0-220">Nesta etapa do tutorial, você vai criar um gráfico com dados da tabela que você criou anteriormente e depois vai formatar o gráfico.</span><span class="sxs-lookup"><span data-stu-id="f62e0-220">In this step of the tutorial, you'll create a chart using data from the table that you created previously, and then format the chart.</span></span>

### <a name="chart-a-chart-using-table-data"></a><span data-ttu-id="f62e0-221">Gráfico de um gráfico com dados de tabela</span><span class="sxs-lookup"><span data-stu-id="f62e0-221">Chart a chart using table data</span></span>

1. <span data-ttu-id="f62e0-222">Abra o arquivo **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-222">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="f62e0-223">Localize o elemento `<button>` para o botão `sort-table` e adicione a seguinte marcação após essa linha:</span><span class="sxs-lookup"><span data-stu-id="f62e0-223">Locate the `<button>` element for the `sort-table` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="create-chart">Create Chart</button><br/><br/>
    ```

3. <span data-ttu-id="f62e0-224">Abra o arquivo **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-224">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="f62e0-225">Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `sort-table` e adicione o seguinte código após ela:</span><span class="sxs-lookup"><span data-stu-id="f62e0-225">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `sort-table` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("create-chart").onclick = createChart;
    ```

5. <span data-ttu-id="f62e0-226">Adicione a seguinte função ao final do arquivo:</span><span class="sxs-lookup"><span data-stu-id="f62e0-226">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="f62e0-227">Na função `createChart()`, substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="f62e0-227">Within the `createChart()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="f62e0-228">Para excluir a linha de cabeçalho, o código usa o método `Table.getDataBodyRange` para acessar o intervalo de dados que você deseja representar graficamente em vez do método `getRange`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-228">Note that in order to exclude the header row, the code uses the `Table.getDataBodyRange` method to get the range of data you want to chart instead of the `getRange` method.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var dataRange = expensesTable.getDataBodyRange();
    ```

7. <span data-ttu-id="f62e0-229">Na função `createChart()`, substitua `TODO2` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="f62e0-229">Within the `createChart()` function, replace `TODO2` with the following code.</span></span> <span data-ttu-id="f62e0-230">Observe os seguintes parâmetros:</span><span class="sxs-lookup"><span data-stu-id="f62e0-230">Note the following parameters:</span></span>

   - <span data-ttu-id="f62e0-p125">O primeiro parâmetro para o método `add` especifica o tipo de gráfico. Há diversos tipos.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p125">The first parameter to the `add` method specifies the type of chart. There are several dozen types.</span></span>

   - <span data-ttu-id="f62e0-233">O segundo parâmetro especifica um intervalo de dados a incluir no gráfico.</span><span class="sxs-lookup"><span data-stu-id="f62e0-233">The second parameter specifies the range of data to include in the chart.</span></span>

   - <span data-ttu-id="f62e0-234">O terceiro parâmetro determina se uma série de pontos de dados da tabela deve estar representada por linha ou por coluna.</span><span class="sxs-lookup"><span data-stu-id="f62e0-234">The third parameter determines whether a series of data points from the table should be charted row-wise or column-wise.</span></span> <span data-ttu-id="f62e0-235">A opção `auto` informa ao Excel para decidir o melhor método.</span><span class="sxs-lookup"><span data-stu-id="f62e0-235">The option `auto` tells Excel to decide the best method.</span></span>

    ```js
    var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ```

8. <span data-ttu-id="f62e0-236">Na função `createChart()`, substitua `TODO3` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="f62e0-236">Within the `createChart()` function, replace `TODO3` with the following code.</span></span> <span data-ttu-id="f62e0-237">A maior parte do código é autoexplicativa.</span><span class="sxs-lookup"><span data-stu-id="f62e0-237">Most of this code is self-explanatory.</span></span> <span data-ttu-id="f62e0-238">Observação:</span><span class="sxs-lookup"><span data-stu-id="f62e0-238">Note:</span></span>
   
   - <span data-ttu-id="f62e0-p128">Os parâmetros do método `setPosition` especificam as células da esquerda superior e da direita inferior da área da planilha que deve conter o gráfico. O Excel ajusta detalhes como a largura da linha para criar uma boa aparência para o gráfico no espaço fornecido.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p128">The parameters to the `setPosition` method specify the upper left and lower right cells of the worksheet area that should contain the chart. Excel can adjust things like line width to make the chart look good in the space it has been given.</span></span>
   
   - <span data-ttu-id="f62e0-p129">"Série" é um conjunto de pontos de dados de uma coluna da tabela. Como há apenas uma coluna sem cadeia de caracteres na tabela, o Excel deduz que essa é a única coluna de pontos de dados no gráfico. Ele interpreta outras colunas como rótulos do gráfico. Portanto, haverá apenas uma série no gráfico e será necessário o índice 0. Ele será rotulado como "Valor em €".</span><span class="sxs-lookup"><span data-stu-id="f62e0-p129">A "series" is a set of data points from a column of the table. Since there is only one non-string column in the table, Excel infers that the column is the only column of data points to chart. It interprets the other columns as chart labels. So there will be just one series in the chart and it will have index 0. This is the one to label with "Value in €".</span></span>

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in €';
    ```

9. <span data-ttu-id="f62e0-246">Verifique se você salvou todas as alterações feitas no projeto.</span><span class="sxs-lookup"><span data-stu-id="f62e0-246">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="f62e0-247">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="f62e0-247">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="f62e0-248">Se o painel de tarefas do suplemento ainda não estiver aberto no Excel, vá para a guia **Página Inicial** e escolha o botão **Mostrar Painel de Tarefas** na faixa de opções para abri-lo.</span><span class="sxs-lookup"><span data-stu-id="f62e0-248">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="f62e0-249">Se a tabela que você adicionou anteriormente neste tutorial não estiver presente na planilha aberta, escolha o botão **Criar tabela** e depois os botões **Filtrar Tabela** e **Classificar Tabela**, em qualquer ordem.</span><span class="sxs-lookup"><span data-stu-id="f62e0-249">If the table you added previously in this tutorial is not present in the open worksheet, choose the **Create Table** button, and then the **Filter Table** button and the **Sort Table** button, in either order.</span></span>

4. <span data-ttu-id="f62e0-p130">Clique no botão **Criar gráfico**. Um gráfico é criado e incluirá somente os dados das linhas que foram filtradas. Os rótulos dos pontos de dados na parte inferior estão na ordem de classificação do gráfico, ou seja, nomes de comerciantes em ordem alfabética inversa.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p130">Choose the **Create Chart** button. A chart is created and only the data from the rows that have been filtered are included. The labels on the data points across the bottom are in the sort order of the chart; that is, merchant names in reverse alphabetical order.</span></span>

    ![Tutorial do Excel - Criar gráfico ](../images/excel-tutorial-create-chart-2.png)

## <a name="freeze-a-table-header"></a><span data-ttu-id="f62e0-254">Congelar um cabeçalho de tabela</span><span class="sxs-lookup"><span data-stu-id="f62e0-254">Freeze a table header</span></span>

<span data-ttu-id="f62e0-p131">Quando uma tabela for longa o suficiente para que um usuário precise rolar para ver algumas linhas, a linha de cabeçalho poderá ficar fora da vista. Nesta etapa do tutorial, você precisará congelar a linha do cabeçalho da tabela que criou anteriormente para que ela permaneça visível, mesmo que o usuário role ao longo da planilha.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p131">When a table is long enough that a user must scroll to see some rows, the header row can scroll out of sight. In this step of the tutorial, you'll freeze the header row of the table that you created previously, so that it remains visible even as the user scrolls down the worksheet.</span></span>

### <a name="freeze-the-tables-header-row"></a><span data-ttu-id="f62e0-257">Congelar a linha de cabeçalho da tabela</span><span class="sxs-lookup"><span data-stu-id="f62e0-257">Freeze the table's header row</span></span>

1. <span data-ttu-id="f62e0-258">Abra o arquivo **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-258">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="f62e0-259">Localize o elemento `<button>` para o botão `create-chart` e adicione a seguinte marcação após essa linha:</span><span class="sxs-lookup"><span data-stu-id="f62e0-259">Locate the `<button>` element for the `create-chart` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="freeze-header">Freeze Header</button><br/><br/>
    ```

3. <span data-ttu-id="f62e0-260">Abra o arquivo **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-260">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="f62e0-261">Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `create-chart` e adicione o seguinte código após ela:</span><span class="sxs-lookup"><span data-stu-id="f62e0-261">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `create-chart` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("freeze-header").onclick = freezeHeader;
    ```

5. <span data-ttu-id="f62e0-262">Adicione a seguinte função ao final do arquivo:</span><span class="sxs-lookup"><span data-stu-id="f62e0-262">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="f62e0-263">Na função `freezeHeader()`, substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="f62e0-263">Within the `freezeHeader()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="f62e0-264">Observação:</span><span class="sxs-lookup"><span data-stu-id="f62e0-264">Note:</span></span>

   - <span data-ttu-id="f62e0-265">A coleção `Worksheet.freezePanes` é um conjunto de painéis da planilha que fica congelado ou fixado no mesmo lugar quando rolamos a planilha.</span><span class="sxs-lookup"><span data-stu-id="f62e0-265">The `Worksheet.freezePanes` collection is a set of panes in the worksheet that are pinned, or frozen, in place when the worksheet is scrolled.</span></span>

   - <span data-ttu-id="f62e0-p133">O método `freezeRows` considera como parâmetro o número de linhas, começando da parte superior, que devem ser fixadas no local. Passamos `1` para fixar a primeira linha no local.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p133">The `freezeRows` method takes as a parameter the number of rows, from the top that are to be pinned in place. We pass `1` to pin the first row in place.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ```

7. <span data-ttu-id="f62e0-268">Verifique se você salvou todas as alterações feitas no projeto.</span><span class="sxs-lookup"><span data-stu-id="f62e0-268">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="f62e0-269">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="f62e0-269">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="f62e0-270">Se o painel de tarefas do suplemento ainda não estiver aberto no Excel, vá para a guia **Página Inicial** e escolha o botão **Mostrar Painel de Tarefas** na faixa de opções para abri-lo.</span><span class="sxs-lookup"><span data-stu-id="f62e0-270">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="f62e0-271">Se a tabela que você adicionou anteriormente neste tutorial estiver presente na planilha, faça a exclusão dela.</span><span class="sxs-lookup"><span data-stu-id="f62e0-271">If the table you added previously in this tutorial is present in the worksheet, delete it.</span></span>

4. <span data-ttu-id="f62e0-272">No painel de tarefas, escolha o botão **Criar tabela**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-272">In the task pane, choose the **Create Table** button.</span></span>

5. <span data-ttu-id="f62e0-273">No painel de tarefas, escolha o botão **Congelar Cabeçalho**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-273">In the task pane, choose the **Freeze Header** button.</span></span>

6. <span data-ttu-id="f62e0-274">Role a planilha para baixo o suficiente para ver que o cabeçalho da tabela permanece visível na parte superior mesmo ao rolar até que as primeiras linhas fiquem fora da vista.</span><span class="sxs-lookup"><span data-stu-id="f62e0-274">Scroll down the worksheet far enough to see that the table header remains visible at the top even when the higher rows scroll out of sight.</span></span>

    ![Tutorial do Excel: congelar cabeçalho](../images/excel-tutorial-freeze-header-2.png)

## <a name="protect-a-worksheet"></a><span data-ttu-id="f62e0-276">Proteger uma planilha</span><span class="sxs-lookup"><span data-stu-id="f62e0-276">Protect a worksheet</span></span>

<span data-ttu-id="f62e0-277">Nesta etapa do tutorial, você adicionará outro botão à faixa de opções que, quando selecionado, executa uma função que você precisará definir para ativar e desativar a proteção da planilha.</span><span class="sxs-lookup"><span data-stu-id="f62e0-277">In this step of the tutorial, you'll add another button to the ribbon that, when chosen, executes a function that you'll define to toggle worksheet protection on and off.</span></span>

### <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a><span data-ttu-id="f62e0-278">Configure o manifesto para adicionar um segundo botão à faixa de opções</span><span class="sxs-lookup"><span data-stu-id="f62e0-278">Configure the manifest to add a second ribbon button</span></span>

1. <span data-ttu-id="f62e0-279">Abra o arquivo de manifesto **./manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-279">Open the manifest file **./manifest.xml**.</span></span>

2. <span data-ttu-id="f62e0-280">Localize o elemento `<Control>`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-280">Locate the `<Control>` element.</span></span> <span data-ttu-id="f62e0-281">Esse elemento define o botão **Mostrar Painel de Tarefas** na faixa de opções **Início** que você usa para iniciar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="f62e0-281">This element defines the **Show Taskpane** button on the **Home** ribbon you have been using to launch the add-in.</span></span> <span data-ttu-id="f62e0-282">Vamos adicionar um segundo botão ao mesmo grupo na faixa de opções **Início**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-282">We're going to add a second button to the same group on the **Home** ribbon.</span></span> <span data-ttu-id="f62e0-283">Entre a marca de Controle final (`</Control>`) e a marca de Grupo final (`</Group>`), adicione a marcação a seguir.</span><span class="sxs-lookup"><span data-stu-id="f62e0-283">In between the end Control tag (`</Control>`) and the end Group tag (`</Group>`), add the following markup.</span></span>

    ```xml
    <Control xsi:type="Button" id="<!--TODO1: Unique (in manifest) name for button -->">
        <Label resid="<!--TODO2: Button label -->" />
        <Supertip>            
            <Title resid="<!-- TODO3: Button tool tip title -->" />
            <Description resid="<!-- TODO4: Button tool tip description -->" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Icon.16x16"/>
            <bt:Image size="32" resid="Icon.32x32"/>
            <bt:Image size="80" resid="Icon.80x80"/>
        </Icon>
        <Action xsi:type="<!-- TODO5: Specify the type of action-->">
            <!-- TODO6: Identify the function.-->
        </Action>
    </Control>
    ```

3. <span data-ttu-id="f62e0-284">No XML que você acabou de adicionar ao arquivo de manifesto, substitua `TODO1` por uma sequência que forneça ao botão um ID exclusivo nesse arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="f62e0-284">Within the XML you just added to the manifest file, replace `TODO1` with a string that gives the button an ID that is unique within this manifest file.</span></span> <span data-ttu-id="f62e0-285">Como nosso botão ativará ou desativará a proteção da planilha, use "ToggleProtection".</span><span class="sxs-lookup"><span data-stu-id="f62e0-285">Since our button is going to toggle protection of the worksheet on and off, use "ToggleProtection".</span></span> <span data-ttu-id="f62e0-286">Quando você terminar, a marca de abertura para o elemento `Control` deverá ficar assim:</span><span class="sxs-lookup"><span data-stu-id="f62e0-286">When you are done, the opening tag for the `Control` element should look like this:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. <span data-ttu-id="f62e0-287">Os próximos três `TODO`s definem “resid”, que significa ID de recurso.</span><span class="sxs-lookup"><span data-stu-id="f62e0-287">The next three `TODO`s set "resid"s, which is short for resource ID.</span></span> <span data-ttu-id="f62e0-288">Um recurso é uma cadeia de caracteres e você criará essas três cadeias de caracteres em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="f62e0-288">A resource is a string, and you'll create these three strings in a later step.</span></span> <span data-ttu-id="f62e0-289">Por enquanto, você precisa fornecer IDs aos recursos.</span><span class="sxs-lookup"><span data-stu-id="f62e0-289">For now, you need to give IDs to the resources.</span></span> <span data-ttu-id="f62e0-290">O rótulo do botão deve ser "Toggle Protection", mas o *ID* dessa cadeia de caracteres deve ser "ProtectionButtonLabel", para que o elemento `Label` fique assim:</span><span class="sxs-lookup"><span data-stu-id="f62e0-290">The button label should read "Toggle Protection", but the *ID* of this string should be "ProtectionButtonLabel", so the `Label` element should look like this:</span></span>

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. <span data-ttu-id="f62e0-291">O elemento `SuperTip` define a dica de ferramenta do botão.</span><span class="sxs-lookup"><span data-stu-id="f62e0-291">The `SuperTip` element defines the tool tip for the button.</span></span> <span data-ttu-id="f62e0-292">O título da dica de ferramenta deve ser o mesmo que o rótulo do botão, por isso, usamos a mesma ID de recurso: "ProtectionButtonLabel".</span><span class="sxs-lookup"><span data-stu-id="f62e0-292">The tool tip title should be the same as the button label, so we use the very same resource ID: "ProtectionButtonLabel".</span></span> <span data-ttu-id="f62e0-293">A descrição da dica de ferramenta será "Click to turn protection of the worksheet on and off".</span><span class="sxs-lookup"><span data-stu-id="f62e0-293">The tool tip description will be "Click to turn protection of the worksheet on and off".</span></span> <span data-ttu-id="f62e0-294">Mas o `ID` será "ProtectionButtonToolTip".</span><span class="sxs-lookup"><span data-stu-id="f62e0-294">But the `ID` should be "ProtectionButtonToolTip".</span></span> <span data-ttu-id="f62e0-295">Portanto, quando você terminar, o elemento `SuperTip` deverá ficar assim:</span><span class="sxs-lookup"><span data-stu-id="f62e0-295">So, when you are done, the `SuperTip` element should look like this:</span></span> 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > <span data-ttu-id="f62e0-p138">Em um suplemento de produção,não é recomendável usar o mesmo ícone para dois botões diferentes; mas, para simplificar este tutorial, faremos isso. Portanto, a marcação `Icon` em nosso novo `Control` é apenas uma cópia do elemento `Icon` do `Control` existente.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p138">In a production add-in, you would not want to use the same icon for two different buttons; but to simplify this tutorial, we'll do that. So the `Icon` markup in our new `Control` is just a copy of the `Icon` element from the existing `Control`.</span></span> 

6. <span data-ttu-id="f62e0-298">O elemento `Action` dentro do elemento `Control` original já está presente no manifesto, tem seu tipo definido como `ShowTaskpane`, mas nosso novo botão não abrirá um painel de tarefas, mas sim executará uma função personalizada criada em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="f62e0-298">The `Action` element inside the original `Control` element that was already present in the manifest, has its type set to `ShowTaskpane`, but our new button isn't going to open a task pane; it's going to run a custom function that you create in a later step.</span></span> <span data-ttu-id="f62e0-299">Portanto, substitua `TODO5` por `ExecuteFunction`, que é o tipo de ação para botões que acionam funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="f62e0-299">So replace `TODO5` with `ExecuteFunction` which is the action type for buttons that trigger custom functions.</span></span> <span data-ttu-id="f62e0-300">A marca de abertura para o elemento `Action` deve ficar assim:</span><span class="sxs-lookup"><span data-stu-id="f62e0-300">The opening tag for the `Action` element should look like this:</span></span>
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. <span data-ttu-id="f62e0-p140">O elemento `Action` original tem elementos filhos que especificam uma ID do painel de tarefas e uma URL da página que deve ser aberta no painel de tarefas. No entanto, um elemento `Action` do tipo `ExecuteFunction` tem um único elemento filho que nomeia a função executada pelo controle. Você criará essa função em uma etapa posterior e ela será chamada de `toggleProtection`. Então, substitua `TODO6` pela marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="f62e0-p140">The original `Action` element has child elements that specify a task pane ID and a URL of the page that should be opened in the task pane. But an `Action` element of the `ExecuteFunction` type has a single child element that names the function that the control executes. You'll create that function in a later step, and it will be called `toggleProtection`. So, replace `TODO6` with the following markup:</span></span>
 
    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    <span data-ttu-id="f62e0-305">A marcação `Control` inteira deve ter a aparência a seguir:</span><span class="sxs-lookup"><span data-stu-id="f62e0-305">The entire `Control` markup should now look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
        <Label resid="ProtectionButtonLabel" />
        <Supertip>            
            <Title resid="ProtectionButtonLabel" />
            <Description resid="ProtectionButtonToolTip" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Icon.16x16"/>
            <bt:Image size="32" resid="Icon.32x32"/>
            <bt:Image size="80" resid="Icon.80x80"/>
        </Icon>
        <Action xsi:type="ExecuteFunction">
           <FunctionName>toggleProtection</FunctionName>
        </Action>
    </Control>
    ```

8. <span data-ttu-id="f62e0-306">Role para baixo até a seção `Resources` do manifesto.</span><span class="sxs-lookup"><span data-stu-id="f62e0-306">Scroll down to the `Resources` section of the manifest.</span></span>

9. <span data-ttu-id="f62e0-307">Adicione a seguinte marcação como filho do elemento `bt:ShortStrings`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-307">Add the following markup as a child of the `bt:ShortStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. <span data-ttu-id="f62e0-308">Adicione a seguinte marcação como filho do elemento `bt:LongStrings`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-308">Add the following markup as a child of the `bt:LongStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. <span data-ttu-id="f62e0-309">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="f62e0-309">Save the file.</span></span>

### <a name="create-the-function-that-protects-the-sheet"></a><span data-ttu-id="f62e0-310">Criar a função que protege a planilha</span><span class="sxs-lookup"><span data-stu-id="f62e0-310">Create the function that protects the sheet</span></span>

1. <span data-ttu-id="f62e0-311">Abra o arquivo **.\commands\commands.js**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-311">Open the file **.\commands\commands.js**.</span></span>

2. <span data-ttu-id="f62e0-312">Adicione a seguinte função imediatamente após a função `action`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-312">Add the following function immediately after the `action` function.</span></span> <span data-ttu-id="f62e0-313">Especificamos um parâmetro `args` para a função, e a última linha da função chama `args.completed`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-313">Note that we specify an `args` parameter to the function and the very last line of the function calls `args.completed`.</span></span> <span data-ttu-id="f62e0-314">Esse é um requisito para todos os comandos de suplemento do tipo **ExecuteFunction**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-314">This is a requirement for all add-in commands of type **ExecuteFunction**.</span></span> <span data-ttu-id="f62e0-315">Ele sinaliza para o aplicativo host do Office que a função terminou e que a interface do usuário podem ficar responsiva novamente.</span><span class="sxs-lookup"><span data-stu-id="f62e0-315">It signals the Office host application that the function has finished and the UI can become responsive again.</span></span>

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

3. <span data-ttu-id="f62e0-316">Adicione a seguinte linha ao final do arquivo:</span><span class="sxs-lookup"><span data-stu-id="f62e0-316">Add the following line to the end of the file:</span></span>

    ```js
    g.toggleProtection = toggleProtection;
    ```

4. <span data-ttu-id="f62e0-317">Na função `toggleProtection`, substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="f62e0-317">Within the `toggleProtection` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="f62e0-318">O código usa propriedade de proteção do objeto de planilha em um padrão de botão de alternância padrão.</span><span class="sxs-lookup"><span data-stu-id="f62e0-318">This code uses the worksheet object's protection property in a standard toggle pattern.</span></span> <span data-ttu-id="f62e0-319">O `TODO2` será explicado na próxima seção.</span><span class="sxs-lookup"><span data-stu-id="f62e0-319">The `TODO2` will be explained in the next section.</span></span>

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

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="f62e0-320">Adicione código para buscar propriedades do documento em objetos de script do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="f62e0-320">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="f62e0-321">Em cada função criada neste tutorial até agora, você enfileirou comandos para *gravar* no documento do Office.</span><span class="sxs-lookup"><span data-stu-id="f62e0-321">In each function that you've created in this tutorial until now, you queued commands to *write* to the Office document.</span></span> <span data-ttu-id="f62e0-322">Cada função terminou com uma chamada para o método `context.sync()`, que envia os comandos em fila para o documento a ser executado.</span><span class="sxs-lookup"><span data-stu-id="f62e0-322">Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed.</span></span> <span data-ttu-id="f62e0-323">Entretanto, o código adicionado na última etapa chama a propriedade `sheet.protection.protected` e essa é uma grande diferença das funções anteriores que você escreveu, pois o objeto `sheet` é apenas um objeto de proxy que existe no script do seu painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="f62e0-323">But the code you added in the last step calls the `sheet.protection.protected` property, and this is a significant difference from the earlier functions you wrote, because the `sheet` object is only a proxy object that exists in your task pane's script.</span></span> <span data-ttu-id="f62e0-324">Ele não sabe qual é o estado real de proteção do documento, portanto, sua propriedade `protection.protected` não pode ter um valor real.</span><span class="sxs-lookup"><span data-stu-id="f62e0-324">It doesn't know what the actual protection state of the document is, so its `protection.protected` property can't have a real value.</span></span> <span data-ttu-id="f62e0-325">É necessário primeiro buscar o status de proteção do documento e definir o valor de `sheet.protection.protected`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-325">It is necessary to first fetch the protection status from the document and use it set the value of `sheet.protection.protected`.</span></span> <span data-ttu-id="f62e0-326">Somente então será possível chamar `sheet.protection.protected` sem causar uma exceção.</span><span class="sxs-lookup"><span data-stu-id="f62e0-326">Only then can `sheet.protection.protected` be called without causing an exception to be thrown.</span></span> <span data-ttu-id="f62e0-327">Esse processo de busca tem três etapas:</span><span class="sxs-lookup"><span data-stu-id="f62e0-327">This fetching process has three steps:</span></span>

   1. <span data-ttu-id="f62e0-328">Coloque em fila um comando para carregar (ou seja, fetch) as propriedades que seu código precisa ler.</span><span class="sxs-lookup"><span data-stu-id="f62e0-328">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>

   2. <span data-ttu-id="f62e0-329">Chame o método `sync` do objeto de contexto para enviar o comando em fila para o documento para execução e retornar as informações solicitadas.</span><span class="sxs-lookup"><span data-stu-id="f62e0-329">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>

   3. <span data-ttu-id="f62e0-330">Como o método `sync` é assíncrono, certifique-se de que ele tenha sido concluído antes que o código chame as propriedades que foram buscadas.</span><span class="sxs-lookup"><span data-stu-id="f62e0-330">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="f62e0-331">Essas etapas devem ser concluídas sempre que seu código precisar *ler* informações do documento do Office.</span><span class="sxs-lookup"><span data-stu-id="f62e0-331">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="f62e0-332">Na função `toggleProtection`, substitua `TODO2` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="f62e0-332">Within the `toggleProtection` function, replace `TODO2` with the following code.</span></span> <span data-ttu-id="f62e0-333">Observação:</span><span class="sxs-lookup"><span data-stu-id="f62e0-333">Note:</span></span>
   
   - <span data-ttu-id="f62e0-p145">Todos os objetos do Excel têm um método `load`. Especifique as propriedades do objeto que você deseja ler no parâmetro como uma cadeia de caracteres de nomes delimitados por vírgulas. Nesse caso, a propriedade que você precisa ler é uma subpropriedade de `protection`. Referencie a subpropriedade quase exatamente como você faria em qualquer lugar do seu código, mas usando uma barra (“/”) em vez de um ponto (".").</span><span class="sxs-lookup"><span data-stu-id="f62e0-p145">Every Excel object has a `load` method. You specify the properties of the object that you want to read in the parameter as a string of comma-delimited names. In this case, the property you need to read is a subproperty of the `protection` property. You reference the subproperty almost exactly as you would anywhere else in your code, with the exception that you use a forward slash ('/') character instead of a "." character.</span></span>

   - <span data-ttu-id="f62e0-338">Para garantir que a lógica de botão de alternância, `sheet.protection.protected`, não seja executada até após `sync` ser concluído e o `sheet.protection.protected` ser atribuída ao valor correto buscado no documento, ele será movido (na próxima etapa) para uma função `then` que não será executada até `sync` ser concluído.</span><span class="sxs-lookup"><span data-stu-id="f62e0-338">To ensure that the toggle logic, which reads `sheet.protection.protected`, does not run until after the `sync` is complete and the `sheet.protection.protected` has been assigned the correct value that is fetched from the document, it will be moved (in the next step) into a `then` function that won't run until the `sync` has completed.</span></span> 

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

2. <span data-ttu-id="f62e0-p146">Você não pode ter duas instruções `return` no mesmo caminho de código sem ramificações, portanto, exclua a linha final `return context.sync();` no final de `Excel.run`. Você adicionará um novo `context.sync` final em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p146">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Excel.run`. You will add a new final `context.sync`, in a later step.</span></span>

3. <span data-ttu-id="f62e0-341">Recorte a estrutura `if ... else` na função `toggleProtection` e a cole no lugar de `TODO3`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-341">Cut the `if ... else` structure in the `toggleProtection` function and paste it in place of `TODO3`.</span></span>

4. <span data-ttu-id="f62e0-p147">Substitua `TODO4` pelo código a seguir. Observação:</span><span class="sxs-lookup"><span data-stu-id="f62e0-p147">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="f62e0-344">Passar o método `sync` para uma função `then` garante que ele não seja executado até que `sheet.protection.unprotect()` ou `sheet.protection.protect()` seja enfileirado.</span><span class="sxs-lookup"><span data-stu-id="f62e0-344">Passing the `sync` method to a `then` function ensures that it does not run until either `sheet.protection.unprotect()` or `sheet.protection.protect()` has been queued.</span></span>

   - <span data-ttu-id="f62e0-345">O método `then` invoca qualquer função que é passada para ele e não é recomendável que `sync` seja chamado duas vezes, portanto, remova os “()” do fim de `context.sync`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-345">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of `context.sync`.</span></span>

    ```js
    .then(context.sync);
    ```

   <span data-ttu-id="f62e0-346">Quando terminar, a função inteira deve se parecer com o seguinte:</span><span class="sxs-lookup"><span data-stu-id="f62e0-346">When you are done, the entire function should look like the following:</span></span>

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

5. <span data-ttu-id="f62e0-347">Verifique se você salvou todas as alterações feitas no projeto.</span><span class="sxs-lookup"><span data-stu-id="f62e0-347">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="f62e0-348">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="f62e0-348">Test the add-in</span></span>

1. <span data-ttu-id="f62e0-349">Feche todos os aplicativos do Office, incluindo o Excel.</span><span class="sxs-lookup"><span data-stu-id="f62e0-349">Close all Office applications, including Excel.</span></span> 

2. <span data-ttu-id="f62e0-p148">Para excluir o cache do Office, exclua o conteúdo da pasta de cache. Isso é necessário para limpar totalmente a versão anterior do suplemento do host.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p148">Delete the Office cache by deleting the contents of the cache folder. This is necessary to completely clear the old version of the add-in from the host.</span></span> 

    - <span data-ttu-id="f62e0-352">No Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-352">For Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

    - <span data-ttu-id="f62e0-353">No Mac: `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-353">For Mac: `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span> 
    
        > [!NOTE]
        > <span data-ttu-id="f62e0-354">Se essa pasta não existir, verifique as seguintes pastas. Se encontrada, exclua o conteúdo da pasta:</span><span class="sxs-lookup"><span data-stu-id="f62e0-354">If that folder doesn't exist, check for the following folders and if found, delete the contents of the folder:</span></span>
        >    - <span data-ttu-id="f62e0-355">`~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/` em que `{host}` é o host do Office (por exemplo, `Excel`)</span><span class="sxs-lookup"><span data-stu-id="f62e0-355">`~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/` where `{host}` is the Office host (e.g., `Excel`)</span></span>
        >    - <span data-ttu-id="f62e0-356">`~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/` em que `{host}` é o host do Office (por exemplo, `Excel`)</span><span class="sxs-lookup"><span data-stu-id="f62e0-356">`~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/` where `{host}` is the Office host (e.g., `Excel`)</span></span>
        >    - `com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`

3. <span data-ttu-id="f62e0-357">Se o servidor da Web local já estiver em execução, feche a janela de comando do nó para interrompê-lo.</span><span class="sxs-lookup"><span data-stu-id="f62e0-357">If the local web server is already running, stop it by closing the node command window.</span></span>

4. <span data-ttu-id="f62e0-358">Como o arquivo de manifesto foi atualizado, você deve carregar o suplemento novamente usando esse arquivo.</span><span class="sxs-lookup"><span data-stu-id="f62e0-358">Because your manifest file has been updated, you must sideload your add-in again, using the updated manifest file.</span></span> <span data-ttu-id="f62e0-359">Inicie o servidor Web local e realize o sideload no seu suplemento:</span><span class="sxs-lookup"><span data-stu-id="f62e0-359">Start the local web server and sideload your add-in:</span></span> 

    - <span data-ttu-id="f62e0-360">Para testar o seu suplemento no Excel, execute o seguinte comando no diretório raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="f62e0-360">To test your add-in in Excel, run the following command in the root directory of your project.</span></span> <span data-ttu-id="f62e0-361">Isso inicia o servidor Web local (caso ainda não esteja em execução) e abre o Excel com o suplemento carregado.</span><span class="sxs-lookup"><span data-stu-id="f62e0-361">This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="f62e0-362">Para testar seu suplemento no Excel na Web, execute o seguinte comando no diretório raiz do seu projeto.</span><span class="sxs-lookup"><span data-stu-id="f62e0-362">To test your add-in in Excel on the web, run the following command in the root directory of your project.</span></span> <span data-ttu-id="f62e0-363">Quando você executar este comando, o servidor da Web local será iniciado (se ainda não estiver em execução).</span><span class="sxs-lookup"><span data-stu-id="f62e0-363">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="f62e0-364">Para usar o seu suplemento, abra um novo documento no Excel na Web e em seguida realize o sideload no suplemento de acordo com as instruções em [Realizar Sideload nos Suplementos do Office no Office na Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="f62e0-364">To use your add-in, open a new document in Excel on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

5. <span data-ttu-id="f62e0-365">Na guia **Página Inicial** no Excel, escolha o botão **Proteger Planilha**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-365">On the **Home** tab in Excel, choose the **Toggle Worksheet Protection** button.</span></span> <span data-ttu-id="f62e0-366">A maioria dos controles da faixa de opções está desabilitada e esmaecida, como mostra a captura de tela a seguir.</span><span class="sxs-lookup"><span data-stu-id="f62e0-366">Note that most of the controls on the ribbon are disabled (and visually grayed-out) as seen in the following screenshot.</span></span> 

    ![Tutorial do Excel - Faixa de Opções com a Proteção Ativada](../images/excel-tutorial-ribbon-with-protection-on-2.png)

6. <span data-ttu-id="f62e0-368">Escolha uma célula como se quisesse alterar o conteúdo.</span><span class="sxs-lookup"><span data-stu-id="f62e0-368">Choose a cell as you would if you wanted to change its content.</span></span> <span data-ttu-id="f62e0-369">O Excel exibe uma mensagem de erro indicando que a planilha está protegida.</span><span class="sxs-lookup"><span data-stu-id="f62e0-369">Excel displays an error message indicating that the worksheet is protected.</span></span>

7. <span data-ttu-id="f62e0-370">Escolha o botão **Proteger Planilha** novamente. Os controles são reabilitados, e você pode alterar os valores das células novamente.</span><span class="sxs-lookup"><span data-stu-id="f62e0-370">Choose the **Toggle Worksheet Protection** button again, and the controls are reenabled, and you can change cell values again.</span></span>

## <a name="open-a-dialog"></a><span data-ttu-id="f62e0-371">Abrir uma caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="f62e0-371">Open a dialog</span></span>

<span data-ttu-id="f62e0-p154">Nesta etapa final do tutorial, você abre uma caixa de diálogo no suplemento, passa uma mensagem do processo de caixa de diálogo para o processo de painel de tarefas e fecha a caixa de diálogo. As caixas de diálogo do Suplemento do Office são *não modais*: o usuário pode continuar a interagir com o documento no aplicativo do Office do host e com a página host no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p154">In this final step of the tutorial, you'll open a dialog in your add-in, pass a message from the dialog process to the task pane process, and close the dialog. Office Add-in dialogs are *nonmodal*: a user can continue to interact with both the document in the host Office application and with the host page in the task pane.</span></span>

### <a name="create-the-dialog-page"></a><span data-ttu-id="f62e0-374">Crie a página da caixa de diálogo</span><span class="sxs-lookup"><span data-stu-id="f62e0-374">Create the dialog page</span></span>

1. <span data-ttu-id="f62e0-375">Na pasta **./src** localizada na raiz do projeto, crie uma pasta chamada **dialogs**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-375">In the **./src** folder that's located at the root of the project, create a new folder named **dialogs**.</span></span>

2. <span data-ttu-id="f62e0-376">Na pasta **./src/dialogs**, crie um novo arquivo chamado **popup.html**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-376">In the **./src/dialogs** folder, create new file named **popup.html**.</span></span>

3. <span data-ttu-id="f62e0-377">Adicione a seguinte marcação a **popup.html**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-377">Add the following markup to **popup.html**.</span></span> <span data-ttu-id="f62e0-378">Observação:</span><span class="sxs-lookup"><span data-stu-id="f62e0-378">Note:</span></span>

   - <span data-ttu-id="f62e0-379">a página tem um `<input>` em que o usuário insere o nome dele e um botão que envia o nome para a página no painel de tarefas onde ele será exibido.</span><span class="sxs-lookup"><span data-stu-id="f62e0-379">The page has a `<input>` where the user will enter their name and a button that will send the name to the page in the task pane where it will be displayed.</span></span>

   - <span data-ttu-id="f62e0-380">a marcação carrega um script chamado **popup.js** que você criará em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="f62e0-380">The markup loads a script named **popup.js** that you will create in a later step.</span></span>

   - <span data-ttu-id="f62e0-381">Ela também carrega a biblioteca Office.js porque esta será usada em **popup.js**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-381">It also loads the Office.js library because it will be used in **popup.js**.</span></span>

    ```html
    <!DOCTYPE html>
    <html>
        <head lang="en">
            <title>Dialog for My Office Add-in</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">

            <!-- For more information on Office UI Fabric, visit https://developer.microsoft.com/fabric. -->
            <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css"/>

            <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
            <script type="text/javascript" src="popup.js"></script>

        </head>
        <body style="display:flex;flex-direction:column;align-items:center;justify-content:center">
            <p class="ms-font-xl">ENTER YOUR NAME</p>
            <input id="name-box" type="text"/><br/><br/>
            <button id="ok-button" class="ms-Button">OK</button>
        </body>
    </html>
    ```

4. <span data-ttu-id="f62e0-382">Na pasta **./src/dialogs**, crie um arquivo chamado **popup.js**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-382">In the **./src/dialogs** folder, create new file named **popup.js**.</span></span>

5. <span data-ttu-id="f62e0-383">Adicione o código a seguir a **popup.js**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-383">Add the following code to **popup.js**.</span></span> <span data-ttu-id="f62e0-384">Observe o seguinte sobre este código:</span><span class="sxs-lookup"><span data-stu-id="f62e0-384">Note the following about this code:</span></span>

   - <span data-ttu-id="f62e0-385">*Todas as páginas que chamam APIs na biblioteca Office.JS devem primeiro garantir que a biblioteca tenha sido totalmente inicializada.*</span><span class="sxs-lookup"><span data-stu-id="f62e0-385">*Every page that calls APIs in the Office.js library must first ensure that the library is fully initialized.*</span></span> <span data-ttu-id="f62e0-386">A melhor maneira de fazer isso é chamando o método `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-386">The best way to do that is to call the `Office.onReady()` method.</span></span> <span data-ttu-id="f62e0-387">Se o suplemento possuir as próprias tarefas de inicialização, o código deverá ser colocado em um método `then()` encadeado à chamada de `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-387">If your add-in has its own initialization tasks, the code should go in a `then()` method that is chained to the call of `Office.onReady()`.</span></span> <span data-ttu-id="f62e0-388">A chamada de `Office.onReady()` deve ser executada antes de qualquer chamada para Office.js; por isso, a tarefa se encontra em um arquivo de script que é carregado pela página, como neste caso.</span><span class="sxs-lookup"><span data-stu-id="f62e0-388">The call of `Office.onReady()` must run before any calls to Office.js; hence the assignment is in a script file that is loaded by the page, as it is in this case.</span></span>

    ```js
    (function () {
    "use strict";

        Office.onReady()
            .then(function() {

                // TODO1: Assign handler to the OK button.

            });

        // TODO2: Create the OK button handler

    }());
    ```

6. <span data-ttu-id="f62e0-p158">Substitua `TODO1` pelo código a seguir. Você criará a função `sendStringToParentPage` na próxima etapa.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p158">Replace `TODO1` with the following code. You'll create the `sendStringToParentPage` function in the next step.</span></span>

    ```js
    document.getElementById("ok-button").onclick = sendStringToParentPage;
    ```

7. <span data-ttu-id="f62e0-p159">Substitua `TODO2` pelo código a seguir. O método `messageParent` passa seu parâmetro para a página pai, neste caso, a página no painel de tarefas. O parâmetro pode ser um booliano ou uma cadeia de caracteres, que inclui tudo o que pode ser serializado como uma cadeia de caracteres, como XML ou JSON.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p159">Replace `TODO2` with the following code. The `messageParent` method passes its parameter to the parent page, in this case, the page in the task pane. The parameter can be a boolean or a string, which includes anything that can be serialized as a string, such as XML or JSON.</span></span>

    ```js
    function sendStringToParentPage() {
        var userName = document.getElementById("name-box").value;
        Office.context.ui.messageParent(userName);
    }
    ```

> [!NOTE]
> <span data-ttu-id="f62e0-394">O arquivo **popup.html** e o arquivo **popup.js** que ele carrega são executados em um processo totalmente separado do Microsoft Edge ou do Internet Explorer 11 no painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="f62e0-394">The **popup.html** file, and the **popup.js** file that it loads, run in an entirely separate Microsoft Edge or Internet Explorer 11 process from the add-in's task pane.</span></span> <span data-ttu-id="f62e0-395">Se o **popup.js** foi transcompilado no mesmo arquivo **bundle.js** que o arquivo **app.js**, o suplemento precisará carregar duas cópias do arquivo **bundle.js**, o que anula o propósito do agrupamento.</span><span class="sxs-lookup"><span data-stu-id="f62e0-395">If **popup.js** was transpiled into the same **bundle.js** file as the **app.js** file, then the add-in would have to load two copies of the **bundle.js** file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="f62e0-396">Portanto, esse suplemento não transcompila o arquivo **popup.js**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-396">Therefore, this add-in does not transpile the **popup.js** file at all.</span></span>

### <a name="update-webpack-config-settings"></a><span data-ttu-id="f62e0-397">Atualizar as configurações webpack config</span><span class="sxs-lookup"><span data-stu-id="f62e0-397">Update webpack config settings</span></span>

<span data-ttu-id="f62e0-398">Abra o arquivo **webpack.config.js** no diretório raiz do projeto e conclua as seguintes etapas.</span><span class="sxs-lookup"><span data-stu-id="f62e0-398">Open the file **webpack.config.js** in the root directory of the project and complete the following steps.</span></span>

1. <span data-ttu-id="f62e0-399">Localize o objeto `entry` dentro do objeto `config` e adicione uma nova entrada para `popup`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-399">Locate the `entry` object within the `config` object and add a new entry for `popup`.</span></span>

    ```js
    popup: "./src/dialogs/popup.js"
    ```

    <span data-ttu-id="f62e0-400">Após fazer isso, o novo objeto `entry` ficará assim:</span><span class="sxs-lookup"><span data-stu-id="f62e0-400">After you've done this, the new `entry` object will look like this:</span></span>

    ```js
    entry: {
      polyfill: "@babel/polyfill",
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
      popup: "./src/dialogs/popup.js"
    },
    ```
  
2. <span data-ttu-id="f62e0-401">Localize a matriz `plugins` no objeto `config` e adicione o seguinte objeto ao final dela.</span><span class="sxs-lookup"><span data-stu-id="f62e0-401">Locate the `plugins` array within the `config` object and add the following object to the end of that array.</span></span>

    ```js
    new HtmlWebpackPlugin({
      filename: "popup.html",
      template: "./src/dialogs/popup.html",
      chunks: ["polyfill", "popup"]
    })
    ```

    <span data-ttu-id="f62e0-402">Após fazer isso, a nova matriz `plugins` ficará assim:</span><span class="sxs-lookup"><span data-stu-id="f62e0-402">After you've done this, the new `plugins` array will look like this:</span></span>

    ```js
    plugins: [
      new CleanWebpackPlugin(),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ['polyfill', 'taskpane']
      }),
      new CopyWebpackPlugin([
      {
        to: "taskpane.css",
        from: "./src/taskpane/taskpane.css"
      }
      ]),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      }),
      new HtmlWebpackPlugin({
        filename: "popup.html",
        template: "./src/dialogs/popup.html",
        chunks: ["polyfill", "popup"]
      })
    ],
    ```

3. <span data-ttu-id="f62e0-403">Se o servidor da Web local estiver em execução, feche a janela de comando do nó para interrompê-lo.</span><span class="sxs-lookup"><span data-stu-id="f62e0-403">If the local web server is running, stop it by closing the node command window.</span></span>

4. <span data-ttu-id="f62e0-404">Execute o seguinte comando para recriar o projeto.</span><span class="sxs-lookup"><span data-stu-id="f62e0-404">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

### <a name="open-the-dialog-from-the-task-pane"></a><span data-ttu-id="f62e0-405">Abra a caixa de diálogo do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="f62e0-405">Open the dialog from the task pane</span></span>

1. <span data-ttu-id="f62e0-406">Abra o arquivo **./src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-406">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="f62e0-407">Localize o elemento `<button>` para o botão `freeze-header` e adicione a seguinte marcação após essa linha:</span><span class="sxs-lookup"><span data-stu-id="f62e0-407">Locate the `<button>` element for the `freeze-header` button, and add the following markup after that line:</span></span>

    ```html
    <button class="ms-Button" id="open-dialog">Open Dialog</button><br/><br/>
    ```

3. <span data-ttu-id="f62e0-408">A caixa de diálogo solicitará que o usuário insira um nome e passará o nome de usuário para o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="f62e0-408">The dialog will prompt the user to enter a name and pass the user's name to the task pane.</span></span> <span data-ttu-id="f62e0-409">O painel de tarefas o exibirá em um rótulo.</span><span class="sxs-lookup"><span data-stu-id="f62e0-409">The task pane will display it in a label.</span></span> <span data-ttu-id="f62e0-410">Imediatamente após o `button` que você adicionou, adicione a marcação a seguir:</span><span class="sxs-lookup"><span data-stu-id="f62e0-410">Immediately after the `button` that you just added, add the following markup:</span></span>

    ```html
    <label id="user-name"></label><br/><br/>
    ```

4. <span data-ttu-id="f62e0-411">Abra o arquivo **./src/taskpane/taskpane.js**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-411">Open the file **./src/taskpane/taskpane.js**.</span></span>

5. <span data-ttu-id="f62e0-412">Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `freeze-header` e adicione o seguinte código após ela.</span><span class="sxs-lookup"><span data-stu-id="f62e0-412">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `freeze-header` button, and add the following code after that line.</span></span> <span data-ttu-id="f62e0-413">Você criará o método `openDialog` em uma etapa posterior.</span><span class="sxs-lookup"><span data-stu-id="f62e0-413">You'll create the `openDialog` method in a later step.</span></span>

    ```js
    document.getElementById("open-dialog").onclick = openDialog;
    ```

6. <span data-ttu-id="f62e0-414">Adicione a seguinte declaração ao final do arquivo:</span><span class="sxs-lookup"><span data-stu-id="f62e0-414">Add the following declaration to the end of the file.</span></span> <span data-ttu-id="f62e0-415">Essa variável é usada para armazenar um objeto no contexto de execução da página pai que atua como um intermediador no contexto de execução da página da caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="f62e0-415">This variable is used to hold an object in the parent page's execution context that acts as an intermediator to the dialog page's execution context.</span></span>

    ```js
    var dialog = null;
    ```

7. <span data-ttu-id="f62e0-416">Adicione a seguinte função ao final do arquivo, após a declaração de `dialog`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-416">Add the following function to the end of the file (after the declaration of `dialog`).</span></span> <span data-ttu-id="f62e0-417">É importante observar o que esse código *não* contém: não há nenhuma chamada de `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-417">The important thing to notice about this code is what is *not* there: there is no call of `Excel.run`.</span></span> <span data-ttu-id="f62e0-418">Isso ocorre porque a API para abrir uma caixa de diálogo é compartilhada com todos os hosts do Office, portanto, ela faz parte da API de Office JavaScript Common, não da API específica do Excel.</span><span class="sxs-lookup"><span data-stu-id="f62e0-418">This is because the API to open a dialog is shared among all Office hosts, so it is part of the Office JavaScript Common API, not the Excel-specific API.</span></span>

    ```js
    function openDialog() {
        // TODO1: Call the Office Common API that opens a dialog
    }
    ```

8. <span data-ttu-id="f62e0-419">Substitua `TODO1` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="f62e0-419">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="f62e0-420">Observação:</span><span class="sxs-lookup"><span data-stu-id="f62e0-420">Note:</span></span>

   - <span data-ttu-id="f62e0-421">O método`displayDialogAsync` abre uma caixa de diálogo no centro da tela.</span><span class="sxs-lookup"><span data-stu-id="f62e0-421">The `displayDialogAsync` method opens a dialog in the center of the screen.</span></span>

   - <span data-ttu-id="f62e0-422">O primeiro parâmetro é a URL da página a ser aberta.</span><span class="sxs-lookup"><span data-stu-id="f62e0-422">The first parameter is the URL of the page to open.</span></span>

   - <span data-ttu-id="f62e0-p166">O segundo parâmetro passa opções. `height` e `width` são porcentagens do tamanho da janela do aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p166">The second parameter passes options. `height` and `width` are percentages of the size of the Office application's window.</span></span>

    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},

        // TODO2: Add callback parameter.
    );
    ```

### <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a><span data-ttu-id="f62e0-425">Processar a mensagem da caixa de diálogo e depois fechá-la</span><span class="sxs-lookup"><span data-stu-id="f62e0-425">Process the message from the dialog and close the dialog</span></span>

1. <span data-ttu-id="f62e0-426">Na função `openDialog` no arquivo **./src/taskpane/taskpane.js**, substitua `TODO2` pelo código a seguir.</span><span class="sxs-lookup"><span data-stu-id="f62e0-426">Within the `openDialog` function in the file **./src/taskpane/taskpane.js**, replace `TODO2` with the following code.</span></span> <span data-ttu-id="f62e0-427">Observação:</span><span class="sxs-lookup"><span data-stu-id="f62e0-427">Note:</span></span>

   - <span data-ttu-id="f62e0-428">O retorno de chamada é executado imediatamente depois que a caixa de diálogo é aberta com êxito e antes de usuário executar a ação na caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="f62e0-428">The callback is executed immediately after the dialog successfully opens and before the user has taken any action in the dialog.</span></span>

   - <span data-ttu-id="f62e0-429">O `result.value` é o objeto que funciona como um tipo de intermediário entre contextos execução das páginas de pai e de caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="f62e0-429">The `result.value` is the object that acts as a kind of middleman between the execution contexts of the parent and dialog pages.</span></span>

   - <span data-ttu-id="f62e0-p168">A função `processMessage` será criada em uma etapa posterior. Esse identificador processará os valores que sejam enviados da página da caixa de diálogo com chamadas da função `messageParent`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p168">The `processMessage` function will be created in a later step. This handler will process any values that are sent from the dialog page with calls of the `messageParent` function.</span></span>

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. <span data-ttu-id="f62e0-432">Adicione a seguinte função após a função `openDialog`.</span><span class="sxs-lookup"><span data-stu-id="f62e0-432">Add the following function after the `openDialog` function.</span></span>

    ```js
    function processMessage(arg) {
        document.getElementById("user-name").innerHTML = arg.message;
        dialog.close();
    }
    ```

3. <span data-ttu-id="f62e0-433">Verifique se você salvou todas as alterações feitas no projeto.</span><span class="sxs-lookup"><span data-stu-id="f62e0-433">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="f62e0-434">Testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="f62e0-434">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="f62e0-435">Se o painel de tarefas do suplemento ainda não estiver aberto no Excel, vá para a guia **Página Inicial** e escolha o botão **Mostrar Painel de Tarefas** na faixa de opções para abri-lo.</span><span class="sxs-lookup"><span data-stu-id="f62e0-435">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="f62e0-436">Escolha o botão **Abrir Caixa de Diálogo** no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="f62e0-436">Choose the **Open Dialog** button in the task pane.</span></span>

4. <span data-ttu-id="f62e0-437">Quando a caixa de diálogo estiver aberta, arraste-a e redimensione-a.</span><span class="sxs-lookup"><span data-stu-id="f62e0-437">While the dialog is open, drag it and resize it.</span></span> <span data-ttu-id="f62e0-438">Observe que você pode interagir com a planilha e pressionar outros botões no painel de tarefas. No entanto, não é possível iniciar uma segunda caixa de diálogo na mesma página do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="f62e0-438">Note that you can interact with the worksheet and press other buttons on the task pane, but you cannot launch a second dialog from the same task pane page.</span></span>

5. <span data-ttu-id="f62e0-439">Na caixa de diálogo, digite um nome e escolha o botão **OK**.</span><span class="sxs-lookup"><span data-stu-id="f62e0-439">In the dialog, enter a name and choose the **OK** button.</span></span> <span data-ttu-id="f62e0-440">O nome aparecerá no painel de tarefas e a caixa de diálogo será fechada.</span><span class="sxs-lookup"><span data-stu-id="f62e0-440">The name appears on the task pane and the dialog closes.</span></span>

6. <span data-ttu-id="f62e0-p171">Opcionalmente, comente a linha `dialog.close();` na função `processMessage`. Em seguida, repita as etapas desta seção. A caixa de diálogo permanece aberta e você pode alterar o nome. É possível fechá-la manualmente pressionando o botão **X** no canto superior direito.</span><span class="sxs-lookup"><span data-stu-id="f62e0-p171">Optionally, comment out the line `dialog.close();` in the `processMessage` function. Then repeat the steps of this section. The dialog stays open and you can change the name. You can close it manually by pressing the **X** button in the upper right corner.</span></span>

    ![Tutorial do Excel - Caixa de diálogo](../images/excel-tutorial-dialog-open-2.png)

## <a name="next-steps"></a><span data-ttu-id="f62e0-446">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="f62e0-446">Next steps</span></span>

<span data-ttu-id="f62e0-447">Neste tutorial você criou um suplemento do Excel que interage com tabelas, gráficos, planilhas e caixas de diálogo em uma pasta de trabalho do Excel.</span><span class="sxs-lookup"><span data-stu-id="f62e0-447">In this tutorial, you've created an Excel task pane add-in that interacts with tables, charts, worksheets, and dialogs in an Excel workbook.</span></span> <span data-ttu-id="f62e0-448">Para saber mais sobre o desenvolvimento de suplementos do Excel, continue no seguinte artigo:</span><span class="sxs-lookup"><span data-stu-id="f62e0-448">To learn more about building Excel add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="f62e0-449">Visão geral dos suplementos do Excel</span><span class="sxs-lookup"><span data-stu-id="f62e0-449">Excel add-ins overview</span></span>](../excel/excel-add-ins-overview.md)

## <a name="see-also"></a><span data-ttu-id="f62e0-450">Confira também</span><span class="sxs-lookup"><span data-stu-id="f62e0-450">See also</span></span>

* [<span data-ttu-id="f62e0-451">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="f62e0-451">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="f62e0-452">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="f62e0-452">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="f62e0-453">Desenvolver Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="f62e0-453">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="f62e0-454">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="f62e0-454">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)