---
title: Tutorial de suplemento do Excel
description: Crie um suplemento do Excel que cria, preenche, filtra e classifica uma tabela, cria um gráfico, congela um cabeçalho de tabela, protege uma planilha e abre uma caixa de diálogo.
ms.date: 02/26/2022
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: 5ca9ea0fdc600d6044cf3a5ef405dd0f3a98e2b3
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746419"
---
# <a name="tutorial-create-an-excel-task-pane-add-in"></a>Tutorial: criar um suplemento do painel de tarefas no Excel

Neste tutorial: você criará um suplemento do painel de tarefas no Excel

> [!div class="checklist"]
>
> - Cria uma tabela
> - Filtra e classifica uma tabela
> - Cria um gráfico
> - Congela um cabeçalho de tabela
> - Protege uma planilha
> - Abre uma caixa de diálogo

> [!TIP]
> Se você já concluiu o inicio rápido do [Criar um suplemento do painel de tarefas no Excel](../quickstarts/excel-quickstart-jquery.md) usando o gerador Yeoman e deseja usar esse projeto como ponto de partida para este tutorial, vá diretamente para a seção [Criar uma tabela](#create-a-table) para iniciar este tutorial.

## <a name="prerequisites"></a>Pré-requisitos

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Office conectado a uma assinatura Microsoft 365 (incluindo o Office na web).

    > [!NOTE]
    > Se você ainda não tem o Office, poderá [ingressar no programa para desenvolvedores do Microsoft 365](https://developer.microsoft.com/office/dev-program) para obter uma assinatura do Microsoft 365 gratuita e renovável por 90 dias para usar durante o desenvolvimento.

## <a name="create-your-add-in-project"></a>Criar seu projeto do suplemento

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Escolha o tipo de projeto:** `Office Add-in Task Pane project`
- **Escolha o tipo de script:** `JavaScript`
- **Qual será o nome do suplemento?** `My Office Add-in`
- **Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?** `Excel`

![Captura de tela da interface de linha de comando do gerador do suplemento Yeoman Office.](../images/yo-office-excel.png)

Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="create-a-table"></a>Criar uma tabela

Nesta etapa do tutorial, você testará no programa se o suplemento é compatível com a versão atual do Excel do usuário, adicionará uma tabela a uma planilha, depois preencherá e formatará a tabela com os dados.

### <a name="code-the-add-in"></a>Codificação do suplemento

1. Abra o projeto em seu editor de código.

1. Abra o arquivo **./src/taskpane/taskpane.html**.  Este arquivo contém a marcação HTML para o painel de tarefas.

1. Localize o elemento `<main>` e exclua todas as linhas que aparecem após a marca de abertura `<main>` e antes da marca de fechamento `</main>`.

1. Adicione a seguinte marcação imediatamente após a marca de abertura `<main>`.

    ```html
    <button class="ms-Button" id="create-table">Create Table</button><br/><br/>
    ```

1. Abra o arquivo **./src/taskpane/taskpane.js**. Esse arquivo contém o código da API do JavaScript do Office que facilita a interação entre o painel de tarefas e o aplicativo cliente do Office.

1. Remova todas as referências ao botão `run` e à função `run()` da seguinte forma:

    - Localize e exclua a linha `document.getElementById("run").onclick = run;`.

    - Localize e exclua toda a função `run()`.

1. Na chamada do método `Office.onReady`, localize a linha `if (info.host === Office.HostType.Excel) {` e adicione o código a seguir imediatamente após essa linha. Observação:

    - A primeira parte desse código determina se a versão do Excel do usuário oferece suporte a uma versão do Excel.js que inclua todas as APIs que essa série de tutoriais usará. Em um suplemento de produção, use o corpo do bloco condicional para ocultar ou desabilitar a IU que chamaria APIs sem suporte. Isso permitirá que o usuário ainda use as partes do suplemento que são compatíveis com sua versão do Excel.

    - A segunda parte desse código adiciona um manipulador de eventos para o botão `create-table`.

    ```js
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("create-table").onclick = createTable;
    ```

1. Adicione a seguinte função ao final do arquivo. Observação:

    - A lógica de negócios de Excel.js será adicionada à função que passar por `Excel.run`. Essa lógica não é executada imediatamente. Em vez disso, ela é adicionada à fila de comandos pendentes.

    - O método `context.sync` envia todos os comandos da fila para execução no Excel.

    - `Excel.run` é seguido por um bloco `catch`. Essa é uma prática recomendada que você sempre deve seguir.

    [!include[Information about the use of ES6 JavaScript](../includes/modern-js-note.md)]

    ```js
    async function createTable() {
        await Excel.run(async (context) => {

            // TODO1: Queue table creation logic here.

            // TODO2: Queue commands to populate the table with data.

            // TODO3: Queue commands to format the table.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. Na função `createTable()`, substitua `TODO1` com o seguinte código. Observação:

    - O código cria uma tabela usando o método `add` da coleção de tabelas de uma planilha, que sempre existe, mesmo que esteja vazia. Essa é a maneira padrão em que os objetos Excel.js são criados. Não há APIs de construtor de classe e você nunca usa um operador `new` para criar um objeto do Excel. Em vez disso, você adiciona a um objeto de coleção pai.

    - O primeiro parâmetro do método `add` é o intervalo apenas da linha superior da tabela, e não de todo o intervalo que a tabela por fim usará. Isso ocorre porque quando o suplemento preenche as linhas de dados (na próxima etapa), ele adiciona novas linhas à tabela em vez de gravar valores nas células das linhas existentes. Esse é um padrão comum, porque o número de linhas que uma tabela terá geralmente é desconhecido quando a tabela é criada.

    - Os nomes de tabelas devem ser exclusivos pela pasta de trabalho inteira, não só na planilha.

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ```

1. Na função `createTable()`, substitua `TODO2` com o seguinte código. Observação:

    - Os valores das células de um intervalo são definidos em uma matriz de matrizes.

    - Novas linhas são criadas em uma tabela ao chamar o método `add` do conjunto de linhas da tabela. Você pode adicionar várias linhas em uma única chamada de `add` ao incluir várias matrizes de valores de células na matriz pai que é passada como segundo parâmetro.

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

1. Na função `createTable()`, substitua `TODO3` com o seguinte código. Observação:

    - O código recebe uma referência para a coluna **quantidade** ao passar o índice com base em zero para o método `getItemAt` do conjunto de colunas da tabela.

        > [!NOTE]
        > Os objetos do conjunto Excel.js, como `TableCollection`, `WorksheetCollection`, e `TableColumnCollection`, têm a propriedade `items` que é como uma matriz dos tipos de objetos filhos, como `Table` ou `Worksheet` ou `TableColumn`; mas um objeto `*Collection` não é uma matriz.

    - O código formata o intervalo da coluna **quantidade** como Euros com um segundo decimal.

    - Por fim, isso garante que a largura das colunas e a altura das linhas sejam grandes o suficiente para o maior (ou o mais alto) item de dados. Observe que o código deve receber os objetos `Range` a formatar. Os objetos `TableColumn` e `TableRow` não têm propriedades de formato.

    ```js
    expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();
    ```

1. Verifique se você salvou todas as alterações feitas no projeto.

### <a name="test-the-add-in"></a>Testar o suplemento

1. Conclua as etapas a seguir para iniciar o servidor Web local e fazer o sideload do seu suplemento.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    > [!TIP]
    > Se você estiver testando seu suplemento no Mac, execute o seguinte comando no diretório raiz do seu projeto antes de continuar. O servidor Web local é iniciado quando este comando é executado.
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - Para testar o seu suplemento no Excel, execute o seguinte comando no diretório raiz do projeto. Isso inicia o servidor Web local (caso ainda não esteja em execução) e abre o Excel com o suplemento carregado.

        ```command&nbsp;line
        npm start
        ```

    - Para testar seu suplemento no Excel na Web, execute o seguinte comando no diretório raiz do seu projeto. O servidor Web local é iniciado quando este comando é executado. Substitua “{url}” pelo URL de um documento do Excel no seu OneDrive ou uma biblioteca do SharePoint para a qual você tenha permissões.

        [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

1. No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.

    ![Captura de tela do menu da página inicial do Excel, com o botão Mostrar Painel de Tarefas realçado.](../images/excel-quickstart-addin-3b.png)

1. No painel de tarefas, escolha o botão **Criar tabela**.

    ![Captura de tela do Excel, exibindo um painel de tarefas de suplemento com um botão Criar Tabela, e uma tabela na planilha preenchida com dados de Data, Comerciante, Categoria e Quantidade.](../images/excel-tutorial-create-table-2.png)

## <a name="filter-and-sort-a-table"></a>Filtrar e classificar uma tabela

Nesta etapa do tutorial, você vai filtrar e classificar a tabela que criou anteriormente.

### <a name="filter-the-table"></a>Filtrar a tabela

1. Abra o arquivo **./src/taskpane/taskpane.html**.

1. Localize o elemento `<button>` do botão `create-table` e adicione a seguinte marcação logo após essa linha.

    ```html
    <button class="ms-Button" id="filter-table">Filter Table</button><br/><br/>
    ```

1. Abra o arquivo **./src/taskpane/taskpane.js**.

1. Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `create-table` e adicione o seguinte código após ela.

    ```js
    document.getElementById("filter-table").onclick = filterTable;
    ```

1. Adicione a seguinte função ao final do arquivo.

    ```js
    async function filterTable() {
        await Excel.run(async (context) => {

            // TODO1: Queue commands to filter out all expense categories except
            //        Groceries and Education.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. Na função `filterTable()`, substitua `TODO1` com o seguinte código. Observação:

   - O código primeiro faz referência à coluna que precisa de filtragem ao passar o nome da coluna para o método `getItem`, em vez de passar o índice para o método `getItemAt` como o método `createTable` faz. Como os usuários podem mover as colunas da tabela, a coluna de um determinado índice pode mudar depois da criação da tabela. Portanto, é mais seguro usar o nome da coluna como referência dela. Usamos de forma segura `getItemAt` em um tutorial anterior porque usamos o mesmo método que cria a tabela. Assim não existe a chance de um usuário mover a coluna.

   - O método `applyValuesFilter` é um dos vários métodos de filtragem do objeto `Filter`.

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(['Education', 'Groceries']);
    ```

### <a name="sort-the-table"></a>Classificar a tabela

1. Abra o arquivo **./src/taskpane/taskpane.html**.

1. Localize o elemento `<button>` do botão `filter-table` e adicione a seguinte marcação logo após essa linha.

    ```html
    <button class="ms-Button" id="sort-table">Sort Table</button><br/><br/>
    ```

1. Abra o arquivo **./src/taskpane/taskpane.js**.

1. Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `filter-table` e adicione o seguinte código após ela.

    ```js
    document.getElementById("sort-table").onclick = sortTable;
    ```

1. Adicione a seguinte função ao final do arquivo.

    ```js
    async function sortTable() {
        await Excel.run(async (context) => {

            // TODO1: Queue commands to sort the table by Merchant name.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. Na função `sortTable()`, substitua `TODO1` com o seguinte código. Observação:

   - O código cria uma matriz de objetos `SortField`, que possui apenas um membro, pois o suplemento é classificado apenas na coluna Merchant.

   - A propriedade `key` de um objeto `SortField` é o índice baseado em zero da coluna usada para classificação. As linhas da tabela são classificadas com base nos valores da coluna referenciada.

   - O `sort` membro de um `Table` é um `TableSort` objeto, não um método. Os `SortField`s são passados ao `TableSort` método do `apply` objeto.

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

1. Verifique se você salvou todas as alterações feitas no projeto.

### <a name="test-the-add-in"></a>Testar o suplemento

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

1. Se o painel de tarefas do suplemento ainda não estiver aberto no Excel, vá para a guia **Página Inicial** e escolha o botão **Mostrar Painel de Tarefas** na faixa de opções para abri-lo.

1. Se a tabela que você adicionou anteriormente neste tutorial não estiver presente na planilha aberta, escolha o botão **Criar tabela** no painel de tarefas.

1. Escolha os botões **Filtrar Tabela** e **Classificar Tabela**, em qualquer ordem.

    ![Captura de tela do Excel, com os botões Filtrar Tabela e Classificar Tabela visíveis no painel de tarefas do suplemento.](../images/excel-tutorial-filter-and-sort-table-2.png)

## <a name="create-a-chart"></a>Criar um gráfico

Nesta etapa do tutorial, você vai criar um gráfico com dados da tabela que você criou anteriormente e depois vai formatar o gráfico.

### <a name="chart-a-chart-using-table-data"></a>Gráfico de um gráfico com dados de tabela

1. Abra o arquivo **./src/taskpane/taskpane.html**.

1. Localize o elemento `<button>` do botão `sort-table` e adicione a seguinte marcação logo após essa linha.

    ```html
    <button class="ms-Button" id="create-chart">Create Chart</button><br/><br/>
    ```

1. Abra o arquivo **./src/taskpane/taskpane.js**.

1. Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `sort-table` e adicione o seguinte código após ela.

    ```js
    document.getElementById("create-chart").onclick = createChart;
    ```

1. Adicione a seguinte função ao final do arquivo.

    ```js
    async function createChart() {
        await Excel.run(async (context) => {

            // TODO1: Queue commands to get the range of data to be charted.

            // TODO2: Queue command to create the chart and define its type.

            // TODO3: Queue commands to position and format the chart.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. Dentro da função `createChart()`, substitua `TODO1` pelo seguinte código. Observe que, para excluir a linha do cabeçalho, o código usa o método `Table.getDataBodyRange` para obter o intervalo de dados que você deseja registrar em vez do método `getRange`.

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const dataRange = expensesTable.getDataBodyRange();
    ```

1. Na função `createChart()`, substitua `TODO2` com o seguinte código. Observe os seguintes parâmetros.

   - O primeiro parâmetro para o método `add` especifica o tipo de gráfico. Há diversos tipos.

   - O segundo parâmetro especifica um intervalo de dados a incluir no gráfico.

   - O terceiro parâmetro determina se uma série de pontos de dados da tabela deve ser representada por linha ou coluna. A opção `auto` informa ao Excel para decidir o melhor método.

    ```js
    const chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'Auto');
    ```

1. Na função `createChart()`, substitua `TODO3` pelo seguinte código. A maior parte desse código é autoexplicativa. Observação:

   - Os parâmetros do método `setPosition` especificam as células da esquerda superior e da direita inferior da área da planilha que deve conter o gráfico. O Excel ajusta detalhes como a largura da linha para criar uma boa aparência para o gráfico no espaço fornecido.

   - “Série” é um conjunto de pontos de dados de uma coluna da tabela. Como há apenas uma coluna sem cadeia de caracteres na tabela, o Excel deduz que essa é a única coluna de pontos de dados no gráfico. Ele interpreta outra colunas como rótulos do gráfico. Portanto, haverá apenas uma série no gráfico e será necessário o índice 0. Ele será rotulado como “Valor em &euro;”.

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "Right";
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in \u20AC';
    ```

1. Verifique se você salvou todas as alterações feitas no projeto.

### <a name="test-the-add-in"></a>Testar o suplemento

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

1. Se o painel de tarefas do suplemento ainda não estiver aberto no Excel, vá para a guia **Página Inicial** e escolha o botão **Mostrar Painel de Tarefas** na faixa de opções para abri-lo.

1. Se a tabela que você adicionou anteriormente neste tutorial não estiver presente na planilha aberta, escolha o botão **Criar tabela** e depois os botões **Filtrar Tabela** e **Classificar Tabela**, em qualquer ordem.

1. Clique no botão **Criar gráfico**. Um gráfico é criado e incluirá somente os dados das linhas que foram filtradas. Os rótulos dos pontos de dados na parte inferior estão na ordem de classificação do gráfico, ou seja, nomes de comerciantes em ordem alfabética inversa.

    ![Captura de tela do Excel, com um botão Criar Gráfico visível no painel de tarefas do suplemento e um gráfico na planilha exibindo dados de despesas com alimentos e educação.](../images/excel-tutorial-create-chart-2.png)

## <a name="freeze-a-table-header"></a>Congelar um cabeçalho de tabela

Quando uma tabela for longa o suficiente para que um usuário precise rolar para ver algumas linhas, a linha de cabeçalho poderá ficar fora da vista. Nesta etapa do tutorial, você precisará congelar a linha do cabeçalho da tabela que criou anteriormente para que ela permaneça visível, mesmo que o usuário role ao longo da planilha.

### <a name="freeze-the-tables-header-row"></a>Congelar a linha de cabeçalho da tabela

1. Abra o arquivo **./src/taskpane/taskpane.html**.

1. Localize o elemento `<button>` do botão `create-chart` e adicione a seguinte marcação logo após essa linha.

    ```html
    <button class="ms-Button" id="freeze-header">Freeze Header</button><br/><br/>
    ```

1. Abra o arquivo **./src/taskpane/taskpane.js**.

1. Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `create-chart` e adicione o seguinte código após ela.

    ```js
    document.getElementById("freeze-header").onclick = freezeHeader;
    ```

1. Adicione a seguinte função ao final do arquivo.

    ```js
    async function freezeHeader() {
        await Excel.run(async (context) => {

            // TODO1: Queue commands to keep the header visible when the user scrolls.

            await context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

1. Na função `freezeHeader()`, substitua `TODO1` com o seguinte código. Observação:

   - A coleção `Worksheet.freezePanes` é um conjunto de painéis da planilha que fica congelado ou fixado no mesmo lugar quando rolamos a planilha.

   - O método `freezeRows` toma como parâmetro o número de linhas, a partir do topo, que devem ser fixadas no lugar. Passamos `1` para fixar a primeira linha no lugar.

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ```

1. Verifique se você salvou todas as alterações feitas no projeto.

### <a name="test-the-add-in"></a>Testar o suplemento

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

1. Se o painel de tarefas do suplemento ainda não estiver aberto no Excel, vá para a guia **Página Inicial** e escolha o botão **Mostrar Painel de Tarefas** na faixa de opções para abri-lo.

1. Se a tabela que você adicionou anteriormente neste tutorial estiver presente na planilha, faça a exclusão dela.

1. No painel de tarefas, escolha o botão **Criar tabela**.

1. No painel de tarefas, escolha o botão **Congelar Cabeçalho**.

1. Role a planilha para baixo o suficiente para ver que o cabeçalho da tabela permanece visível na parte superior mesmo ao rolar até que as primeiras linhas fiquem fora da vista.

    ![Captura de tela mostrando uma planilha do Excel com um cabeçalho de tabela congelado.](../images/excel-tutorial-freeze-header-2.png)

## <a name="protect-a-worksheet"></a>Proteger uma planilha

Nesta etapa do tutorial, você adicionará um botão à faixa de opções que ativa ou desativa a proteção da planilha.

### <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a>Configure o manifesto para adicionar um segundo botão à faixa de opções

1. Abra o arquivo de manifesto **./manifest.xml**.

1. Localize o elemento `<Control>`. Esse elemento define o botão **Mostrar painel de tarefas** na faixa **Página inicial** que você está usando para iniciar o suplemento. Vamos adicionar um segundo botão ao mesmo grupo na faixa **Página inicial**. Entre os rótulos `</Control>` de fechamento `</Group>`adicione a seguinte marcação.

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

1. No XML que você acabou de adicionar ao arquivo de manifesto, substitua `TODO1` por uma sequência que forneça ao botão um ID exclusivo nesse arquivo de manifesto. Como nosso botão ativará ou desativará a proteção da planilha, use "ToggleProtection". Quando você terminar, a marca de abertura para o elemento `Control` deverá ficar assim:

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

1. Os três `TODO`s seguintes definem as IDs dos recursos, ou `resid`s. Um recurso é uma cadeia de caracteres (com um comprimento máximo de 32 caracteres) e você criará essas três cadeias de caracteres em uma etapa posterior. Por enquanto, você precisa fornecer IDs aos recursos. O rótulo do botão deve ser "Toggle Protection", mas a *ID* dessa cadeia de caracteres deve ser "ProtectionButtonLabel", então o elemento `Label` deve ser semelhante a este:

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

1. O elemento `SuperTip` define a dica de ferramenta do botão. O título da dica de ferramenta deve ser o mesmo que o rótulo do botão, por isso, usamos a mesma ID de recurso: "ProtectionButtonLabel". A descrição da dica de ferramenta será "Click to turn protection of the worksheet on and off". Mas o `resid` será "ProtectionButtonToolTip". Portanto, quando você terminar, o elemento `SuperTip` deverá ficar assim:

    ```xml
    <Supertip>
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE]
   > Em um suplemento de produção,não é recomendável usar o mesmo ícone para dois botões diferentes; mas, para simplificar este tutorial, faremos isso. Portanto, a marcação `Icon` em nosso novo `Control` é apenas uma cópia do elemento `Icon` do `Control` existente.

1. O elemento `Action` dentro do elemento original `Control` está com seu tipo definido para `ShowTaskpane`, mas nosso novo botão não abrirá um painel de tarefas; ele executará uma função personalizada que você criará em uma etapa posterior. Portanto, substitua `TODO5` por `ExecuteFunction`, pois é o tipo de ação para botões que acionam funções personalizadas. O rótulo de abertura do elemento `Action` deve ser semelhante a este:

    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

1. O elemento `Action` original tem elementos filhos que especificam uma ID do painel de tarefas e uma URL da página que deve ser aberta no painel de tarefas. No entanto, um elemento `Action` do tipo `ExecuteFunction` possui um único elemento filho que nomeia a função executada pelo controle. Você criará essa função em uma etapa posterior e ela será chamada de `toggleProtection`. Então, substitua `TODO6` pela marcação a seguir.

    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    A marcação `Control` inteira deve ter a aparência a seguir:

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

1. Role para baixo até a seção `Resources` do manifesto.

1. Adicione a seguinte marcação como filho do elemento `bt:ShortStrings`.

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

1. Adicione a seguinte marcação como filho do elemento `bt:LongStrings`.

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

1. Salve o arquivo.

### <a name="create-the-function-that-protects-the-sheet"></a>Criar a função que protege a planilha

1. Abra o arquivo **.\commands\commands.js**.

1. Adicione a seguinte função imediatamente após a função `action`. Especificamos um parâmetro `args` para a função, e a última linha da função chama `args.completed`. Esse é um requisito para todos os comandos de suplemento do tipo **ExecuteFunction**. Ele sinaliza para o aplicativo do cliente Office que a função terminou e que a interface do usuário podem ficar responsiva novamente.

    ```js
    async function toggleProtection(args) {
        await Excel.run(async (context) => {

            // TODO1: Queue commands to reverse the protection status of the current worksheet.

            await context.sync();
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

1. Adicione a seguinte linha ao final do arquivo:

    ```js
    g.toggleProtection = toggleProtection;
    ```

1. Na função `toggleProtection`, substitua `TODO1` pelo seguinte código. Esse código usa a propriedade de proteção do objeto de planilha em um padrão de alternância padrão. O `TODO2` será explicado na próxima seção.

    ```js
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // TODO2: Queue command to load the sheet's "protection.protected" property from
    //        the document and re-synchronize the document and task pane.

    if (sheet.protection.protected) {
        sheet.protection.unprotect();
    } else {
        sheet.protection.protect();
    }
    ```

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a>Adicione código para buscar propriedades do documento em objetos de script do painel de tarefas

Em cada função que você criou neste tutorial até agora, você enfileirou comandos para *escrever* no documento do Office. Cada função termina com uma chamada ao método `context.sync()`, que envia os comandos enfileirados ao documento a ser executado. No entanto, o código que você adicionou na última etapa chama o `sheet.protection.protected property`. Essa é uma diferença significativa em relação às funções anteriores que você escreveu, porque o objeto `sheet` é apenas um objeto proxy que existe no script do painel de tarefas. O objeto proxy não conhece o estado real de proteção do documento, portanto, sua propriedade `protection.protected` não pode ter um valor real. Para evitar um erro de exceção, você deve primeiro buscar o status de proteção do documento e usá-lo para definir o valor de `sheet.protection.protected`. Esse processo de busca possui três etapas.

   1. Coloque em fila um comando para carregar (ou seja, fetch) as propriedades que seu código precisa ler.

   1. Chame o método `sync` do objeto de contexto para enviar o comando em fila para o documento para execução e retornar as informações solicitadas.

   1. Como o método `sync` é assíncrono, certifique-se de que ele tenha sido concluído antes que o código chame as propriedades que foram buscadas.

Essas etapas devem ser concluídas sempre que seu código precisar *ler* informações do documento do Office.

1. Na função `toggleProtection`, substitua `TODO2` com o seguinte código. Observação:

   - Todos os objetos do Excel têm um método `load`. Especifique as propriedades do objeto que você deseja ler no parâmetro como uma cadeia de caracteres de nomes delimitados por vírgulas. Nesse caso, a propriedade que você precisa ler é uma subpropriedade de `protection`. Referencie a subpropriedade quase exatamente como você faria em qualquer lugar do seu código, mas usando uma barra (“/”) em vez de um ponto (".").

   - Para garantir que a lógica de alternância, que lê `sheet.protection.protected`, não seja executada até que o `sync` seja concluído e o `sheet.protection.protected` tenha recebido o valor correto obtido do documento, ele deverá vir depois que o operador `await` garantir que `sync` tenha sido concluído.

    ```js
    sheet.load('protection/protected');
    await context.sync();
    ```

   Quando terminar, a função inteira deve se parecer com o seguinte:

    ```js
    async function toggleProtection(args) {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.load('protection/protected');

            await context.sync();

            if (sheet.protection.protected) {
                sheet.protection.unprotect();
            } else {
                sheet.protection.protect();
            }
            
            await context.sync();
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

1. Verifique se você salvou todas as alterações feitas no projeto.

### <a name="test-the-add-in"></a>Testar o suplemento

1. Feche todos os aplicativos do Office, incluindo o Excel.

1. Exclua o cache do Office, excluindo o conteúdo (todos os arquivos e subpastas) da pasta de cache. Isso é necessário para limpar completamente a versão antiga do suplemento do aplicativo cliente.

    - No Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.

    - No Mac: `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.

      > [!NOTE]
      > Se essa pasta não existir, verifique as pastas a seguir e, se encontradas, exclua o conteúdo da pasta.
      >
      >  - `~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/` onde `{host}` é o aplicativo do Office (por exemplo, `Excel`)
      >  - `~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/` onde `{host}` é o aplicativo do Office (por exemplo, `Excel`)
      >  - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`
      >  - `~/Library/Containers/com.microsoft.Office365ServiceV2/Data/Library/Caches/com.microsoft.Office365ServiceV2/`

1. Se o servidor da Web local já estiver em execução, feche a janela de comando do nó para interrompê-lo.

1. Como o arquivo de manifesto foi atualizado, você deve carregar o suplemento novamente usando esse arquivo. Inicie o servidor Web local e realize o sideload no seu suplemento:

    - Para testar o seu suplemento no Excel, execute o seguinte comando no diretório raiz do projeto. Isso inicia o servidor Web local (caso ainda não esteja em execução) e abre o Excel com o suplemento carregado.

        ```command&nbsp;line
        npm start
        ```

    - Para testar seu suplemento no Excel na Web, execute o seguinte comando no diretório raiz do seu projeto. Quando você executar este comando, o servidor da Web local será iniciado (se ainda não estiver em execução).

        ```command&nbsp;line
        npm run start:web
        ```

        Para usar seu suplemento, abra um documento no Excel na Web e realize o sideload do suplemento seguindo as instruções em [Realizar Sideload de Suplementos do Office na Web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).

1. Na aba **Página inicial** do Excel, escolha o botão **Ativar a proteção da planilha**. Observe que a maioria dos controles da faixa de opções estão desabilitados (e esmaecidos visualmente) como visto na captura de tela a seguir.

    ![Captura de tela da faixa de opções do Excel com o botão Alternar Proteção de Planilha realçado e habilitado. A maioria dos outros botões aparecem cinza e desabilitados.](../images/excel-tutorial-ribbon-with-protection-on-2.png)

1. Escolha uma célula como faria se quisesse alterar seu conteúdo. O Excel exibe uma mensagem de erro indicando que a planilha está protegida.

1. Escolha o botão **Proteger Planilha** novamente. Os controles são reabilitados, e você pode alterar os valores das células novamente.

## <a name="open-a-dialog"></a>Abrir uma caixa de diálogo

Nesta etapa final do tutorial, você abre uma caixa de diálogo no suplemento, passa uma mensagem do processo de caixa de diálogo para o processo de painel de tarefas e fecha a caixa de diálogo. As caixas de diálogo do Suplemento do Office são *não modais*: o usuário pode continuar a interagir com o documento no aplicativo do Office e com a página host no painel de tarefas.

### <a name="create-the-dialog-page"></a>Crie a página da caixa de diálogo

1. Na pasta **./src** localizada na raiz do projeto, crie uma pasta chamada **dialogs**.

1. Na pasta **./src/dialogs**, crie um novo arquivo chamado **popup.html**.

1. Adicione a seguinte marcação a **popup.html**. Observação:

   - A página possui um campo `<input>` onde o usuário digitará seu nome e um botão que enviará esse nome para o painel de tarefas em que será exibido.

   - a marcação carrega um script chamado **popup.js** que você criará em uma etapa posterior.

   - Ela também carrega a biblioteca Office.js porque esta será usada em **popup.js**.

    ```html
    <!DOCTYPE html>
    <html>
        <head lang="en">
            <title>Dialog for My Office Add-in</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">

            <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui. -->
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

1. Na pasta **./src/dialogs**, crie um arquivo chamado **popup.js**.

1. Adicione o código a seguir a **popup.js**. Observe o seguinte sobre este código.

   - *Todas as páginas que chamam APIs na biblioteca Office.JS devem primeiro garantir que a biblioteca tenha sido totalmente inicializada.* A melhor maneira de fazer isso é chamando o método `Office.onReady()`. Se o suplemento possuir as próprias tarefas de inicialização, o código deverá ser colocado em um método `then()` encadeado à chamada de `Office.onReady()`. A chamada de `Office.onReady()` deve ser executada antes de qualquer chamada para Office.js; por isso, a tarefa se encontra em um arquivo de script que é carregado pela página, como neste caso.

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

1. Substitua `TODO1` pelo código a seguir. Você criará a função `sendStringToParentPage` na próxima etapa.

    ```js
    document.getElementById("ok-button").onclick = sendStringToParentPage;
    ```

1. Substitua `TODO2` pelo código a seguir. O método `messageParent` passa seu parâmetro para a página pai, nesse caso, a página no painel de tarefas. O parâmetro pode ser um booliano ou uma cadeia de caracteres, que inclui tudo o que pode ser serializado como uma cadeia de caracteres, como XML ou JSON., ou qualquer tipo que possa ser convertido em uma cadeia de caracteres.

    ```js
    function sendStringToParentPage() {
        const userName = document.getElementById("name-box").value;
        Office.context.ui.messageParent(userName);
    }
    ```

> [!NOTE]
> O arquivo **popup.html** e o arquivo **popup.js** que ele carrega são executados em um processo do Internet Explorer completamente separado do painel de tarefas do suplemento. Se o **popup.js** foi transcompilado no mesmo arquivo **bundle.js** que o arquivo **app.js**, o suplemento precisará carregar duas cópias do arquivo **bundle.js**, o que anula o propósito do agrupamento. Portanto, esse suplemento não transcompila o arquivo **popup.js**.

### <a name="update-webpack-config-settings"></a>Atualizar as configurações webpack config

Abra o arquivo **webpack.config.js** no diretório raiz do projeto e conclua as seguintes etapas.

1. Localize o objeto `entry` dentro do objeto `config` e adicione uma nova entrada para `popup`.

    ```js
    popup: "./src/dialogs/popup.js"
    ```

    Após fazer isso, o novo objeto `entry` terá a seguinte aparência.

    ```js
    entry: {
      polyfill: "@babel/polyfill",
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
      popup: "./src/dialogs/popup.js"
    },
    ```
  
1. Localize a matriz `plugins` no objeto `config` e adicione o seguinte objeto ao final dela.

    ```js
    new HtmlWebpackPlugin({
      filename: "popup.html",
      template: "./src/dialogs/popup.html",
      chunks: ["polyfill", "popup"]
    })
    ```

    Após fazer isso, a nova `plugins` matriz terá a seguinte aparência.

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

1. Se o servidor da Web local estiver em execução, feche a janela de comando do nó para interrompê-lo.

1. Execute o seguinte comando para recriar o projeto.

    ```command&nbsp;line
    npm run build
    ```

### <a name="open-the-dialog-from-the-task-pane"></a>Abra a caixa de diálogo do painel de tarefas

1. Abra o arquivo **./src/taskpane/taskpane.html**.

1. Localize o elemento `<button>` do botão `freeze-header` e adicione a seguinte marcação logo após essa linha.

    ```html
    <button class="ms-Button" id="open-dialog">Open Dialog</button><br/><br/>
    ```

1. A caixa de diálogo pedirá que o usuário insira um nome e passe o nome do usuário para o painel de tarefas. O painel de tarefas o exibirá em um rótulo. Imediatamente após o `button` que você acabou de adicionar, adicione a seguinte marcação.

    ```html
    <label id="user-name"></label><br/><br/>
    ```

1. Abra o arquivo **./src/taskpane/taskpane.js**.

1. Na chamada do método `Office.onReady`, localize a linha que atribui um manipulador de cliques ao botão `freeze-header` e adicione o código a seguir logo após essa linha. Você criará o método `openDialog` em uma etapa posterior.

    ```js
    document.getElementById("open-dialog").onclick = openDialog;
    ```

1. Adicione a seguinte declaração ao final do arquivo. Esta variável é usada para armazenar um objeto no contexto de execução da página pai que atua como um intermediário ao contexto de execução da página de diálogo.

    ```js
    let dialog = null;
    ```

1. Adicione a seguinte função ao final do arquivo (depois da declaração de `dialog`). O importante a observar sobre esse código é o que *não* está lá: não há chamada de `Excel.run`. Isso ocorre porque a API para abrir uma caixa de diálogo é compartilhada entre todos os aplicativos do Office, por isso, faz parte da API comum do JavaScript do Office, não da API específica do Excel.

    ```js
    function openDialog() {
        // TODO1: Call the Office Common API that opens a dialog
    }
    ```

1. Substitua `TODO1` pelo código a seguir. Observação:

   - O método`displayDialogAsync` abre uma caixa de diálogo no centro da tela.

   - O primeiro parâmetro é a URL da página a ser aberta.

   - O segundo parâmetro passa opções. `height` e `width` são porcentagens do tamanho da janela do aplicativo do Office.

    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},

        // TODO2: Add callback parameter.
    );
    ```

### <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a>Processar a mensagem da caixa de diálogo e depois fechá-la

1. Na função `openDialog` no arquivo **./src/taskpane/taskpane.js**, substitua `TODO2` pelo seguinte código. Observação:

   - O retorno de chamada é executado imediatamente depois que a caixa de diálogo é aberta com êxito e antes de usuário executar a ação na caixa de diálogo.

   - O `result.value` é o objeto que atua como intermediário entre os contextos de execução das páginas pai e de diálogo.

   - A função `processMessage` será criada em uma etapa posterior. Esse identificador processará os valores que sejam enviados da página da caixa de diálogo com chamadas da função `messageParent`.

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
    ```

1. Adicione a seguinte função após a função `openDialog`.

    ```js
    function processMessage(arg) {
        document.getElementById("user-name").innerHTML = arg.message;
        dialog.close();
    }
    ```

1. Verifique se você salvou todas as alterações feitas no projeto.

### <a name="test-the-add-in"></a>Testar o suplemento

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

1. Se o painel de tarefas do suplemento ainda não estiver aberto no Excel, vá para a guia **Página Inicial** e escolha o botão **Mostrar Painel de Tarefas** na faixa de opções para abri-lo.

1. Escolha o botão **Abrir Caixa de Diálogo** no painel de tarefas.

1. Enquanto a caixa de diálogo estiver aberta, arraste-a e redimensione-a. Observe que você pode interagir com a planilha e pressionar outros botões no painel de tarefas, mas não pode iniciar uma segunda caixa de diálogo a partir da mesma página do painel de tarefas.

1. Na caixa de diálogo, insira um nome e selecione o botão **OK**. O nome é exibido no painel de tarefas e a caixa de diálogo é fechada.

1. Opcionalmente, comente a linha `dialog.close();` na função `processMessage`. Em seguida, repita as etapas desta seção. A caixa de diálogo permanece aberta e você pode alterar o nome. É possível fechá-la manualmente pressionando o botão **X** no canto superior direito.

    ![Captura de tela do Excel, com um botão Abrir Caixa de Diálogo visível no painel de tarefas do suplemento e uma caixa de diálogo exibida sobre a planilha.](../images/excel-tutorial-dialog-open-2.png)

## <a name="next-steps"></a>Próximas etapas

Neste tutorial você criou um suplemento do Excel que interage com tabelas, gráficos, planilhas e caixas de diálogo em uma pasta de trabalho do Excel. Para saber mais sobre o desenvolvimento de suplementos do Excel, continue no artigo a seguir.

> [!div class="nextstepaction"]
> [Visão geral dos suplementos do Excel](../excel/excel-add-ins-overview.md)

## <a name="see-also"></a>Confira também

- [Visão geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
- [Desenvolver Suplementos do Office](../develop/develop-overview.md)
- [Modelo de objeto JavaScript do Excel em Suplementos do Office](../excel/excel-add-ins-core-concepts.md)
