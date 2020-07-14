A API JavaScript do Office inclui dois modelos diferentes:

- As APIs **Específicas do host** fornecem objetos fortemente tipados que podem ser usados para interagir com objetos que são nativos de um aplicativos do Office específico. Por exemplo, você pode usar as APIs JavaScript do Excel para acessar planilhas, intervalos, tabelas, gráficos e mais. Atualmente, as APIs específicas de host estão disponíveis para os seguintes hosts:

    - [Excel](../reference/overview/excel-add-ins-reference-overview.md)

    - [Word](../reference/overview/word-add-ins-reference-overview.md)

    - [OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)

    Esse modelo de API usa [promessas](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) e permite que você especifique várias operações em cada solicitação enviada ao host do Office. Dessa maneira, operações de envio em lote podem melhorar significativamente o desempenho do suplemento em aplicativos do Office na Web. As APIs específicas do host foram introduzidas com o Office 2016 e não pode ser usadas para interagir com o Office 2013.

    > [!NOTE]
    > Também há uma API específica do host para o [Visio](../reference/overview/visio-javascript-reference-overview.md), mas você só pode usá-la nas páginas do SharePoint Online para interagir com os diagramas do Visio que foram incorporados na página. Os suplementos da Web do Office não são compatíveis com o Visio.

- As APIs** Comuns** pode ser usada para acessar recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office. Esse modelo de API usa [retornos de chamada](https://developer.mozilla.org/docs/Glossary/Callback_function), que permitem especificar apenas uma operação em cada solicitação enviada ao host do Office. As APIs comuns foram introduzidas com o Office 2013 e podem ser usadas para interagir com o Office 2013 ou posterior. Para saber mais sobre o modelo de objeto da API Comum, que inclui APIs para interagir com o Outlook, o PowerPoint e o Project, confira [Modelo de objeto da API JavaScript comum](../develop/office-javascript-api-object-model.md).

> [!NOTE]
> Algumas funções personalizadas do Excel são executadas em um tempo de execução exclusivo que prioriza a execução de cálculos e não têm um painel de tarefas. Essas funções usam um modelo de programação ligeiramente diferente e são chamadas de funções sem interface do usuário.
