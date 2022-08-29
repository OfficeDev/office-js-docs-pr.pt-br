A API JavaScript do Office inclui dois modelos diferentes:

- As APIs **específicas do aplicativo** fornecem objetos fortemente tipados que podem ser usados para interagir com objetos que são nativos de um aplicativo específico do Office. Por exemplo, você pode usar as APIs JavaScript do Excel para acessar planilhas, intervalos, tabelas, gráficos e mais. As APIs específicas do aplicativo estão disponíveis atualmente para os seguintes aplicativos do Office.

    - [Excel](../reference/overview/excel-add-ins-reference-overview.md)
    - [OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)
    - [PowerPoint](../reference/overview/powerpoint-add-ins-reference-overview.md)
    - [Word](../reference/overview/word-add-ins-reference-overview.md)

    Esse modelo de API usa [promessas](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) e permite que você especifique várias operações em cada solicitação enviada ao aplicativo do Office. Dessa maneira, operações de envio em lote podem melhorar significativamente o desempenho do suplemento em aplicativos do Office na Web. As APIs específicas do aplicativo foram introduzidas com o Office 2016 e não podem ser usadas para interagir com o Office 2013.

    > [!NOTE]
    > Também há uma API específica do aplicativo para o [Visio](../reference/overview/visio-javascript-reference-overview.md), mas você só pode usá-la nas páginas do SharePoint Online para interagir com os diagramas do Visio que foram incorporados na página. Os suplementos da Web do Office não são compatíveis com o Visio.

    Confira [Usando o modelo de API específico do aplicativo](../develop/application-specific-api-model.md) para saber mais sobre esse modelo de API.

- As APIs **Comuns** pode ser usada para acessar recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office. Esse modelo de API usa [retornos de chamada](https://developer.mozilla.org/docs/Glossary/Callback_function), que permitem especificar apenas uma operação em cada solicitação enviada ao aplicativo do Office. As APIs comuns foram introduzidas com o Office 2013 e podem ser usadas para interagir com o Office 2013 ou posterior. Para saber mais sobre o modelo de objeto da API Comum, que inclui APIs para interagir com o Outlook, o PowerPoint e o Project, confira [Modelo de objeto da API JavaScript comum](../develop/office-javascript-api-object-model.md).

> [!NOTE]
>Funções personalizadas sem uma [execução de runtime compartilhado](../testing/runtimes.md#shared-runtime) em um [runtime somente JavaScript](../testing/runtimes.md#javascript-only-runtime) que prioriza a execução de cálculos. Essas funções usam um modelo de programação ligeiramente diferente.
