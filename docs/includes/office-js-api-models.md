A API JavaScript do Office inclui dois modelos diferentes:

- As APIs **Específicas do host** fornecem objetos fortemente tipados que podem ser usados para interagir com objetos que são nativos de um aplicativos do Office específico. Por exemplo, você pode usar as APIs JavaScript do Excel para acessar planilhas, intervalos, tabelas, gráficos e mais. Atualmente, as APIs específicas de host estão disponíveis para os seguintes hosts:

    - [Excel](../reference/overview/excel-add-ins-reference-overview.md)

    - [Word](../reference/overview/word-add-ins-reference-overview.md)

    - [OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)

    Esse modelo de API usa [promessas](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) e permite que você especifique várias operações em cada solicitação enviada ao host do Office. Operações de envio em lote dessa maneira podem melhorar significativamente o desempenho do suplemento no Office nos aplicativos Web. As APIs específicas do host foram introduzidas com o Office 2016 e não pode ser usadas para interagir com o Office 2013.

- As APIs **Comuns** podem ser usadas para acessar recursos como interface do usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office. Este modelo de API usa [retornos de chamada](https://developer.mozilla.org/docs/Glossary/Callback_function), em que você pode especificar apenas uma operação em cada solicitação enviada ao host do Office. As APIs comuns foram introduzidas com o Office 2013 e podem ser usadas para interagir com o Office 2013 ou posterior. Para obter mais detalhes do modelo de objeto API comum, que inclui APIs para interagir com o Outlook e o PowerPoint, consulte [modelo do objeto JavaScript API comum](../develop/office-javascript-api-object-model.md).

> [!NOTE]
> As funções personalizadas do Excel são executadas dentro de um tempo de execução único que prioriza a execução de cálculos e, portanto, usa um modelo de programação ligeiramente diferente. Para saber mais, confira [Arquitetura de funções personalizadas](../excel/custom-functions-architecture.md).