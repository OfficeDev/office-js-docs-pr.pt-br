---
title: Excel conjunto de requisitos da API JavaScript 1.14
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.14.
ms.date: 12/08/2021
ms.prod: excel
ms.localizationpriority: medium
---

# <a name="whats-new-in-excel-javascript-api-114"></a>Novidades na API JavaScript 1.14 Excel JavaScript

O ExcelApi 1.14 adicionou objetos para controlar o recurso de tabela de dados de um gráfico, um método para localizar todas as células precedentes de uma fórmula e eventos de proteção de planilha para rastrear alterações no estado de proteção de uma planilha. Ele também adicionou vários [`getItemOrNullObject`](../../develop/application-specific-api-model.md#ornullobject-methods-and-properties) métodos para objetos como `CommentCollection`, `ShapeCollection`e para melhorar `StyleCollection` o tratamento de erros.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Tabelas de dados de gráfico](../../excel/excel-add-ins-charts.md#add-and-format-a-chart-data-table) | Controlar a aparência, a formatação e a visibilidade das tabelas de dados nos gráficos. | [Chart](/javascript/api/excel/excel.chart), [ChartDataTable](/javascript/api/excel/excel.chartdatatable), [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| [Precedentes de fórmula](../../excel/excel-add-ins-ranges-precedents-dependents.md#get-the-precedents-of-a-formula) | Retorne todas as células precedentes de uma fórmula. | [Range](/javascript/api/excel/excel.range) |
| Consultas | Recupere atributos de Consulta do Power, como nome, data de atualização e contagem de consultas. | [Consulta](/javascript/api/excel/excel.query), [QueryCollection](/javascript/api/excel/excel.querycollection)|
| [Eventos de proteção de planilha](../../excel/excel-add-ins-worksheets.md#detect-changes-to-the-worksheet-protection-state) | Acompanhe as alterações no estado de proteção de uma planilha e na origem dessas alterações. | [WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs), [Worksheet](/javascript/api/excel/excel.worksheet), [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) |

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.14. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.14 ou anterior, consulte Excel APIs no conjunto de requisitos [1.14](/javascript/api/excel?view=excel-js-1.14&preserve-view=true) ou anterior.

| Classe | Campos | Descrição |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-clearcolumncriteria-member(1))|Limpa os critérios de filtro de coluna do AutoFilter.|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#excel-excel-changedirectionstate-deleteshiftdirection-member)|Representa a direção (como para cima ou para a esquerda) que as células restantes serão deslocadas quando uma célula ou células são excluídas.|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#excel-excel-changedirectionstate-insertshiftdirection-member)|Representa a direção (como para baixo ou para a direita) que as células existentes mudarão quando uma nova célula ou células são inseridas.|
|[Chart](/javascript/api/excel/excel.chart)|[getDataTable()](/javascript/api/excel/excel.chart#excel-excel-chart-getdatatable-member(1))|Obtém a tabela de dados no gráfico.|
||[getDataTableOrNullObject()](/javascript/api/excel/excel.chart#excel-excel-chart-getdatatableornullobject-member(1))|Obtém a tabela de dados no gráfico.|
|[ChartDataTable](/javascript/api/excel/excel.chartdatatable)|[format](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-format-member)|Representa o formato de uma tabela de dados do gráfico, que inclui o formato de preenchimento, fonte e borda.|
||[showHorizontalBorder](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showhorizontalborder-member)|Especifica se a borda horizontal da tabela de dados deve ser exibida.|
||[showLegendKey](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showlegendkey-member)|Especifica se a chave de legenda da tabela de dados deve ser apresentada.|
||[showOutlineBorder](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showoutlineborder-member)|Especifica se a borda de contorno da tabela de dados deve ser exibida.|
||[showVerticalBorder](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showverticalborder-member)|Especifica se a borda vertical da tabela de dados deve ser exibida.|
||[visible](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-visible-member)|Especifica se é preciso mostrar a tabela de dados do gráfico.|
|[ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat)|[borda](/javascript/api/excel/excel.chartdatatableformat#excel-excel-chartdatatableformat-border-member)|Representa o formato de borda da tabela de dados do gráfico, que inclui cor, estilo de linha e peso.|
||[fill](/javascript/api/excel/excel.chartdatatableformat#excel-excel-chartdatatableformat-fill-member)|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo.|
||[font](/javascript/api/excel/excel.chartdatatableformat#excel-excel-chartdatatableformat-font-member)|Representa os atributos de fonte (como nome da fonte, tamanho da fonte e cor) do objeto atual.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitemornullobject-member(1))|Obtém um comentário da coleção com base em seu ID.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getitemornullobject-member(1))|Retorna uma resposta de comentário identificada pela respectiva ID.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-getitemornullobject-member(1))|Retorna um formato condicional identificado por sua ID.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getitemornullobject-member(1))|Obtém uma forma usando seu nome ou ID.|
|[Query](/javascript/api/excel/excel.query)|[error](/javascript/api/excel/excel.query#excel-excel-query-error-member)|Obtém a mensagem de erro de consulta de quando a consulta foi atualizada pela última vez.|
||[loadedTo](/javascript/api/excel/excel.query#excel-excel-query-loadedto-member)|Obtém a consulta carregada para o tipo de objeto.|
||[loadedToDataModel](/javascript/api/excel/excel.query#excel-excel-query-loadedtodatamodel-member)|Especifica se a consulta foi carregada para o modelo de dados.|
||[name](/javascript/api/excel/excel.query#excel-excel-query-name-member)|Obtém o nome da consulta.|
||[refreshDate](/javascript/api/excel/excel.query#excel-excel-query-refreshdate-member)|Obtém a data e a hora em que a consulta foi atualizada pela última vez.|
||[rowsLoadedCount](/javascript/api/excel/excel.query#excel-excel-query-rowsloadedcount-member)|Obtém o número de linhas que foram carregadas quando a consulta foi atualizada pela última vez.|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[getCount()](/javascript/api/excel/excel.querycollection#excel-excel-querycollection-getcount-member(1))|Obtém o número de consultas na guia de trabalho.|
||[getItem(key: string)](/javascript/api/excel/excel.querycollection#excel-excel-querycollection-getitem-member(1))|Obtém uma consulta da coleção com base em seu nome.|
||[items](/javascript/api/excel/excel.querycollection#excel-excel-querycollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents()](/javascript/api/excel/excel.range#excel-excel-range-getprecedents-member(1))|Retorna um `WorkbookRangeAreas` objeto que representa o intervalo que contém todos os precedentes de uma célula na mesma planilha ou em várias planilhas.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getitemornullobject-member(1))|Obtém uma forma usando seu nome ou ID.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getitemornullobject-member(1))|Obtém um estilo por nome.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getitemornullobject-member(1))|Obtém uma tabela pelo nome ou ID.|
|[Workbook](/javascript/api/excel/excel.workbook)|[consultas](/javascript/api/excel/excel.workbook#excel-excel-workbook-queries-member)|Retorna uma coleção de consultas do Power Query que fazem parte da workbook.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onProtectionChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onprotectionchanged-member)|Ocorre quando o estado de proteção da planilha é alterado.|
||[tabId](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tabid-member)|Retorna um valor que representa essa planilha que pode ser lido por Open Office XML.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-changedirectionstate-member)|Representa uma alteração na direção em que as células de uma planilha serão deslocadas quando uma célula ou células são excluídas ou inseridas.|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-triggersource-member)|Representa a origem do gatilho do evento.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onprotectionchanged-member)|Ocorre quando o estado de proteção da planilha é alterado.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-isprotected-member)|Obtém o status de proteção atual da planilha.|
||[source](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-source-member)|A origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-worksheetid-member)|Obtém a ID da planilha na qual o status da proteção é alterado.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.14&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
