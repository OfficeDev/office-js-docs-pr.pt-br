---
title: APIs de visualização do JavaScript para Excel
description: Detalhes sobre as próximas Excel APIs JavaScript.
ms.date: 07/23/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d90c5e8bb2c344cb3bb297a3cd793613f017e910ab99df6dfffc456c3f715d20
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092639"
---
# <a name="excel-javascript-preview-apis"></a>APIs de visualização do JavaScript para Excel

As novas APIs do JavaScript para Excel são introduzidas pela primeira vez na "visualização" e, posteriormente, tornam-se parte de um conjunto específico de requisitos numerados, após passarem por vários testes e após a recolha das opiniões de usuários.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

A tabela a seguir fornece um resumo conciso das APIs, enquanto a tabela de lista [de API](#api-list) subsequente fornece uma lista detalhada.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| Tabelas de dados de gráfico | Controlar a aparência, a formatação e a visibilidade das tabelas de dados nos gráficos. | [Chart,](/javascript/api/excel/excel.chart) [ChartDataTable,](/javascript/api/excel/excel.chartdatatable) [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| Tarefas do documento | Transforme os comentários em tarefas atribuídas aos usuários. | [DocumentTask](/javascript/api/excel/excel.documenttask) |
| Identidades | Gerencie identidades de usuário, incluindo nome de exibição e endereço de email. | [Identity](/javascript/api/excel/excel.identity), [IdentityCollection](/javascript/api/excel/excel.identitycollection), [IdentityEntity](/javascript/api/excel/excel.identityentity) |
| Tipos de dados vinculados | Adiciona suporte para tipos de dados conectados Excel de fontes externas. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| Guias de trabalho vinculadas | Gerencie links entre as guias de trabalho, incluindo o suporte para atualizar e quebrar links de livros de trabalho. | [LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook), [LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection) |
| Table styles | Fornece controle para fonte, borda, cor de preenchimento e outros aspectos dos estilos de tabela. | [Tabela,](/javascript/api/excel/excel.table) [Tabela Dinâmica,](/javascript/api/excel/excel.pivottable) [Slicer](/javascript/api/excel/excel.slicer) |
| Consultas | Recupere atributos de consulta, como nome, data de atualização e contagem de consultas. | [Consulta](/javascript/api/excel/excel.query), [QueryCollection](/javascript/api/excel/excel.querycollection)|

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as Excel APIs JavaScript atualmente em visualização. Para uma lista completa de todas as EXCEL JavaScript (incluindo APIs de visualização e APIs lançadas anteriormente), consulte todas as [APIs JavaScript](/javascript/api/excel?view=excel-js-preview&preserve-view=true)Excel JavaScript .

| Classe | Campos | Descrição |
|:---|:---|:---|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#deleteShiftDirection)|Representa a direção (como para cima ou para a esquerda) que as células restantes serão deslocadas quando uma célula ou células são excluídas.|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#insertShiftDirection)|Representa a direção (como para baixo ou para a direita) que as células existentes mudarão quando uma nova célula ou células são inseridas.|
|[Gráfico](/javascript/api/excel/excel.chart)|[getDataTable()](/javascript/api/excel/excel.chart#getDataTable__)|Obtém a tabela de dados no gráfico.|
||[getDataTableOrNullObject()](/javascript/api/excel/excel.chart#getDataTableOrNullObject__)|Obtém a tabela de dados no gráfico.|
|[ChartDataTable](/javascript/api/excel/excel.chartdatatable)|[format](/javascript/api/excel/excel.chartdatatable#format)|Representa o formato de uma tabela de dados do gráfico, que inclui o formato de preenchimento, fonte e borda.|
||[showHorizontalBorder](/javascript/api/excel/excel.chartdatatable#showHorizontalBorder)|Especifica se será exibida a borda horizontal da tabela de dados.|
||[showLegendKey](/javascript/api/excel/excel.chartdatatable#showLegendKey)|Especifica se a chave de legenda da tabela de dados deve ser apresentada.|
||[showOutlineBorder](/javascript/api/excel/excel.chartdatatable#showOutlineBorder)|Especifica se será exibida a borda de contorno da tabela de dados.|
||[showVerticalBorder](/javascript/api/excel/excel.chartdatatable#showVerticalBorder)|Especifica se será exibida a borda vertical da tabela de dados.|
||[visible](/javascript/api/excel/excel.chartdatatable#visible)|Especifica se é preciso mostrar a tabela de dados do gráfico.|
|[ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat)|[border](/javascript/api/excel/excel.chartdatatableformat#border)|Representa o formato de borda da tabela de dados do gráfico, que inclui cor, estilo de linha e peso.|
||[fill](/javascript/api/excel/excel.chartdatatableformat#fill)|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo.|
||[font](/javascript/api/excel/excel.chartdatatableformat#font)|Representa os atributos de fonte (como nome da fonte, tamanho da fonte e cor) do objeto atual.|
|[Comentário](/javascript/api/excel/excel.comment)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.comment#assignTask_assignee_)|Atribui a tarefa anexada ao comentário ao usuário dado como um destinatário.|
||[getTask()](/javascript/api/excel/excel.comment#getTask__)|Obtém a tarefa associada a este comentário.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#getTaskOrNullObject__)|Obtém a tarefa associada a este comentário.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#getItemOrNullObject_commentId_)|Obtém um comentário da coleção com base em seu ID.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(assignee: Identity)](/javascript/api/excel/excel.commentreply#assignTask_assignee_)|Atribui a tarefa anexada ao comentário ao usuário determinado como o único destinatário.|
||[getTask()](/javascript/api/excel/excel.commentreply#getTask__)|Obtém a tarefa associada ao thread desta resposta de comentário.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#getTaskOrNullObject__)|Obtém a tarefa associada ao thread desta resposta de comentário.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getItemOrNullObject_commentReplyId_)|Retorna uma resposta de comentário identificada pela respectiva ID.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: string)](/javascript/api/excel/excel.conditionalformatcollection#getItemOrNullObject_id_)|Retorna um formato condicional identificado por sua ID.|
|[DocumentTask](/javascript/api/excel/excel.documenttask)|[percentComplete](/javascript/api/excel/excel.documenttask#percentComplete)|Especifica a porcentagem de conclusão da tarefa.|
||[prioridade](/javascript/api/excel/excel.documenttask#priority)|Especifica a prioridade da tarefa.|
||[assignees](/javascript/api/excel/excel.documenttask#assignees)|Retorna uma coleção de atribuídos da tarefa.|
||[changes](/javascript/api/excel/excel.documenttask#changes)|Obtém os registros de alteração da tarefa.|
||[comment](/javascript/api/excel/excel.documenttask#comment)|Obtém o comentário associado à tarefa.|
||[completedBy](/javascript/api/excel/excel.documenttask#completedBy)|Obtém o usuário mais recente para ter concluído a tarefa.|
||[completedDateTime](/javascript/api/excel/excel.documenttask#completedDateTime)|Obtém a data e a hora em que a tarefa foi concluída.|
||[createdBy](/javascript/api/excel/excel.documenttask#createdBy)|Obtém o usuário que criou a tarefa.|
||[createdDateTime](/javascript/api/excel/excel.documenttask#createdDateTime)|Obtém a data e a hora em que a tarefa foi criada.|
||[id](/javascript/api/excel/excel.documenttask#id)|Obtém a ID da tarefa.|
||[setStartAndDueDateTime(startDateTime: Date, dueDateTime: Date)](/javascript/api/excel/excel.documenttask#setStartAndDueDateTime_startDateTime__dueDateTime_)|Altera o início e as datas de vencimento da tarefa.|
||[startAndDueDateTime](/javascript/api/excel/excel.documenttask#startAndDueDateTime)|Obtém ou define a data e a hora em que a tarefa deve começar e deve ser final.|
||[title](/javascript/api/excel/excel.documenttask#title)|Especifica o título da tarefa.|
|[DocumentTaskChange](/javascript/api/excel/excel.documenttaskchange)|[assignee](/javascript/api/excel/excel.documenttaskchange#assignee)|Representa o usuário atribuído à tarefa para um tipo de registro de alteração ou o usuário não atribuído da tarefa para um tipo `assign` `unassign` de registro de alteração.|
||[changedBy](/javascript/api/excel/excel.documenttaskchange#changedBy)|Representa o usuário que criou ou alterou a tarefa.|
||[commentId](/javascript/api/excel/excel.documenttaskchange#commentId)|Representa a ID do `Comment` ou ao qual a alteração da tarefa está `CommentReply` ancorada.|
||[createdDateTime](/javascript/api/excel/excel.documenttaskchange#createdDateTime)|Representa a data e a hora de criação do registro de alteração de tarefa.|
||[dueDateTime](/javascript/api/excel/excel.documenttaskchange#dueDateTime)|Representa a data e a hora de vencimento da tarefa, no fuso horário UTC.|
||[id](/javascript/api/excel/excel.documenttaskchange#id)|ID do registro de alteração de tarefa.|
||[percentComplete](/javascript/api/excel/excel.documenttaskchange#percentComplete)|Representa a porcentagem de conclusão da tarefa.|
||[prioridade](/javascript/api/excel/excel.documenttaskchange#priority)|Representa a prioridade da tarefa.|
||[startDateTime](/javascript/api/excel/excel.documenttaskchange#startDateTime)|Representa a data e a hora de início da tarefa, no fuso horário UTC.|
||[title](/javascript/api/excel/excel.documenttaskchange#title)|Representa o título da tarefa.|
||[type](/javascript/api/excel/excel.documenttaskchange#type)|Representa o tipo de ação do registro de alteração de tarefa.|
||[undoHistoryId](/javascript/api/excel/excel.documenttaskchange#undoHistoryId)|Representa a `DocumentTaskChange.id` propriedade que foi desfeita para o tipo `undo` de registro de alteração.|
|[DocumentTaskChangeCollection](/javascript/api/excel/excel.documenttaskchangecollection)|[getCount()](/javascript/api/excel/excel.documenttaskchangecollection#getCount__)|Obtém o número de registros de alteração na coleção da tarefa.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskchangecollection#getItemAt_index_)|Obtém um registro de alteração de tarefa usando seu índice na coleção.|
||[items](/javascript/api/excel/excel.documenttaskchangecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[DocumentTaskCollection](/javascript/api/excel/excel.documenttaskcollection)|[getCount()](/javascript/api/excel/excel.documenttaskcollection#getCount__)|Obtém o número de tarefas na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.documenttaskcollection#getItem_key_)|Obtém uma tarefa usando sua ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.documenttaskcollection#getItemAt_index_)|Obtém uma tarefa pelo índice na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.documenttaskcollection#getItemOrNullObject_key_)|Obtém uma tarefa usando sua ID.|
||[items](/javascript/api/excel/excel.documenttaskcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[DocumentTaskSchedule](/javascript/api/excel/excel.documenttaskschedule)|[dueDateTime](/javascript/api/excel/excel.documenttaskschedule#dueDateTime)|Obtém a data e a hora de vencimento da tarefa.|
||[startDateTime](/javascript/api/excel/excel.documenttaskschedule#startDateTime)|Obtém a data e a hora em que a tarefa deve começar.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getItemOrNullObject_key_)|Obtém uma forma usando seu nome ou ID.|
|[Identidade](/javascript/api/excel/excel.identity)|[displayName](/javascript/api/excel/excel.identity#displayName)|Representa o nome para exibição do usuário.|
||[email](/javascript/api/excel/excel.identity#email)|Representa o endereço de email do usuário.|
||[id](/javascript/api/excel/excel.identity#id)|Representa a ID exclusiva do usuário.|
|[IdentityCollection](/javascript/api/excel/excel.identitycollection)|[add(assignee: Identity)](/javascript/api/excel/excel.identitycollection#add_assignee_)|Adiciona uma identidade de usuário à coleção.|
||[clear()](/javascript/api/excel/excel.identitycollection#clear__)|Remove todas as identidades de usuário da coleção.|
||[getCount()](/javascript/api/excel/excel.identitycollection#getCount__)|Obtém o número de itens na coleção.|
||[getItemAt(index: number)](/javascript/api/excel/excel.identitycollection#getItemAt_index_)|Obtém uma identidade de usuário de documento usando seu índice na coleção.|
||[items](/javascript/api/excel/excel.identitycollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[remove(assignee: Identity)](/javascript/api/excel/excel.identitycollection#remove_assignee_)|Remove uma identidade de usuário da coleção.|
|[IdentityEntity](/javascript/api/excel/excel.identityentity)|[displayName](/javascript/api/excel/excel.identityentity#displayName)|Representa o nome para exibição do usuário.|
||[email](/javascript/api/excel/excel.identityentity#email)|Representa o endereço de email do usuário.|
||[id](/javascript/api/excel/excel.identityentity#id)|Representa a ID exclusiva do usuário.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataProvider)|O nome do provedor de dados do tipo de dados vinculado.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastRefreshed)|A data e a hora do fuso horário local desde que a lista de trabalho foi aberta quando o tipo de dados vinculado foi atualizado pela última vez.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|O nome do tipo de dados vinculado.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicRefreshInterval)|A frequência, em segundos, na qual o tipo de dados vinculado é atualizado se `refreshMode` estiver definido como "Periódico".|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshMode)|O mecanismo pelo qual os dados do tipo de dados vinculados são recuperados.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceId)|A ID exclusiva do tipo de dados vinculado.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedRefreshModes)|Retorna uma matriz com todos os modos de atualização suportados pelo tipo de dados vinculado.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestRefresh__)|Faz uma solicitação para atualizar o tipo de dados vinculado.|
||[requestSetRefreshMode(refreshMode: Excel. LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestSetRefreshMode_refreshMode_)|Faz uma solicitação para alterar o modo de atualização para esse tipo de dados vinculado.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceId)|A ID exclusiva do novo tipo de dados vinculado.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Obtém o tipo do evento.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getCount__)|Obtém o número de tipos de dados vinculados na coleção.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItem_key_)|Obtém um tipo de dados vinculado por ID de serviço.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemAt_index_)|Obtém um tipo de dados vinculado pelo índice na coleção.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getItemOrNullObject_key_)|Obtém um tipo de dados vinculado por ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestRefreshAll__)|Faz uma solicitação para atualizar todos os tipos de dados vinculados na coleção.|
|[LinkedWorkbook](/javascript/api/excel/excel.linkedworkbook)|[breakLinks()](/javascript/api/excel/excel.linkedworkbook#breakLinks__)|Faz uma solicitação para quebrar os links apontando para a lista de trabalho vinculada.|
||[id](/javascript/api/excel/excel.linkedworkbook#id)|A URL original apontando para a lista de trabalho vinculada.|
||[refresh()](/javascript/api/excel/excel.linkedworkbook#refresh__)|Faz uma solicitação para atualizar os dados recuperados da lista de trabalho vinculada.|
|[LinkedWorkbookCollection](/javascript/api/excel/excel.linkedworkbookcollection)|[breakAllLinks()](/javascript/api/excel/excel.linkedworkbookcollection#breakAllLinks__)|Quebra todos os links para as guias de trabalho vinculadas.|
||[getItem(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItem_key_)|Obtém informações sobre uma lista de trabalho vinculada por sua URL.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.linkedworkbookcollection#getItemOrNullObject_key_)|Obtém informações sobre uma lista de trabalho vinculada por sua URL.|
||[items](/javascript/api/excel/excel.linkedworkbookcollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[refreshAll()](/javascript/api/excel/excel.linkedworkbookcollection#refreshAll__)|Faz uma solicitação para atualizar todos os links da workbook.|
||[workbookLinksRefreshMode](/javascript/api/excel/excel.linkedworkbookcollection#workbookLinksRefreshMode)|Representa o modo de atualização dos links da agenda de trabalho.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getItemOrNullObject_key_)|Obtém uma exibição de planilha usando seu nome.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getCell_dataHierarchy__rowItems__columnItems_)|Obtém uma célula exclusiva na tabela dinâmica com base em uma hierarquia de dados, bem como os itens de linha e coluna de suas respectivas hierarquias.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotStyle)|O estilo aplicado à Tabela Dinâmica.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setStyle_style_)|Define o estilo aplicado à Tabela Dinâmica.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getFirstOrNullObject()](/javascript/api/excel/excel.pivottablescopedcollection#getFirstOrNullObject__)|Obtém a primeira Tabela Dinâmica da coleção.|
|[Query](/javascript/api/excel/excel.query)|[error](/javascript/api/excel/excel.query#error)|Obtém a mensagem de erro de consulta de quando a consulta foi atualizada pela última vez.|
||[loadedTo](/javascript/api/excel/excel.query#loadedTo)|Obtém o tipo de objeto de consulta "carregado para".|
||[loadedToDataModel](/javascript/api/excel/excel.query#loadedToDataModel)|Especifica se a consulta foi carregada para o Modelo de Dados.|
||[name](/javascript/api/excel/excel.query#name)|Obtém o nome da consulta.|
||[refreshDate](/javascript/api/excel/excel.query#refreshDate)|Obtém a data e a hora em que a consulta foi atualizada pela última vez.|
||[rowsLoadedCount](/javascript/api/excel/excel.query#rowsLoadedCount)|Obtém o número de linhas que foram carregadas quando a consulta foi atualizada pela última vez.|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[getCount()](/javascript/api/excel/excel.querycollection#getCount__)|Obtém o número de consultas na guia de trabalho.|
||[getItem(key: string)](/javascript/api/excel/excel.querycollection#getItem_key_)|Obtém uma consulta da coleção com base em seu nome.|
||[items](/javascript/api/excel/excel.querycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Range](/javascript/api/excel/excel.range)|[getDependents()](/javascript/api/excel/excel.range#getDependents__)|Retorna um objeto que representa o intervalo que contém todos os dependentes de uma célula na mesma planilha ou `WorkbookRangeAreas` em várias planilhas.|
||[getPrecedents()](/javascript/api/excel/excel.range#getPrecedents__)|Retorna um objeto que representa o intervalo que contém todos os precedentes de uma célula na mesma planilha ou `WorkbookRangeAreas` em várias planilhas.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshMode)|O modo de atualização do tipo de dados vinculado.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceId)|A ID exclusiva do objeto cujo modo de atualização foi alterado.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Obtém o tipo do evento.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[atualizado](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Indica se a solicitação de atualização foi bem-sucedida.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceId)|A ID exclusiva do objeto cuja solicitação de atualização foi concluída.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Obtém o tipo do evento.|
||[avisos](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Uma matriz que contém quaisquer avisos gerados a partir da solicitação de atualização.|
|[Formato](/javascript/api/excel/excel.shape)|[displayName](/javascript/api/excel/excel.shape#displayName)|Obtém o nome de exibição da forma.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addSvg_xml_)|Cria um gráfico vetorial escalável (SVG) de uma cadeia de caracteres XML e a adiciona à planilha.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getItemOrNullObject_key_)|Obtém uma forma usando seu nome ou ID.|
|[Segmentação de dados](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameInFormula)|Representa o nome da segmentação de dados usada na fórmula.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerStyle)|O estilo aplicado à slicer.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setStyle_style_)|Define o estilo aplicado à slicer.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(name: string)](/javascript/api/excel/excel.stylecollection#getItemOrNullObject_name_)|Obtém um estilo por nome.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearStyle__)|Altera a tabela para usar o estilo de tabela padrão.|
||[onFiltered](/javascript/api/excel/excel.table#onFiltered)|Ocorre quando um filtro é aplicado em uma tabela específica.|
||[tableStyle](/javascript/api/excel/excel.table#tableStyle)|O estilo aplicado à tabela.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setStyle_style_)|Define o estilo aplicado à tabela.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onFiltered)|Ocorre quando um filtro é aplicado em qualquer tabela em uma pasta de trabalho ou em uma planilha.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableId)|Obtém a ID da tabela na qual o filtro é aplicado.|
||[tipo](/javascript/api/excel/excel.tablefilteredeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetId)|Obtém a ID da planilha que contém a tabela.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[deleteRows(rows: number[] \| TableRow[])](/javascript/api/excel/excel.tablerowcollection#deleteRows_rows_)|Exclua várias linhas de uma tabela.|
||[deleteRowsAt(index: number, count?: number)](/javascript/api/excel/excel.tablerowcollection#deleteRowsAt_index__count_)|Exclua um número especificado de linhas de uma tabela, começando em um determinado índice.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getItemOrNullObject_key_)|Obtém uma tabela pelo nome ou ID.|
|[Pasta de trabalho](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkedDataTypes)|Retorna uma coleção de tipos de dados vinculados que fazem parte da lista de trabalho.|
||[linkedWorkbooks](/javascript/api/excel/excel.workbook#linkedWorkbooks)|Retorna uma coleção de guias de trabalho vinculadas.|
||[consultas](/javascript/api/excel/excel.workbook#queries)|Retorna uma coleção de consultas do Power Query que fazem parte da workbook.|
||[tasks](/javascript/api/excel/excel.workbook#tasks)|Retorna uma coleção de tarefas que estão presentes na workbook.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showPivotFieldList)|Especifica se o painel de lista de campos da Tabela Dinâmica é mostrado no nível da lista de trabalho.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904DateSystem)|True se a pasta de trabalho usar o sistema de dados 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFiltered](/javascript/api/excel/excel.worksheet#onFiltered)|Ocorre quando um filtro é aplicado em uma planilha específica.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheet#onProtectionChanged)|Ocorre quando o estado de proteção da planilha é alterado.|
||[tabId](/javascript/api/excel/excel.worksheet#tabId)|Retorna um valor que representa essa planilha que pode ser lido por Open Office XML.|
||[tasks](/javascript/api/excel/excel.worksheet#tasks)|Retorna uma coleção de tarefas presentes na planilha.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#changeDirectionState)|Representa uma alteração na direção em que as células de uma planilha serão deslocadas quando uma célula ou células são excluídas ou inseridas.|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#triggerSource)|Representa a origem do gatilho do evento.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addFromBase64_base64File__sheetNamesToInsert__positionType__relativeTo_)|Insere as planilhas especificadas de uma pasta de trabalho na pasta de trabalho atual.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onFiltered)|Ocorre quando filtro de uma planilha é aplicado na pasta de trabalho.|
||[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#onProtectionChanged)|Ocorre quando o estado de proteção da planilha é alterado.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[tipo](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetId)|Obtém a ID da planilha na qual o filtro é aplicado.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#isProtected)|Obtém o status de proteção atual da planilha.|
||[source](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#source)|A origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#worksheetId)|Obtém a ID da planilha na qual o status da proteção é alterado.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
