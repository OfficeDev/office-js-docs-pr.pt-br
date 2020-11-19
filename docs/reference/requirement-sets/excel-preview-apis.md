---
title: APIs de visualização do JavaScript para Excel
description: Detalhes sobre as futuras APIs JavaScript do Excel.
ms.date: 11/17/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 083741d35d3e881c2e46b186c4e93591bf7f4834
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49131763"
---
# <a name="excel-javascript-preview-apis"></a>APIs de visualização do JavaScript para Excel

As novas APIs do JavaScript para Excel são introduzidas pela primeira vez na "visualização" e, posteriormente, tornam-se parte de um conjunto específico de requisitos numerados, após passarem por vários testes e após a recolha das opiniões de usuários.

A primeira tabela fornece um resumo conciso das APIs e, a tabela subsequente, fornece uma lista detalhada.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| Tipos de dados vinculados | Adiciona suporte para tipos de dados conectados ao Excel a partir de fontes externas. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| Exibições de planilha nomeadas | Fornece controle programático de modos de exibição de planilha por usuário. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |
| Tarefas | Transforme comentários em tarefas atribuídas aos usuários. | [Tarefa](/javascript/api/excel/excel.task) |

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs JavaScript do Excel atualmente em versão prévia. Para obter uma lista completa de todas as APIs JavaScript do Excel (incluindo APIs de visualização e APIs previamente lançadas), consulte [todas as APIs JavaScript do Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask (email: cadeia de caracteres)](/javascript/api/excel/excel.comment#assigntask-email-)|Atribui a tarefa anexada ao comentário para o usuário fornecido como o único destinatário.|
||[getTask ()](/javascript/api/excel/excel.comment#gettask--)|Obtém a tarefa associada a este comentário.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#gettaskornullobject--)|Obtém a tarefa associada a este comentário.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask (email: cadeia de caracteres)](/javascript/api/excel/excel.commentreply#assigntask-email-)|Atribui a tarefa anexada ao comentário para o usuário fornecido como o único destinatário.|
||[getTask ()](/javascript/api/excel/excel.commentreply#gettask--)|Obtém a tarefa associada a este comentário.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|Obtém a tarefa associada a este comentário.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[DataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|O nome do provedor de dados para o tipo de dados vinculados.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|A data e a hora da zona de tempo local desde que a pasta de trabalho foi aberta quando o tipo de dados vinculados foi atualizado pela última vez.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|O nome do tipo de dados vinculados.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|A frequência, em segundos, em que o tipo de dados vinculado é atualizado, se `refreshMode` estiver definido como "periódico".|
||[RefreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|O mecanismo pelo qual os dados para o tipo de dados vinculados são recuperados.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|A identificação exclusiva do tipo de dados vinculados.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|Retorna uma matriz com todos os modos de atualização compatíveis com o tipo de dados vinculados.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|Faz uma solicitação para atualizar o tipo de dados vinculados.|
||[requestSetRefreshMode (RefreshMode: Excel. LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|Faz uma solicitação para alterar o modo de atualização para esse tipo de dados vinculados.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|A identificação exclusiva do novo tipo de dados vinculados.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Obtém o tipo do evento.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|Obtém o número de tipos de dados vinculados na coleção.|
||[getItem (Key: Number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|Obtém um tipo de dados vinculado por ID de serviço.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|Obtém um tipo de dados vinculado por seu índice na coleção.|
||[getItemOrNullObject (Key: Number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|Obtém um tipo de dados vinculado por ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|Faz uma solicitação para atualizar todos os tipos de dados vinculados na coleção.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|Ativa este modo de exibição de planilha.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|Remove o modo de exibição de planilha da planilha.|
||[Duplicate (Name?: String)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|Cria uma cópia deste modo de exibição de planilha.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Obtém ou define o nome do modo de exibição de planilha.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|Cria um novo modo de exibição de planilha com o nome fornecido.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|Cria e ativa um novo modo de exibição de planilha temporária.|
||[Exit ()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|Sai do modo de exibição de planilha ativo no momento.|
||[getactive ()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|Obtém o modo de exibição de planilha atualmente ativo da planilha.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|Obtém o número de modos de exibição de planilha nesta planilha.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|Obtém um modo de exibição de planilha usando seu nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|Obtém um modo de exibição de planilha por seu índice na coleção.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|A descrição de texto alt da tabela dinâmica.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|O título do texto alt da tabela dinâmica.|
||[displayBlankLineAfterEachItem (exibição: Boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Define se deve ou não exibir uma linha em branco após cada item.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|O texto que é preenchido automaticamente em qualquer célula vazia da tabela dinâmica se `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Especifica se as células vazias da tabela dinâmica devem ser preenchidas com o `emptyCellText` .|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtém uma célula exclusiva na tabela dinâmica com base em uma hierarquia de dados, bem como os itens de linha e coluna de suas respectivas hierarquias.|
||[tabela dinâmica](/javascript/api/excel/excel.pivotlayout#pivotstyle)|O estilo aplicado à tabela dinâmica.|
||[repeatAllItemLabels (repeatLabels: Boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Define a configuração "repetir todos os rótulos de item" em todos os campos da tabela dinâmica.|
||[setStyle (Style: String \| pivotstyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Define o estilo aplicado à tabela dinâmica.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Especifica se a tabela dinâmica exibe cabeçalhos de campos (legendas de campos e suspensas de filtro).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Especifica se a tabela dinâmica é atualizada quando a pasta de trabalho é aberta.|
|[Range](/javascript/api/excel/excel.range)|[getprecedentes ()](/javascript/api/excel/excel.range#getprecedents--)|Retorna um `WorkbookRangeAreas` objeto que representa o intervalo que contém todos os precedentes de uma célula na mesma planilha ou em várias planilhas.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[RefreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|O modo de atualização do tipo de dados vinculado.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|A identificação exclusiva do objeto cujo modo de atualização foi alterado.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Obtém o tipo do evento.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[atualizado](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Indica se a solicitação para atualizar foi bem-sucedida.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|A identificação exclusiva do objeto cuja solicitação de atualização foi concluída.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Obtém o tipo do evento.|
||[alerta](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Uma matriz que contém quaisquer avisos gerados a partir da solicitação de atualização.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Cria um gráfico vetorial escalável (SVG) de uma cadeia de caracteres XML e a adiciona à planilha.|
|[Segmentação de dados](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Representa o nome da segmentação de dados usada na fórmula.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|O estilo aplicado à segmentação de,.|
||[setStyle (Style: String \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Define o estilo aplicado à segmentação de,.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Altera a tabela para usar o estilo de tabela padrão.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Ocorre quando o filtro é aplicado em uma tabela específica.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|O estilo aplicado à tabela.|
||[setStyle (Style: String \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Define o estilo aplicado à tabela.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Ocorre quando o filtro é aplicado em uma tabela localizada em uma pasta de trabalho ou em uma planilha.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Obtém a ID da tabela na qual o filtro é aplicado.|
||[tipo](/javascript/api/excel/excel.tablefilteredeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Obtém a ID da planilha que contém a tabela.|
|[Tarefa](/javascript/api/excel/excel.task)|[AddAssigne (email: cadeia de caracteres)](/javascript/api/excel/excel.task#addassignee-email-)|Adiciona um destinatário à tarefa.|
||[applyChanges (taskChanges: Excel. TaskChanges)](/javascript/api/excel/excel.task#applychanges-taskchanges-)|Aplica as alterações determinadas à tarefa.|
||[destinatários](/javascript/api/excel/excel.task#assignees)|Obtém os usuários aos quais a tarefa é atribuída.|
||[Retire](/javascript/api/excel/excel.task#comment)|Obtém o comentário associado à tarefa.|
||[dueDate](/javascript/api/excel/excel.task#duedate)|Obtém a data e hora de conclusão da tarefa.|
||[historyRecords](/javascript/api/excel/excel.task#historyrecords)|Obtém os registros de histórico da tarefa.|
||[id](/javascript/api/excel/excel.task#id)|Obtém a ID da tarefa.|
||[percentComplete](/javascript/api/excel/excel.task#percentcomplete)|Obtém a porcentagem de conclusão da tarefa.|
||[prioridade](/javascript/api/excel/excel.task#priority)|Obtém a prioridade da tarefa.|
||[startDate](/javascript/api/excel/excel.task#startdate)|Obtém a data e hora de início da tarefa.|
||[title](/javascript/api/excel/excel.task#title)|Obtém o título da tarefa.|
||[removeAllAssignees()](/javascript/api/excel/excel.task#removeallassignees--)|Remove todos os destinatários da tarefa.|
||[removeAssignee (email: cadeia de caracteres)](/javascript/api/excel/excel.task#removeassignee-email-)|Remove um destinatário da tarefa.|
||[setPercentComplete (PorcentagemConcluída: número)](/javascript/api/excel/excel.task#setpercentcomplete-percentcomplete-)|Altera a conclusão da tarefa.|
||[SetPriority (prioridade: número)](/javascript/api/excel/excel.task#setpriority-priority-)|Altera a prioridade da tarefa.|
||[setStartDateAndDueDate (startDate: Date, dueDate: Date)](/javascript/api/excel/excel.task#setstartdateandduedate-startdate--duedate-)|Altera as datas de início e de conclusão da tarefa.|
||[setTitle (título: cadeia de caracteres)](/javascript/api/excel/excel.task#settitle-title-)|Altera o título da tarefa.|
|[TaskChanges](/javascript/api/excel/excel.taskchanges)|[dueDate](/javascript/api/excel/excel.taskchanges#duedate)|Define uma nova data de conclusão para a tarefa, no fuso horário UTC.|
||[emailsToAssign](/javascript/api/excel/excel.taskchanges#emailstoassign)|Define endereços de email dos usuários a serem atribuídos à tarefa.|
||[emailsToUnassign](/javascript/api/excel/excel.taskchanges#emailstounassign)|Define endereços de email dos usuários para cancelar a atribuição da tarefa.|
||[percentComplete](/javascript/api/excel/excel.taskchanges#percentcomplete)|Define uma nova porcentagem de conclusão para a tarefa.|
||[prioridade](/javascript/api/excel/excel.taskchanges#priority)|Define uma nova prioridade para a tarefa.|
||[removeAllPreviousAssignees](/javascript/api/excel/excel.taskchanges#removeallpreviousassignees)|Define se a alteração deve remover todos os destinatários anteriores da tarefa.|
||[startDate](/javascript/api/excel/excel.taskchanges#startdate)|Define uma nova data de início para a tarefa, no fuso horário UTC.|
||[title](/javascript/api/excel/excel.taskchanges#title)|Define um novo título para a tarefa.|
|[Taskcollection](/javascript/api/excel/excel.taskcollection)|[getCount()](/javascript/api/excel/excel.taskcollection#getcount--)|Obtém o número de tarefas na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.taskcollection#getitem-key-)|Obtém uma tarefa usando sua ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskcollection#getitemat-index-)|Obtém uma tarefa por seu índice na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.taskcollection#getitemornullobject-key-)|Obtém uma tarefa usando sua ID.|
||[items](/javascript/api/excel/excel.taskcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[TaskHistoryRecord](/javascript/api/excel/excel.taskhistoryrecord)|[AnchorID](/javascript/api/excel/excel.taskhistoryrecord#anchorid)|Representa a ID do objeto para o qual a tarefa é ancorada (por exemplo, commentId para tarefas anexadas a comentários).|
||[destinatário](/javascript/api/excel/excel.taskhistoryrecord#assignee)|Representa o usuário atribuído à tarefa para um tipo de registro de "Assign", ou o usuário para cancelar a atribuição da tarefa para um tipo de registro de "cancelamento de atribuição".|
||[attributionUser](/javascript/api/excel/excel.taskhistoryrecord#attributionuser)|Representa o usuário que criou ou alterou a tarefa.|
||[dueDate](/javascript/api/excel/excel.taskhistoryrecord#duedate)|Representa a data de conclusão da tarefa.|
||[historyRecordCreatedDate](/javascript/api/excel/excel.taskhistoryrecord#historyrecordcreateddate)|Representa a data de criação do registro de histórico de tarefas.|
||[id](/javascript/api/excel/excel.taskhistoryrecord#id)|ID do registro de histórico.|
||[percentComplete](/javascript/api/excel/excel.taskhistoryrecord#percentcomplete)|Representa a porcentagem de conclusão da tarefa.|
||[prioridade](/javascript/api/excel/excel.taskhistoryrecord#priority)|Representa a prioridade da tarefa.|
||[startDate](/javascript/api/excel/excel.taskhistoryrecord#startdate)|Representa a data de início da tarefa.|
||[title](/javascript/api/excel/excel.taskhistoryrecord#title)|Representa o título da tarefa.|
||[type](/javascript/api/excel/excel.taskhistoryrecord#type)|Representa o tipo de registro do histórico de tarefas.|
||[undoHistoryId](/javascript/api/excel/excel.taskhistoryrecord#undohistoryid)|Representa a propriedade TaskHistoryRecord.id que foi desfeita para o tipo de registro de histórico "desfazer".|
|[TaskHistoryRecordCollection](/javascript/api/excel/excel.taskhistoryrecordcollection)|[getCount()](/javascript/api/excel/excel.taskhistoryrecordcollection#getcount--)|Obtém o número de registros de histórico na coleção para a tarefa.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskhistoryrecordcollection#getitemat-index-)|Obtém um registro de histórico de tarefas usando seu índice na coleção.|
||[items](/javascript/api/excel/excel.taskhistoryrecordcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Usuário](/javascript/api/excel/excel.user)|[displayName](/javascript/api/excel/excel.user#displayname)|Representa o nome para exibição do usuário.|
||[email](/javascript/api/excel/excel.user#email)|Representa o endereço de email do usuário.|
||[uid](/javascript/api/excel/excel.user#uid)|Representa a ID exclusiva do usuário.|
|[Workbook](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|Retorna uma coleção de tipos de dados vinculados que fazem parte da pasta de trabalho.|
||[tarefas](/javascript/api/excel/excel.workbook#tasks)|Retorna uma coleção de tarefas que estão presentes na pasta de trabalho.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Especifica se o painel de lista de campos da tabela dinâmica é mostrado no nível da pasta de trabalho.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True se a pasta de trabalho usar o sistema de dados 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Retorna uma coleção de modos de exibição de planilha que estão presentes na planilha.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Ocorre quando o filtro é aplicado em uma planilha específica.|
||[tarefas](/javascript/api/excel/excel.worksheet#tasks)|Retorna uma coleção de tarefas que estão presentes na planilha.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insere as planilhas especificadas de uma pasta de trabalho na pasta de trabalho atual.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Ocorre quando filtro de uma planilha é aplicado na pasta de trabalho.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[tipo](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Obtém a ID da planilha na qual o filtro é aplicado.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
