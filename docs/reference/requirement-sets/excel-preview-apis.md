---
title: APIs de visualização do JavaScript para Excel
description: Detalhes sobre as futuras APIs JavaScript do Excel.
ms.date: 02/24/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0663b6330c402f64e7ed7e8f598a52848bbe1319
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505532"
---
# <a name="excel-javascript-preview-apis"></a>APIs de visualização do JavaScript para Excel

As novas APIs do JavaScript para Excel são introduzidas pela primeira vez na "visualização" e, posteriormente, tornam-se parte de um conjunto específico de requisitos numerados, após passarem por vários testes e após a recolha das opiniões de usuários.

A primeira tabela fornece um resumo conciso das APIs e, a tabela subsequente, fornece uma lista detalhada.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| Fórmula de eventos alterados | Acompanhe as alterações nas fórmulas, incluindo a origem e o tipo de evento que causou uma alteração. | [Worksheet.onFormulaChanged](/javascript/api/excel/excel.worksheet#onFormulaChanged)|
| Tipos de dados vinculados | Adiciona suporte para tipos de dados conectados ao Excel de fontes externas. | [LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|
| Exibições de planilha nomeadas | Fornece controle programático de exibições de planilha por usuário. | [NamedSheetView](/javascript/api/excel/excel.namedsheetview) |
| Tarefas | Transforme os comentários em tarefas atribuídas aos usuários. | [Tarefa](/javascript/api/excel/excel.task) |

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs JavaScript do Excel atualmente em visualização. Para ver uma lista completa de todas as APIs JavaScript do Excel (incluindo APIs de visualização e APIs lançadas anteriormente), consulte todas as [APIs JavaScript do Excel.](/javascript/api/excel?view=excel-js-preview&preserve-view=true)

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[assignTask(email: string)](/javascript/api/excel/excel.comment#assigntask-email-)|Atribui a tarefa anexada ao comentário ao usuário determinado como o único destinatário.|
||[getTask()](/javascript/api/excel/excel.comment#gettask--)|Obtém a tarefa associada a este comentário.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.comment#gettaskornullobject--)|Obtém a tarefa associada a este comentário.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[assignTask(email: string)](/javascript/api/excel/excel.commentreply#assigntask-email-)|Atribui a tarefa anexada ao comentário ao usuário determinado como o único destinatário.|
||[getTask()](/javascript/api/excel/excel.commentreply#gettask--)|Obtém a tarefa associada a este comentário.|
||[getTaskOrNullObject()](/javascript/api/excel/excel.commentreply#gettaskornullobject--)|Obtém a tarefa associada a este comentário.|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|O endereço da célula que contém a fórmula alterada.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|Representa a fórmula anterior, antes de ser alterada.|
|[LinkedDataType](/javascript/api/excel/excel.linkeddatatype)|[dataProvider](/javascript/api/excel/excel.linkeddatatype#dataprovider)|O nome do provedor de dados do tipo de dados vinculado.|
||[lastRefreshed](/javascript/api/excel/excel.linkeddatatype#lastrefreshed)|A data e a hora do fuso horário local desde que a lista de trabalho foi aberta quando o tipo de dados vinculado foi atualizado pela última vez.|
||[name](/javascript/api/excel/excel.linkeddatatype#name)|O nome do tipo de dados vinculado.|
||[periodicRefreshInterval](/javascript/api/excel/excel.linkeddatatype#periodicrefreshinterval)|A frequência, em segundos, na qual o tipo de dados vinculado é atualizado se `refreshMode` estiver definido como "Periódico".|
||[refreshMode](/javascript/api/excel/excel.linkeddatatype#refreshmode)|O mecanismo pelo qual os dados do tipo de dados vinculados são recuperados.|
||[serviceId](/javascript/api/excel/excel.linkeddatatype#serviceid)|A id exclusiva do tipo de dados vinculado.|
||[supportedRefreshModes](/javascript/api/excel/excel.linkeddatatype#supportedrefreshmodes)|Retorna uma matriz com todos os modos de atualização suportados pelo tipo de dados vinculado.|
||[requestRefresh()](/javascript/api/excel/excel.linkeddatatype#requestrefresh--)|Faz uma solicitação para atualizar o tipo de dados vinculado.|
||[requestSetRefreshMode(refreshMode: Excel.LinkedDataTypeRefreshMode)](/javascript/api/excel/excel.linkeddatatype#requestsetrefreshmode-refreshmode-)|Faz uma solicitação para alterar o modo de atualização para esse tipo de dados vinculado.|
|[LinkedDataTypeAddedEventArgs](/javascript/api/excel/excel.linkeddatatypeaddedeventargs)|[serviceId](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#serviceid)|A id exclusiva do novo tipo de dados vinculado.|
||[source](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.linkeddatatypeaddedeventargs#type)|Obtém o tipo do evento.|
|[LinkedDataTypeCollection](/javascript/api/excel/excel.linkeddatatypecollection)|[getCount()](/javascript/api/excel/excel.linkeddatatypecollection#getcount--)|Obtém o número de tipos de dados vinculados na coleção.|
||[getItem(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitem-key-)|Obtém um tipo de dados vinculado por id de serviço.|
||[getItemAt(index: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemat-index-)|Obtém um tipo de dados vinculado pelo índice na coleção.|
||[getItemOrNullObject(key: number)](/javascript/api/excel/excel.linkeddatatypecollection#getitemornullobject-key-)|Obtém um tipo de dados vinculado por ID.|
||[items](/javascript/api/excel/excel.linkeddatatypecollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[requestRefreshAll()](/javascript/api/excel/excel.linkeddatatypecollection#requestrefreshall--)|Faz uma solicitação para atualizar todos os tipos de dados vinculados na coleção.|
|[NamedSheetView](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|Ativa esse modo de exibição de planilha.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|Remove o exibição de planilha da planilha.|
||[duplicate(name?: string)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|Cria uma cópia desse exibição de planilha.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Obtém ou define o nome do exibição de planilha.|
|[NamedSheetViewCollection](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|Cria um novo exibição de planilha com o nome determinado.|
||[enterTemporary()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|Cria e ativa um novo modo de exibição de planilha temporária.|
||[exit()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|Sai do exibição de planilha ativa no momento.|
||[getActive()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|Obtém a exibição de planilha ativa da planilha no momento.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|Obtém o número de exibições de planilha nesta planilha.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|Obtém uma exibição de planilha usando seu nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|Obtém uma exibição de planilha pelo índice na coleção.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|A descrição de texto alt da Tabela Dinâmica.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|O título de texto alt da Tabela Dinâmica.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Define se uma linha em branco deve ou não ser exibida após cada item.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|O texto que é preenchido automaticamente em qualquer célula vazia na Tabela Dinâmica se `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Especifica se células vazias na Tabela Dinâmica devem ser preenchidas com `emptyCellText` o .|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Obtém uma célula exclusiva na tabela dinâmica com base em uma hierarquia de dados, bem como os itens de linha e coluna de suas respectivas hierarquias.|
||[pivotStyle](/javascript/api/excel/excel.pivotlayout#pivotstyle)|O estilo aplicado à Tabela Dinâmica.|
||[repeatAllItemLabels(repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Define a configuração "repetir todos os rótulos de item" em todos os campos da Tabela Dinâmica.|
||[setStyle(style: string \| PivotTableStyle \| BuiltInPivotTableStyle)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Define o estilo aplicado à Tabela Dinâmica.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Especifica se a Tabela Dinâmica exibe os headers de campo (legendas de campo e drop-downs de filtro).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Especifica se a Tabela Dinâmica é atualizada quando a workbook é aberta.|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents()](/javascript/api/excel/excel.range#getprecedents--)|Retorna um objeto que representa o intervalo que contém todos os precedentes de uma célula na mesma planilha ou `WorkbookRangeAreas` em várias planilhas.|
|[RefreshModeChangedEventArgs](/javascript/api/excel/excel.refreshmodechangedeventargs)|[refreshMode](/javascript/api/excel/excel.refreshmodechangedeventargs#refreshmode)|O modo de atualização do tipo de dados vinculado.|
||[serviceId](/javascript/api/excel/excel.refreshmodechangedeventargs#serviceid)|A ID exclusiva do objeto cujo modo de atualização foi alterado.|
||[source](/javascript/api/excel/excel.refreshmodechangedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.refreshmodechangedeventargs#type)|Obtém o tipo do evento.|
|[RefreshRequestCompletedEventArgs](/javascript/api/excel/excel.refreshrequestcompletedeventargs)|[atualizado](/javascript/api/excel/excel.refreshrequestcompletedeventargs#refreshed)|Indica se a solicitação de atualização foi bem-sucedida.|
||[serviceId](/javascript/api/excel/excel.refreshrequestcompletedeventargs#serviceid)|A id exclusiva do objeto cuja solicitação de atualização foi concluída.|
||[source](/javascript/api/excel/excel.refreshrequestcompletedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.refreshrequestcompletedeventargs#type)|Obtém o tipo do evento.|
||[avisos](/javascript/api/excel/excel.refreshrequestcompletedeventargs#warnings)|Uma matriz que contém quaisquer avisos gerados a partir da solicitação de atualização.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Cria um gráfico vetorial escalável (SVG) de uma cadeia de caracteres XML e a adiciona à planilha.|
|[Segmentação de dados](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Representa o nome da segmentação de dados usada na fórmula.|
||[slicerStyle](/javascript/api/excel/excel.slicer#slicerstyle)|O estilo aplicado à Slicer.|
||[setStyle(style: string \| SlicerStyle \| BuiltInSlicerStyle)](/javascript/api/excel/excel.slicer#setstyle-style-)|Define o estilo aplicado à slicer.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Altera a tabela para usar o estilo de tabela padrão.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Ocorre quando o filtro é aplicado em uma tabela específica.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|O estilo aplicado à Tabela.|
||[setStyle(style: string \| TableStyle \| BuiltInTableStyle)](/javascript/api/excel/excel.table#setstyle-style-)|Define o estilo aplicado à tabela.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Ocorre quando o filtro é aplicado em uma tabela localizada em uma pasta de trabalho ou em uma planilha.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Obtém a id da tabela na qual o filtro é aplicado.|
||[tipo](/javascript/api/excel/excel.tablefilteredeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Obtém a id da planilha que contém a tabela.|
|[Tarefa](/javascript/api/excel/excel.task)|[addAssignee(email: string)](/javascript/api/excel/excel.task#addassignee-email-)|Adiciona um destinatário à tarefa.|
||[applyChanges(taskChanges: Excel.TaskChanges)](/javascript/api/excel/excel.task#applychanges-taskchanges-)|Aplica as alterações fornecidas à tarefa.|
||[assignees](/javascript/api/excel/excel.task#assignees)|Obtém os usuários aos quais a tarefa é atribuída.|
||[comment](/javascript/api/excel/excel.task#comment)|Obtém o comentário associado à tarefa.|
||[dueDate](/javascript/api/excel/excel.task#duedate)|Obtém a data e a hora de vencimento da tarefa.|
||[historyRecords](/javascript/api/excel/excel.task#historyrecords)|Obtém os registros de histórico da tarefa.|
||[id](/javascript/api/excel/excel.task#id)|Obtém a id da tarefa.|
||[percentComplete](/javascript/api/excel/excel.task#percentcomplete)|Obtém a porcentagem de conclusão da tarefa.|
||[prioridade](/javascript/api/excel/excel.task#priority)|Obtém a prioridade da tarefa.|
||[startDate](/javascript/api/excel/excel.task#startdate)|Obtém a data e a hora em que a tarefa deve começar.|
||[title](/javascript/api/excel/excel.task#title)|Obtém o título da tarefa.|
||[removeAllAssignees()](/javascript/api/excel/excel.task#removeallassignees--)|Remove todos os atribuídos da tarefa.|
||[removeAssignee(email: string)](/javascript/api/excel/excel.task#removeassignee-email-)|Remove um destinatário da tarefa.|
||[setPercentComplete(percentComplete: number)](/javascript/api/excel/excel.task#setpercentcomplete-percentcomplete-)|Altera a conclusão da tarefa.|
||[setPriority(priority: number)](/javascript/api/excel/excel.task#setpriority-priority-)|Altera a prioridade da tarefa.|
||[setStartDateAndDueDate(startDate: Date, dueDate: Date)](/javascript/api/excel/excel.task#setstartdateandduedate-startdate--duedate-)|Altera o início e as datas de vencimento da tarefa.|
||[setTitle(title: string)](/javascript/api/excel/excel.task#settitle-title-)|Altera o título da tarefa.|
|[TaskChanges](/javascript/api/excel/excel.taskchanges)|[dueDate](/javascript/api/excel/excel.taskchanges#duedate)|Define uma nova data de vencimento para a tarefa, no fuso horário UTC.|
||[emailsToAssign](/javascript/api/excel/excel.taskchanges#emailstoassign)|Define endereços de email dos usuários a atribuir à tarefa.|
||[emailsToUnassign](/javascript/api/excel/excel.taskchanges#emailstounassign)|Define endereços de email dos usuários para desaignar da tarefa.|
||[percentComplete](/javascript/api/excel/excel.taskchanges#percentcomplete)|Define um novo percentual de conclusão para a tarefa.|
||[prioridade](/javascript/api/excel/excel.taskchanges#priority)|Define uma nova prioridade para a tarefa.|
||[removeAllPreviousAssignees](/javascript/api/excel/excel.taskchanges#removeallpreviousassignees)|Define se a alteração deve remover todos os atribuídos anteriores da tarefa.|
||[startDate](/javascript/api/excel/excel.taskchanges#startdate)|Define uma nova data de início para a tarefa, no fuso horário UTC.|
||[title](/javascript/api/excel/excel.taskchanges#title)|Define um novo título para a tarefa.|
|[TaskCollection](/javascript/api/excel/excel.taskcollection)|[getCount()](/javascript/api/excel/excel.taskcollection#getcount--)|Obtém o número de tarefas na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.taskcollection#getitem-key-)|Obtém uma tarefa usando sua id.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskcollection#getitemat-index-)|Obtém uma tarefa pelo índice na coleção.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.taskcollection#getitemornullobject-key-)|Obtém uma tarefa usando sua id.|
||[items](/javascript/api/excel/excel.taskcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[TaskHistoryRecord](/javascript/api/excel/excel.taskhistoryrecord)|[anchorId](/javascript/api/excel/excel.taskhistoryrecord#anchorid)|Representa a ID do objeto ao qual a tarefa está ancorada (por exemplo, commentId para tarefas anexadas a comentários).|
||[assignee](/javascript/api/excel/excel.taskhistoryrecord#assignee)|Representa o usuário atribuído à tarefa para um tipo de registro de histórico "Atribuir" ou o usuário a desatribuição da tarefa para um tipo de registro de histórico "Unassign".|
||[attributionUser](/javascript/api/excel/excel.taskhistoryrecord#attributionuser)|Representa o usuário que criou ou alterou a tarefa.|
||[dueDate](/javascript/api/excel/excel.taskhistoryrecord#duedate)|Representa a data de vencimento da tarefa.|
||[historyRecordCreatedDate](/javascript/api/excel/excel.taskhistoryrecord#historyrecordcreateddate)|Representa a data de criação do registro de histórico de tarefas.|
||[id](/javascript/api/excel/excel.taskhistoryrecord#id)|ID do registro de histórico.|
||[percentComplete](/javascript/api/excel/excel.taskhistoryrecord#percentcomplete)|Representa a porcentagem de conclusão da tarefa.|
||[prioridade](/javascript/api/excel/excel.taskhistoryrecord#priority)|Representa a prioridade da tarefa.|
||[startDate](/javascript/api/excel/excel.taskhistoryrecord#startdate)|Representa a data de início da tarefa.|
||[title](/javascript/api/excel/excel.taskhistoryrecord#title)|Representa o título da tarefa.|
||[type](/javascript/api/excel/excel.taskhistoryrecord#type)|Representa o tipo do registro do histórico de tarefas.|
||[undoHistoryId](/javascript/api/excel/excel.taskhistoryrecord#undohistoryid)|Representa a TaskHistoryRecord.id que foi desfeita para o tipo de registro de histórico "Desfazer".|
|[TaskHistoryRecordCollection](/javascript/api/excel/excel.taskhistoryrecordcollection)|[getCount()](/javascript/api/excel/excel.taskhistoryrecordcollection#getcount--)|Obtém o número de registros de histórico na coleção da tarefa.|
||[getItemAt(index: number)](/javascript/api/excel/excel.taskhistoryrecordcollection#getitemat-index-)|Obtém um registro de histórico de tarefas usando seu índice na coleção.|
||[items](/javascript/api/excel/excel.taskhistoryrecordcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[User](/javascript/api/excel/excel.user)|[displayName](/javascript/api/excel/excel.user#displayname)|Representa o nome para exibição do usuário.|
||[email](/javascript/api/excel/excel.user#email)|Representa o endereço de email do usuário.|
||[uid](/javascript/api/excel/excel.user#uid)|Representa a ID exclusiva do usuário.|
|[Pasta de trabalho](/javascript/api/excel/excel.workbook)|[linkedDataTypes](/javascript/api/excel/excel.workbook#linkeddatatypes)|Retorna uma coleção de tipos de dados vinculados que fazem parte da lista de trabalho.|
||[tasks](/javascript/api/excel/excel.workbook#tasks)|Retorna uma coleção de tarefas que estão presentes na workbook.|
||[showPivotFieldList](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Especifica se o painel de lista de campos da Tabela Dinâmica é mostrado no nível da lista de trabalho.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|True se a pasta de trabalho usar o sistema de dados 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[namedSheetViews](/javascript/api/excel/excel.worksheet#namedsheetviews)|Retorna uma coleção de exibições de planilha presentes na planilha.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Ocorre quando o filtro é aplicado em uma planilha específica.|
||[onFormulaChanged](/javascript/api/excel/excel.worksheet#onformulachanged)|Ocorre quando uma ou mais fórmulas são alteradas nesta planilha.|
||[tasks](/javascript/api/excel/excel.worksheet#tasks)|Retorna uma coleção de tarefas presentes na planilha.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Insere as planilhas especificadas de uma pasta de trabalho na pasta de trabalho atual.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Ocorre quando filtro de uma planilha é aplicado na pasta de trabalho.|
||[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|Ocorre quando uma ou mais fórmulas são alteradas em qualquer planilha dessa coleção.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[tipo](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Obtém a id da planilha na qual o filtro é aplicado.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|Obtém uma matriz de objetos FormulaChangedEventDetail, que contêm os detalhes sobre todas as fórmulas alteradas.|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|A origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|Obtém a ID da planilha na qual a fórmula foi alterada.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
